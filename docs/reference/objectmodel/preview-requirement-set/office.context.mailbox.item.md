---
title: Office.Context.Mailbox.Item - наборы требований предварительного просмотра
description: ''
ms.date: 01/16/2019
localization_priority: Normal
ms.openlocfilehash: b4b2ec9c735270d9b1bfca3d1c24ef6b0f1ca1cb
ms.sourcegitcommit: d1aa7201820176ed986b9f00bb9c88e055906c77
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 01/23/2019
ms.locfileid: "29389601"
---
# <a name="item"></a><span data-ttu-id="fa854-102">item</span><span class="sxs-lookup"><span data-stu-id="fa854-102">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="fa854-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="fa854-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="fa854-p101">Пространство имен `item` используется для доступа к выбранному в данный момент сообщению, приглашению на собрание или описанию встречи. Вы можете определить тип пространства имен `item` с помощью свойства [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="fa854-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="fa854-106">Requirements</span><span class="sxs-lookup"><span data-stu-id="fa854-106">Requirements</span></span>

|<span data-ttu-id="fa854-107">Требование</span><span class="sxs-lookup"><span data-stu-id="fa854-107">Requirement</span></span>|<span data-ttu-id="fa854-108">Значение</span><span class="sxs-lookup"><span data-stu-id="fa854-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="fa854-109">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="fa854-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="fa854-110">1.0</span><span class="sxs-lookup"><span data-stu-id="fa854-110">1.0</span></span>|
|[<span data-ttu-id="fa854-111">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="fa854-111">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="fa854-112">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="fa854-112">Restricted</span></span>|
|[<span data-ttu-id="fa854-113">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="fa854-113">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="fa854-114">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="fa854-114">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="fa854-115">Элементы и методы</span><span class="sxs-lookup"><span data-stu-id="fa854-115">Members and methods</span></span>

| <span data-ttu-id="fa854-116">Элемент</span><span class="sxs-lookup"><span data-stu-id="fa854-116">Member</span></span> | <span data-ttu-id="fa854-117">Тип</span><span class="sxs-lookup"><span data-stu-id="fa854-117">Type</span></span> |
|--------|------|
| [<span data-ttu-id="fa854-118">attachments</span><span class="sxs-lookup"><span data-stu-id="fa854-118">attachments</span></span>](#attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails) | <span data-ttu-id="fa854-119">Элемент</span><span class="sxs-lookup"><span data-stu-id="fa854-119">Member</span></span> |
| [<span data-ttu-id="fa854-120">bcc</span><span class="sxs-lookup"><span data-stu-id="fa854-120">bcc</span></span>](#bcc-recipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="fa854-121">Элемент</span><span class="sxs-lookup"><span data-stu-id="fa854-121">Member</span></span> |
| [<span data-ttu-id="fa854-122">body</span><span class="sxs-lookup"><span data-stu-id="fa854-122">body</span></span>](#body-bodyjavascriptapioutlookofficebody) | <span data-ttu-id="fa854-123">Элемент</span><span class="sxs-lookup"><span data-stu-id="fa854-123">Member</span></span> |
| [<span data-ttu-id="fa854-124">cc</span><span class="sxs-lookup"><span data-stu-id="fa854-124">cc</span></span>](#cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="fa854-125">Элемент</span><span class="sxs-lookup"><span data-stu-id="fa854-125">Member</span></span> |
| [<span data-ttu-id="fa854-126">conversationId</span><span class="sxs-lookup"><span data-stu-id="fa854-126">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="fa854-127">Элемент</span><span class="sxs-lookup"><span data-stu-id="fa854-127">Member</span></span> |
| [<span data-ttu-id="fa854-128">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="fa854-128">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="fa854-129">Элемент</span><span class="sxs-lookup"><span data-stu-id="fa854-129">Member</span></span> |
| [<span data-ttu-id="fa854-130">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="fa854-130">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="fa854-131">Элемент</span><span class="sxs-lookup"><span data-stu-id="fa854-131">Member</span></span> |
| [<span data-ttu-id="fa854-132">end</span><span class="sxs-lookup"><span data-stu-id="fa854-132">end</span></span>](#end-datetimejavascriptapioutlookofficetime) | <span data-ttu-id="fa854-133">Элемент</span><span class="sxs-lookup"><span data-stu-id="fa854-133">Member</span></span> |
| [<span data-ttu-id="fa854-134">from</span><span class="sxs-lookup"><span data-stu-id="fa854-134">from</span></span>](#from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom) | <span data-ttu-id="fa854-135">Элемент</span><span class="sxs-lookup"><span data-stu-id="fa854-135">Member</span></span> |
| [<span data-ttu-id="fa854-136">internetHeaders</span><span class="sxs-lookup"><span data-stu-id="fa854-136">internetHeaders</span></span>](#internetheaders-internetheadersjavascriptapioutlookofficeinternetheaders) | <span data-ttu-id="fa854-137">Элемент</span><span class="sxs-lookup"><span data-stu-id="fa854-137">Member</span></span> |
| [<span data-ttu-id="fa854-138">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="fa854-138">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="fa854-139">Элемент</span><span class="sxs-lookup"><span data-stu-id="fa854-139">Member</span></span> |
| [<span data-ttu-id="fa854-140">itemClass</span><span class="sxs-lookup"><span data-stu-id="fa854-140">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="fa854-141">Элемент</span><span class="sxs-lookup"><span data-stu-id="fa854-141">Member</span></span> |
| [<span data-ttu-id="fa854-142">itemId</span><span class="sxs-lookup"><span data-stu-id="fa854-142">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="fa854-143">Элемент</span><span class="sxs-lookup"><span data-stu-id="fa854-143">Member</span></span> |
| [<span data-ttu-id="fa854-144">itemType</span><span class="sxs-lookup"><span data-stu-id="fa854-144">itemType</span></span>](#itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype) | <span data-ttu-id="fa854-145">Элемент</span><span class="sxs-lookup"><span data-stu-id="fa854-145">Member</span></span> |
| [<span data-ttu-id="fa854-146">location</span><span class="sxs-lookup"><span data-stu-id="fa854-146">location</span></span>](#location-stringlocationjavascriptapioutlookofficelocation) | <span data-ttu-id="fa854-147">Элемент</span><span class="sxs-lookup"><span data-stu-id="fa854-147">Member</span></span> |
| [<span data-ttu-id="fa854-148">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="fa854-148">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="fa854-149">Элемент</span><span class="sxs-lookup"><span data-stu-id="fa854-149">Member</span></span> |
| [<span data-ttu-id="fa854-150">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="fa854-150">notificationMessages</span></span>](#notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessages) | <span data-ttu-id="fa854-151">Элемент</span><span class="sxs-lookup"><span data-stu-id="fa854-151">Member</span></span> |
| [<span data-ttu-id="fa854-152">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="fa854-152">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="fa854-153">Элемент</span><span class="sxs-lookup"><span data-stu-id="fa854-153">Member</span></span> |
| [<span data-ttu-id="fa854-154">organizer</span><span class="sxs-lookup"><span data-stu-id="fa854-154">organizer</span></span>](#organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsorganizerjavascriptapioutlookofficeorganizer) | <span data-ttu-id="fa854-155">Элемент</span><span class="sxs-lookup"><span data-stu-id="fa854-155">Member</span></span> |
| [<span data-ttu-id="fa854-156">recurrence</span><span class="sxs-lookup"><span data-stu-id="fa854-156">recurrence</span></span>](#nullable-recurrence-recurrencejavascriptapioutlookofficerecurrence) | <span data-ttu-id="fa854-157">Элемент</span><span class="sxs-lookup"><span data-stu-id="fa854-157">Member</span></span> |
| [<span data-ttu-id="fa854-158">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="fa854-158">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="fa854-159">Элемент</span><span class="sxs-lookup"><span data-stu-id="fa854-159">Member</span></span> |
| [<span data-ttu-id="fa854-160">sender</span><span class="sxs-lookup"><span data-stu-id="fa854-160">sender</span></span>](#sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails) | <span data-ttu-id="fa854-161">Элемент</span><span class="sxs-lookup"><span data-stu-id="fa854-161">Member</span></span> |
| [<span data-ttu-id="fa854-162">seriesId</span><span class="sxs-lookup"><span data-stu-id="fa854-162">seriesId</span></span>](#nullable-seriesid-string) | <span data-ttu-id="fa854-163">Элемент</span><span class="sxs-lookup"><span data-stu-id="fa854-163">Member</span></span> |
| [<span data-ttu-id="fa854-164">start</span><span class="sxs-lookup"><span data-stu-id="fa854-164">start</span></span>](#start-datetimejavascriptapioutlookofficetime) | <span data-ttu-id="fa854-165">Элемент</span><span class="sxs-lookup"><span data-stu-id="fa854-165">Member</span></span> |
| [<span data-ttu-id="fa854-166">subject</span><span class="sxs-lookup"><span data-stu-id="fa854-166">subject</span></span>](#subject-stringsubjectjavascriptapioutlookofficesubject) | <span data-ttu-id="fa854-167">Элемент</span><span class="sxs-lookup"><span data-stu-id="fa854-167">Member</span></span> |
| [<span data-ttu-id="fa854-168">to</span><span class="sxs-lookup"><span data-stu-id="fa854-168">to</span></span>](#to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="fa854-169">Элемент</span><span class="sxs-lookup"><span data-stu-id="fa854-169">Member</span></span> |
| [<span data-ttu-id="fa854-170">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="fa854-170">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="fa854-171">Метод</span><span class="sxs-lookup"><span data-stu-id="fa854-171">Method</span></span> |
| [<span data-ttu-id="fa854-172">addFileAttachmentFromBase64Async</span><span class="sxs-lookup"><span data-stu-id="fa854-172">addFileAttachmentFromBase64Async</span></span>](#addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback) | <span data-ttu-id="fa854-173">Метод</span><span class="sxs-lookup"><span data-stu-id="fa854-173">Method</span></span> |
| [<span data-ttu-id="fa854-174">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="fa854-174">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="fa854-175">Метод</span><span class="sxs-lookup"><span data-stu-id="fa854-175">Method</span></span> |
| [<span data-ttu-id="fa854-176">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="fa854-176">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="fa854-177">Метод</span><span class="sxs-lookup"><span data-stu-id="fa854-177">Method</span></span> |
| [<span data-ttu-id="fa854-178">close</span><span class="sxs-lookup"><span data-stu-id="fa854-178">close</span></span>](#close) | <span data-ttu-id="fa854-179">Метод</span><span class="sxs-lookup"><span data-stu-id="fa854-179">Method</span></span> |
| [<span data-ttu-id="fa854-180">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="fa854-180">displayReplyAllForm</span></span>](#displayreplyallformformdata) | <span data-ttu-id="fa854-181">Метод</span><span class="sxs-lookup"><span data-stu-id="fa854-181">Method</span></span> |
| [<span data-ttu-id="fa854-182">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="fa854-182">displayReplyForm</span></span>](#displayreplyformformdata) | <span data-ttu-id="fa854-183">Метод</span><span class="sxs-lookup"><span data-stu-id="fa854-183">Method</span></span> |
| [<span data-ttu-id="fa854-184">getAttachmentContentAsync</span><span class="sxs-lookup"><span data-stu-id="fa854-184">getAttachmentContentAsync</span></span>](#getattachmentcontentasyncattachmentid-options-callback--attachmentcontentjavascriptapioutlookofficeattachmentcontent) | <span data-ttu-id="fa854-185">Метод</span><span class="sxs-lookup"><span data-stu-id="fa854-185">Method</span></span> |
| [<span data-ttu-id="fa854-186">getAttachmentsAsync</span><span class="sxs-lookup"><span data-stu-id="fa854-186">getAttachmentsAsync</span></span>](#getattachmentsasyncoptions-callback--arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails) | <span data-ttu-id="fa854-187">Метод</span><span class="sxs-lookup"><span data-stu-id="fa854-187">Method</span></span> |
| [<span data-ttu-id="fa854-188">getEntities</span><span class="sxs-lookup"><span data-stu-id="fa854-188">getEntities</span></span>](#getentities--entitiesjavascriptapioutlookofficeentities) | <span data-ttu-id="fa854-189">Метод</span><span class="sxs-lookup"><span data-stu-id="fa854-189">Method</span></span> |
| [<span data-ttu-id="fa854-190">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="fa854-190">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion) | <span data-ttu-id="fa854-191">Метод</span><span class="sxs-lookup"><span data-stu-id="fa854-191">Method</span></span> |
| [<span data-ttu-id="fa854-192">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="fa854-192">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion) | <span data-ttu-id="fa854-193">Метод</span><span class="sxs-lookup"><span data-stu-id="fa854-193">Method</span></span> |
| [<span data-ttu-id="fa854-194">getInitializationContextAsync</span><span class="sxs-lookup"><span data-stu-id="fa854-194">getInitializationContextAsync</span></span>](#getinitializationcontextasyncoptions-callback) | <span data-ttu-id="fa854-195">Метод</span><span class="sxs-lookup"><span data-stu-id="fa854-195">Method</span></span> |
| [<span data-ttu-id="fa854-196">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="fa854-196">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="fa854-197">Метод</span><span class="sxs-lookup"><span data-stu-id="fa854-197">Method</span></span> |
| [<span data-ttu-id="fa854-198">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="fa854-198">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="fa854-199">Метод</span><span class="sxs-lookup"><span data-stu-id="fa854-199">Method</span></span> |
| [<span data-ttu-id="fa854-200">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="fa854-200">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="fa854-201">Метод</span><span class="sxs-lookup"><span data-stu-id="fa854-201">Method</span></span> |
| [<span data-ttu-id="fa854-202">getSelectedEntities</span><span class="sxs-lookup"><span data-stu-id="fa854-202">getSelectedEntities</span></span>](#getselectedentities--entitiesjavascriptapioutlookofficeentities) | <span data-ttu-id="fa854-203">Метод</span><span class="sxs-lookup"><span data-stu-id="fa854-203">Method</span></span> |
| [<span data-ttu-id="fa854-204">getSelectedRegExMatches</span><span class="sxs-lookup"><span data-stu-id="fa854-204">getSelectedRegExMatches</span></span>](#getselectedregexmatches--object) | <span data-ttu-id="fa854-205">Метод</span><span class="sxs-lookup"><span data-stu-id="fa854-205">Method</span></span> |
| [<span data-ttu-id="fa854-206">getSharedPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="fa854-206">getSharedPropertiesAsync</span></span>](#getsharedpropertiesasyncoptions-callback) | <span data-ttu-id="fa854-207">Метод</span><span class="sxs-lookup"><span data-stu-id="fa854-207">Method</span></span> |
| [<span data-ttu-id="fa854-208">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="fa854-208">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="fa854-209">Метод</span><span class="sxs-lookup"><span data-stu-id="fa854-209">Method</span></span> |
| [<span data-ttu-id="fa854-210">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="fa854-210">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="fa854-211">Метод</span><span class="sxs-lookup"><span data-stu-id="fa854-211">Method</span></span> |
| [<span data-ttu-id="fa854-212">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="fa854-212">removeHandlerAsync</span></span>](#removehandlerasynceventtype-options-callback) | <span data-ttu-id="fa854-213">Метод</span><span class="sxs-lookup"><span data-stu-id="fa854-213">Method</span></span> |
| [<span data-ttu-id="fa854-214">saveAsync</span><span class="sxs-lookup"><span data-stu-id="fa854-214">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="fa854-215">Метод</span><span class="sxs-lookup"><span data-stu-id="fa854-215">Method</span></span> |
| [<span data-ttu-id="fa854-216">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="fa854-216">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="fa854-217">Метод</span><span class="sxs-lookup"><span data-stu-id="fa854-217">Method</span></span> |

### <a name="example"></a><span data-ttu-id="fa854-218">Пример</span><span class="sxs-lookup"><span data-stu-id="fa854-218">Example</span></span>

<span data-ttu-id="fa854-219">В примере кода JavaScript, приведенном ниже, показано, как получить доступ к свойству `subject` текущего элемента в Outlook.</span><span class="sxs-lookup"><span data-stu-id="fa854-219">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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
}
```

### <a name="members"></a><span data-ttu-id="fa854-220">Элементы</span><span class="sxs-lookup"><span data-stu-id="fa854-220">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails"></a><span data-ttu-id="fa854-221">attachments :Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="fa854-221">attachments :Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

<span data-ttu-id="fa854-222">Получает вложения элемента в качестве массива.</span><span class="sxs-lookup"><span data-stu-id="fa854-222">Gets the item's attachments as an array.</span></span> <span data-ttu-id="fa854-223">Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="fa854-223">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="fa854-224">Outlook блокирует определенные типы файлов из-за потенциальных проблем с безопасностью, поэтому они не возвращаются.</span><span class="sxs-lookup"><span data-stu-id="fa854-224">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="fa854-225">Дополнительные сведения см. в статье [Блокировка вложений в Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="fa854-225">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="fa854-226">Тип:</span><span class="sxs-lookup"><span data-stu-id="fa854-226">Type:</span></span>

*   <span data-ttu-id="fa854-227">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="fa854-227">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="fa854-228">Требования</span><span class="sxs-lookup"><span data-stu-id="fa854-228">Requirements</span></span>

|<span data-ttu-id="fa854-229">Требование</span><span class="sxs-lookup"><span data-stu-id="fa854-229">Requirement</span></span>|<span data-ttu-id="fa854-230">Значение</span><span class="sxs-lookup"><span data-stu-id="fa854-230">Value</span></span>|
|---|---|
|[<span data-ttu-id="fa854-231">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="fa854-231">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="fa854-232">1.0</span><span class="sxs-lookup"><span data-stu-id="fa854-232">1.0</span></span>|
|[<span data-ttu-id="fa854-233">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="fa854-233">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="fa854-234">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fa854-234">ReadItem</span></span>|
|[<span data-ttu-id="fa854-235">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="fa854-235">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="fa854-236">Чтение</span><span class="sxs-lookup"><span data-stu-id="fa854-236">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="fa854-237">Пример</span><span class="sxs-lookup"><span data-stu-id="fa854-237">Example</span></span>

<span data-ttu-id="fa854-238">С помощью приведенного ниже кода можно создать HTML-строку с подробными сведениями обо всех вложениях для текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="fa854-238">The following code builds an HTML string with details of all attachments on the current item.</span></span>

```javascript
var _Item = Office.context.mailbox.item;
var outputString = "";

if (_Item.attachments.length > 0) {
  for (i = 0 ; i < _Item.attachments.length ; i++) {
    var _att = _Item.attachments[i];
    outputString += "<BR>" + i + ". Name: ";
    outputString += _att.name;
    outputString += "<BR>ID: " + _att.id;
    outputString += "<BR>contentType: " + _att.contentType;
    outputString += "<BR>size: " + _att.size;
    outputString += "<BR>attachmentType: " + _att.attachmentType;
    outputString += "<BR>isInline: " + _att.isInline;
  }
}

// Do something with outputString
```

####  <a name="bcc-recipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="fa854-239">bcc :[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="fa854-239">bcc :[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="fa854-240">Получает объект, который предоставляет методы для получения или обновления получателей скрытой копии сообщения.</span><span class="sxs-lookup"><span data-stu-id="fa854-240">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="fa854-241">Только в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="fa854-241">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="fa854-242">Тип:</span><span class="sxs-lookup"><span data-stu-id="fa854-242">Type:</span></span>

*   [<span data-ttu-id="fa854-243">Recipients</span><span class="sxs-lookup"><span data-stu-id="fa854-243">Recipients</span></span>](/javascript/api/outlook/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="fa854-244">Требования</span><span class="sxs-lookup"><span data-stu-id="fa854-244">Requirements</span></span>

|<span data-ttu-id="fa854-245">Требование</span><span class="sxs-lookup"><span data-stu-id="fa854-245">Requirement</span></span>|<span data-ttu-id="fa854-246">Значение</span><span class="sxs-lookup"><span data-stu-id="fa854-246">Value</span></span>|
|---|---|
|[<span data-ttu-id="fa854-247">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="fa854-247">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="fa854-248">1.1</span><span class="sxs-lookup"><span data-stu-id="fa854-248">1.1</span></span>|
|[<span data-ttu-id="fa854-249">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="fa854-249">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="fa854-250">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fa854-250">ReadItem</span></span>|
|[<span data-ttu-id="fa854-251">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="fa854-251">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="fa854-252">Создание</span><span class="sxs-lookup"><span data-stu-id="fa854-252">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="fa854-253">Пример</span><span class="sxs-lookup"><span data-stu-id="fa854-253">Example</span></span>

```javascript
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlookofficebody"></a><span data-ttu-id="fa854-254">body :[Body](/javascript/api/outlook/office.body)</span><span class="sxs-lookup"><span data-stu-id="fa854-254">body :[Body](/javascript/api/outlook/office.body)</span></span>

<span data-ttu-id="fa854-255">Получает объект, предоставляющий методы для работы с основным текстом элемента.</span><span class="sxs-lookup"><span data-stu-id="fa854-255">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="fa854-256">Тип:</span><span class="sxs-lookup"><span data-stu-id="fa854-256">Type:</span></span>

*   [<span data-ttu-id="fa854-257">Body</span><span class="sxs-lookup"><span data-stu-id="fa854-257">Body</span></span>](/javascript/api/outlook/office.body)

##### <a name="requirements"></a><span data-ttu-id="fa854-258">Требования</span><span class="sxs-lookup"><span data-stu-id="fa854-258">Requirements</span></span>

|<span data-ttu-id="fa854-259">Требование</span><span class="sxs-lookup"><span data-stu-id="fa854-259">Requirement</span></span>|<span data-ttu-id="fa854-260">Значение</span><span class="sxs-lookup"><span data-stu-id="fa854-260">Value</span></span>|
|---|---|
|[<span data-ttu-id="fa854-261">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="fa854-261">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="fa854-262">1.1</span><span class="sxs-lookup"><span data-stu-id="fa854-262">1.1</span></span>|
|[<span data-ttu-id="fa854-263">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="fa854-263">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="fa854-264">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fa854-264">ReadItem</span></span>|
|[<span data-ttu-id="fa854-265">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="fa854-265">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="fa854-266">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="fa854-266">Compose or read</span></span>|

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="fa854-267">cc :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="fa854-267">cc :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="fa854-268">Предоставляет доступ к получателям копии сообщения.</span><span class="sxs-lookup"><span data-stu-id="fa854-268">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="fa854-269">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="fa854-269">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="fa854-270">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="fa854-270">Read mode</span></span>

<span data-ttu-id="fa854-p106">Свойство `cc` возвращает массив, который содержит объект `EmailAddressDetails` для каждого получателя, указанного в строке **Копия** сообщения. Коллекция может включать не более 100 элементов.</span><span class="sxs-lookup"><span data-stu-id="fa854-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="fa854-273">Режим создания</span><span class="sxs-lookup"><span data-stu-id="fa854-273">Compose mode</span></span>

<span data-ttu-id="fa854-274">Свойство `cc` возвращает объект `Recipients`, предоставляющий методы для получения или обновления получателей, которые указаны в строке **Копия** сообщения.</span><span class="sxs-lookup"><span data-stu-id="fa854-274">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="fa854-275">Тип:</span><span class="sxs-lookup"><span data-stu-id="fa854-275">Type:</span></span>

*   <span data-ttu-id="fa854-276">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="fa854-276">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="fa854-277">Требования</span><span class="sxs-lookup"><span data-stu-id="fa854-277">Requirements</span></span>

|<span data-ttu-id="fa854-278">Требование</span><span class="sxs-lookup"><span data-stu-id="fa854-278">Requirement</span></span>|<span data-ttu-id="fa854-279">Значение</span><span class="sxs-lookup"><span data-stu-id="fa854-279">Value</span></span>|
|---|---|
|[<span data-ttu-id="fa854-280">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="fa854-280">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="fa854-281">1.0</span><span class="sxs-lookup"><span data-stu-id="fa854-281">1.0</span></span>|
|[<span data-ttu-id="fa854-282">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="fa854-282">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="fa854-283">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fa854-283">ReadItem</span></span>|
|[<span data-ttu-id="fa854-284">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="fa854-284">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="fa854-285">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="fa854-285">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="fa854-286">Пример</span><span class="sxs-lookup"><span data-stu-id="fa854-286">Example</span></span>

```javascript
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="fa854-287">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="fa854-287">(nullable) conversationId :String</span></span>

<span data-ttu-id="fa854-288">Получает идентификатор разговора по электронной почте, содержащего конкретное сообщение.</span><span class="sxs-lookup"><span data-stu-id="fa854-288">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="fa854-p107">Вы можете получить целочисленное значение этого свойства, если ваше почтовое приложение активируется в формах просмотра или формах создания ответов. Если пользователь изменит тему ответа, после его отправки идентификатор беседы будет изменен, и полученное ранее значение будет недействительным.</span><span class="sxs-lookup"><span data-stu-id="fa854-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="fa854-p108">Это свойство имеет значение NULL для нового элемента в форме создания. Свойство `conversationId` вернет значение, если пользователь задаст тему и сохранит элемент.</span><span class="sxs-lookup"><span data-stu-id="fa854-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="fa854-293">Тип:</span><span class="sxs-lookup"><span data-stu-id="fa854-293">Type:</span></span>

*   <span data-ttu-id="fa854-294">String</span><span class="sxs-lookup"><span data-stu-id="fa854-294">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="fa854-295">Требования</span><span class="sxs-lookup"><span data-stu-id="fa854-295">Requirements</span></span>

|<span data-ttu-id="fa854-296">Требование</span><span class="sxs-lookup"><span data-stu-id="fa854-296">Requirement</span></span>|<span data-ttu-id="fa854-297">Значение</span><span class="sxs-lookup"><span data-stu-id="fa854-297">Value</span></span>|
|---|---|
|[<span data-ttu-id="fa854-298">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="fa854-298">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="fa854-299">1.0</span><span class="sxs-lookup"><span data-stu-id="fa854-299">1.0</span></span>|
|[<span data-ttu-id="fa854-300">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="fa854-300">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="fa854-301">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fa854-301">ReadItem</span></span>|
|[<span data-ttu-id="fa854-302">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="fa854-302">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="fa854-303">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="fa854-303">Compose or read</span></span>|

#### <a name="datetimecreated-date"></a><span data-ttu-id="fa854-304">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="fa854-304">dateTimeCreated :Date</span></span>

<span data-ttu-id="fa854-p109">Получает дату и время создания элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="fa854-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="fa854-307">Тип:</span><span class="sxs-lookup"><span data-stu-id="fa854-307">Type:</span></span>

*   <span data-ttu-id="fa854-308">Date</span><span class="sxs-lookup"><span data-stu-id="fa854-308">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="fa854-309">Требования</span><span class="sxs-lookup"><span data-stu-id="fa854-309">Requirements</span></span>

|<span data-ttu-id="fa854-310">Требование</span><span class="sxs-lookup"><span data-stu-id="fa854-310">Requirement</span></span>|<span data-ttu-id="fa854-311">Значение</span><span class="sxs-lookup"><span data-stu-id="fa854-311">Value</span></span>|
|---|---|
|[<span data-ttu-id="fa854-312">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="fa854-312">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="fa854-313">1.0</span><span class="sxs-lookup"><span data-stu-id="fa854-313">1.0</span></span>|
|[<span data-ttu-id="fa854-314">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="fa854-314">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="fa854-315">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fa854-315">ReadItem</span></span>|
|[<span data-ttu-id="fa854-316">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="fa854-316">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="fa854-317">Чтение</span><span class="sxs-lookup"><span data-stu-id="fa854-317">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="fa854-318">Пример</span><span class="sxs-lookup"><span data-stu-id="fa854-318">Example</span></span>

```javascript
var created = Office.context.mailbox.item.dateTimeCreated;
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="fa854-319">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="fa854-319">dateTimeModified :Date</span></span>

<span data-ttu-id="fa854-p110">Получает дату и время последнего изменения элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="fa854-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="fa854-322">Этот элемент не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="fa854-322">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="fa854-323">Тип:</span><span class="sxs-lookup"><span data-stu-id="fa854-323">Type:</span></span>

*   <span data-ttu-id="fa854-324">Date</span><span class="sxs-lookup"><span data-stu-id="fa854-324">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="fa854-325">Требования</span><span class="sxs-lookup"><span data-stu-id="fa854-325">Requirements</span></span>

|<span data-ttu-id="fa854-326">Требование</span><span class="sxs-lookup"><span data-stu-id="fa854-326">Requirement</span></span>|<span data-ttu-id="fa854-327">Значение</span><span class="sxs-lookup"><span data-stu-id="fa854-327">Value</span></span>|
|---|---|
|[<span data-ttu-id="fa854-328">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="fa854-328">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="fa854-329">1.0</span><span class="sxs-lookup"><span data-stu-id="fa854-329">1.0</span></span>|
|[<span data-ttu-id="fa854-330">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="fa854-330">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="fa854-331">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fa854-331">ReadItem</span></span>|
|[<span data-ttu-id="fa854-332">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="fa854-332">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="fa854-333">Чтение</span><span class="sxs-lookup"><span data-stu-id="fa854-333">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="fa854-334">Пример</span><span class="sxs-lookup"><span data-stu-id="fa854-334">Example</span></span>

```javascript
var modified = Office.context.mailbox.item.dateTimeModified;
```

####  <a name="end-datetimejavascriptapioutlookofficetime"></a><span data-ttu-id="fa854-335">end :Date|[Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="fa854-335">end :Date|[Time](/javascript/api/outlook/office.time)</span></span>

<span data-ttu-id="fa854-336">Получает или задает дату и время окончания встречи.</span><span class="sxs-lookup"><span data-stu-id="fa854-336">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="fa854-p111">Свойство `end` представлено в виде значения даты и времени в формате UTC. Преобразовать значение свойства end в местные значения даты и времени клиента можно с помощью метода [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime).</span><span class="sxs-lookup"><span data-stu-id="fa854-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="fa854-339">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="fa854-339">Read mode</span></span>

<span data-ttu-id="fa854-340">Свойство `end` возвращает объект `Date`.</span><span class="sxs-lookup"><span data-stu-id="fa854-340">The `end` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="fa854-341">Режим создания</span><span class="sxs-lookup"><span data-stu-id="fa854-341">Compose mode</span></span>

<span data-ttu-id="fa854-342">Свойство `end` возвращает объект `Time`.</span><span class="sxs-lookup"><span data-stu-id="fa854-342">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="fa854-343">Если вы задаете время окончания с помощью метода [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-), необходимо использовать метод [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) для преобразования местного времени на клиенте в формат UTC для сервера.</span><span class="sxs-lookup"><span data-stu-id="fa854-343">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="fa854-344">Тип:</span><span class="sxs-lookup"><span data-stu-id="fa854-344">Type:</span></span>

*   <span data-ttu-id="fa854-345">Date | [Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="fa854-345">Date | [Time](/javascript/api/outlook/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="fa854-346">Требования</span><span class="sxs-lookup"><span data-stu-id="fa854-346">Requirements</span></span>

|<span data-ttu-id="fa854-347">Требование</span><span class="sxs-lookup"><span data-stu-id="fa854-347">Requirement</span></span>|<span data-ttu-id="fa854-348">Значение</span><span class="sxs-lookup"><span data-stu-id="fa854-348">Value</span></span>|
|---|---|
|[<span data-ttu-id="fa854-349">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="fa854-349">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="fa854-350">1.0</span><span class="sxs-lookup"><span data-stu-id="fa854-350">1.0</span></span>|
|[<span data-ttu-id="fa854-351">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="fa854-351">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="fa854-352">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fa854-352">ReadItem</span></span>|
|[<span data-ttu-id="fa854-353">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="fa854-353">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="fa854-354">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="fa854-354">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="fa854-355">Пример</span><span class="sxs-lookup"><span data-stu-id="fa854-355">Example</span></span>

<span data-ttu-id="fa854-356">В примере ниже показано, как с помощью метода [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) объекта `Time` задать время окончания встречи в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="fa854-356">The following example sets the end time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

```javascript
var endTime = new Date("3/14/2015");
var options = {
  // Pass information that can be used
  // in the callback
  asyncContext: {verb:"Set"}
}
Office.context.mailbox.item.end.setAsync(endTime, options, function(result) {
  if (result.error) {
    console.debug(result.error);
  } else {
    // Access the asyncContext that was passed to the setAsync function
    console.debug("End Time " + result.asyncContext.verb);
  }
});
```

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom"></a><span data-ttu-id="fa854-357">from :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[From](/javascript/api/outlook/office.from)</span><span class="sxs-lookup"><span data-stu-id="fa854-357">from :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[From](/javascript/api/outlook/office.from)</span></span>

<span data-ttu-id="fa854-358">Получает адрес электронной почты отправителя сообщения.</span><span class="sxs-lookup"><span data-stu-id="fa854-358">Gets the email address of the sender of a message.</span></span>

<span data-ttu-id="fa854-p112">Свойства `from` и [`sender`](#sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails) представляют одно лицо, если сообщение не отправлено представителем. В противном случае свойство `from` представляет лицо, делегировавшее полномочия, а свойство sender — представителя.</span><span class="sxs-lookup"><span data-stu-id="fa854-p112">The `from` and [`sender`](#sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="fa854-361">Свойству `recipientType`, принадлежащему объекту `EmailAddressDetails` в свойстве `from`, задано значение `undefined`.</span><span class="sxs-lookup"><span data-stu-id="fa854-361">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="fa854-362">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="fa854-362">Read mode</span></span>

<span data-ttu-id="fa854-363">Свойство `from` возвращает объект `EmailAddressDetails`.</span><span class="sxs-lookup"><span data-stu-id="fa854-363">The `from` property returns an `EmailAddressDetails` object.</span></span>

```javascript
var subject = Office.context.mailbox.item.from;
```

##### <a name="compose-mode"></a><span data-ttu-id="fa854-364">Режим создания</span><span class="sxs-lookup"><span data-stu-id="fa854-364">Compose mode</span></span>

<span data-ttu-id="fa854-365">Свойство `from` возвращает объект `From`, который предоставляет метод для получения значения отправителя.</span><span class="sxs-lookup"><span data-stu-id="fa854-365">The `from` property returns a `From` object that provides a method to get the from value.</span></span>

```javascript
Office.context.mailbox.item.from.getAsync(callback);

function callback(asyncResult) {
  var from = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="fa854-366">Тип:</span><span class="sxs-lookup"><span data-stu-id="fa854-366">Type:</span></span>

*   <span data-ttu-id="fa854-367">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [From](/javascript/api/outlook/office.from)</span><span class="sxs-lookup"><span data-stu-id="fa854-367">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [From](/javascript/api/outlook/office.from)</span></span>

##### <a name="requirements"></a><span data-ttu-id="fa854-368">Требования</span><span class="sxs-lookup"><span data-stu-id="fa854-368">Requirements</span></span>

|<span data-ttu-id="fa854-369">Требование</span><span class="sxs-lookup"><span data-stu-id="fa854-369">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="fa854-370">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="fa854-370">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="fa854-371">1.0</span><span class="sxs-lookup"><span data-stu-id="fa854-371">1.0</span></span>|<span data-ttu-id="fa854-372">1.7</span><span class="sxs-lookup"><span data-stu-id="fa854-372">1.7</span></span>|
|[<span data-ttu-id="fa854-373">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="fa854-373">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="fa854-374">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fa854-374">ReadItem</span></span>|<span data-ttu-id="fa854-375">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="fa854-375">ReadWriteItem</span></span>|
|[<span data-ttu-id="fa854-376">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="fa854-376">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="fa854-377">Чтение</span><span class="sxs-lookup"><span data-stu-id="fa854-377">Read</span></span>|<span data-ttu-id="fa854-378">Создание</span><span class="sxs-lookup"><span data-stu-id="fa854-378">Compose</span></span>|

#### <a name="internetheaders-internetheadersjavascriptapioutlookofficeinternetheaders"></a><span data-ttu-id="fa854-379">internetHeaders :[InternetHeaders](/javascript/api/outlook/office.internetheaders)</span><span class="sxs-lookup"><span data-stu-id="fa854-379">internetHeaders :[InternetHeaders](/javascript/api/outlook/office.internetheaders)</span></span>

<span data-ttu-id="fa854-380">Получает или задает заголовки Интернета в сообщении.</span><span class="sxs-lookup"><span data-stu-id="fa854-380">Gets or sets the internet headers of a message.</span></span>

##### <a name="type"></a><span data-ttu-id="fa854-381">Тип:</span><span class="sxs-lookup"><span data-stu-id="fa854-381">Type:</span></span>

*   [<span data-ttu-id="fa854-382">InternetHeaders</span><span class="sxs-lookup"><span data-stu-id="fa854-382">InternetHeaders</span></span>](/javascript/api/outlook/office.internetheaders)

##### <a name="requirements"></a><span data-ttu-id="fa854-383">Требования</span><span class="sxs-lookup"><span data-stu-id="fa854-383">Requirements</span></span>

|<span data-ttu-id="fa854-384">Требование</span><span class="sxs-lookup"><span data-stu-id="fa854-384">Requirement</span></span>|<span data-ttu-id="fa854-385">Значение</span><span class="sxs-lookup"><span data-stu-id="fa854-385">Value</span></span>|
|---|---|
|[<span data-ttu-id="fa854-386">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="fa854-386">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="fa854-387">Предварительная версия</span><span class="sxs-lookup"><span data-stu-id="fa854-387">Preview</span></span>|
|[<span data-ttu-id="fa854-388">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="fa854-388">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="fa854-389">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fa854-389">ReadItem</span></span>|
|[<span data-ttu-id="fa854-390">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="fa854-390">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="fa854-391">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="fa854-391">Compose or read</span></span>|

#### <a name="internetmessageid-string"></a><span data-ttu-id="fa854-392">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="fa854-392">internetMessageId :String</span></span>

<span data-ttu-id="fa854-p113">Получает идентификатор интернет-сообщения для электронного сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="fa854-p113">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="fa854-395">Тип:</span><span class="sxs-lookup"><span data-stu-id="fa854-395">Type:</span></span>

*   <span data-ttu-id="fa854-396">String</span><span class="sxs-lookup"><span data-stu-id="fa854-396">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="fa854-397">Требования</span><span class="sxs-lookup"><span data-stu-id="fa854-397">Requirements</span></span>

|<span data-ttu-id="fa854-398">Требование</span><span class="sxs-lookup"><span data-stu-id="fa854-398">Requirement</span></span>|<span data-ttu-id="fa854-399">Значение</span><span class="sxs-lookup"><span data-stu-id="fa854-399">Value</span></span>|
|---|---|
|[<span data-ttu-id="fa854-400">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="fa854-400">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="fa854-401">1.0</span><span class="sxs-lookup"><span data-stu-id="fa854-401">1.0</span></span>|
|[<span data-ttu-id="fa854-402">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="fa854-402">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="fa854-403">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fa854-403">ReadItem</span></span>|
|[<span data-ttu-id="fa854-404">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="fa854-404">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="fa854-405">Чтение</span><span class="sxs-lookup"><span data-stu-id="fa854-405">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="fa854-406">Пример</span><span class="sxs-lookup"><span data-stu-id="fa854-406">Example</span></span>

```javascript
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="fa854-407">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="fa854-407">itemClass :String</span></span>

<span data-ttu-id="fa854-p114">Получает класс элемента веб-служб Exchange для выбранного элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="fa854-p114">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="fa854-p115">Свойство `itemClass` указывает класс сообщения выбранного элемента. Ниже приводятся классы сообщения по умолчанию для элемента сообщения или встречи.</span><span class="sxs-lookup"><span data-stu-id="fa854-p115">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

|<span data-ttu-id="fa854-412">Тип</span><span class="sxs-lookup"><span data-stu-id="fa854-412">Type</span></span>|<span data-ttu-id="fa854-413">Описание</span><span class="sxs-lookup"><span data-stu-id="fa854-413">Description</span></span>|<span data-ttu-id="fa854-414">Класс элемента</span><span class="sxs-lookup"><span data-stu-id="fa854-414">item class</span></span>|
|---|---|---|
|<span data-ttu-id="fa854-415">Элементы встречи</span><span class="sxs-lookup"><span data-stu-id="fa854-415">Appointment items</span></span>|<span data-ttu-id="fa854-416">Это элементы календаря для класса элемента `IPM.Appointment` или `IPM.Appointment.Occurence`.</span><span class="sxs-lookup"><span data-stu-id="fa854-416">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurence`.</span></span>|`IPM.Appointment`<br />`IPM.Appointment.Occurence`|
|<span data-ttu-id="fa854-417">Элементы сообщения</span><span class="sxs-lookup"><span data-stu-id="fa854-417">Message items</span></span>|<span data-ttu-id="fa854-418">Сюда входят электронные сообщения, для которых по умолчанию задан класс сообщения `IPM.Note`, а также приглашения на собрания, ответы на них и уведомления об их отмене, использующие `IPM.Schedule.Meeting` в качестве базового класса сообщения.</span><span class="sxs-lookup"><span data-stu-id="fa854-418">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span>|`IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled`|

<span data-ttu-id="fa854-419">Можно создавать настраиваемые классы сообщения, расширяющие классы сообщения по умолчанию, например настраиваемый класс сообщения о встрече `IPM.Appointment.Contoso`.</span><span class="sxs-lookup"><span data-stu-id="fa854-419">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="fa854-420">Тип:</span><span class="sxs-lookup"><span data-stu-id="fa854-420">Type:</span></span>

*   <span data-ttu-id="fa854-421">String</span><span class="sxs-lookup"><span data-stu-id="fa854-421">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="fa854-422">Требования</span><span class="sxs-lookup"><span data-stu-id="fa854-422">Requirements</span></span>

|<span data-ttu-id="fa854-423">Требование</span><span class="sxs-lookup"><span data-stu-id="fa854-423">Requirement</span></span>|<span data-ttu-id="fa854-424">Значение</span><span class="sxs-lookup"><span data-stu-id="fa854-424">Value</span></span>|
|---|---|
|[<span data-ttu-id="fa854-425">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="fa854-425">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="fa854-426">1.0</span><span class="sxs-lookup"><span data-stu-id="fa854-426">1.0</span></span>|
|[<span data-ttu-id="fa854-427">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="fa854-427">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="fa854-428">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fa854-428">ReadItem</span></span>|
|[<span data-ttu-id="fa854-429">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="fa854-429">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="fa854-430">Чтение</span><span class="sxs-lookup"><span data-stu-id="fa854-430">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="fa854-431">Пример</span><span class="sxs-lookup"><span data-stu-id="fa854-431">Example</span></span>

```javascript
var itemClass = Office.context.mailbox.item.itemClass;
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="fa854-432">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="fa854-432">(nullable) itemId :String</span></span>

<span data-ttu-id="fa854-p116">Получает идентификатор элемента веб-служб Exchange для текущего элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="fa854-p116">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="fa854-435">Идентификатор, возвращаемый свойством `itemId`, совпадает с идентификатором элемента веб-служб Exchange.</span><span class="sxs-lookup"><span data-stu-id="fa854-435">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="fa854-436">Свойство `itemId` не совпадает с идентификатором записи Outlook, а также идентификатором, который используется REST API Outlook.</span><span class="sxs-lookup"><span data-stu-id="fa854-436">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="fa854-437">Прежде чем совершать вызовы REST API, используя это значение, его необходимо преобразовать с помощью [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="fa854-437">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="fa854-438">Дополнительные сведения см. в статье [Использование REST API Outlook из надстройки Outlook](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="fa854-438">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="fa854-p118">Свойство `itemId` недоступно в режиме создания. Если требуется идентификатор элемента, с помощью метода [`saveAsync`](#saveasyncoptions-callback) можно сохранить элемент в хранилище. При этом в параметре [`AsyncResult.value`](/javascript/api/office/office.asyncresult) функции обратного вызова возвращается идентификатор элемента.</span><span class="sxs-lookup"><span data-stu-id="fa854-p118">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="fa854-441">Тип:</span><span class="sxs-lookup"><span data-stu-id="fa854-441">Type:</span></span>

*   <span data-ttu-id="fa854-442">String</span><span class="sxs-lookup"><span data-stu-id="fa854-442">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="fa854-443">Требования</span><span class="sxs-lookup"><span data-stu-id="fa854-443">Requirements</span></span>

|<span data-ttu-id="fa854-444">Требование</span><span class="sxs-lookup"><span data-stu-id="fa854-444">Requirement</span></span>|<span data-ttu-id="fa854-445">Значение</span><span class="sxs-lookup"><span data-stu-id="fa854-445">Value</span></span>|
|---|---|
|[<span data-ttu-id="fa854-446">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="fa854-446">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="fa854-447">1.0</span><span class="sxs-lookup"><span data-stu-id="fa854-447">1.0</span></span>|
|[<span data-ttu-id="fa854-448">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="fa854-448">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="fa854-449">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fa854-449">ReadItem</span></span>|
|[<span data-ttu-id="fa854-450">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="fa854-450">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="fa854-451">Чтение</span><span class="sxs-lookup"><span data-stu-id="fa854-451">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="fa854-452">Пример</span><span class="sxs-lookup"><span data-stu-id="fa854-452">Example</span></span>

<span data-ttu-id="fa854-p119">Указанный ниже код проверяет наличие идентификатора элемента. Если свойство `itemId` возвращает значение `null` или `undefined`, элемент будет сохранен в хранилище, а из асинхронного результата будет получен идентификатор элемента.</span><span class="sxs-lookup"><span data-stu-id="fa854-p119">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```javascript
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result){
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype"></a><span data-ttu-id="fa854-455">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="fa854-455">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="fa854-456">Получает тип элемента, который представляет экземпляр.</span><span class="sxs-lookup"><span data-stu-id="fa854-456">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="fa854-457">Свойство `itemType` возвращает одно из значений перечисления `ItemType`, которое указывает, является ли экземпляр объекта `item` сообщением или встречей.</span><span class="sxs-lookup"><span data-stu-id="fa854-457">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="fa854-458">Тип:</span><span class="sxs-lookup"><span data-stu-id="fa854-458">Type:</span></span>

*   [<span data-ttu-id="fa854-459">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="fa854-459">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="fa854-460">Требования</span><span class="sxs-lookup"><span data-stu-id="fa854-460">Requirements</span></span>

|<span data-ttu-id="fa854-461">Требование</span><span class="sxs-lookup"><span data-stu-id="fa854-461">Requirement</span></span>|<span data-ttu-id="fa854-462">Значение</span><span class="sxs-lookup"><span data-stu-id="fa854-462">Value</span></span>|
|---|---|
|[<span data-ttu-id="fa854-463">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="fa854-463">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="fa854-464">1.0</span><span class="sxs-lookup"><span data-stu-id="fa854-464">1.0</span></span>|
|[<span data-ttu-id="fa854-465">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="fa854-465">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="fa854-466">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fa854-466">ReadItem</span></span>|
|[<span data-ttu-id="fa854-467">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="fa854-467">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="fa854-468">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="fa854-468">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="fa854-469">Пример</span><span class="sxs-lookup"><span data-stu-id="fa854-469">Example</span></span>

```javascript
if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Message)
  // do something
else
  // do something else
```

####  <a name="location-stringlocationjavascriptapioutlookofficelocation"></a><span data-ttu-id="fa854-470">location :String|[Location](/javascript/api/outlook/office.location)</span><span class="sxs-lookup"><span data-stu-id="fa854-470">location :String|[Location](/javascript/api/outlook/office.location)</span></span>

<span data-ttu-id="fa854-471">Получает или задает место встречи.</span><span class="sxs-lookup"><span data-stu-id="fa854-471">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="fa854-472">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="fa854-472">Read mode</span></span>

<span data-ttu-id="fa854-473">Свойство `location` возвращает строку, содержащую сведения о месте встречи.</span><span class="sxs-lookup"><span data-stu-id="fa854-473">The `location` property returns a string that contains the location of the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="fa854-474">Режим создания</span><span class="sxs-lookup"><span data-stu-id="fa854-474">Compose mode</span></span>

<span data-ttu-id="fa854-475">Свойство `location` возвращает объект `Location`, предоставляющий методы, которые используются для получения и задания места встречи.</span><span class="sxs-lookup"><span data-stu-id="fa854-475">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="fa854-476">Тип:</span><span class="sxs-lookup"><span data-stu-id="fa854-476">Type:</span></span>

*   <span data-ttu-id="fa854-477">String | [Location](/javascript/api/outlook/office.location)</span><span class="sxs-lookup"><span data-stu-id="fa854-477">String | [Location](/javascript/api/outlook/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="fa854-478">Требования</span><span class="sxs-lookup"><span data-stu-id="fa854-478">Requirements</span></span>

|<span data-ttu-id="fa854-479">Требование</span><span class="sxs-lookup"><span data-stu-id="fa854-479">Requirement</span></span>|<span data-ttu-id="fa854-480">Значение</span><span class="sxs-lookup"><span data-stu-id="fa854-480">Value</span></span>|
|---|---|
|[<span data-ttu-id="fa854-481">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="fa854-481">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="fa854-482">1.0</span><span class="sxs-lookup"><span data-stu-id="fa854-482">1.0</span></span>|
|[<span data-ttu-id="fa854-483">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="fa854-483">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="fa854-484">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fa854-484">ReadItem</span></span>|
|[<span data-ttu-id="fa854-485">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="fa854-485">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="fa854-486">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="fa854-486">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="fa854-487">Пример</span><span class="sxs-lookup"><span data-stu-id="fa854-487">Example</span></span>

```javascript
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

#### <a name="normalizedsubject-string"></a><span data-ttu-id="fa854-488">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="fa854-488">normalizedSubject :String</span></span>

<span data-ttu-id="fa854-p120">Получает тему элемента со всеми удаленными префиксами (включая `RE:` и `FWD:`). Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="fa854-p120">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="fa854-p121">Свойство normalizedSubject получает тему элемента со стандартными префиксами (такими как `RE:` и `FW:`), добавляемыми почтовыми программами. Для получения темы элемента с неизмененными префиксами используйте свойство [`subject`](#subject-stringsubjectjavascriptapioutlookofficesubject).</span><span class="sxs-lookup"><span data-stu-id="fa854-p121">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubjectjavascriptapioutlookofficesubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="fa854-493">Тип:</span><span class="sxs-lookup"><span data-stu-id="fa854-493">Type:</span></span>

*   <span data-ttu-id="fa854-494">String</span><span class="sxs-lookup"><span data-stu-id="fa854-494">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="fa854-495">Требования</span><span class="sxs-lookup"><span data-stu-id="fa854-495">Requirements</span></span>

|<span data-ttu-id="fa854-496">Требование</span><span class="sxs-lookup"><span data-stu-id="fa854-496">Requirement</span></span>|<span data-ttu-id="fa854-497">Значение</span><span class="sxs-lookup"><span data-stu-id="fa854-497">Value</span></span>|
|---|---|
|[<span data-ttu-id="fa854-498">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="fa854-498">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="fa854-499">1.0</span><span class="sxs-lookup"><span data-stu-id="fa854-499">1.0</span></span>|
|[<span data-ttu-id="fa854-500">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="fa854-500">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="fa854-501">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fa854-501">ReadItem</span></span>|
|[<span data-ttu-id="fa854-502">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="fa854-502">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="fa854-503">Чтение</span><span class="sxs-lookup"><span data-stu-id="fa854-503">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="fa854-504">Пример</span><span class="sxs-lookup"><span data-stu-id="fa854-504">Example</span></span>

```javascript
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
```

####  <a name="notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessages"></a><span data-ttu-id="fa854-505">notificationMessages :[NotificationMessages](/javascript/api/outlook/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="fa854-505">notificationMessages :[NotificationMessages](/javascript/api/outlook/office.notificationmessages)</span></span>

<span data-ttu-id="fa854-506">Получает сообщения уведомления для элемента.</span><span class="sxs-lookup"><span data-stu-id="fa854-506">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="fa854-507">Тип:</span><span class="sxs-lookup"><span data-stu-id="fa854-507">Type:</span></span>

*   [<span data-ttu-id="fa854-508">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="fa854-508">NotificationMessages</span></span>](/javascript/api/outlook/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="fa854-509">Требования</span><span class="sxs-lookup"><span data-stu-id="fa854-509">Requirements</span></span>

|<span data-ttu-id="fa854-510">Требование</span><span class="sxs-lookup"><span data-stu-id="fa854-510">Requirement</span></span>|<span data-ttu-id="fa854-511">Значение</span><span class="sxs-lookup"><span data-stu-id="fa854-511">Value</span></span>|
|---|---|
|[<span data-ttu-id="fa854-512">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="fa854-512">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="fa854-513">1.3</span><span class="sxs-lookup"><span data-stu-id="fa854-513">1.3</span></span>|
|[<span data-ttu-id="fa854-514">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="fa854-514">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="fa854-515">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fa854-515">ReadItem</span></span>|
|[<span data-ttu-id="fa854-516">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="fa854-516">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="fa854-517">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="fa854-517">Compose or read</span></span>|

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="fa854-518">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="fa854-518">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="fa854-519">Предоставляет доступ к необязательным участникам события.</span><span class="sxs-lookup"><span data-stu-id="fa854-519">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="fa854-520">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="fa854-520">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="fa854-521">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="fa854-521">Read mode</span></span>

<span data-ttu-id="fa854-522">Свойство `optionalAttendees` возвращает массив, содержащий объект `EmailAddressDetails` для каждого необязательного участника собрания.</span><span class="sxs-lookup"><span data-stu-id="fa854-522">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="fa854-523">Режим создания</span><span class="sxs-lookup"><span data-stu-id="fa854-523">Compose mode</span></span>

<span data-ttu-id="fa854-524">Свойство `optionalAttendees` возвращает объект `Recipients`, который предоставляет методы для получения или обновления необязательных участников собрания.</span><span class="sxs-lookup"><span data-stu-id="fa854-524">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="fa854-525">Тип:</span><span class="sxs-lookup"><span data-stu-id="fa854-525">Type:</span></span>

*   <span data-ttu-id="fa854-526">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="fa854-526">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="fa854-527">Требования</span><span class="sxs-lookup"><span data-stu-id="fa854-527">Requirements</span></span>

|<span data-ttu-id="fa854-528">Требование</span><span class="sxs-lookup"><span data-stu-id="fa854-528">Requirement</span></span>|<span data-ttu-id="fa854-529">Значение</span><span class="sxs-lookup"><span data-stu-id="fa854-529">Value</span></span>|
|---|---|
|[<span data-ttu-id="fa854-530">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="fa854-530">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="fa854-531">1.0</span><span class="sxs-lookup"><span data-stu-id="fa854-531">1.0</span></span>|
|[<span data-ttu-id="fa854-532">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="fa854-532">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="fa854-533">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fa854-533">ReadItem</span></span>|
|[<span data-ttu-id="fa854-534">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="fa854-534">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="fa854-535">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="fa854-535">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="fa854-536">Пример</span><span class="sxs-lookup"><span data-stu-id="fa854-536">Example</span></span>

```javascript
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsorganizerjavascriptapioutlookofficeorganizer"></a><span data-ttu-id="fa854-537">organizer :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[Organizer](/javascript/api/outlook/office.organizer)</span><span class="sxs-lookup"><span data-stu-id="fa854-537">organizer :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[Organizer](/javascript/api/outlook/office.organizer)</span></span>

<span data-ttu-id="fa854-538">Получает адрес электронной почты организатора указанного собрания.</span><span class="sxs-lookup"><span data-stu-id="fa854-538">Gets the email address of the organizer for a specified meeting.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="fa854-539">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="fa854-539">Read mode</span></span>

<span data-ttu-id="fa854-540">Свойство `organizer` возвращает объект [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails), представляющий организатора собрания.</span><span class="sxs-lookup"><span data-stu-id="fa854-540">The `organizer` property returns an [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) object that represents the meeting organizer.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="fa854-541">Режим создания</span><span class="sxs-lookup"><span data-stu-id="fa854-541">Compose mode</span></span>

<span data-ttu-id="fa854-542">Свойство `organizer` возвращает объект [Organizer](/javascript/api/outlook/office.organizer), который предоставляет метод для получения значения организатора.</span><span class="sxs-lookup"><span data-stu-id="fa854-542">The `organizer` property returns an [Organizer](/javascript/api/outlook/office.organizer) object that provides a method to get the organizer value.</span></span>

##### <a name="type"></a><span data-ttu-id="fa854-543">Тип:</span><span class="sxs-lookup"><span data-stu-id="fa854-543">Type:</span></span>

*   <span data-ttu-id="fa854-544">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [Organizer](/javascript/api/outlook/office.organizer)</span><span class="sxs-lookup"><span data-stu-id="fa854-544">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [Organizer](/javascript/api/outlook/office.organizer)</span></span>

##### <a name="requirements"></a><span data-ttu-id="fa854-545">Требования</span><span class="sxs-lookup"><span data-stu-id="fa854-545">Requirements</span></span>

|<span data-ttu-id="fa854-546">Требование</span><span class="sxs-lookup"><span data-stu-id="fa854-546">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="fa854-547">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="fa854-547">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="fa854-548">1.0</span><span class="sxs-lookup"><span data-stu-id="fa854-548">1.0</span></span>|<span data-ttu-id="fa854-549">1.7</span><span class="sxs-lookup"><span data-stu-id="fa854-549">1.7</span></span>|
|[<span data-ttu-id="fa854-550">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="fa854-550">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="fa854-551">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fa854-551">ReadItem</span></span>|<span data-ttu-id="fa854-552">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="fa854-552">ReadWriteItem</span></span>|
|[<span data-ttu-id="fa854-553">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="fa854-553">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="fa854-554">Чтение</span><span class="sxs-lookup"><span data-stu-id="fa854-554">Read</span></span>|<span data-ttu-id="fa854-555">Создание</span><span class="sxs-lookup"><span data-stu-id="fa854-555">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="fa854-556">Пример</span><span class="sxs-lookup"><span data-stu-id="fa854-556">Example</span></span>

```javascript
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
```

#### <a name="nullable-recurrence-recurrencejavascriptapioutlookofficerecurrence"></a><span data-ttu-id="fa854-557">(nullable) recurrence :[Recurrence](/javascript/api/outlook/office.recurrence)</span><span class="sxs-lookup"><span data-stu-id="fa854-557">(nullable) recurrence :[Recurrence](/javascript/api/outlook/office.recurrence)</span></span>

<span data-ttu-id="fa854-558">Получает или задает расписание повторения для встречи.</span><span class="sxs-lookup"><span data-stu-id="fa854-558">Gets or sets the recurrence pattern of an appointment.</span></span> <span data-ttu-id="fa854-559">Получает расписание повторения для приглашения на собрание.</span><span class="sxs-lookup"><span data-stu-id="fa854-559">Gets the recurrence pattern of a meeting request.</span></span> <span data-ttu-id="fa854-560">Доступно в режимах чтения и создания для элементов встречи.</span><span class="sxs-lookup"><span data-stu-id="fa854-560">Read and compose modes for appointment items.</span></span> <span data-ttu-id="fa854-561">Доступно в режиме чтения для элементов приглашения на собрание.</span><span class="sxs-lookup"><span data-stu-id="fa854-561">Read mode for meeting request items.</span></span>

<span data-ttu-id="fa854-562">Свойство `recurrence` возвращает объект [recurrence](/javascript/api/outlook/office.recurrence) для повторяющихся встреч или приглашений на собрание, если элемент представляет собой серию или экземпляр в пределах серии.</span><span class="sxs-lookup"><span data-stu-id="fa854-562">The `recurrence` property returns a [recurrence](/javascript/api/outlook/office.recurrence) object for recurring appointments or meetings requests if an item is a series or an instance in a series.</span></span> <span data-ttu-id="fa854-563">Значение `null` возвращается для отдельных встреч и приглашений на собрания, связанных с одной встречей.</span><span class="sxs-lookup"><span data-stu-id="fa854-563">`null` is returned for single appointments and meeting requests of single appointments.</span></span> <span data-ttu-id="fa854-564">Значение `undefined` возвращается для сообщений, которые не являются приглашениями на собрания.</span><span class="sxs-lookup"><span data-stu-id="fa854-564">`undefined` is returned for messages that are not meeting requests.</span></span>

> <span data-ttu-id="fa854-565">Примечание. Приглашения на собрания имеют значение `itemClass` для класса IPM.Schedule.Meeting.Request.</span><span class="sxs-lookup"><span data-stu-id="fa854-565">Note: Meeting requests have an `itemClass` value of IPM.Schedule.Meeting.Request.</span></span>

> <span data-ttu-id="fa854-566">Примечание. Если объект recurrence имеет значение `null`, он представляет собой отдельную встречу или приглашение на собрание, связанное с одной встречей, и НЕ входит в серию.</span><span class="sxs-lookup"><span data-stu-id="fa854-566">Note: If the recurrence object is `null`, this indicates that the object is a single appointment or a meeting request of a single appointment and NOT a part of a series.</span></span>

##### <a name="type"></a><span data-ttu-id="fa854-567">Тип:</span><span class="sxs-lookup"><span data-stu-id="fa854-567">Type:</span></span>

* [<span data-ttu-id="fa854-568">Recurrence</span><span class="sxs-lookup"><span data-stu-id="fa854-568">Recurrence</span></span>](/javascript/api/outlook/office.recurrence)

|<span data-ttu-id="fa854-569">Требование</span><span class="sxs-lookup"><span data-stu-id="fa854-569">Requirement</span></span>|<span data-ttu-id="fa854-570">Значение</span><span class="sxs-lookup"><span data-stu-id="fa854-570">Value</span></span>|
|---|---|
|[<span data-ttu-id="fa854-571">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="fa854-571">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="fa854-572">1.7</span><span class="sxs-lookup"><span data-stu-id="fa854-572">1.7</span></span>|
|[<span data-ttu-id="fa854-573">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="fa854-573">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="fa854-574">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fa854-574">ReadItem</span></span>|
|[<span data-ttu-id="fa854-575">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="fa854-575">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="fa854-576">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="fa854-576">Compose or read</span></span>|

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="fa854-577">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="fa854-577">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="fa854-578">Предоставляет доступ к обязательным участникам события.</span><span class="sxs-lookup"><span data-stu-id="fa854-578">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="fa854-579">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="fa854-579">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="fa854-580">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="fa854-580">Read mode</span></span>

<span data-ttu-id="fa854-581">Свойство `requiredAttendees` возвращает массив, содержащий объект `EmailAddressDetails` для каждого обязательного участника собрания.</span><span class="sxs-lookup"><span data-stu-id="fa854-581">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="fa854-582">Режим создания</span><span class="sxs-lookup"><span data-stu-id="fa854-582">Compose mode</span></span>

<span data-ttu-id="fa854-583">Свойство `requiredAttendees` возвращает объект `Recipients`, предоставляющий методы, с помощью которых можно получить или обновить сведения об обязательных участниках собрания.</span><span class="sxs-lookup"><span data-stu-id="fa854-583">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="fa854-584">Тип:</span><span class="sxs-lookup"><span data-stu-id="fa854-584">Type:</span></span>

*   <span data-ttu-id="fa854-585">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="fa854-585">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="fa854-586">Требования</span><span class="sxs-lookup"><span data-stu-id="fa854-586">Requirements</span></span>

|<span data-ttu-id="fa854-587">Требование</span><span class="sxs-lookup"><span data-stu-id="fa854-587">Requirement</span></span>|<span data-ttu-id="fa854-588">Значение</span><span class="sxs-lookup"><span data-stu-id="fa854-588">Value</span></span>|
|---|---|
|[<span data-ttu-id="fa854-589">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="fa854-589">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="fa854-590">1.0</span><span class="sxs-lookup"><span data-stu-id="fa854-590">1.0</span></span>|
|[<span data-ttu-id="fa854-591">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="fa854-591">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="fa854-592">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fa854-592">ReadItem</span></span>|
|[<span data-ttu-id="fa854-593">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="fa854-593">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="fa854-594">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="fa854-594">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="fa854-595">Пример</span><span class="sxs-lookup"><span data-stu-id="fa854-595">Example</span></span>

```javascript
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
}
```

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails"></a><span data-ttu-id="fa854-596">sender :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="fa854-596">sender :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)</span></span>

<span data-ttu-id="fa854-p126">Получает электронный адрес отправителя электронного сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="fa854-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="fa854-p127">Свойства [`from`](#from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom) и `sender` представляют одно лицо, если сообщение не отправлено представителем. В противном случае свойство `from` представляет лицо, делегировавшее полномочия, а свойство sender — представителя.</span><span class="sxs-lookup"><span data-stu-id="fa854-p127">The [`from`](#from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="fa854-601">Свойству `recipientType`, принадлежащему объекту `EmailAddressDetails` в свойстве `sender`, задано значение `undefined`.</span><span class="sxs-lookup"><span data-stu-id="fa854-601">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="fa854-602">Тип:</span><span class="sxs-lookup"><span data-stu-id="fa854-602">Type:</span></span>

*   [<span data-ttu-id="fa854-603">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="fa854-603">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="fa854-604">Требования</span><span class="sxs-lookup"><span data-stu-id="fa854-604">Requirements</span></span>

|<span data-ttu-id="fa854-605">Требование</span><span class="sxs-lookup"><span data-stu-id="fa854-605">Requirement</span></span>|<span data-ttu-id="fa854-606">Значение</span><span class="sxs-lookup"><span data-stu-id="fa854-606">Value</span></span>|
|---|---|
|[<span data-ttu-id="fa854-607">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="fa854-607">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="fa854-608">1.0</span><span class="sxs-lookup"><span data-stu-id="fa854-608">1.0</span></span>|
|[<span data-ttu-id="fa854-609">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="fa854-609">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="fa854-610">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fa854-610">ReadItem</span></span>|
|[<span data-ttu-id="fa854-611">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="fa854-611">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="fa854-612">Чтение</span><span class="sxs-lookup"><span data-stu-id="fa854-612">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="fa854-613">Пример</span><span class="sxs-lookup"><span data-stu-id="fa854-613">Example</span></span>

```javascript
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
```

#### <a name="nullable-seriesid-string"></a><span data-ttu-id="fa854-614">(nullable) seriesId :String</span><span class="sxs-lookup"><span data-stu-id="fa854-614">(nullable) seriesId :String</span></span>

<span data-ttu-id="fa854-615">Получает идентификатор серии, к которой относится экземпляр.</span><span class="sxs-lookup"><span data-stu-id="fa854-615">Gets the id of the series that an instance belongs to.</span></span>

<span data-ttu-id="fa854-616">В Outlook Web App и Outlook свойство `seriesId` возвращает идентификатор веб-служб Exchange (EWS) родительского элемента (серии), к которому относится этот элемент.</span><span class="sxs-lookup"><span data-stu-id="fa854-616">In OWA and Outlook, the `seriesId` returns the Exchange Web Services (EWS) ID of the parent (series) item that this item belongs to.</span></span> <span data-ttu-id="fa854-617">Однако в iOS и Android свойство `seriesId` возвращает идентификатор REST родительского элемента.</span><span class="sxs-lookup"><span data-stu-id="fa854-617">However, in iOS and Android, the `seriesId` returns the REST ID of the parent item.</span></span>

> [!NOTE]
> <span data-ttu-id="fa854-618">Идентификатор, возвращаемый свойством `seriesId`, совпадает с идентификатором элемента веб-служб Exchange.</span><span class="sxs-lookup"><span data-stu-id="fa854-618">The identifier returned by the `seriesId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="fa854-619">Свойство `seriesId` не совпадает с идентификаторами Outlook, которые используются в REST API Outlook.</span><span class="sxs-lookup"><span data-stu-id="fa854-619">The `seriesId` property is not identical to the Outlook IDs used by the Outlook REST API.</span></span> <span data-ttu-id="fa854-620">Прежде чем совершать вызовы REST API, используя это значение, его необходимо преобразовать с помощью [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="fa854-620">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="fa854-621">Дополнительные сведения см. в статье [Использование REST API Outlook из надстройки Outlook](https://docs.microsoft.com/outlook/add-ins/use-rest-api).</span><span class="sxs-lookup"><span data-stu-id="fa854-621">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api).</span></span>

<span data-ttu-id="fa854-622">Свойство `seriesId` возвращает значение `null` для элементов, у которых нет родительских элементов, например отдельных встреч, элементов серий или приглашений на собрания, и возвращает значение `undefined` для всех других элементов, которые не представляют собой приглашения на собрания.</span><span class="sxs-lookup"><span data-stu-id="fa854-622">The `seriesId` property returns `null` for items that do not have parent items such as single appointments, series items, or meeting requests and returns `undefined` for any other items that are not meeting requests.</span></span>

##### <a name="type"></a><span data-ttu-id="fa854-623">Тип:</span><span class="sxs-lookup"><span data-stu-id="fa854-623">Type:</span></span>

* <span data-ttu-id="fa854-624">String</span><span class="sxs-lookup"><span data-stu-id="fa854-624">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="fa854-625">Требования</span><span class="sxs-lookup"><span data-stu-id="fa854-625">Requirements</span></span>

|<span data-ttu-id="fa854-626">Требование</span><span class="sxs-lookup"><span data-stu-id="fa854-626">Requirement</span></span>|<span data-ttu-id="fa854-627">Значение</span><span class="sxs-lookup"><span data-stu-id="fa854-627">Value</span></span>|
|---|---|
|[<span data-ttu-id="fa854-628">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="fa854-628">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="fa854-629">1.7</span><span class="sxs-lookup"><span data-stu-id="fa854-629">1.7</span></span>|
|[<span data-ttu-id="fa854-630">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="fa854-630">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="fa854-631">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fa854-631">ReadItem</span></span>|
|[<span data-ttu-id="fa854-632">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="fa854-632">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="fa854-633">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="fa854-633">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="fa854-634">Пример</span><span class="sxs-lookup"><span data-stu-id="fa854-634">Example</span></span>

```javascript
var seriesId = Office.context.mailbox.item.seriesId;
var isSeries = (seriesId == null);
```

####  <a name="start-datetimejavascriptapioutlookofficetime"></a><span data-ttu-id="fa854-635">start :Date|[Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="fa854-635">start :Date|[Time](/javascript/api/outlook/office.time)</span></span>

<span data-ttu-id="fa854-636">Получает или задает дату и время начала встречи.</span><span class="sxs-lookup"><span data-stu-id="fa854-636">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="fa854-p130">Свойство `start` представлено в виде значения даты и времени в формате UTC. Это значение можно преобразовать в местные значения даты и времени клиента с помощью метода [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime).</span><span class="sxs-lookup"><span data-stu-id="fa854-p130">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="fa854-639">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="fa854-639">Read mode</span></span>

<span data-ttu-id="fa854-640">Свойство `start` возвращает объект `Date`.</span><span class="sxs-lookup"><span data-stu-id="fa854-640">The `start` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="fa854-641">Режим создания</span><span class="sxs-lookup"><span data-stu-id="fa854-641">Compose mode</span></span>

<span data-ttu-id="fa854-642">Свойство `start` возвращает объект `Time`.</span><span class="sxs-lookup"><span data-stu-id="fa854-642">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="fa854-643">Если вы задаете время начала с помощью метода [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-), необходимо использовать метод [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) для преобразования местного времени на клиенте в формат UTC для сервера.</span><span class="sxs-lookup"><span data-stu-id="fa854-643">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="fa854-644">Тип:</span><span class="sxs-lookup"><span data-stu-id="fa854-644">Type:</span></span>

*   <span data-ttu-id="fa854-645">Date | [Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="fa854-645">Date | [Time](/javascript/api/outlook/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="fa854-646">Требования</span><span class="sxs-lookup"><span data-stu-id="fa854-646">Requirements</span></span>

|<span data-ttu-id="fa854-647">Требование</span><span class="sxs-lookup"><span data-stu-id="fa854-647">Requirement</span></span>|<span data-ttu-id="fa854-648">Значение</span><span class="sxs-lookup"><span data-stu-id="fa854-648">Value</span></span>|
|---|---|
|[<span data-ttu-id="fa854-649">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="fa854-649">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="fa854-650">1.0</span><span class="sxs-lookup"><span data-stu-id="fa854-650">1.0</span></span>|
|[<span data-ttu-id="fa854-651">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="fa854-651">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="fa854-652">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fa854-652">ReadItem</span></span>|
|[<span data-ttu-id="fa854-653">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="fa854-653">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="fa854-654">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="fa854-654">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="fa854-655">Пример</span><span class="sxs-lookup"><span data-stu-id="fa854-655">Example</span></span>

<span data-ttu-id="fa854-656">В примере ниже с помощью метода [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) объекта `Time` задается время начала встречи в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="fa854-656">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

```javascript
var startTime = new Date("3/14/2015");
var options = {
  // Pass information that can be used
  // in the callback
  asyncContext: {verb:"Set"}
}
Office.context.mailbox.item.start.setAsync(startTime, options, function(result) {
  if (result.error) {
    console.debug(result.error);
  } else {
    // Access the asyncContext that was passed to the setAsync function
    console.debug("Start Time " + result.asyncContext.verb);
  }
});
```

####  <a name="subject-stringsubjectjavascriptapioutlookofficesubject"></a><span data-ttu-id="fa854-657">subject :String|[Subject](/javascript/api/outlook/office.subject)</span><span class="sxs-lookup"><span data-stu-id="fa854-657">subject :String|[Subject](/javascript/api/outlook/office.subject)</span></span>

<span data-ttu-id="fa854-658">Получает или задает описание, которое отображается в поле темы элемента.</span><span class="sxs-lookup"><span data-stu-id="fa854-658">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="fa854-659">Свойство `subject` получает или задает всю тему элемента для отправки с почтового сервера.</span><span class="sxs-lookup"><span data-stu-id="fa854-659">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="fa854-660">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="fa854-660">Read mode</span></span>

<span data-ttu-id="fa854-p131">Свойство `subject` возвращает строку. С помощью свойства [`normalizedSubject`](#normalizedsubject-string) можно получить тему без начальных префиксов, таких как `RE:` и `FW:`.</span><span class="sxs-lookup"><span data-stu-id="fa854-p131">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```javascript
var subject = Office.context.mailbox.item.subject;
```

##### <a name="compose-mode"></a><span data-ttu-id="fa854-663">Режим создания</span><span class="sxs-lookup"><span data-stu-id="fa854-663">Compose mode</span></span>

<span data-ttu-id="fa854-664">Свойство `subject` возвращает объект `Subject`, который предоставляет методы для получения и задания темы.</span><span class="sxs-lookup"><span data-stu-id="fa854-664">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```javascript
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="fa854-665">Тип:</span><span class="sxs-lookup"><span data-stu-id="fa854-665">Type:</span></span>

*   <span data-ttu-id="fa854-666">String | [Subject](/javascript/api/outlook/office.subject)</span><span class="sxs-lookup"><span data-stu-id="fa854-666">String | [Subject](/javascript/api/outlook/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="fa854-667">Требования</span><span class="sxs-lookup"><span data-stu-id="fa854-667">Requirements</span></span>

|<span data-ttu-id="fa854-668">Требование</span><span class="sxs-lookup"><span data-stu-id="fa854-668">Requirement</span></span>|<span data-ttu-id="fa854-669">Значение</span><span class="sxs-lookup"><span data-stu-id="fa854-669">Value</span></span>|
|---|---|
|[<span data-ttu-id="fa854-670">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="fa854-670">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="fa854-671">1.0</span><span class="sxs-lookup"><span data-stu-id="fa854-671">1.0</span></span>|
|[<span data-ttu-id="fa854-672">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="fa854-672">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="fa854-673">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fa854-673">ReadItem</span></span>|
|[<span data-ttu-id="fa854-674">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="fa854-674">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="fa854-675">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="fa854-675">Compose or read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="fa854-676">to :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="fa854-676">to :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="fa854-677">Предоставляет доступ к получателям, указанным в строке **Кому** сообщения.</span><span class="sxs-lookup"><span data-stu-id="fa854-677">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="fa854-678">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="fa854-678">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="fa854-679">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="fa854-679">Read mode</span></span>

<span data-ttu-id="fa854-p133">Свойство `to` возвращает массив, содержащий объект `EmailAddressDetails` для каждого получателя в строке **Кому** сообщения. Коллекция может включать не более 100 элементов.</span><span class="sxs-lookup"><span data-stu-id="fa854-p133">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="fa854-682">Режим создания</span><span class="sxs-lookup"><span data-stu-id="fa854-682">Compose mode</span></span>

<span data-ttu-id="fa854-683">Свойство `to` возвращает объект `Recipients`, предоставляющий методы для получения или обновления получателей, которые указаны в строке **Кому** сообщения.</span><span class="sxs-lookup"><span data-stu-id="fa854-683">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="fa854-684">Тип:</span><span class="sxs-lookup"><span data-stu-id="fa854-684">Type:</span></span>

*   <span data-ttu-id="fa854-685">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="fa854-685">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="fa854-686">Требования</span><span class="sxs-lookup"><span data-stu-id="fa854-686">Requirements</span></span>

|<span data-ttu-id="fa854-687">Требование</span><span class="sxs-lookup"><span data-stu-id="fa854-687">Requirement</span></span>|<span data-ttu-id="fa854-688">Значение</span><span class="sxs-lookup"><span data-stu-id="fa854-688">Value</span></span>|
|---|---|
|[<span data-ttu-id="fa854-689">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="fa854-689">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="fa854-690">1.0</span><span class="sxs-lookup"><span data-stu-id="fa854-690">1.0</span></span>|
|[<span data-ttu-id="fa854-691">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="fa854-691">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="fa854-692">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fa854-692">ReadItem</span></span>|
|[<span data-ttu-id="fa854-693">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="fa854-693">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="fa854-694">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="fa854-694">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="fa854-695">Пример</span><span class="sxs-lookup"><span data-stu-id="fa854-695">Example</span></span>

```javascript
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

### <a name="methods"></a><span data-ttu-id="fa854-696">Методы</span><span class="sxs-lookup"><span data-stu-id="fa854-696">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="fa854-697">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="fa854-697">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="fa854-698">Добавляет файл в сообщение или встречу в качестве вложения.</span><span class="sxs-lookup"><span data-stu-id="fa854-698">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="fa854-699">Метод `addFileAttachmentAsync` передает файл по указанному универсальному коду ресурса (URI) и вкладывает его в элемент в форме создания.</span><span class="sxs-lookup"><span data-stu-id="fa854-699">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="fa854-700">Идентификатор можно последовательно использовать с методом [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback), чтобы удалить вложение, добавленное во время текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="fa854-700">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="fa854-701">Параметры</span><span class="sxs-lookup"><span data-stu-id="fa854-701">Parameters:</span></span>
|<span data-ttu-id="fa854-702">Имя</span><span class="sxs-lookup"><span data-stu-id="fa854-702">Name</span></span>|<span data-ttu-id="fa854-703">Тип</span><span class="sxs-lookup"><span data-stu-id="fa854-703">Type</span></span>|<span data-ttu-id="fa854-704">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="fa854-704">Attributes</span></span>|<span data-ttu-id="fa854-705">Описание</span><span class="sxs-lookup"><span data-stu-id="fa854-705">Description</span></span>|
|---|---|---|---|
|`uri`|<span data-ttu-id="fa854-706">String</span><span class="sxs-lookup"><span data-stu-id="fa854-706">String</span></span>||<span data-ttu-id="fa854-p134">Универсальный код ресурса (URI), представляющий расположение файла, который нужно вложить в сообщение или встречу. Максимальная длина — 2048 символов.</span><span class="sxs-lookup"><span data-stu-id="fa854-p134">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="fa854-709">String</span><span class="sxs-lookup"><span data-stu-id="fa854-709">String</span></span>||<span data-ttu-id="fa854-p135">Имя вложения, которое отображается при передаче вложения. Максимальная длина — 255 символов.</span><span class="sxs-lookup"><span data-stu-id="fa854-p135">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="fa854-712">Object</span><span class="sxs-lookup"><span data-stu-id="fa854-712">Object</span></span>|<span data-ttu-id="fa854-713">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="fa854-713">&lt;optional&gt;</span></span>|<span data-ttu-id="fa854-714">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="fa854-714">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="fa854-715">Object</span><span class="sxs-lookup"><span data-stu-id="fa854-715">Object</span></span>|<span data-ttu-id="fa854-716">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="fa854-716">&lt;optional&gt;</span></span>|<span data-ttu-id="fa854-717">В методе обратного вызова разработчики могут указать любой объект, к которому необходимо получить доступ.</span><span class="sxs-lookup"><span data-stu-id="fa854-717">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="fa854-718">Boolean</span><span class="sxs-lookup"><span data-stu-id="fa854-718">Boolean</span></span>|<span data-ttu-id="fa854-719">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="fa854-719">&lt;optional&gt;</span></span>|<span data-ttu-id="fa854-720">Значение `true` указывает на то, что вложение будет встроено в текст сообщения и не должно отображаться в списке вложений.</span><span class="sxs-lookup"><span data-stu-id="fa854-720">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="fa854-721">function</span><span class="sxs-lookup"><span data-stu-id="fa854-721">function</span></span>|<span data-ttu-id="fa854-722">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="fa854-722">&lt;optional&gt;</span></span>|<span data-ttu-id="fa854-723">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="fa854-723">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="fa854-724">После успешного выполнения идентификатор вложения будет представлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="fa854-724">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="fa854-725">Если добавить вложение не удастся, объект `asyncResult` будет содержать объект `Error` с описанием ошибки.</span><span class="sxs-lookup"><span data-stu-id="fa854-725">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="fa854-726">Ошибки</span><span class="sxs-lookup"><span data-stu-id="fa854-726">Errors</span></span>

|<span data-ttu-id="fa854-727">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="fa854-727">Error code</span></span>|<span data-ttu-id="fa854-728">Описание</span><span class="sxs-lookup"><span data-stu-id="fa854-728">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="fa854-729">Вложение превышает максимальный размер.</span><span class="sxs-lookup"><span data-stu-id="fa854-729">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="fa854-730">Расширение вложения не поддерживается.</span><span class="sxs-lookup"><span data-stu-id="fa854-730">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="fa854-731">Сообщение или встреча содержат слишком много вложений.</span><span class="sxs-lookup"><span data-stu-id="fa854-731">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="fa854-732">Требования</span><span class="sxs-lookup"><span data-stu-id="fa854-732">Requirements</span></span>

|<span data-ttu-id="fa854-733">Требование</span><span class="sxs-lookup"><span data-stu-id="fa854-733">Requirement</span></span>|<span data-ttu-id="fa854-734">Значение</span><span class="sxs-lookup"><span data-stu-id="fa854-734">Value</span></span>|
|---|---|
|[<span data-ttu-id="fa854-735">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="fa854-735">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="fa854-736">1.1</span><span class="sxs-lookup"><span data-stu-id="fa854-736">1.1</span></span>|
|[<span data-ttu-id="fa854-737">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="fa854-737">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="fa854-738">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="fa854-738">ReadWriteItem</span></span>|
|[<span data-ttu-id="fa854-739">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="fa854-739">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="fa854-740">Создание</span><span class="sxs-lookup"><span data-stu-id="fa854-740">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="fa854-741">Примеры</span><span class="sxs-lookup"><span data-stu-id="fa854-741">Examples</span></span>

```js
function callback(result) {
  if (result.error) {
    showMessage(result.error);
  } else {
    showMessage("Attachment added");
  }
}

function addAttachment() {
  // The values in asyncContext can be accessed in the callback
  var options = { 'asyncContext': { var1: 1, var2: 2 } };

  var attachmentURL = "https://contoso.com/rtm/icon.png";
  Office.context.mailbox.item.addFileAttachmentAsync(attachmentURL, attachmentURL, options, callback);
}
```

<span data-ttu-id="fa854-742">В приведенном ниже примере файл изображения добавляется в качестве встроенного вложения, а ссылка на вложение добавляется в текст сообщения.</span><span class="sxs-lookup"><span data-stu-id="fa854-742">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

```js
Office.context.mailbox.item.addFileAttachmentAsync
(
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
        
      }
    );
  }
);
```

#### <a name="addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback"></a><span data-ttu-id="fa854-743">addFileAttachmentFromBase64Async(base64File, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="fa854-743">addFileAttachmentFromBase64Async(base64File, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="fa854-744">Добавляет файл из кодирования base64 в сообщение или встречу в качестве вложения.</span><span class="sxs-lookup"><span data-stu-id="fa854-744">Adds a file from the base64 encoding to a message or appointment as an attachment.</span></span>

<span data-ttu-id="fa854-745">Метод `addFileAttachmentFromBase64Async` передает файл из кодировки base64 и вкладывает его в элемент в форме создания.</span><span class="sxs-lookup"><span data-stu-id="fa854-745">The `addFileAttachmentFromBase64Async` method uploads the file from the base64 encoding and attaches it to the item in the compose form.</span></span> <span data-ttu-id="fa854-746">Этот способ возвращает идентификатор вложения в объекте AsyncResult.value.</span><span class="sxs-lookup"><span data-stu-id="fa854-746">This method returns the attachment identifier in the AsyncResult.value object.</span></span>

<span data-ttu-id="fa854-747">Идентификатор можно использовать с методом [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback), чтобы удалить вложение, добавленное во время текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="fa854-747">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="fa854-748">Параметры</span><span class="sxs-lookup"><span data-stu-id="fa854-748">Parameters:</span></span>
|<span data-ttu-id="fa854-749">Имя</span><span class="sxs-lookup"><span data-stu-id="fa854-749">Name</span></span>|<span data-ttu-id="fa854-750">Тип</span><span class="sxs-lookup"><span data-stu-id="fa854-750">Type</span></span>|<span data-ttu-id="fa854-751">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="fa854-751">Attributes</span></span>|<span data-ttu-id="fa854-752">Описание</span><span class="sxs-lookup"><span data-stu-id="fa854-752">Description</span></span>|
|---|---|---|---|
|`base64File`|<span data-ttu-id="fa854-753">Строка</span><span class="sxs-lookup"><span data-stu-id="fa854-753">String</span></span>||<span data-ttu-id="fa854-754">Закодированное содержимое base64 изображения или файла, которое следует добавить в сообщение электронной почты или событие.</span><span class="sxs-lookup"><span data-stu-id="fa854-754">The base64 encoded content of an image or file to be added to an email or event.</span></span>|
|`attachmentName`|<span data-ttu-id="fa854-755">Строка</span><span class="sxs-lookup"><span data-stu-id="fa854-755">String</span></span>||<span data-ttu-id="fa854-p137">Имя вложения, которое отображается при передаче вложения. Максимальная длина — 255 символов.</span><span class="sxs-lookup"><span data-stu-id="fa854-p137">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="fa854-758">Object</span><span class="sxs-lookup"><span data-stu-id="fa854-758">Object</span></span>|<span data-ttu-id="fa854-759">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="fa854-759">&lt;optional&gt;</span></span>|<span data-ttu-id="fa854-760">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="fa854-760">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="fa854-761">Object</span><span class="sxs-lookup"><span data-stu-id="fa854-761">Object</span></span>|<span data-ttu-id="fa854-762">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="fa854-762">&lt;optional&gt;</span></span>|<span data-ttu-id="fa854-763">В методе обратного вызова разработчики могут указать любой объект, к которому необходимо получить доступ.</span><span class="sxs-lookup"><span data-stu-id="fa854-763">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="fa854-764">Boolean</span><span class="sxs-lookup"><span data-stu-id="fa854-764">Boolean</span></span>|<span data-ttu-id="fa854-765">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="fa854-765">&lt;optional&gt;</span></span>|<span data-ttu-id="fa854-766">Значение `true` указывает на то, что вложение будет встроено в текст сообщения и не должно отображаться в списке вложений.</span><span class="sxs-lookup"><span data-stu-id="fa854-766">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="fa854-767">function</span><span class="sxs-lookup"><span data-stu-id="fa854-767">function</span></span>|<span data-ttu-id="fa854-768">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="fa854-768">&lt;optional&gt;</span></span>|<span data-ttu-id="fa854-769">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="fa854-769">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="fa854-770">После успешного выполнения идентификатор вложения будет представлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="fa854-770">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="fa854-771">Если добавить вложение не удастся, объект `asyncResult` будет содержать объект `Error` с описанием ошибки.</span><span class="sxs-lookup"><span data-stu-id="fa854-771">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="fa854-772">Ошибки</span><span class="sxs-lookup"><span data-stu-id="fa854-772">Errors</span></span>

|<span data-ttu-id="fa854-773">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="fa854-773">Error code</span></span>|<span data-ttu-id="fa854-774">Описание</span><span class="sxs-lookup"><span data-stu-id="fa854-774">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="fa854-775">Вложение превышает максимальный размер.</span><span class="sxs-lookup"><span data-stu-id="fa854-775">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="fa854-776">Расширение вложения не поддерживается.</span><span class="sxs-lookup"><span data-stu-id="fa854-776">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="fa854-777">Сообщение или встреча содержат слишком много вложений.</span><span class="sxs-lookup"><span data-stu-id="fa854-777">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="fa854-778">Требования</span><span class="sxs-lookup"><span data-stu-id="fa854-778">Requirements</span></span>

|<span data-ttu-id="fa854-779">Требование</span><span class="sxs-lookup"><span data-stu-id="fa854-779">Requirement</span></span>|<span data-ttu-id="fa854-780">Значение</span><span class="sxs-lookup"><span data-stu-id="fa854-780">Value</span></span>|
|---|---|
|[<span data-ttu-id="fa854-781">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="fa854-781">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="fa854-782">Предварительная версия</span><span class="sxs-lookup"><span data-stu-id="fa854-782">Preview</span></span>|
|[<span data-ttu-id="fa854-783">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="fa854-783">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="fa854-784">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="fa854-784">ReadWriteItem</span></span>|
|[<span data-ttu-id="fa854-785">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="fa854-785">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="fa854-786">Создание</span><span class="sxs-lookup"><span data-stu-id="fa854-786">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="fa854-787">Примеры</span><span class="sxs-lookup"><span data-stu-id="fa854-787">Examples</span></span>

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
      }
    );
  }
);
```

####  <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="fa854-788">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="fa854-788">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="fa854-789">Добавляет обработчик для поддерживаемого события.</span><span class="sxs-lookup"><span data-stu-id="fa854-789">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="fa854-790">Сейчас поддерживаются следующие типы событий: `Office.EventType.AttachmentsChanged`, `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged` и `Office.EventType.RecurrenceChanged`.</span><span class="sxs-lookup"><span data-stu-id="fa854-790">Currently the supported event types are `Office.EventType.AttachmentsChanged`, `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="fa854-791">Параметры:</span><span class="sxs-lookup"><span data-stu-id="fa854-791">Parameters:</span></span>

| <span data-ttu-id="fa854-792">Имя</span><span class="sxs-lookup"><span data-stu-id="fa854-792">Name</span></span> | <span data-ttu-id="fa854-793">Тип</span><span class="sxs-lookup"><span data-stu-id="fa854-793">Type</span></span> | <span data-ttu-id="fa854-794">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="fa854-794">Attributes</span></span> | <span data-ttu-id="fa854-795">Описание</span><span class="sxs-lookup"><span data-stu-id="fa854-795">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="fa854-796">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="fa854-796">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="fa854-797">Событие, которое должно вызвать обработчик.</span><span class="sxs-lookup"><span data-stu-id="fa854-797">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="fa854-798">Function</span><span class="sxs-lookup"><span data-stu-id="fa854-798">Function</span></span> || <span data-ttu-id="fa854-p138">Функция для обработки события. Функция должна принимать один параметр, представляющий собой объектный литерал. Значение свойства `type` параметра совпадет со значением параметра `eventType`, переданного методу `addHandlerAsync`.</span><span class="sxs-lookup"><span data-stu-id="fa854-p138">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="fa854-802">Объект</span><span class="sxs-lookup"><span data-stu-id="fa854-802">Object</span></span> | <span data-ttu-id="fa854-803">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="fa854-803">&lt;optional&gt;</span></span> | <span data-ttu-id="fa854-804">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="fa854-804">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="fa854-805">Object</span><span class="sxs-lookup"><span data-stu-id="fa854-805">Object</span></span> | <span data-ttu-id="fa854-806">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="fa854-806">&lt;optional&gt;</span></span> | <span data-ttu-id="fa854-807">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="fa854-807">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="fa854-808">функция</span><span class="sxs-lookup"><span data-stu-id="fa854-808">function</span></span>| <span data-ttu-id="fa854-809">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="fa854-809">&lt;optional&gt;</span></span>|<span data-ttu-id="fa854-810">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="fa854-810">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="fa854-811">Требования</span><span class="sxs-lookup"><span data-stu-id="fa854-811">Requirements</span></span>

|<span data-ttu-id="fa854-812">Требование</span><span class="sxs-lookup"><span data-stu-id="fa854-812">Requirement</span></span>| <span data-ttu-id="fa854-813">Значение</span><span class="sxs-lookup"><span data-stu-id="fa854-813">Value</span></span>|
|---|---|
|[<span data-ttu-id="fa854-814">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="fa854-814">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fa854-815">1.7</span><span class="sxs-lookup"><span data-stu-id="fa854-815">1.7</span></span> |
|[<span data-ttu-id="fa854-816">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="fa854-816">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fa854-817">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fa854-817">ReadItem</span></span> |
|[<span data-ttu-id="fa854-818">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="fa854-818">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fa854-819">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="fa854-819">Compose or read</span></span> |

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="fa854-820">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="fa854-820">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="fa854-821">Добавляет к сообщению элемент Exchange, например сообщение, в виде вложения.</span><span class="sxs-lookup"><span data-stu-id="fa854-821">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="fa854-p139">С помощью метода `addItemAttachmentAsync` можно в элемент формы создания вложить элемент с указанным идентификатором Exchange. Если указать метод обратного вызова, то этот метод вызывается с помощью параметра `asyncResult`, который содержит идентификатор вложения или код, указывающий на ошибки, которые произошли при вложении элемента. При необходимости можно использовать параметр `options` для передачи сведений о состоянии методу обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="fa854-p139">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="fa854-825">Идентификатор можно использовать с методом [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback), чтобы удалить вложение, добавленное во время текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="fa854-825">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="fa854-826">Если ваша надстройка Office выполняется в Outlook Web App, метод `addItemAttachmentAsync` обеспечивает вложение элементов в элементы, отличные от редактируемого. Однако это действие не рекомендуем выполнять, так как оно не поддерживается.</span><span class="sxs-lookup"><span data-stu-id="fa854-826">If your Office Add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="fa854-827">Параметры:</span><span class="sxs-lookup"><span data-stu-id="fa854-827">Parameters:</span></span>

|<span data-ttu-id="fa854-828">Имя</span><span class="sxs-lookup"><span data-stu-id="fa854-828">Name</span></span>|<span data-ttu-id="fa854-829">Тип</span><span class="sxs-lookup"><span data-stu-id="fa854-829">Type</span></span>|<span data-ttu-id="fa854-830">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="fa854-830">Attributes</span></span>|<span data-ttu-id="fa854-831">Описание</span><span class="sxs-lookup"><span data-stu-id="fa854-831">Description</span></span>|
|---|---|---|---|
|`itemId`|<span data-ttu-id="fa854-832">String</span><span class="sxs-lookup"><span data-stu-id="fa854-832">String</span></span>||<span data-ttu-id="fa854-p140">Идентификатор Exchange для вкладываемого элемента. Максимальная длина — 100 символов.</span><span class="sxs-lookup"><span data-stu-id="fa854-p140">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="fa854-835">String</span><span class="sxs-lookup"><span data-stu-id="fa854-835">String</span></span>||<span data-ttu-id="fa854-p141">Тема вкладываемого элемента. Максимальная длина — 255 символов.</span><span class="sxs-lookup"><span data-stu-id="fa854-p141">The sujbect of the item to be attached. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="fa854-838">Object</span><span class="sxs-lookup"><span data-stu-id="fa854-838">Object</span></span>|<span data-ttu-id="fa854-839">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="fa854-839">&lt;optional&gt;</span></span>|<span data-ttu-id="fa854-840">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="fa854-840">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="fa854-841">Object</span><span class="sxs-lookup"><span data-stu-id="fa854-841">Object</span></span>|<span data-ttu-id="fa854-842">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="fa854-842">&lt;optional&gt;</span></span>|<span data-ttu-id="fa854-843">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="fa854-843">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="fa854-844">функция</span><span class="sxs-lookup"><span data-stu-id="fa854-844">function</span></span>|<span data-ttu-id="fa854-845">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="fa854-845">&lt;optional&gt;</span></span>|<span data-ttu-id="fa854-846">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="fa854-846">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="fa854-847">После успешного выполнения идентификатор вложения будет представлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="fa854-847">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="fa854-848">Если добавить вложение не удастся, объект `asyncResult` будет содержать объект `Error` с описанием ошибки.</span><span class="sxs-lookup"><span data-stu-id="fa854-848">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="fa854-849">Ошибки</span><span class="sxs-lookup"><span data-stu-id="fa854-849">Errors</span></span>

|<span data-ttu-id="fa854-850">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="fa854-850">Error code</span></span>|<span data-ttu-id="fa854-851">Описание</span><span class="sxs-lookup"><span data-stu-id="fa854-851">Description</span></span>|
|------------|-------------|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="fa854-852">Сообщение или встреча содержат слишком много вложений.</span><span class="sxs-lookup"><span data-stu-id="fa854-852">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="fa854-853">Требования</span><span class="sxs-lookup"><span data-stu-id="fa854-853">Requirements</span></span>

|<span data-ttu-id="fa854-854">Требование</span><span class="sxs-lookup"><span data-stu-id="fa854-854">Requirement</span></span>|<span data-ttu-id="fa854-855">Значение</span><span class="sxs-lookup"><span data-stu-id="fa854-855">Value</span></span>|
|---|---|
|[<span data-ttu-id="fa854-856">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="fa854-856">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="fa854-857">1.1</span><span class="sxs-lookup"><span data-stu-id="fa854-857">1.1</span></span>|
|[<span data-ttu-id="fa854-858">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="fa854-858">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="fa854-859">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="fa854-859">ReadWriteItem</span></span>|
|[<span data-ttu-id="fa854-860">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="fa854-860">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="fa854-861">Создание</span><span class="sxs-lookup"><span data-stu-id="fa854-861">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="fa854-862">Пример</span><span class="sxs-lookup"><span data-stu-id="fa854-862">Example</span></span>

<span data-ttu-id="fa854-863">В следующем примере существующий элемент Outlook добавляется в виде вложения с именем `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="fa854-863">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

```javascript
function callback(result) {
  if (result.error) {
    showMessage(result.error);
  } else {
    showMessage("Attachment added");
  }
}

function addAttachment() {
  // EWS ID of item to attach
  // (Shortened for readability)
  var itemId = "AAMkADI1...AAA=";

  // The values in asyncContext can be accessed in the callback
  var options = { 'asyncContext': { var1: 1, var2: 2 } };

  Office.context.mailbox.item.addItemAttachmentAsync(itemId, "My Attachment", options, callback);
}
```

####  <a name="close"></a><span data-ttu-id="fa854-864">close()</span><span class="sxs-lookup"><span data-stu-id="fa854-864">close()</span></span>

<span data-ttu-id="fa854-865">Закрывает текущий создаваемый элемент.</span><span class="sxs-lookup"><span data-stu-id="fa854-865">Closes the current item that is being composed.</span></span>

<span data-ttu-id="fa854-p142">Работа метода `close` зависит от текущего состояния создаваемого элемента. Если элемент содержит несохраненные изменения, клиент предложит пользователю сохранить или отклонить их либо отменить действие закрытия.</span><span class="sxs-lookup"><span data-stu-id="fa854-p142">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="fa854-868">Если в Outlook в Интернете элемент представляет собой встречу, ранее сохраненную с помощью метода `saveAsync`, пользователю предлагается сохранить, отклонить или отменить действие, даже если с момента последнего сохранения элемента не вносились какие-либо изменения.</span><span class="sxs-lookup"><span data-stu-id="fa854-868">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="fa854-869">Если в клиенте Outlook для настольных ПК сообщение представляет собой ответ в тексте, метод `close` не работает.</span><span class="sxs-lookup"><span data-stu-id="fa854-869">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="fa854-870">Требования</span><span class="sxs-lookup"><span data-stu-id="fa854-870">Requirements</span></span>

|<span data-ttu-id="fa854-871">Требование</span><span class="sxs-lookup"><span data-stu-id="fa854-871">Requirement</span></span>|<span data-ttu-id="fa854-872">Значение</span><span class="sxs-lookup"><span data-stu-id="fa854-872">Value</span></span>|
|---|---|
|[<span data-ttu-id="fa854-873">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="fa854-873">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="fa854-874">1.3</span><span class="sxs-lookup"><span data-stu-id="fa854-874">1.3</span></span>|
|[<span data-ttu-id="fa854-875">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="fa854-875">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="fa854-876">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="fa854-876">Restricted</span></span>|
|[<span data-ttu-id="fa854-877">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="fa854-877">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="fa854-878">Создание</span><span class="sxs-lookup"><span data-stu-id="fa854-878">Compose</span></span>|

#### <a name="displayreplyallformformdata"></a><span data-ttu-id="fa854-879">displayReplyAllForm(formData)</span><span class="sxs-lookup"><span data-stu-id="fa854-879">displayReplyAllForm(formData)</span></span>

<span data-ttu-id="fa854-880">Отображает форму ответа, включающую отправителя и всех получателей выбранного сообщения или организатора и всех участников выбранной встречи.</span><span class="sxs-lookup"><span data-stu-id="fa854-880">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="fa854-881">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="fa854-881">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="fa854-882">В Outlook Web App форма ответа отображается в виде всплывающей формы в представлении с 3 либо 1 или 2 колонками.</span><span class="sxs-lookup"><span data-stu-id="fa854-882">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="fa854-883">Если любой строковый параметр превышает указанные для него ограничения, `displayReplyAllForm` возвращает исключение.</span><span class="sxs-lookup"><span data-stu-id="fa854-883">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="fa854-p143">Если в параметре `formData.attachments` указаны вложения, Outlook и Outlook Web App пытаются скачать их и вложить в форму ответа. Если какие-либо вложения добавить не удается, в форме отображается сообщение об ошибке. Если сообщения об ошибках не предусмотрены, то они не отображаются.</span><span class="sxs-lookup"><span data-stu-id="fa854-p143">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="fa854-887">Параметры</span><span class="sxs-lookup"><span data-stu-id="fa854-887">Parameters:</span></span>

|<span data-ttu-id="fa854-888">Имя</span><span class="sxs-lookup"><span data-stu-id="fa854-888">Name</span></span>|<span data-ttu-id="fa854-889">Тип</span><span class="sxs-lookup"><span data-stu-id="fa854-889">Type</span></span>|<span data-ttu-id="fa854-890">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="fa854-890">Attributes</span></span>|<span data-ttu-id="fa854-891">Описание</span><span class="sxs-lookup"><span data-stu-id="fa854-891">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="fa854-892">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="fa854-892">String &#124; Object</span></span>||<span data-ttu-id="fa854-p144">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="fa854-p144">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="fa854-895">**ИЛИ**</span><span class="sxs-lookup"><span data-stu-id="fa854-895">**OR**</span></span><br/><span data-ttu-id="fa854-p145">Объект, который содержит текст или данные вложения и функцию обратного вызова. Ниже представлено определение этого объекта.</span><span class="sxs-lookup"><span data-stu-id="fa854-p145">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="fa854-898">String</span><span class="sxs-lookup"><span data-stu-id="fa854-898">String</span></span>|<span data-ttu-id="fa854-899">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="fa854-899">&lt;optional&gt;</span></span>|<span data-ttu-id="fa854-p146">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="fa854-p146">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="fa854-902">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="fa854-902">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="fa854-903">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="fa854-903">&lt;optional&gt;</span></span>|<span data-ttu-id="fa854-904">Массив объектов JSON, представляющих собой вложенные файлы или элементы.</span><span class="sxs-lookup"><span data-stu-id="fa854-904">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="fa854-905">String</span><span class="sxs-lookup"><span data-stu-id="fa854-905">String</span></span>||<span data-ttu-id="fa854-p147">Указывает тип вложения. Допустимые значения: `file` для вложенного файла и `item` для вложенного элемента.</span><span class="sxs-lookup"><span data-stu-id="fa854-p147">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="fa854-908">String</span><span class="sxs-lookup"><span data-stu-id="fa854-908">String</span></span>||<span data-ttu-id="fa854-909">Строка, содержащая имя вложения, длиной до 255 символов.</span><span class="sxs-lookup"><span data-stu-id="fa854-909">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="fa854-910">String</span><span class="sxs-lookup"><span data-stu-id="fa854-910">String</span></span>||<span data-ttu-id="fa854-p148">Используется, только если свойству `type` задано значение `file`. URI расположения файла.</span><span class="sxs-lookup"><span data-stu-id="fa854-p148">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="fa854-913">Boolean</span><span class="sxs-lookup"><span data-stu-id="fa854-913">Boolean</span></span>||<span data-ttu-id="fa854-p149">Используется, только если свойству `type` задано значение `file`. Значение `true` указывает на то, что вложение будет встроено в текст сообщения и не должно отображаться в списке вложений.</span><span class="sxs-lookup"><span data-stu-id="fa854-p149">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="fa854-916">String</span><span class="sxs-lookup"><span data-stu-id="fa854-916">String</span></span>||<span data-ttu-id="fa854-p150">Используется, только если свойству `type` задано значение `item`. Идентификатор вложения EWS. Это строка длиной до 100 символов.</span><span class="sxs-lookup"><span data-stu-id="fa854-p150">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="fa854-920">function</span><span class="sxs-lookup"><span data-stu-id="fa854-920">function</span></span>|<span data-ttu-id="fa854-921">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="fa854-921">&lt;optional&gt;</span></span>|<span data-ttu-id="fa854-922">По завершении работы метода функция, переданная параметру `callback`, вызывается с помощью одного параметра `asyncResult`, представляющего собой объект [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="fa854-922">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="fa854-923">Требования</span><span class="sxs-lookup"><span data-stu-id="fa854-923">Requirements</span></span>

|<span data-ttu-id="fa854-924">Требование</span><span class="sxs-lookup"><span data-stu-id="fa854-924">Requirement</span></span>|<span data-ttu-id="fa854-925">Значение</span><span class="sxs-lookup"><span data-stu-id="fa854-925">Value</span></span>|
|---|---|
|[<span data-ttu-id="fa854-926">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="fa854-926">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="fa854-927">1.0</span><span class="sxs-lookup"><span data-stu-id="fa854-927">1.0</span></span>|
|[<span data-ttu-id="fa854-928">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="fa854-928">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="fa854-929">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fa854-929">ReadItem</span></span>|
|[<span data-ttu-id="fa854-930">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="fa854-930">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="fa854-931">Чтение</span><span class="sxs-lookup"><span data-stu-id="fa854-931">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="fa854-932">Примеры</span><span class="sxs-lookup"><span data-stu-id="fa854-932">Examples</span></span>

<span data-ttu-id="fa854-933">Приведенный ниже код передает строку в функцию `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="fa854-933">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="fa854-934">Ответ с пустым текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="fa854-934">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="fa854-935">Ответ только с текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="fa854-935">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="fa854-936">Ответ с текстом сообщения и вложенным файлом.</span><span class="sxs-lookup"><span data-stu-id="fa854-936">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="fa854-937">Ответ с текстом сообщения и вложенным элементом.</span><span class="sxs-lookup"><span data-stu-id="fa854-937">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="fa854-938">Ответ с текстом сообщения, вложенным файлом, вложенным элементом и обратным вызовом.</span><span class="sxs-lookup"><span data-stu-id="fa854-938">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata"></a><span data-ttu-id="fa854-939">displayReplyForm(formData)</span><span class="sxs-lookup"><span data-stu-id="fa854-939">displayReplyForm(formData)</span></span>

<span data-ttu-id="fa854-940">Отображает форму ответа, включающую только отправителя выбранного сообщения или организатора выбранной встречи.</span><span class="sxs-lookup"><span data-stu-id="fa854-940">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="fa854-941">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="fa854-941">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="fa854-942">В Outlook Web App форма ответа отображается в виде всплывающей формы в представлении с 3 либо 1 или 2 колонками.</span><span class="sxs-lookup"><span data-stu-id="fa854-942">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="fa854-943">Если любой строковый параметр превышает указанные для него ограничения, `displayReplyForm` возвращает исключение.</span><span class="sxs-lookup"><span data-stu-id="fa854-943">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="fa854-p151">Если в параметре `formData.attachments` указаны вложения, Outlook и Outlook Web App пытаются скачать их и вложить в форму ответа. Если какие-либо вложения добавить не удается, в форме отображается сообщение об ошибке. Если сообщения об ошибках не предусмотрены, то они не отображаются.</span><span class="sxs-lookup"><span data-stu-id="fa854-p151">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="fa854-947">Параметры</span><span class="sxs-lookup"><span data-stu-id="fa854-947">Parameters:</span></span>

|<span data-ttu-id="fa854-948">Имя</span><span class="sxs-lookup"><span data-stu-id="fa854-948">Name</span></span>|<span data-ttu-id="fa854-949">Тип</span><span class="sxs-lookup"><span data-stu-id="fa854-949">Type</span></span>|<span data-ttu-id="fa854-950">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="fa854-950">Attributes</span></span>|<span data-ttu-id="fa854-951">Описание</span><span class="sxs-lookup"><span data-stu-id="fa854-951">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="fa854-952">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="fa854-952">String &#124; Object</span></span>||<span data-ttu-id="fa854-p152">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="fa854-p152">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="fa854-955">**ИЛИ**</span><span class="sxs-lookup"><span data-stu-id="fa854-955">**OR**</span></span><br/><span data-ttu-id="fa854-p153">Объект, который содержит текст или данные вложения и функцию обратного вызова. Ниже представлено определение этого объекта.</span><span class="sxs-lookup"><span data-stu-id="fa854-p153">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="fa854-958">String</span><span class="sxs-lookup"><span data-stu-id="fa854-958">String</span></span>|<span data-ttu-id="fa854-959">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="fa854-959">&lt;optional&gt;</span></span>|<span data-ttu-id="fa854-p154">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="fa854-p154">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="fa854-962">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="fa854-962">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="fa854-963">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="fa854-963">&lt;optional&gt;</span></span>|<span data-ttu-id="fa854-964">Массив объектов JSON, представляющих собой вложенные файлы или элементы.</span><span class="sxs-lookup"><span data-stu-id="fa854-964">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="fa854-965">String</span><span class="sxs-lookup"><span data-stu-id="fa854-965">String</span></span>||<span data-ttu-id="fa854-p155">Указывает тип вложения. Допустимые значения: `file` для вложенного файла и `item` для вложенного элемента.</span><span class="sxs-lookup"><span data-stu-id="fa854-p155">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="fa854-968">String</span><span class="sxs-lookup"><span data-stu-id="fa854-968">String</span></span>||<span data-ttu-id="fa854-969">Строка, содержащая имя вложения, длиной до 255 символов.</span><span class="sxs-lookup"><span data-stu-id="fa854-969">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="fa854-970">String</span><span class="sxs-lookup"><span data-stu-id="fa854-970">String</span></span>||<span data-ttu-id="fa854-p156">Используется, только если свойству `type` задано значение `file`. URI расположения файла.</span><span class="sxs-lookup"><span data-stu-id="fa854-p156">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="fa854-973">Boolean</span><span class="sxs-lookup"><span data-stu-id="fa854-973">Boolean</span></span>||<span data-ttu-id="fa854-p157">Используется, только если свойству `type` задано значение `file`. Значение `true` указывает на то, что вложение будет встроено в текст сообщения и не должно отображаться в списке вложений.</span><span class="sxs-lookup"><span data-stu-id="fa854-p157">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="fa854-976">String</span><span class="sxs-lookup"><span data-stu-id="fa854-976">String</span></span>||<span data-ttu-id="fa854-p158">Используется, только если свойству `type` задано значение `item`. Идентификатор вложения EWS. Это строка длиной до 100 символов.</span><span class="sxs-lookup"><span data-stu-id="fa854-p158">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="fa854-980">function</span><span class="sxs-lookup"><span data-stu-id="fa854-980">function</span></span>|<span data-ttu-id="fa854-981">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="fa854-981">&lt;optional&gt;</span></span>|<span data-ttu-id="fa854-982">По завершении работы метода функция, переданная параметру `callback`, вызывается с помощью одного параметра `asyncResult`, представляющего собой объект [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="fa854-982">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="fa854-983">Требования</span><span class="sxs-lookup"><span data-stu-id="fa854-983">Requirements</span></span>

|<span data-ttu-id="fa854-984">Требование</span><span class="sxs-lookup"><span data-stu-id="fa854-984">Requirement</span></span>|<span data-ttu-id="fa854-985">Значение</span><span class="sxs-lookup"><span data-stu-id="fa854-985">Value</span></span>|
|---|---|
|[<span data-ttu-id="fa854-986">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="fa854-986">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="fa854-987">1.0</span><span class="sxs-lookup"><span data-stu-id="fa854-987">1.0</span></span>|
|[<span data-ttu-id="fa854-988">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="fa854-988">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="fa854-989">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fa854-989">ReadItem</span></span>|
|[<span data-ttu-id="fa854-990">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="fa854-990">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="fa854-991">Чтение</span><span class="sxs-lookup"><span data-stu-id="fa854-991">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="fa854-992">Примеры</span><span class="sxs-lookup"><span data-stu-id="fa854-992">Examples</span></span>

<span data-ttu-id="fa854-993">Приведенный ниже код передает строку в функцию `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="fa854-993">The following code passes a string to the `displayReplyForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="fa854-994">Ответ с пустым текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="fa854-994">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="fa854-995">Ответ только с текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="fa854-995">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="fa854-996">Ответ с текстом сообщения и вложенным файлом.</span><span class="sxs-lookup"><span data-stu-id="fa854-996">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="fa854-997">Ответ с текстом сообщения и вложенным элементом.</span><span class="sxs-lookup"><span data-stu-id="fa854-997">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="fa854-998">Ответ с текстом сообщения, вложенным файлом, вложенным элементом и обратным вызовом.</span><span class="sxs-lookup"><span data-stu-id="fa854-998">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getattachmentcontentasyncattachmentid-options-callback--attachmentcontentjavascriptapioutlookofficeattachmentcontent"></a><span data-ttu-id="fa854-999">getAttachmentContentAsync(attachmentId, [options], callback) → [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span><span class="sxs-lookup"><span data-stu-id="fa854-999">getAttachmentContentAsync(attachmentId, [options], callback) → [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span></span>

<span data-ttu-id="fa854-1000">Получает указанное вложение из сообщения или встречи и возвращает в качестве объекта `AttachmentContent`.</span><span class="sxs-lookup"><span data-stu-id="fa854-1000">Gets the specified attachment from a message or appointment and returns it as an `AttachmentContent` object.</span></span>

<span data-ttu-id="fa854-1001">Метод `getAttachmentContentAsync` получает вложение с указанным идентификатором из элемента.</span><span class="sxs-lookup"><span data-stu-id="fa854-1001">The `getAttachmentContentAsync` method gets the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="fa854-1002">Рекомендуется использовать идентификатор для получения вложения в том же сеансе, в котором были получены идентификаторы вложений attachmentIds посредством вызова `getAttachmentsAsync` или `item.attachments`.</span><span class="sxs-lookup"><span data-stu-id="fa854-1002">As a best practice, you should use the identifier to retrieve an attachment in the same session that the attachmentIds were retrieved with the `getAttachmentsAsync` or `item.attachments` call.</span></span> <span data-ttu-id="fa854-1003">В Outlook Web App и Outlook Web App для устройств идентификатор вложения действителен только в рамках одного сеанса.</span><span class="sxs-lookup"><span data-stu-id="fa854-1003">In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="fa854-1004">Сеанс завершается, когда пользователь закрывает приложение или начинает создавать элемент во встроенной форме, а затем переходит из формы в отдельное окно.</span><span class="sxs-lookup"><span data-stu-id="fa854-1004">A session is over when the user closes the app, or if the user starts composing an inline form then subsequently pops out the form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="fa854-1005">Параметры:</span><span class="sxs-lookup"><span data-stu-id="fa854-1005">Parameters:</span></span>

|<span data-ttu-id="fa854-1006">Имя</span><span class="sxs-lookup"><span data-stu-id="fa854-1006">Name</span></span>|<span data-ttu-id="fa854-1007">Тип</span><span class="sxs-lookup"><span data-stu-id="fa854-1007">Type</span></span>|<span data-ttu-id="fa854-1008">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="fa854-1008">Attributes</span></span>|<span data-ttu-id="fa854-1009">Описание</span><span class="sxs-lookup"><span data-stu-id="fa854-1009">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="fa854-1010">Строка</span><span class="sxs-lookup"><span data-stu-id="fa854-1010">String</span></span>||<span data-ttu-id="fa854-1011">Идентификатор вложения, который необходимо получить.</span><span class="sxs-lookup"><span data-stu-id="fa854-1011">The identifier of the attachment you want to get.</span></span>|
|`options`|<span data-ttu-id="fa854-1012">Object</span><span class="sxs-lookup"><span data-stu-id="fa854-1012">Object</span></span>|<span data-ttu-id="fa854-1013">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="fa854-1013">&lt;optional&gt;</span></span>|<span data-ttu-id="fa854-1014">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="fa854-1014">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="fa854-1015">Object</span><span class="sxs-lookup"><span data-stu-id="fa854-1015">Object</span></span>|<span data-ttu-id="fa854-1016">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="fa854-1016">&lt;optional&gt;</span></span>|<span data-ttu-id="fa854-1017">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="fa854-1017">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="fa854-1018">функция</span><span class="sxs-lookup"><span data-stu-id="fa854-1018">function</span></span>|<span data-ttu-id="fa854-1019">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="fa854-1019">&lt;optional&gt;</span></span>|<span data-ttu-id="fa854-1020">По завершении работы метода функция, переданная параметру `callback`, вызывается с помощью одного параметра `asyncResult`, представляющего собой объект [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="fa854-1020">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="fa854-1021">Требования</span><span class="sxs-lookup"><span data-stu-id="fa854-1021">Requirements</span></span>

|<span data-ttu-id="fa854-1022">Требование</span><span class="sxs-lookup"><span data-stu-id="fa854-1022">Requirement</span></span>|<span data-ttu-id="fa854-1023">Значение</span><span class="sxs-lookup"><span data-stu-id="fa854-1023">Value</span></span>|
|---|---|
|[<span data-ttu-id="fa854-1024">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="fa854-1024">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="fa854-1025">Предварительная версия</span><span class="sxs-lookup"><span data-stu-id="fa854-1025">Preview</span></span>|
|[<span data-ttu-id="fa854-1026">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="fa854-1026">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="fa854-1027">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fa854-1027">ReadItem</span></span>|
|[<span data-ttu-id="fa854-1028">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="fa854-1028">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="fa854-1029">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="fa854-1029">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="fa854-1030">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="fa854-1030">Returns:</span></span>

<span data-ttu-id="fa854-1031">Тип: [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span><span class="sxs-lookup"><span data-stu-id="fa854-1031">Type: [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span></span>

##### <a name="example"></a><span data-ttu-id="fa854-1032">Пример</span><span class="sxs-lookup"><span data-stu-id="fa854-1032">Example</span></span>

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
    // parse string to be a url, an .eml file, a base64-encoded string, or an .icalendar file
    if (result.format == Office.MailboxEnums.AttachmentContentFormat.Base64) {
        // handle file attachment
    }
    else if (result.format == Office.MailboxEnums.AttachmentContentFormat.Eml) {
        // handle item attachment
    }
    else if (result.format == Office.MailboxEnums.AttachmentContentFormat.ICalendar) {
        // handle .icalender attachment
    }
    else {
        // handle cloud attachment  
    }
}
```

#### <a name="getattachmentsasyncoptions-callback--arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails"></a><span data-ttu-id="fa854-1033">getAttachmentsAsync([options], callback) → Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="fa854-1033">getAttachmentsAsync([options], callback) → Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

<span data-ttu-id="fa854-1034">Получает вложения элемента в качестве массива.</span><span class="sxs-lookup"><span data-stu-id="fa854-1034">Gets the item's attachments as an array.</span></span> <span data-ttu-id="fa854-1035">Только в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="fa854-1035">Compose mode only.</span></span>

##### <a name="parameters"></a><span data-ttu-id="fa854-1036">Параметры:</span><span class="sxs-lookup"><span data-stu-id="fa854-1036">Parameters:</span></span>

|<span data-ttu-id="fa854-1037">Имя</span><span class="sxs-lookup"><span data-stu-id="fa854-1037">Name</span></span>|<span data-ttu-id="fa854-1038">Тип</span><span class="sxs-lookup"><span data-stu-id="fa854-1038">Type</span></span>|<span data-ttu-id="fa854-1039">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="fa854-1039">Attributes</span></span>|<span data-ttu-id="fa854-1040">Описание</span><span class="sxs-lookup"><span data-stu-id="fa854-1040">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="fa854-1041">Object</span><span class="sxs-lookup"><span data-stu-id="fa854-1041">Object</span></span>|<span data-ttu-id="fa854-1042">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="fa854-1042">&lt;optional&gt;</span></span>|<span data-ttu-id="fa854-1043">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="fa854-1043">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="fa854-1044">Object</span><span class="sxs-lookup"><span data-stu-id="fa854-1044">Object</span></span>|<span data-ttu-id="fa854-1045">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="fa854-1045">&lt;optional&gt;</span></span>|<span data-ttu-id="fa854-1046">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="fa854-1046">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="fa854-1047">функция</span><span class="sxs-lookup"><span data-stu-id="fa854-1047">function</span></span>|<span data-ttu-id="fa854-1048">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="fa854-1048">&lt;optional&gt;</span></span>|<span data-ttu-id="fa854-1049">По завершении работы метода функция, переданная параметру `callback`, вызывается с помощью одного параметра `asyncResult`, представляющего собой объект [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="fa854-1049">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="fa854-1050">Требования</span><span class="sxs-lookup"><span data-stu-id="fa854-1050">Requirements</span></span>

|<span data-ttu-id="fa854-1051">Требование</span><span class="sxs-lookup"><span data-stu-id="fa854-1051">Requirement</span></span>|<span data-ttu-id="fa854-1052">Значение</span><span class="sxs-lookup"><span data-stu-id="fa854-1052">Value</span></span>|
|---|---|
|[<span data-ttu-id="fa854-1053">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="fa854-1053">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="fa854-1054">Предварительная версия</span><span class="sxs-lookup"><span data-stu-id="fa854-1054">Preview</span></span>|
|[<span data-ttu-id="fa854-1055">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="fa854-1055">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="fa854-1056">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fa854-1056">ReadItem</span></span>|
|[<span data-ttu-id="fa854-1057">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="fa854-1057">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="fa854-1058">Создание</span><span class="sxs-lookup"><span data-stu-id="fa854-1058">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="fa854-1059">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="fa854-1059">Returns:</span></span>

<span data-ttu-id="fa854-1060">Тип: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="fa854-1060">Type: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

##### <a name="example"></a><span data-ttu-id="fa854-1061">Пример</span><span class="sxs-lookup"><span data-stu-id="fa854-1061">Example</span></span>

<span data-ttu-id="fa854-1062">В приведенном ниже примере создается HTML-строка с подробными сведениями обо всех вложениях для текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="fa854-1062">The following example builds an HTML string with details of all attachments on the current item.</span></span>

```javascript
var item = Office.context.mailbox.item;
var outputString = "";
item.getAttachmentsAsync(callback);  
function callback(result) {
    if (result.value.length > 0) {
        for (i = 0 ; i < result.value.length ; i++) {
            var _att = result.value [i];
            outputString += "<BR>" + i + ". Name: ";
            outputString += _att.name;
            outputString += "<BR>ID: " + _att.id;
            outputString += "<BR>contentType: " + _att.contentType;
            outputString += "<BR>size: " + _att.size;
            outputString += "<BR>attachmentType: " + _att.attachmentType;
            outputString += "<BR>isInline: " + _att.isInline;
        }
    }
}
```

#### <a name="getentities--entitiesjavascriptapioutlookofficeentities"></a><span data-ttu-id="fa854-1063">getEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="fa854-1063">getEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span></span>

<span data-ttu-id="fa854-1064">Получает сущности, обнаруженные в теле выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="fa854-1064">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="fa854-1065">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="fa854-1065">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="fa854-1066">Требования</span><span class="sxs-lookup"><span data-stu-id="fa854-1066">Requirements</span></span>

|<span data-ttu-id="fa854-1067">Требование</span><span class="sxs-lookup"><span data-stu-id="fa854-1067">Requirement</span></span>|<span data-ttu-id="fa854-1068">Значение</span><span class="sxs-lookup"><span data-stu-id="fa854-1068">Value</span></span>|
|---|---|
|[<span data-ttu-id="fa854-1069">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="fa854-1069">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="fa854-1070">1.0</span><span class="sxs-lookup"><span data-stu-id="fa854-1070">1.0</span></span>|
|[<span data-ttu-id="fa854-1071">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="fa854-1071">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="fa854-1072">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fa854-1072">ReadItem</span></span>|
|[<span data-ttu-id="fa854-1073">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="fa854-1073">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="fa854-1074">Чтение</span><span class="sxs-lookup"><span data-stu-id="fa854-1074">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="fa854-1075">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="fa854-1075">Returns:</span></span>

<span data-ttu-id="fa854-1076">Тип: [Entities](/javascript/api/outlook/office.entities)</span><span class="sxs-lookup"><span data-stu-id="fa854-1076">Type: [Entities](/javascript/api/outlook/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="fa854-1077">Пример</span><span class="sxs-lookup"><span data-stu-id="fa854-1077">Example</span></span>

<span data-ttu-id="fa854-1078">Ниже приведен пример получения доступа к сущностям контактов в теле текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="fa854-1078">The following example accesses the contacts entities in the current item's body.</span></span>

```javascript
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion"></a><span data-ttu-id="fa854-1079">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="fa854-1079">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span></span>

<span data-ttu-id="fa854-1080">Получает массив всех сущностей указанного типа, обнаруженных в теле выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="fa854-1080">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="fa854-1081">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="fa854-1081">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="fa854-1082">Параметры</span><span class="sxs-lookup"><span data-stu-id="fa854-1082">Parameters:</span></span>

|<span data-ttu-id="fa854-1083">Имя</span><span class="sxs-lookup"><span data-stu-id="fa854-1083">Name</span></span>|<span data-ttu-id="fa854-1084">Тип</span><span class="sxs-lookup"><span data-stu-id="fa854-1084">Type</span></span>|<span data-ttu-id="fa854-1085">Описание</span><span class="sxs-lookup"><span data-stu-id="fa854-1085">Description</span></span>|
|---|---|---|
|`entityType`|[<span data-ttu-id="fa854-1086">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="fa854-1086">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook/office.mailboxenums.entitytype)|<span data-ttu-id="fa854-1087">Одно из значений перечисления EntityType.</span><span class="sxs-lookup"><span data-stu-id="fa854-1087">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="fa854-1088">Требования</span><span class="sxs-lookup"><span data-stu-id="fa854-1088">Requirements</span></span>

|<span data-ttu-id="fa854-1089">Требование</span><span class="sxs-lookup"><span data-stu-id="fa854-1089">Requirement</span></span>|<span data-ttu-id="fa854-1090">Значение</span><span class="sxs-lookup"><span data-stu-id="fa854-1090">Value</span></span>|
|---|---|
|[<span data-ttu-id="fa854-1091">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="fa854-1091">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="fa854-1092">1.0</span><span class="sxs-lookup"><span data-stu-id="fa854-1092">1.0</span></span>|
|[<span data-ttu-id="fa854-1093">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="fa854-1093">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="fa854-1094">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="fa854-1094">Restricted</span></span>|
|[<span data-ttu-id="fa854-1095">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="fa854-1095">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="fa854-1096">Чтение</span><span class="sxs-lookup"><span data-stu-id="fa854-1096">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="fa854-1097">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="fa854-1097">Returns:</span></span>

<span data-ttu-id="fa854-1098">Если значение, переданное в `entityType`, не является допустимым членом перечисления `EntityType`, метод возвращает значение NULL.</span><span class="sxs-lookup"><span data-stu-id="fa854-1098">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="fa854-1099">Если в теле элемента отсутствуют сущности указанного типа, метод возвращает пустой массив.</span><span class="sxs-lookup"><span data-stu-id="fa854-1099">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="fa854-1100">В противном случае тип объектов в возвращаемом массиве зависит от типа сущности, запрошенной в параметре `entityType`.</span><span class="sxs-lookup"><span data-stu-id="fa854-1100">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="fa854-1101">Хотя минимальный уровень разрешений для использования этого метода — **Restricted**, для некоторых типов сущностей требуется доступ на уровне **ReadItem**, как указано в приведенной ниже таблице.</span><span class="sxs-lookup"><span data-stu-id="fa854-1101">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

|<span data-ttu-id="fa854-1102">Значение параметра `entityType`</span><span class="sxs-lookup"><span data-stu-id="fa854-1102">Value of `entityType`</span></span>|<span data-ttu-id="fa854-1103">Тип объектов в возвращаемом массиве</span><span class="sxs-lookup"><span data-stu-id="fa854-1103">Type of objects in returned array</span></span>|<span data-ttu-id="fa854-1104">Необходимый уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="fa854-1104">Required Permission Level</span></span>|
|---|---|---|
|`Address`|<span data-ttu-id="fa854-1105">String</span><span class="sxs-lookup"><span data-stu-id="fa854-1105">String</span></span>|<span data-ttu-id="fa854-1106">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="fa854-1106">**Restricted**</span></span>|
|`Contact`|<span data-ttu-id="fa854-1107">Contact</span><span class="sxs-lookup"><span data-stu-id="fa854-1107">Contact</span></span>|<span data-ttu-id="fa854-1108">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="fa854-1108">**ReadItem**</span></span>|
|`EmailAddress`|<span data-ttu-id="fa854-1109">String</span><span class="sxs-lookup"><span data-stu-id="fa854-1109">String</span></span>|<span data-ttu-id="fa854-1110">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="fa854-1110">**ReadItem**</span></span>|
|`MeetingSuggestion`|<span data-ttu-id="fa854-1111">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="fa854-1111">MeetingSuggestion</span></span>|<span data-ttu-id="fa854-1112">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="fa854-1112">**ReadItem**</span></span>|
|`PhoneNumber`|<span data-ttu-id="fa854-1113">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="fa854-1113">PhoneNumber</span></span>|<span data-ttu-id="fa854-1114">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="fa854-1114">**Restricted**</span></span>|
|`TaskSuggestion`|<span data-ttu-id="fa854-1115">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="fa854-1115">TaskSuggestion</span></span>|<span data-ttu-id="fa854-1116">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="fa854-1116">**ReadItem**</span></span>|
|`URL`|<span data-ttu-id="fa854-1117">String</span><span class="sxs-lookup"><span data-stu-id="fa854-1117">String</span></span>|<span data-ttu-id="fa854-1118">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="fa854-1118">**Restricted**</span></span>|

<span data-ttu-id="fa854-1119">Тип: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="fa854-1119">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="fa854-1120">Пример</span><span class="sxs-lookup"><span data-stu-id="fa854-1120">Example</span></span>

<span data-ttu-id="fa854-1121">В примере ниже показано, как получить доступ к массиву строк, которые представляют собой почтовые адреса в теле текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="fa854-1121">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion"></a><span data-ttu-id="fa854-1122">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="fa854-1122">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span></span>

<span data-ttu-id="fa854-1123">Возвращает известные сущности в выбранном элементе, которые проходят через именованный фильтр, определяемый в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="fa854-1123">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="fa854-1124">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="fa854-1124">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="fa854-1125">Метод `getFilteredEntitiesByName` возвращает сущности, соответствующие регулярному выражению, которое определяется в элементе правила [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) в XML-файле манифеста, с использованием указанного значения элемента `FilterName`.</span><span class="sxs-lookup"><span data-stu-id="fa854-1125">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="fa854-1126">Параметры</span><span class="sxs-lookup"><span data-stu-id="fa854-1126">Parameters:</span></span>

|<span data-ttu-id="fa854-1127">Имя</span><span class="sxs-lookup"><span data-stu-id="fa854-1127">Name</span></span>|<span data-ttu-id="fa854-1128">Тип</span><span class="sxs-lookup"><span data-stu-id="fa854-1128">Type</span></span>|<span data-ttu-id="fa854-1129">Описание</span><span class="sxs-lookup"><span data-stu-id="fa854-1129">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="fa854-1130">String</span><span class="sxs-lookup"><span data-stu-id="fa854-1130">String</span></span>|<span data-ttu-id="fa854-1131">Имя элемента правила `ItemHasKnownEntity`, определяющее соответствующий фильтр.</span><span class="sxs-lookup"><span data-stu-id="fa854-1131">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="fa854-1132">Требования</span><span class="sxs-lookup"><span data-stu-id="fa854-1132">Requirements</span></span>

|<span data-ttu-id="fa854-1133">Требование</span><span class="sxs-lookup"><span data-stu-id="fa854-1133">Requirement</span></span>|<span data-ttu-id="fa854-1134">Значение</span><span class="sxs-lookup"><span data-stu-id="fa854-1134">Value</span></span>|
|---|---|
|[<span data-ttu-id="fa854-1135">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="fa854-1135">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="fa854-1136">1.0</span><span class="sxs-lookup"><span data-stu-id="fa854-1136">1.0</span></span>|
|[<span data-ttu-id="fa854-1137">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="fa854-1137">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="fa854-1138">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fa854-1138">ReadItem</span></span>|
|[<span data-ttu-id="fa854-1139">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="fa854-1139">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="fa854-1140">Чтение</span><span class="sxs-lookup"><span data-stu-id="fa854-1140">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="fa854-1141">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="fa854-1141">Returns:</span></span>

<span data-ttu-id="fa854-p162">Если в манифесте нет элемента `ItemHasKnownEntity` со значением `FilterName`, соответствующим параметру `name`, метод возвращает `null`. Если параметр `name` соответствует элементу `ItemHasKnownEntity` в манифесте, но при этом в текущем элементе нет соответствующих сущностей, метод возвращает пустой массив.</span><span class="sxs-lookup"><span data-stu-id="fa854-p162">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="fa854-1144">Тип: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="fa854-1144">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span></span>

#### <a name="getinitializationcontextasyncoptions-callback"></a><span data-ttu-id="fa854-1145">getInitializationContextAsync([options], [callback])</span><span class="sxs-lookup"><span data-stu-id="fa854-1145">getInitializationContextAsync([options], [callback])</span></span>

<span data-ttu-id="fa854-1146">Получает данные инициализации, передаваемые при [активации надстройки интерактивным сообщением](https://docs.microsoft.com/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span><span class="sxs-lookup"><span data-stu-id="fa854-1146">Gets initialization data passed when the add-in is [activated by an actionable message](https://docs.microsoft.com/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span></span>

> [!NOTE]
> <span data-ttu-id="fa854-1147">Этот метод поддерживается только версией Outlook 2016 для Windows или более поздней (версии "нажми и работай" с номером больше 16.0.8413.1000) и Outlook в Интернете для Office 365.</span><span class="sxs-lookup"><span data-stu-id="fa854-1147">This method is only supported by Outlook 2016 or later for Windows (Click-to-Run versions later than 16.0.8413.1000) and Outlook on the web for Office 365.</span></span>

##### <a name="parameters"></a><span data-ttu-id="fa854-1148">Параметры:</span><span class="sxs-lookup"><span data-stu-id="fa854-1148">Parameters:</span></span>
|<span data-ttu-id="fa854-1149">Имя</span><span class="sxs-lookup"><span data-stu-id="fa854-1149">Name</span></span>|<span data-ttu-id="fa854-1150">Тип</span><span class="sxs-lookup"><span data-stu-id="fa854-1150">Type</span></span>|<span data-ttu-id="fa854-1151">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="fa854-1151">Attributes</span></span>|<span data-ttu-id="fa854-1152">Описание</span><span class="sxs-lookup"><span data-stu-id="fa854-1152">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="fa854-1153">Object</span><span class="sxs-lookup"><span data-stu-id="fa854-1153">Object</span></span>|<span data-ttu-id="fa854-1154">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="fa854-1154">&lt;optional&gt;</span></span>|<span data-ttu-id="fa854-1155">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="fa854-1155">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="fa854-1156">Object</span><span class="sxs-lookup"><span data-stu-id="fa854-1156">Object</span></span>|<span data-ttu-id="fa854-1157">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="fa854-1157">&lt;optional&gt;</span></span>|<span data-ttu-id="fa854-1158">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="fa854-1158">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="fa854-1159">функция</span><span class="sxs-lookup"><span data-stu-id="fa854-1159">function</span></span>|<span data-ttu-id="fa854-1160">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="fa854-1160">&lt;optional&gt;</span></span>|<span data-ttu-id="fa854-1161">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="fa854-1161">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="fa854-1162">В случае успешного выполнения данные инициализации предоставляются в свойстве `asyncResult.value` как строка.</span><span class="sxs-lookup"><span data-stu-id="fa854-1162">On success, the initialization data is provided in the `asyncResult.value` property as a string.</span></span><br/><span data-ttu-id="fa854-1163">Если контекст инициализации отсутствует, объект `asyncResult` будет содержать объект `Error`, одному свойству которого (`code`) будет присвоено значение `9020`, а другому (`name`) — значение `GenericResponseError`.</span><span class="sxs-lookup"><span data-stu-id="fa854-1163">If there is no initialization context, the `asyncResult` object will contain an `Error` object with its `code` property set to `9020` and its `name` property set to `GenericResponseError`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="fa854-1164">Требования</span><span class="sxs-lookup"><span data-stu-id="fa854-1164">Requirements</span></span>

|<span data-ttu-id="fa854-1165">Требование</span><span class="sxs-lookup"><span data-stu-id="fa854-1165">Requirement</span></span>|<span data-ttu-id="fa854-1166">Значение</span><span class="sxs-lookup"><span data-stu-id="fa854-1166">Value</span></span>|
|---|---|
|[<span data-ttu-id="fa854-1167">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="fa854-1167">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="fa854-1168">Предварительная версия</span><span class="sxs-lookup"><span data-stu-id="fa854-1168">Preview</span></span>|
|[<span data-ttu-id="fa854-1169">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="fa854-1169">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="fa854-1170">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fa854-1170">ReadItem</span></span>|
|[<span data-ttu-id="fa854-1171">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="fa854-1171">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="fa854-1172">Чтение</span><span class="sxs-lookup"><span data-stu-id="fa854-1172">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="fa854-1173">Пример</span><span class="sxs-lookup"><span data-stu-id="fa854-1173">Example</span></span>

```javascript
// Get the initialization context (if present)
Office.context.mailbox.item.getInitializationContextAsync(
  function(asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
      if (asyncResult.value != null && asyncResult.value.length > 0) {
        // The value is a string, parse to an object
        var context = JSON.parse(asyncResult.value);
        // Do something with context
      } else {
        // Empty context, treat as no context
      }
    } else {
      if (asyncResult.error.code == 9020) {
        // GenericResponseError returned when there is
        // no context
        // Treat as no context
      } else {
        // Handle the error
      }
    }
  }
);
```

#### <a name="getregexmatches--object"></a><span data-ttu-id="fa854-1174">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="fa854-1174">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="fa854-1175">Возвращает строковые значения в выбранном элементе, которые соответствуют регулярным выражениям, определенным в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="fa854-1175">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="fa854-1176">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="fa854-1176">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="fa854-p163">Метод `getRegExMatches` возвращает строки, соответствующие регулярному выражению, которое определяется в каждом элементе правила `ItemHasRegularExpressionMatch` или `ItemHasKnownEntity` в XML-файле манифеста. Для правила `ItemHasRegularExpressionMatch` соответствующую строку должно содержать свойство элемента, указанного этим правилом. Простой тип `PropertyName` определяет поддерживаемые свойства.</span><span class="sxs-lookup"><span data-stu-id="fa854-p163">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="fa854-1180">Например, рассмотрим манифест надстройки, который содержит указанный ниже элемент `Rule`.</span><span class="sxs-lookup"><span data-stu-id="fa854-1180">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="fa854-1181">Объект, возвращаемый методом `getRegExMatches`, будет содержать два свойства: `fruits` и `veggies`.</span><span class="sxs-lookup"><span data-stu-id="fa854-1181">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="fa854-p164">Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты. Лучше используйте метод [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) для этого.</span><span class="sxs-lookup"><span data-stu-id="fa854-p164">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="fa854-1185">Requirements</span><span class="sxs-lookup"><span data-stu-id="fa854-1185">Requirements</span></span>

|<span data-ttu-id="fa854-1186">Требование</span><span class="sxs-lookup"><span data-stu-id="fa854-1186">Requirement</span></span>|<span data-ttu-id="fa854-1187">Значение</span><span class="sxs-lookup"><span data-stu-id="fa854-1187">Value</span></span>|
|---|---|
|[<span data-ttu-id="fa854-1188">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="fa854-1188">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="fa854-1189">1.0</span><span class="sxs-lookup"><span data-stu-id="fa854-1189">1.0</span></span>|
|[<span data-ttu-id="fa854-1190">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="fa854-1190">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="fa854-1191">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fa854-1191">ReadItem</span></span>|
|[<span data-ttu-id="fa854-1192">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="fa854-1192">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="fa854-1193">Чтение</span><span class="sxs-lookup"><span data-stu-id="fa854-1193">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="fa854-1194">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="fa854-1194">Returns:</span></span>

<span data-ttu-id="fa854-p165">Объект, содержащий массив строк, которые соответствуют регулярным выражениям, определяемым в XML-файле манифеста. Имя каждого массива равно соответствующему значению атрибута `RegExName` подходящего правила `ItemHasRegularExpressionMatch` или атрибута `FilterName` соответствующего правила `ItemHasKnownEntity`.</span><span class="sxs-lookup"><span data-stu-id="fa854-p165">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="fa854-1197">

<dt>Тип</dt>

</span><span class="sxs-lookup"><span data-stu-id="fa854-1197">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="fa854-1198">Object</span><span class="sxs-lookup"><span data-stu-id="fa854-1198">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="fa854-1199">Пример</span><span class="sxs-lookup"><span data-stu-id="fa854-1199">Example</span></span>

<span data-ttu-id="fa854-1200">В приведенном ниже примере показано, как получить доступ к массиву совпадений с элементами `fruits` и `veggies` правил активации регулярных выражений, указанными в манифесте.</span><span class="sxs-lookup"><span data-stu-id="fa854-1200">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```javascript
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veges = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="fa854-1201">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="fa854-1201">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="fa854-1202">Возвращает строковые значения в выбранном элементе, которые соответствуют именованному регулярному выражению, определенному в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="fa854-1202">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="fa854-1203">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="fa854-1203">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="fa854-1204">Метод `getRegExMatchesByName` возвращает строки, соответствующие регулярному выражению, которое определяется в элементе правила `ItemHasRegularExpressionMatch` в XML-файле манифеста, с использованием указанного значения элемента `RegExName`.</span><span class="sxs-lookup"><span data-stu-id="fa854-1204">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="fa854-p166">Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты.</span><span class="sxs-lookup"><span data-stu-id="fa854-p166">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="fa854-1207">Параметры</span><span class="sxs-lookup"><span data-stu-id="fa854-1207">Parameters:</span></span>

|<span data-ttu-id="fa854-1208">Имя</span><span class="sxs-lookup"><span data-stu-id="fa854-1208">Name</span></span>|<span data-ttu-id="fa854-1209">Тип</span><span class="sxs-lookup"><span data-stu-id="fa854-1209">Type</span></span>|<span data-ttu-id="fa854-1210">Описание</span><span class="sxs-lookup"><span data-stu-id="fa854-1210">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="fa854-1211">String</span><span class="sxs-lookup"><span data-stu-id="fa854-1211">String</span></span>|<span data-ttu-id="fa854-1212">Имя элемента правила `ItemHasRegularExpressionMatch`, определяющее соответствующий фильтр.</span><span class="sxs-lookup"><span data-stu-id="fa854-1212">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="fa854-1213">Требования</span><span class="sxs-lookup"><span data-stu-id="fa854-1213">Requirements</span></span>

|<span data-ttu-id="fa854-1214">Требование</span><span class="sxs-lookup"><span data-stu-id="fa854-1214">Requirement</span></span>|<span data-ttu-id="fa854-1215">Значение</span><span class="sxs-lookup"><span data-stu-id="fa854-1215">Value</span></span>|
|---|---|
|[<span data-ttu-id="fa854-1216">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="fa854-1216">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="fa854-1217">1.0</span><span class="sxs-lookup"><span data-stu-id="fa854-1217">1.0</span></span>|
|[<span data-ttu-id="fa854-1218">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="fa854-1218">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="fa854-1219">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fa854-1219">ReadItem</span></span>|
|[<span data-ttu-id="fa854-1220">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="fa854-1220">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="fa854-1221">Чтение</span><span class="sxs-lookup"><span data-stu-id="fa854-1221">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="fa854-1222">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="fa854-1222">Returns:</span></span>

<span data-ttu-id="fa854-1223">Массив строк, соответствующих регулярному выражению, определяемому в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="fa854-1223">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="fa854-1224">

<dt>Тип</dt>

</span><span class="sxs-lookup"><span data-stu-id="fa854-1224">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="fa854-1225">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="fa854-1225">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="fa854-1226">Пример</span><span class="sxs-lookup"><span data-stu-id="fa854-1226">Example</span></span>

```javascript
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="fa854-1227">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="fa854-1227">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="fa854-1228">Асинхронно возвращает данные, выбранные в теме или тексте сообщения.</span><span class="sxs-lookup"><span data-stu-id="fa854-1228">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="fa854-p167">Если выделенный фрагмент отсутствует, но курсор находится в тексте или теме, метод возвращает значение NULL для выбранных данных. Если выбраны не текст и не тема, метод возвращает ошибку `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="fa854-p167">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="fa854-1231">Параметры</span><span class="sxs-lookup"><span data-stu-id="fa854-1231">Parameters:</span></span>

|<span data-ttu-id="fa854-1232">Имя</span><span class="sxs-lookup"><span data-stu-id="fa854-1232">Name</span></span>|<span data-ttu-id="fa854-1233">Тип</span><span class="sxs-lookup"><span data-stu-id="fa854-1233">Type</span></span>|<span data-ttu-id="fa854-1234">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="fa854-1234">Attributes</span></span>|<span data-ttu-id="fa854-1235">Описание</span><span class="sxs-lookup"><span data-stu-id="fa854-1235">Description</span></span>|
|---|---|---|---|
|`coercionType`|[<span data-ttu-id="fa854-1236">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="fa854-1236">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="fa854-p168">Запрашивает формат данных. Если задано значение Text, метод возвращает обычный текст как строку, удаляя все имеющиеся HTML-теги. Если задано значение HTML, метод возвращает выделенный текст (обычный текст или HTML).</span><span class="sxs-lookup"><span data-stu-id="fa854-p168">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`|<span data-ttu-id="fa854-1240">Object</span><span class="sxs-lookup"><span data-stu-id="fa854-1240">Object</span></span>|<span data-ttu-id="fa854-1241">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="fa854-1241">&lt;optional&gt;</span></span>|<span data-ttu-id="fa854-1242">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="fa854-1242">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="fa854-1243">Object</span><span class="sxs-lookup"><span data-stu-id="fa854-1243">Object</span></span>|<span data-ttu-id="fa854-1244">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="fa854-1244">&lt;optional&gt;</span></span>|<span data-ttu-id="fa854-1245">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="fa854-1245">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="fa854-1246">функция</span><span class="sxs-lookup"><span data-stu-id="fa854-1246">function</span></span>||<span data-ttu-id="fa854-1247">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="fa854-1247">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="fa854-1248">Чтобы получить доступ к выбранным данным из метода обратного вызова, вызовите `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="fa854-1248">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="fa854-1249">Для доступа к исходному свойству, представляющему собой источник выбранных данных, вызовите параметр `asyncResult.value.sourceProperty`, который может иметь значение `body` или `subject`.</span><span class="sxs-lookup"><span data-stu-id="fa854-1249">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="fa854-1250">Требования</span><span class="sxs-lookup"><span data-stu-id="fa854-1250">Requirements</span></span>

|<span data-ttu-id="fa854-1251">Требование</span><span class="sxs-lookup"><span data-stu-id="fa854-1251">Requirement</span></span>|<span data-ttu-id="fa854-1252">Значение</span><span class="sxs-lookup"><span data-stu-id="fa854-1252">Value</span></span>|
|---|---|
|[<span data-ttu-id="fa854-1253">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="fa854-1253">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="fa854-1254">1.2</span><span class="sxs-lookup"><span data-stu-id="fa854-1254">1.2</span></span>|
|[<span data-ttu-id="fa854-1255">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="fa854-1255">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="fa854-1256">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="fa854-1256">ReadWriteItem</span></span>|
|[<span data-ttu-id="fa854-1257">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="fa854-1257">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="fa854-1258">Создание</span><span class="sxs-lookup"><span data-stu-id="fa854-1258">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="fa854-1259">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="fa854-1259">Returns:</span></span>

<span data-ttu-id="fa854-1260">Выбранные данные в виде строки с форматом, определенным в параметре `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="fa854-1260">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="fa854-1261">

<dt>Тип</dt>

</span><span class="sxs-lookup"><span data-stu-id="fa854-1261">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="fa854-1262">String</span><span class="sxs-lookup"><span data-stu-id="fa854-1262">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="fa854-1263">Пример</span><span class="sxs-lookup"><span data-stu-id="fa854-1263">Example</span></span>

```javascript
// getting selected data
Office.initialize = function () {
    Office.context.mailbox.item.getSelectedDataAsync(Office.CoercionType.Text, {}, getCallback);
}

function getCallback(asyncResult) {
    var text = asyncResult.value.data;
    var prop = asyncResult.value.sourceProperty;

    Office.context.mailbox.item.setSelectedDataAsync('Setting ' + prop + ': ' + text, {}, setCallback);
}

function setCallback(asyncResult) {
    // check for errors
}
```

#### <a name="getselectedentities--entitiesjavascriptapioutlookofficeentities"></a><span data-ttu-id="fa854-1264">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="fa854-1264">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span></span>

<span data-ttu-id="fa854-p170">Возвращает сущности, найденные в выделенном совпадении, выбранном пользователем. Выделенные совпадения применяются к [контекстным надстройкам](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="fa854-p170">Gets the entities found in a highlighted match a user has selected. Highlighted matches apply to [contextual add-ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="fa854-1267">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="fa854-1267">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="fa854-1268">Требования</span><span class="sxs-lookup"><span data-stu-id="fa854-1268">Requirements</span></span>

|<span data-ttu-id="fa854-1269">Требование</span><span class="sxs-lookup"><span data-stu-id="fa854-1269">Requirement</span></span>|<span data-ttu-id="fa854-1270">Значение</span><span class="sxs-lookup"><span data-stu-id="fa854-1270">Value</span></span>|
|---|---|
|[<span data-ttu-id="fa854-1271">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="fa854-1271">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="fa854-1272">1.6</span><span class="sxs-lookup"><span data-stu-id="fa854-1272">1.6</span></span>|
|[<span data-ttu-id="fa854-1273">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="fa854-1273">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="fa854-1274">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fa854-1274">ReadItem</span></span>|
|[<span data-ttu-id="fa854-1275">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="fa854-1275">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="fa854-1276">Чтение</span><span class="sxs-lookup"><span data-stu-id="fa854-1276">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="fa854-1277">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="fa854-1277">Returns:</span></span>

<span data-ttu-id="fa854-1278">Тип: [Entities](/javascript/api/outlook/office.entities)</span><span class="sxs-lookup"><span data-stu-id="fa854-1278">Type: [Entities](/javascript/api/outlook/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="fa854-1279">Пример</span><span class="sxs-lookup"><span data-stu-id="fa854-1279">Example</span></span>

<span data-ttu-id="fa854-1280">В приведенном ниже примере показано, как получить доступ к сущностям адресов в выделенном совпадении, выбранном пользователем.</span><span class="sxs-lookup"><span data-stu-id="fa854-1280">The following example accesses the addresses entities in the highlighted match selected by the user.</span></span>

```javascript
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

#### <a name="getselectedregexmatches--object"></a><span data-ttu-id="fa854-1281">getSelectedRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="fa854-1281">getSelectedRegExMatches() → {Object}</span></span>

<span data-ttu-id="fa854-p171">Возвращает строковые значения в выделенном совпадении, которые соответствуют регулярным выражениям, определенным в XML-файле манифеста. Выделенные совпадения применяются к [контекстным надстройкам](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="fa854-p171">Returns string values in a highlighted match that match the regular expressions defined in the manifest XML file. Highlighted matches apply to [contextual add-ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="fa854-1284">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="fa854-1284">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="fa854-p172">Метод `getSelectedRegExMatches` возвращает строки, соответствующие регулярному выражению, которое определяется в каждом элементе правила `ItemHasRegularExpressionMatch` или `ItemHasKnownEntity` в XML-файле манифеста. Для правила `ItemHasRegularExpressionMatch` соответствующую строку должно содержать свойство элемента, указанного этим правилом. Простой тип `PropertyName` определяет поддерживаемые свойства.</span><span class="sxs-lookup"><span data-stu-id="fa854-p172">The `getSelectedRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="fa854-1288">Например, рассмотрим манифест надстройки, который содержит указанный ниже элемент `Rule`.</span><span class="sxs-lookup"><span data-stu-id="fa854-1288">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="fa854-1289">Объект, возвращаемый методом `getRegExMatches`, будет содержать два свойства: `fruits` и `veggies`.</span><span class="sxs-lookup"><span data-stu-id="fa854-1289">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="fa854-p173">Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты. Лучше используйте метод [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) для этого.</span><span class="sxs-lookup"><span data-stu-id="fa854-p173">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="fa854-1293">Requirements</span><span class="sxs-lookup"><span data-stu-id="fa854-1293">Requirements</span></span>

|<span data-ttu-id="fa854-1294">Требование</span><span class="sxs-lookup"><span data-stu-id="fa854-1294">Requirement</span></span>|<span data-ttu-id="fa854-1295">Значение</span><span class="sxs-lookup"><span data-stu-id="fa854-1295">Value</span></span>|
|---|---|
|[<span data-ttu-id="fa854-1296">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="fa854-1296">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="fa854-1297">1.6</span><span class="sxs-lookup"><span data-stu-id="fa854-1297">1.6</span></span>|
|[<span data-ttu-id="fa854-1298">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="fa854-1298">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="fa854-1299">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fa854-1299">ReadItem</span></span>|
|[<span data-ttu-id="fa854-1300">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="fa854-1300">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="fa854-1301">Чтение</span><span class="sxs-lookup"><span data-stu-id="fa854-1301">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="fa854-1302">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="fa854-1302">Returns:</span></span>

<span data-ttu-id="fa854-p174">Объект, содержащий массив строк, которые соответствуют регулярным выражениям, определяемым в XML-файле манифеста. Имя каждого массива равно соответствующему значению атрибута `RegExName` подходящего правила `ItemHasRegularExpressionMatch` или атрибута `FilterName` соответствующего правила `ItemHasKnownEntity`.</span><span class="sxs-lookup"><span data-stu-id="fa854-p174">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

##### <a name="example"></a><span data-ttu-id="fa854-1305">Пример</span><span class="sxs-lookup"><span data-stu-id="fa854-1305">Example</span></span>

<span data-ttu-id="fa854-1306">В приведенном ниже примере показано, как получить доступ к массиву совпадений с элементами `fruits` и `veggies` правил активации регулярных выражений, указанными в манифесте.</span><span class="sxs-lookup"><span data-stu-id="fa854-1306">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```javascript
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

#### <a name="getsharedpropertiesasyncoptions-callback"></a><span data-ttu-id="fa854-1307">getSharedPropertiesAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="fa854-1307">getSharedPropertiesAsync([options], callback)</span></span>

<span data-ttu-id="fa854-1308">Получает свойства выбранного встречи или сообщения в общей папке, календаре или почтовом ящике.</span><span class="sxs-lookup"><span data-stu-id="fa854-1308">Gets the properties of the selected appointment or message in a shared folder, calendar, or mailbox.</span></span>

##### <a name="parameters"></a><span data-ttu-id="fa854-1309">Параметры:</span><span class="sxs-lookup"><span data-stu-id="fa854-1309">Parameters:</span></span>

|<span data-ttu-id="fa854-1310">Имя</span><span class="sxs-lookup"><span data-stu-id="fa854-1310">Name</span></span>|<span data-ttu-id="fa854-1311">Тип</span><span class="sxs-lookup"><span data-stu-id="fa854-1311">Type</span></span>|<span data-ttu-id="fa854-1312">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="fa854-1312">Attributes</span></span>|<span data-ttu-id="fa854-1313">Описание</span><span class="sxs-lookup"><span data-stu-id="fa854-1313">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="fa854-1314">Object</span><span class="sxs-lookup"><span data-stu-id="fa854-1314">Object</span></span>|<span data-ttu-id="fa854-1315">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="fa854-1315">&lt;optional&gt;</span></span>|<span data-ttu-id="fa854-1316">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="fa854-1316">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="fa854-1317">Object</span><span class="sxs-lookup"><span data-stu-id="fa854-1317">Object</span></span>|<span data-ttu-id="fa854-1318">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="fa854-1318">&lt;optional&gt;</span></span>|<span data-ttu-id="fa854-1319">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="fa854-1319">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="fa854-1320">функция</span><span class="sxs-lookup"><span data-stu-id="fa854-1320">function</span></span>||<span data-ttu-id="fa854-1321">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="fa854-1321">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="fa854-1322">Общие свойства предоставляются в виде объекта [`SharedProperties`](/javascript/api/outlook/office.sharedproperties) в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="fa854-1322">The shared properties are provided as a [`SharedProperties`](/javascript/api/outlook/office.sharedproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="fa854-1323">Этот объект можно использовать для получения общих свойств элемента.</span><span class="sxs-lookup"><span data-stu-id="fa854-1323">This object can be used to get the item's shared properties.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="fa854-1324">Требования</span><span class="sxs-lookup"><span data-stu-id="fa854-1324">Requirements</span></span>

|<span data-ttu-id="fa854-1325">Требование</span><span class="sxs-lookup"><span data-stu-id="fa854-1325">Requirement</span></span>|<span data-ttu-id="fa854-1326">Значение</span><span class="sxs-lookup"><span data-stu-id="fa854-1326">Value</span></span>|
|---|---|
|[<span data-ttu-id="fa854-1327">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="fa854-1327">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="fa854-1328">Предварительная версия</span><span class="sxs-lookup"><span data-stu-id="fa854-1328">Preview</span></span>|
|[<span data-ttu-id="fa854-1329">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="fa854-1329">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="fa854-1330">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fa854-1330">ReadItem</span></span>|
|[<span data-ttu-id="fa854-1331">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="fa854-1331">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="fa854-1332">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="fa854-1332">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="fa854-1333">Пример</span><span class="sxs-lookup"><span data-stu-id="fa854-1333">Example</span></span>

```js
Office.context.mailbox.item.getSharedPropertiesAsync(callback);
function callback (asyncResult) {
  var context=asyncResult.context;
  var sharedProperties = asyncResult.value;
}
```

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="fa854-1334">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="fa854-1334">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="fa854-1335">Асинхронно загружает настраиваемые свойства для надстройки для выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="fa854-1335">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="fa854-p176">Настраиваемые свойства сохраняются в виде пар "ключ-значение" для каждого приложения и каждого элемента. Этот метод возвращает объект `CustomProperties` при обратном вызове, который предоставляет методы для доступа к настраиваемым свойствам, характерным для текущего элемента и текущей надстройки. Настраиваемые свойства не шифруются для элемента, поэтому этот способ хранения не является безопасным.</span><span class="sxs-lookup"><span data-stu-id="fa854-p176">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="fa854-1339">Параметры</span><span class="sxs-lookup"><span data-stu-id="fa854-1339">Parameters:</span></span>

|<span data-ttu-id="fa854-1340">Имя</span><span class="sxs-lookup"><span data-stu-id="fa854-1340">Name</span></span>|<span data-ttu-id="fa854-1341">Тип</span><span class="sxs-lookup"><span data-stu-id="fa854-1341">Type</span></span>|<span data-ttu-id="fa854-1342">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="fa854-1342">Attributes</span></span>|<span data-ttu-id="fa854-1343">Описание</span><span class="sxs-lookup"><span data-stu-id="fa854-1343">Description</span></span>|
|---|---|---|---|
|`callback`|<span data-ttu-id="fa854-1344">function</span><span class="sxs-lookup"><span data-stu-id="fa854-1344">function</span></span>||<span data-ttu-id="fa854-1345">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="fa854-1345">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="fa854-1346">Настраиваемые свойства предоставляются в виде объекта [`CustomProperties`](/javascript/api/outlook/office.customproperties) в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="fa854-1346">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="fa854-1347">Этот объект позволяет получить, задать и удалить настраиваемые свойства для элемента, а также сохранить изменения, внесенные в набор настраиваемых свойств, на сервере.</span><span class="sxs-lookup"><span data-stu-id="fa854-1347">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`|<span data-ttu-id="fa854-1348">Object</span><span class="sxs-lookup"><span data-stu-id="fa854-1348">Object</span></span>|<span data-ttu-id="fa854-1349">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="fa854-1349">&lt;optional&gt;</span></span>|<span data-ttu-id="fa854-1350">Разработчики могут указать любой объект, к которому необходимо получить доступ, в функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="fa854-1350">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="fa854-1351">Доступ к этому объекту можно получить с помощью свойства `asyncResult.asyncContext` в функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="fa854-1351">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="fa854-1352">Требования</span><span class="sxs-lookup"><span data-stu-id="fa854-1352">Requirements</span></span>

|<span data-ttu-id="fa854-1353">Требование</span><span class="sxs-lookup"><span data-stu-id="fa854-1353">Requirement</span></span>|<span data-ttu-id="fa854-1354">Значение</span><span class="sxs-lookup"><span data-stu-id="fa854-1354">Value</span></span>|
|---|---|
|[<span data-ttu-id="fa854-1355">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="fa854-1355">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="fa854-1356">1.0</span><span class="sxs-lookup"><span data-stu-id="fa854-1356">1.0</span></span>|
|[<span data-ttu-id="fa854-1357">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="fa854-1357">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="fa854-1358">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fa854-1358">ReadItem</span></span>|
|[<span data-ttu-id="fa854-1359">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="fa854-1359">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="fa854-1360">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="fa854-1360">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="fa854-1361">Пример</span><span class="sxs-lookup"><span data-stu-id="fa854-1361">Example</span></span>

<span data-ttu-id="fa854-p179">Приведенный ниже пример кода показывает, как асинхронно загружать настраиваемые свойства, характерные для текущего элемента, с помощью метода `loadCustomPropertiesAsync`. Этот пример также показывает, как сохранять эти свойства на сервере с помощью метода `CustomProperties.saveAsync`. После загрузки настраиваемых свойств в этом примере кода метод `CustomProperties.get` используется для считывания настраиваемого свойства `myProp`, метод `CustomProperties.set` — для записи настраиваемого свойства `otherProp`, а метод `saveAsync` — для сохранения настраиваемых свойств.</span><span class="sxs-lookup"><span data-stu-id="fa854-p179">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="fa854-1365">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="fa854-1365">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="fa854-1366">Удаляет вложение из сообщения или встречи.</span><span class="sxs-lookup"><span data-stu-id="fa854-1366">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="fa854-1367">Метод `removeAttachmentAsync` удаляет из элемента вложение с указанным идентификатором.</span><span class="sxs-lookup"><span data-stu-id="fa854-1367">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="fa854-1368">Идентификатор вложения рекомендуется использовать для удаления вложения, только если оно добавлено тем же почтовым приложением в ходе текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="fa854-1368">As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session.</span></span> <span data-ttu-id="fa854-1369">В Outlook Web App и Outlook Web App для устройств идентификатор вложения действителен только в рамках одного сеанса.</span><span class="sxs-lookup"><span data-stu-id="fa854-1369">In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="fa854-1370">Сеанс завершается, когда пользователь закрывает приложение или начинает создавать элемент во встроенной форме, а затем переходит из формы в отдельное окно.</span><span class="sxs-lookup"><span data-stu-id="fa854-1370">A session is over when the user closes the app, or if the user starts composing an inline form then subsequently pops out the form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="fa854-1371">Параметры:</span><span class="sxs-lookup"><span data-stu-id="fa854-1371">Parameters:</span></span>

|<span data-ttu-id="fa854-1372">Имя</span><span class="sxs-lookup"><span data-stu-id="fa854-1372">Name</span></span>|<span data-ttu-id="fa854-1373">Тип</span><span class="sxs-lookup"><span data-stu-id="fa854-1373">Type</span></span>|<span data-ttu-id="fa854-1374">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="fa854-1374">Attributes</span></span>|<span data-ttu-id="fa854-1375">Описание</span><span class="sxs-lookup"><span data-stu-id="fa854-1375">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="fa854-1376">String</span><span class="sxs-lookup"><span data-stu-id="fa854-1376">String</span></span>||<span data-ttu-id="fa854-1377">Идентификатор удаляемого вложения.</span><span class="sxs-lookup"><span data-stu-id="fa854-1377">The identifier of the attachment to remove.</span></span>|
|`options`|<span data-ttu-id="fa854-1378">Object</span><span class="sxs-lookup"><span data-stu-id="fa854-1378">Object</span></span>|<span data-ttu-id="fa854-1379">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="fa854-1379">&lt;optional&gt;</span></span>|<span data-ttu-id="fa854-1380">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="fa854-1380">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="fa854-1381">Object</span><span class="sxs-lookup"><span data-stu-id="fa854-1381">Object</span></span>|<span data-ttu-id="fa854-1382">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="fa854-1382">&lt;optional&gt;</span></span>|<span data-ttu-id="fa854-1383">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="fa854-1383">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="fa854-1384">функция</span><span class="sxs-lookup"><span data-stu-id="fa854-1384">function</span></span>|<span data-ttu-id="fa854-1385">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="fa854-1385">&lt;optional&gt;</span></span>|<span data-ttu-id="fa854-1386">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="fa854-1386">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="fa854-1387">Если удалить вложение не удается, свойство `asyncResult.error` содержит код ошибки с указанием ее причины.</span><span class="sxs-lookup"><span data-stu-id="fa854-1387">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="fa854-1388">Ошибки</span><span class="sxs-lookup"><span data-stu-id="fa854-1388">Errors</span></span>

|<span data-ttu-id="fa854-1389">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="fa854-1389">Error code</span></span>|<span data-ttu-id="fa854-1390">Описание</span><span class="sxs-lookup"><span data-stu-id="fa854-1390">Description</span></span>|
|------------|-------------|
|`InvalidAttachmentId`|<span data-ttu-id="fa854-1391">Идентификатор вложения не существует.</span><span class="sxs-lookup"><span data-stu-id="fa854-1391">The attachment identifier does not exist.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="fa854-1392">Требования</span><span class="sxs-lookup"><span data-stu-id="fa854-1392">Requirements</span></span>

|<span data-ttu-id="fa854-1393">Требование</span><span class="sxs-lookup"><span data-stu-id="fa854-1393">Requirement</span></span>|<span data-ttu-id="fa854-1394">Значение</span><span class="sxs-lookup"><span data-stu-id="fa854-1394">Value</span></span>|
|---|---|
|[<span data-ttu-id="fa854-1395">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="fa854-1395">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="fa854-1396">1.1</span><span class="sxs-lookup"><span data-stu-id="fa854-1396">1.1</span></span>|
|[<span data-ttu-id="fa854-1397">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="fa854-1397">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="fa854-1398">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="fa854-1398">ReadWriteItem</span></span>|
|[<span data-ttu-id="fa854-1399">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="fa854-1399">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="fa854-1400">Создание</span><span class="sxs-lookup"><span data-stu-id="fa854-1400">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="fa854-1401">Пример</span><span class="sxs-lookup"><span data-stu-id="fa854-1401">Example</span></span>

<span data-ttu-id="fa854-1402">Указанный ниже код удаляет вложение с идентификатором "0".</span><span class="sxs-lookup"><span data-stu-id="fa854-1402">The following code removes an attachment with an identifier of '0'.</span></span>

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

####  <a name="removehandlerasynceventtype-options-callback"></a><span data-ttu-id="fa854-1403">removeHandlerAsync(eventType, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="fa854-1403">removeHandlerAsync(eventType, [options], [callback])</span></span>

<span data-ttu-id="fa854-1404">Удаляет обработчиков для поддерживаемого типа события.</span><span class="sxs-lookup"><span data-stu-id="fa854-1404">Removes the event handlers for a supported event type.</span></span>

<span data-ttu-id="fa854-1405">Сейчас поддерживаются следующие типы событий: `Office.EventType.AttachmentsChanged`, `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged` и `Office.EventType.RecurrenceChanged`.</span><span class="sxs-lookup"><span data-stu-id="fa854-1405">Currently the supported event types are `Office.EventType.AttachmentsChanged`, `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="fa854-1406">Параметры:</span><span class="sxs-lookup"><span data-stu-id="fa854-1406">Parameters:</span></span>

| <span data-ttu-id="fa854-1407">Имя</span><span class="sxs-lookup"><span data-stu-id="fa854-1407">Name</span></span> | <span data-ttu-id="fa854-1408">Тип</span><span class="sxs-lookup"><span data-stu-id="fa854-1408">Type</span></span> | <span data-ttu-id="fa854-1409">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="fa854-1409">Attributes</span></span> | <span data-ttu-id="fa854-1410">Описание</span><span class="sxs-lookup"><span data-stu-id="fa854-1410">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="fa854-1411">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="fa854-1411">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="fa854-1412">Событие, которое должно отменить обработчик.</span><span class="sxs-lookup"><span data-stu-id="fa854-1412">The event that should revoke the handler.</span></span> |
| `options` | <span data-ttu-id="fa854-1413">Объект</span><span class="sxs-lookup"><span data-stu-id="fa854-1413">Object</span></span> | <span data-ttu-id="fa854-1414">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="fa854-1414">&lt;optional&gt;</span></span> | <span data-ttu-id="fa854-1415">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="fa854-1415">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="fa854-1416">Object</span><span class="sxs-lookup"><span data-stu-id="fa854-1416">Object</span></span> | <span data-ttu-id="fa854-1417">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="fa854-1417">&lt;optional&gt;</span></span> | <span data-ttu-id="fa854-1418">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="fa854-1418">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="fa854-1419">функция</span><span class="sxs-lookup"><span data-stu-id="fa854-1419">function</span></span>| <span data-ttu-id="fa854-1420">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="fa854-1420">&lt;optional&gt;</span></span>|<span data-ttu-id="fa854-1421">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="fa854-1421">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="fa854-1422">Требования</span><span class="sxs-lookup"><span data-stu-id="fa854-1422">Requirements</span></span>

|<span data-ttu-id="fa854-1423">Требование</span><span class="sxs-lookup"><span data-stu-id="fa854-1423">Requirement</span></span>| <span data-ttu-id="fa854-1424">Значение</span><span class="sxs-lookup"><span data-stu-id="fa854-1424">Value</span></span>|
|---|---|
|[<span data-ttu-id="fa854-1425">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="fa854-1425">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fa854-1426">1.7</span><span class="sxs-lookup"><span data-stu-id="fa854-1426">1.7</span></span> |
|[<span data-ttu-id="fa854-1427">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="fa854-1427">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fa854-1428">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fa854-1428">ReadItem</span></span> |
|[<span data-ttu-id="fa854-1429">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="fa854-1429">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fa854-1430">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="fa854-1430">Compose or read</span></span> |

####  <a name="saveasyncoptions-callback"></a><span data-ttu-id="fa854-1431">saveAsync([options], обратный вызов)</span><span class="sxs-lookup"><span data-stu-id="fa854-1431">saveAsync([options], callback)</span></span>

<span data-ttu-id="fa854-1432">Асинхронно сохраняет элемент.</span><span class="sxs-lookup"><span data-stu-id="fa854-1432">Asynchronously saves an item.</span></span>

<span data-ttu-id="fa854-p181">При вызове этот метод сохраняет текущее сообщение в виде черновика и возвращает идентификатор элемента с помощью метода обратного вызова. В Outlook Web App или интерактивном режиме Outlook этот элемент сохраняется на сервере. В Outlook в режиме кэширования этот элемент сохраняется в локальном кэше.</span><span class="sxs-lookup"><span data-stu-id="fa854-p181">When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="fa854-1436">Если в режиме создания надстройка вызывает для элемента метод `saveAsync`, чтобы получить параметр `itemId` для использования с EWS или REST API, необходимо помнить, что синхронизация элемента с сервером может занять много времени, если Outlook работает в режиме кэширования данных.</span><span class="sxs-lookup"><span data-stu-id="fa854-1436">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="fa854-1437">До окончания синхронизации элемента применение параметра `itemId` будет приводить к ошибке.</span><span class="sxs-lookup"><span data-stu-id="fa854-1437">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="fa854-p183">Если метод `saveAsync` вызывается для встречи в режиме создания, она сохраняется как обычная встреча в календаре пользователя, а не как черновик. При сохранении новой встречи приглашения не отправляются. При сохранении существующей встречи уведомления отправляются добавленным или удаленным участникам.</span><span class="sxs-lookup"><span data-stu-id="fa854-p183">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="fa854-1441">Следующие клиенты отличаются другим поведением `saveAsync` в отношении встреч в режиме создания:</span><span class="sxs-lookup"><span data-stu-id="fa854-1441">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="fa854-1442">Outlook для Mac не поддерживает `saveAsync` для собраний в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="fa854-1442">Mac Outlook does not support `saveAsync` on a meeting in compose mode.</span></span> <span data-ttu-id="fa854-1443">При вызове `saveAsync` для собрания в Outlook для Mac возвращается ошибка.</span><span class="sxs-lookup"><span data-stu-id="fa854-1443">Calling `saveAsync` on a meeting in Mac Outlook will return an error.</span></span>
> - <span data-ttu-id="fa854-1444">Outlook в Интернете всегда отправляет приглашение или обновление при вызове `saveAsync` для встречи в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="fa854-1444">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="fa854-1445">Параметры:</span><span class="sxs-lookup"><span data-stu-id="fa854-1445">Parameters:</span></span>

|<span data-ttu-id="fa854-1446">Имя</span><span class="sxs-lookup"><span data-stu-id="fa854-1446">Name</span></span>|<span data-ttu-id="fa854-1447">Тип</span><span class="sxs-lookup"><span data-stu-id="fa854-1447">Type</span></span>|<span data-ttu-id="fa854-1448">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="fa854-1448">Attributes</span></span>|<span data-ttu-id="fa854-1449">Описание</span><span class="sxs-lookup"><span data-stu-id="fa854-1449">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="fa854-1450">Object</span><span class="sxs-lookup"><span data-stu-id="fa854-1450">Object</span></span>|<span data-ttu-id="fa854-1451">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="fa854-1451">&lt;optional&gt;</span></span>|<span data-ttu-id="fa854-1452">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="fa854-1452">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="fa854-1453">Object</span><span class="sxs-lookup"><span data-stu-id="fa854-1453">Object</span></span>|<span data-ttu-id="fa854-1454">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="fa854-1454">&lt;optional&gt;</span></span>|<span data-ttu-id="fa854-1455">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="fa854-1455">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="fa854-1456">функция</span><span class="sxs-lookup"><span data-stu-id="fa854-1456">function</span></span>||<span data-ttu-id="fa854-1457">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="fa854-1457">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="fa854-1458">После успешного выполнения идентификатор элемента будет предоставлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="fa854-1458">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="fa854-1459">Требования</span><span class="sxs-lookup"><span data-stu-id="fa854-1459">Requirements</span></span>

|<span data-ttu-id="fa854-1460">Требование</span><span class="sxs-lookup"><span data-stu-id="fa854-1460">Requirement</span></span>|<span data-ttu-id="fa854-1461">Значение</span><span class="sxs-lookup"><span data-stu-id="fa854-1461">Value</span></span>|
|---|---|
|[<span data-ttu-id="fa854-1462">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="fa854-1462">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="fa854-1463">1.3</span><span class="sxs-lookup"><span data-stu-id="fa854-1463">1.3</span></span>|
|[<span data-ttu-id="fa854-1464">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="fa854-1464">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="fa854-1465">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="fa854-1465">ReadWriteItem</span></span>|
|[<span data-ttu-id="fa854-1466">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="fa854-1466">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="fa854-1467">Создание</span><span class="sxs-lookup"><span data-stu-id="fa854-1467">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="fa854-1468">Примеры</span><span class="sxs-lookup"><span data-stu-id="fa854-1468">Examples</span></span>

```javascript
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result
  });
```

<span data-ttu-id="fa854-p185">Ниже приведен пример параметра `result`, переданного функции обратного вызова. Свойство `value` содержит идентификатор элемента.</span><span class="sxs-lookup"><span data-stu-id="fa854-p185">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

####  <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="fa854-1471">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="fa854-1471">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="fa854-1472">Асинхронно вставляет данные в текст или тему сообщения.</span><span class="sxs-lookup"><span data-stu-id="fa854-1472">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="fa854-p186">Метод `setSelectedDataAsync` вставляет указанную строку в местоположение курсора в теме или тексте элемента либо, если текст выделен в редакторе, он заменяет выделенный текст. Если курсор находится вне текста или темы элемента, возвращается ошибка. После вставки курсор помещается в конец вставленного содержимого.</span><span class="sxs-lookup"><span data-stu-id="fa854-p186">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="fa854-1476">Параметры:</span><span class="sxs-lookup"><span data-stu-id="fa854-1476">Parameters:</span></span>

|<span data-ttu-id="fa854-1477">Имя</span><span class="sxs-lookup"><span data-stu-id="fa854-1477">Name</span></span>|<span data-ttu-id="fa854-1478">Тип</span><span class="sxs-lookup"><span data-stu-id="fa854-1478">Type</span></span>|<span data-ttu-id="fa854-1479">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="fa854-1479">Attributes</span></span>|<span data-ttu-id="fa854-1480">Описание</span><span class="sxs-lookup"><span data-stu-id="fa854-1480">Description</span></span>|
|---|---|---|---|
|`data`|<span data-ttu-id="fa854-1481">String</span><span class="sxs-lookup"><span data-stu-id="fa854-1481">String</span></span>||<span data-ttu-id="fa854-p187">Вставляемые данные. Объем данных не должен превышать 1 000 000 символов. Если передано больше 1 000 000 символов, возвращается исключение `ArgumentOutOfRange`.</span><span class="sxs-lookup"><span data-stu-id="fa854-p187">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`|<span data-ttu-id="fa854-1485">Object</span><span class="sxs-lookup"><span data-stu-id="fa854-1485">Object</span></span>|<span data-ttu-id="fa854-1486">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="fa854-1486">&lt;optional&gt;</span></span>|<span data-ttu-id="fa854-1487">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="fa854-1487">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="fa854-1488">Object</span><span class="sxs-lookup"><span data-stu-id="fa854-1488">Object</span></span>|<span data-ttu-id="fa854-1489">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="fa854-1489">&lt;optional&gt;</span></span>|<span data-ttu-id="fa854-1490">В методе обратного вызова разработчики могут указать любой объект, к которому необходимо получить доступ.</span><span class="sxs-lookup"><span data-stu-id="fa854-1490">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="fa854-1491">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="fa854-1491">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="fa854-1492">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="fa854-1492">&lt;optional&gt;</span></span>|<span data-ttu-id="fa854-p188">Если задано значение `text`, текущий стиль применяется в Outlook Web App и Outlook. Если поле представляет собой редактор HTML, вставляются только текстовые данные, даже если они имеют формат HTML.</span><span class="sxs-lookup"><span data-stu-id="fa854-p188">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="fa854-p189">Если задано значение `html` и поле (не тема) поддерживает HTML, в Outlook Web App применяется текущий стиль, а в Outlook — стиль по умолчанию. Если поле является текстовым, возвращается ошибка `InvalidDataFormat`.</span><span class="sxs-lookup"><span data-stu-id="fa854-p189">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="fa854-1497">Если свойство `coercionType` не задано, результат зависит от поля: если поле имеет формат HTML, используется текст в формате HTML, а если поле текстовое, применяется обычный текст.</span><span class="sxs-lookup"><span data-stu-id="fa854-1497">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`|<span data-ttu-id="fa854-1498">функция</span><span class="sxs-lookup"><span data-stu-id="fa854-1498">function</span></span>||<span data-ttu-id="fa854-1499">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="fa854-1499">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="fa854-1500">Требования</span><span class="sxs-lookup"><span data-stu-id="fa854-1500">Requirements</span></span>

|<span data-ttu-id="fa854-1501">Требование</span><span class="sxs-lookup"><span data-stu-id="fa854-1501">Requirement</span></span>|<span data-ttu-id="fa854-1502">Значение</span><span class="sxs-lookup"><span data-stu-id="fa854-1502">Value</span></span>|
|---|---|
|[<span data-ttu-id="fa854-1503">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="fa854-1503">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="fa854-1504">1.2</span><span class="sxs-lookup"><span data-stu-id="fa854-1504">1.2</span></span>|
|[<span data-ttu-id="fa854-1505">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="fa854-1505">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="fa854-1506">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="fa854-1506">ReadWriteItem</span></span>|
|[<span data-ttu-id="fa854-1507">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="fa854-1507">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="fa854-1508">Создание</span><span class="sxs-lookup"><span data-stu-id="fa854-1508">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="fa854-1509">Пример</span><span class="sxs-lookup"><span data-stu-id="fa854-1509">Example</span></span>

```javascript
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
