---
title: Office. Context. Mailbox. Item — Предварительная версия набора требований
description: ''
ms.date: 06/03/2019
localization_priority: Normal
ms.openlocfilehash: 3dad9133fb23f6190e58eab94dc1724c18ac9d40
ms.sourcegitcommit: 567aa05d6ee6b3639f65c50188df2331b7685857
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/04/2019
ms.locfileid: "34706360"
---
# <a name="item"></a><span data-ttu-id="0b5b9-102">item</span><span class="sxs-lookup"><span data-stu-id="0b5b9-102">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="0b5b9-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="0b5b9-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="0b5b9-p101">Пространство имен `item` используется для доступа к выбранному в данный момент сообщению, приглашению на собрание или описанию встречи. Вы можете определить тип пространства имен `item` с помощью свойства [itemType](#itemtype-officemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="0b5b9-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="0b5b9-106">Requirements</span><span class="sxs-lookup"><span data-stu-id="0b5b9-106">Requirements</span></span>

|<span data-ttu-id="0b5b9-107">Требование</span><span class="sxs-lookup"><span data-stu-id="0b5b9-107">Requirement</span></span>|<span data-ttu-id="0b5b9-108">Значение</span><span class="sxs-lookup"><span data-stu-id="0b5b9-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="0b5b9-109">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="0b5b9-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0b5b9-110">1.0</span><span class="sxs-lookup"><span data-stu-id="0b5b9-110">1.0</span></span>|
|[<span data-ttu-id="0b5b9-111">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="0b5b9-111">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0b5b9-112">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="0b5b9-112">Restricted</span></span>|
|[<span data-ttu-id="0b5b9-113">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="0b5b9-113">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0b5b9-114">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="0b5b9-114">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="0b5b9-115">Элементы и методы</span><span class="sxs-lookup"><span data-stu-id="0b5b9-115">Members and methods</span></span>

| <span data-ttu-id="0b5b9-116">Элемент</span><span class="sxs-lookup"><span data-stu-id="0b5b9-116">Member</span></span> | <span data-ttu-id="0b5b9-117">Тип</span><span class="sxs-lookup"><span data-stu-id="0b5b9-117">Type</span></span> |
|--------|------|
| [<span data-ttu-id="0b5b9-118">attachments</span><span class="sxs-lookup"><span data-stu-id="0b5b9-118">attachments</span></span>](#attachments-arrayattachmentdetails) | <span data-ttu-id="0b5b9-119">Элемент</span><span class="sxs-lookup"><span data-stu-id="0b5b9-119">Member</span></span> |
| [<span data-ttu-id="0b5b9-120">bcc</span><span class="sxs-lookup"><span data-stu-id="0b5b9-120">bcc</span></span>](#bcc-recipients) | <span data-ttu-id="0b5b9-121">Элемент</span><span class="sxs-lookup"><span data-stu-id="0b5b9-121">Member</span></span> |
| [<span data-ttu-id="0b5b9-122">body</span><span class="sxs-lookup"><span data-stu-id="0b5b9-122">body</span></span>](#body-body) | <span data-ttu-id="0b5b9-123">Элемент</span><span class="sxs-lookup"><span data-stu-id="0b5b9-123">Member</span></span> |
| [<span data-ttu-id="0b5b9-124">разделов</span><span class="sxs-lookup"><span data-stu-id="0b5b9-124">categories</span></span>](#categories-categories) | <span data-ttu-id="0b5b9-125">Элемент</span><span class="sxs-lookup"><span data-stu-id="0b5b9-125">Member</span></span> |
| [<span data-ttu-id="0b5b9-126">cc</span><span class="sxs-lookup"><span data-stu-id="0b5b9-126">cc</span></span>](#cc-arrayemailaddressdetailsrecipients) | <span data-ttu-id="0b5b9-127">Элемент</span><span class="sxs-lookup"><span data-stu-id="0b5b9-127">Member</span></span> |
| [<span data-ttu-id="0b5b9-128">conversationId</span><span class="sxs-lookup"><span data-stu-id="0b5b9-128">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="0b5b9-129">Элемент</span><span class="sxs-lookup"><span data-stu-id="0b5b9-129">Member</span></span> |
| [<span data-ttu-id="0b5b9-130">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="0b5b9-130">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="0b5b9-131">Элемент</span><span class="sxs-lookup"><span data-stu-id="0b5b9-131">Member</span></span> |
| [<span data-ttu-id="0b5b9-132">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="0b5b9-132">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="0b5b9-133">Элемент</span><span class="sxs-lookup"><span data-stu-id="0b5b9-133">Member</span></span> |
| [<span data-ttu-id="0b5b9-134">end</span><span class="sxs-lookup"><span data-stu-id="0b5b9-134">end</span></span>](#end-datetime) | <span data-ttu-id="0b5b9-135">Элемент</span><span class="sxs-lookup"><span data-stu-id="0b5b9-135">Member</span></span> |
| [<span data-ttu-id="0b5b9-136">Енханцедлокатион</span><span class="sxs-lookup"><span data-stu-id="0b5b9-136">enhancedLocation</span></span>](#enhancedlocation-enhancedlocation) | <span data-ttu-id="0b5b9-137">Элемент</span><span class="sxs-lookup"><span data-stu-id="0b5b9-137">Member</span></span> |
| [<span data-ttu-id="0b5b9-138">from</span><span class="sxs-lookup"><span data-stu-id="0b5b9-138">from</span></span>](#from-emailaddressdetailsfrom) | <span data-ttu-id="0b5b9-139">Элемент</span><span class="sxs-lookup"><span data-stu-id="0b5b9-139">Member</span></span> |
| [<span data-ttu-id="0b5b9-140">Internetheaders:</span><span class="sxs-lookup"><span data-stu-id="0b5b9-140">internetHeaders</span></span>](#internetheaders-internetheaders) | <span data-ttu-id="0b5b9-141">Элемент</span><span class="sxs-lookup"><span data-stu-id="0b5b9-141">Member</span></span> |
| [<span data-ttu-id="0b5b9-142">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="0b5b9-142">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="0b5b9-143">Элемент</span><span class="sxs-lookup"><span data-stu-id="0b5b9-143">Member</span></span> |
| [<span data-ttu-id="0b5b9-144">itemClass</span><span class="sxs-lookup"><span data-stu-id="0b5b9-144">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="0b5b9-145">Элемент</span><span class="sxs-lookup"><span data-stu-id="0b5b9-145">Member</span></span> |
| [<span data-ttu-id="0b5b9-146">itemId</span><span class="sxs-lookup"><span data-stu-id="0b5b9-146">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="0b5b9-147">Элемент</span><span class="sxs-lookup"><span data-stu-id="0b5b9-147">Member</span></span> |
| [<span data-ttu-id="0b5b9-148">itemType</span><span class="sxs-lookup"><span data-stu-id="0b5b9-148">itemType</span></span>](#itemtype-officemailboxenumsitemtype) | <span data-ttu-id="0b5b9-149">Элемент</span><span class="sxs-lookup"><span data-stu-id="0b5b9-149">Member</span></span> |
| [<span data-ttu-id="0b5b9-150">location</span><span class="sxs-lookup"><span data-stu-id="0b5b9-150">location</span></span>](#location-stringlocation) | <span data-ttu-id="0b5b9-151">Элемент</span><span class="sxs-lookup"><span data-stu-id="0b5b9-151">Member</span></span> |
| [<span data-ttu-id="0b5b9-152">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="0b5b9-152">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="0b5b9-153">Элемент</span><span class="sxs-lookup"><span data-stu-id="0b5b9-153">Member</span></span> |
| [<span data-ttu-id="0b5b9-154">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="0b5b9-154">notificationMessages</span></span>](#notificationmessages-notificationmessages) | <span data-ttu-id="0b5b9-155">Member</span><span class="sxs-lookup"><span data-stu-id="0b5b9-155">Member</span></span> |
| [<span data-ttu-id="0b5b9-156">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="0b5b9-156">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="0b5b9-157">Элемент</span><span class="sxs-lookup"><span data-stu-id="0b5b9-157">Member</span></span> |
| [<span data-ttu-id="0b5b9-158">organizer</span><span class="sxs-lookup"><span data-stu-id="0b5b9-158">organizer</span></span>](#organizer-emailaddressdetailsorganizer) | <span data-ttu-id="0b5b9-159">Элемент</span><span class="sxs-lookup"><span data-stu-id="0b5b9-159">Member</span></span> |
| [<span data-ttu-id="0b5b9-160">recurrence</span><span class="sxs-lookup"><span data-stu-id="0b5b9-160">recurrence</span></span>](#nullable-recurrence-recurrence) | <span data-ttu-id="0b5b9-161">Элемент</span><span class="sxs-lookup"><span data-stu-id="0b5b9-161">Member</span></span> |
| [<span data-ttu-id="0b5b9-162">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="0b5b9-162">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="0b5b9-163">Элемент</span><span class="sxs-lookup"><span data-stu-id="0b5b9-163">Member</span></span> |
| [<span data-ttu-id="0b5b9-164">sender</span><span class="sxs-lookup"><span data-stu-id="0b5b9-164">sender</span></span>](#sender-emailaddressdetails) | <span data-ttu-id="0b5b9-165">Member</span><span class="sxs-lookup"><span data-stu-id="0b5b9-165">Member</span></span> |
| [<span data-ttu-id="0b5b9-166">seriesId</span><span class="sxs-lookup"><span data-stu-id="0b5b9-166">seriesId</span></span>](#nullable-seriesid-string) | <span data-ttu-id="0b5b9-167">Member</span><span class="sxs-lookup"><span data-stu-id="0b5b9-167">Member</span></span> |
| [<span data-ttu-id="0b5b9-168">start</span><span class="sxs-lookup"><span data-stu-id="0b5b9-168">start</span></span>](#start-datetime) | <span data-ttu-id="0b5b9-169">Member</span><span class="sxs-lookup"><span data-stu-id="0b5b9-169">Member</span></span> |
| [<span data-ttu-id="0b5b9-170">subject</span><span class="sxs-lookup"><span data-stu-id="0b5b9-170">subject</span></span>](#subject-stringsubject) | <span data-ttu-id="0b5b9-171">Элемент</span><span class="sxs-lookup"><span data-stu-id="0b5b9-171">Member</span></span> |
| [<span data-ttu-id="0b5b9-172">to</span><span class="sxs-lookup"><span data-stu-id="0b5b9-172">to</span></span>](#to-arrayemailaddressdetailsrecipients) | <span data-ttu-id="0b5b9-173">Элемент</span><span class="sxs-lookup"><span data-stu-id="0b5b9-173">Member</span></span> |
| [<span data-ttu-id="0b5b9-174">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="0b5b9-174">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="0b5b9-175">Метод</span><span class="sxs-lookup"><span data-stu-id="0b5b9-175">Method</span></span> |
| [<span data-ttu-id="0b5b9-176">addFileAttachmentFromBase64Async</span><span class="sxs-lookup"><span data-stu-id="0b5b9-176">addFileAttachmentFromBase64Async</span></span>](#addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback) | <span data-ttu-id="0b5b9-177">Метод</span><span class="sxs-lookup"><span data-stu-id="0b5b9-177">Method</span></span> |
| [<span data-ttu-id="0b5b9-178">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="0b5b9-178">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="0b5b9-179">Метод</span><span class="sxs-lookup"><span data-stu-id="0b5b9-179">Method</span></span> |
| [<span data-ttu-id="0b5b9-180">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="0b5b9-180">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="0b5b9-181">Метод</span><span class="sxs-lookup"><span data-stu-id="0b5b9-181">Method</span></span> |
| [<span data-ttu-id="0b5b9-182">close</span><span class="sxs-lookup"><span data-stu-id="0b5b9-182">close</span></span>](#close) | <span data-ttu-id="0b5b9-183">Метод</span><span class="sxs-lookup"><span data-stu-id="0b5b9-183">Method</span></span> |
| [<span data-ttu-id="0b5b9-184">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="0b5b9-184">displayReplyAllForm</span></span>](#displayreplyallformformdata-callback) | <span data-ttu-id="0b5b9-185">Метод</span><span class="sxs-lookup"><span data-stu-id="0b5b9-185">Method</span></span> |
| [<span data-ttu-id="0b5b9-186">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="0b5b9-186">displayReplyForm</span></span>](#displayreplyformformdata-callback) | <span data-ttu-id="0b5b9-187">Метод</span><span class="sxs-lookup"><span data-stu-id="0b5b9-187">Method</span></span> |
| [<span data-ttu-id="0b5b9-188">Жетаттачментконтентасинк</span><span class="sxs-lookup"><span data-stu-id="0b5b9-188">getAttachmentContentAsync</span></span>](#getattachmentcontentasyncattachmentid-options-callback--attachmentcontent) | <span data-ttu-id="0b5b9-189">Метод</span><span class="sxs-lookup"><span data-stu-id="0b5b9-189">Method</span></span> |
| [<span data-ttu-id="0b5b9-190">Жетаттачментсасинк</span><span class="sxs-lookup"><span data-stu-id="0b5b9-190">getAttachmentsAsync</span></span>](#getattachmentsasyncoptions-callback--arrayattachmentdetails) | <span data-ttu-id="0b5b9-191">Метод</span><span class="sxs-lookup"><span data-stu-id="0b5b9-191">Method</span></span> |
| [<span data-ttu-id="0b5b9-192">getEntities</span><span class="sxs-lookup"><span data-stu-id="0b5b9-192">getEntities</span></span>](#getentities--entities) | <span data-ttu-id="0b5b9-193">Метод</span><span class="sxs-lookup"><span data-stu-id="0b5b9-193">Method</span></span> |
| [<span data-ttu-id="0b5b9-194">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="0b5b9-194">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="0b5b9-195">Метод</span><span class="sxs-lookup"><span data-stu-id="0b5b9-195">Method</span></span> |
| [<span data-ttu-id="0b5b9-196">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="0b5b9-196">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="0b5b9-197">Метод</span><span class="sxs-lookup"><span data-stu-id="0b5b9-197">Method</span></span> |
| [<span data-ttu-id="0b5b9-198">getInitializationContextAsync</span><span class="sxs-lookup"><span data-stu-id="0b5b9-198">getInitializationContextAsync</span></span>](#getinitializationcontextasyncoptions-callback) | <span data-ttu-id="0b5b9-199">Метод</span><span class="sxs-lookup"><span data-stu-id="0b5b9-199">Method</span></span> |
| [<span data-ttu-id="0b5b9-200">Жетитемидасинк</span><span class="sxs-lookup"><span data-stu-id="0b5b9-200">getItemIdAsync</span></span>](#getitemidasyncoptions-callback) | <span data-ttu-id="0b5b9-201">Метод</span><span class="sxs-lookup"><span data-stu-id="0b5b9-201">Method</span></span> |
| [<span data-ttu-id="0b5b9-202">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="0b5b9-202">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="0b5b9-203">Метод</span><span class="sxs-lookup"><span data-stu-id="0b5b9-203">Method</span></span> |
| [<span data-ttu-id="0b5b9-204">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="0b5b9-204">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="0b5b9-205">Метод</span><span class="sxs-lookup"><span data-stu-id="0b5b9-205">Method</span></span> |
| [<span data-ttu-id="0b5b9-206">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="0b5b9-206">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="0b5b9-207">Метод</span><span class="sxs-lookup"><span data-stu-id="0b5b9-207">Method</span></span> |
| [<span data-ttu-id="0b5b9-208">Жетселектедентитиес</span><span class="sxs-lookup"><span data-stu-id="0b5b9-208">getSelectedEntities</span></span>](#getselectedentities--entities) | <span data-ttu-id="0b5b9-209">Метод</span><span class="sxs-lookup"><span data-stu-id="0b5b9-209">Method</span></span> |
| [<span data-ttu-id="0b5b9-210">Жетселектедрежексматчес</span><span class="sxs-lookup"><span data-stu-id="0b5b9-210">getSelectedRegExMatches</span></span>](#getselectedregexmatches--object) | <span data-ttu-id="0b5b9-211">Метод</span><span class="sxs-lookup"><span data-stu-id="0b5b9-211">Method</span></span> |
| [<span data-ttu-id="0b5b9-212">Жетшаредпропертиесасинк</span><span class="sxs-lookup"><span data-stu-id="0b5b9-212">getSharedPropertiesAsync</span></span>](#getsharedpropertiesasyncoptions-callback) | <span data-ttu-id="0b5b9-213">Метод</span><span class="sxs-lookup"><span data-stu-id="0b5b9-213">Method</span></span> |
| [<span data-ttu-id="0b5b9-214">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="0b5b9-214">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="0b5b9-215">Метод</span><span class="sxs-lookup"><span data-stu-id="0b5b9-215">Method</span></span> |
| [<span data-ttu-id="0b5b9-216">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="0b5b9-216">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="0b5b9-217">Метод</span><span class="sxs-lookup"><span data-stu-id="0b5b9-217">Method</span></span> |
| [<span data-ttu-id="0b5b9-218">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="0b5b9-218">removeHandlerAsync</span></span>](#removehandlerasynceventtype-options-callback) | <span data-ttu-id="0b5b9-219">Метод</span><span class="sxs-lookup"><span data-stu-id="0b5b9-219">Method</span></span> |
| [<span data-ttu-id="0b5b9-220">saveAsync</span><span class="sxs-lookup"><span data-stu-id="0b5b9-220">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="0b5b9-221">Метод</span><span class="sxs-lookup"><span data-stu-id="0b5b9-221">Method</span></span> |
| [<span data-ttu-id="0b5b9-222">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="0b5b9-222">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="0b5b9-223">Метод</span><span class="sxs-lookup"><span data-stu-id="0b5b9-223">Method</span></span> |

### <a name="example"></a><span data-ttu-id="0b5b9-224">Пример</span><span class="sxs-lookup"><span data-stu-id="0b5b9-224">Example</span></span>

<span data-ttu-id="0b5b9-225">В примере кода JavaScript, приведенном ниже, показано, как получить доступ к свойству `subject` текущего элемента в Outlook.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-225">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="0b5b9-226">Элементы</span><span class="sxs-lookup"><span data-stu-id="0b5b9-226">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails"></a><span data-ttu-id="0b5b9-227">вложения: Array. <[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="0b5b9-227">attachments: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

<span data-ttu-id="0b5b9-228">Получает вложения элемента в виде массива.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-228">Gets the item's attachments as an array.</span></span> <span data-ttu-id="0b5b9-229">Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-229">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="0b5b9-230">Outlook блокирует определенные типы файлов из-за потенциальных проблем с безопасностью, поэтому они не возвращаются.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-230">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="0b5b9-231">Дополнительные сведения см. в статье [Блокировка вложений в Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="0b5b9-231">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="0b5b9-232">Тип</span><span class="sxs-lookup"><span data-stu-id="0b5b9-232">Type</span></span>

*   <span data-ttu-id="0b5b9-233">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="0b5b9-233">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="0b5b9-234">Требования</span><span class="sxs-lookup"><span data-stu-id="0b5b9-234">Requirements</span></span>

|<span data-ttu-id="0b5b9-235">Требование</span><span class="sxs-lookup"><span data-stu-id="0b5b9-235">Requirement</span></span>|<span data-ttu-id="0b5b9-236">Значение</span><span class="sxs-lookup"><span data-stu-id="0b5b9-236">Value</span></span>|
|---|---|
|[<span data-ttu-id="0b5b9-237">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="0b5b9-237">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0b5b9-238">1.0</span><span class="sxs-lookup"><span data-stu-id="0b5b9-238">1.0</span></span>|
|[<span data-ttu-id="0b5b9-239">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="0b5b9-239">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0b5b9-240">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0b5b9-240">ReadItem</span></span>|
|[<span data-ttu-id="0b5b9-241">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="0b5b9-241">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0b5b9-242">Чтение</span><span class="sxs-lookup"><span data-stu-id="0b5b9-242">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0b5b9-243">Пример</span><span class="sxs-lookup"><span data-stu-id="0b5b9-243">Example</span></span>

<span data-ttu-id="0b5b9-244">С помощью приведенного ниже кода можно создать HTML-строку с подробными сведениями обо всех вложениях для текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-244">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

#### <a name="bcc-recipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="0b5b9-245">СК: [получатели](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="0b5b9-245">bcc: [Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="0b5b9-246">Получает объект, который предоставляет методы для получения или обновления получателей скрытой копии сообщения.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-246">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="0b5b9-247">Только в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-247">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="0b5b9-248">Тип</span><span class="sxs-lookup"><span data-stu-id="0b5b9-248">Type</span></span>

*   [<span data-ttu-id="0b5b9-249">Получатели</span><span class="sxs-lookup"><span data-stu-id="0b5b9-249">Recipients</span></span>](/javascript/api/outlook/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="0b5b9-250">Требования</span><span class="sxs-lookup"><span data-stu-id="0b5b9-250">Requirements</span></span>

|<span data-ttu-id="0b5b9-251">Требование</span><span class="sxs-lookup"><span data-stu-id="0b5b9-251">Requirement</span></span>|<span data-ttu-id="0b5b9-252">Значение</span><span class="sxs-lookup"><span data-stu-id="0b5b9-252">Value</span></span>|
|---|---|
|[<span data-ttu-id="0b5b9-253">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="0b5b9-253">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0b5b9-254">1.1</span><span class="sxs-lookup"><span data-stu-id="0b5b9-254">1.1</span></span>|
|[<span data-ttu-id="0b5b9-255">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="0b5b9-255">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0b5b9-256">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0b5b9-256">ReadItem</span></span>|
|[<span data-ttu-id="0b5b9-257">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="0b5b9-257">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0b5b9-258">Создание</span><span class="sxs-lookup"><span data-stu-id="0b5b9-258">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="0b5b9-259">Пример</span><span class="sxs-lookup"><span data-stu-id="0b5b9-259">Example</span></span>

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

#### <a name="body-bodyjavascriptapioutlookofficebody"></a><span data-ttu-id="0b5b9-260">основной текст: [Body](/javascript/api/outlook/office.body)</span><span class="sxs-lookup"><span data-stu-id="0b5b9-260">body: [Body](/javascript/api/outlook/office.body)</span></span>

<span data-ttu-id="0b5b9-261">Получает объект, предоставляющий методы для работы с основным текстом элемента.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-261">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="0b5b9-262">Тип</span><span class="sxs-lookup"><span data-stu-id="0b5b9-262">Type</span></span>

*   [<span data-ttu-id="0b5b9-263">Body</span><span class="sxs-lookup"><span data-stu-id="0b5b9-263">Body</span></span>](/javascript/api/outlook/office.body)

##### <a name="requirements"></a><span data-ttu-id="0b5b9-264">Требования</span><span class="sxs-lookup"><span data-stu-id="0b5b9-264">Requirements</span></span>

|<span data-ttu-id="0b5b9-265">Требование</span><span class="sxs-lookup"><span data-stu-id="0b5b9-265">Requirement</span></span>|<span data-ttu-id="0b5b9-266">Значение</span><span class="sxs-lookup"><span data-stu-id="0b5b9-266">Value</span></span>|
|---|---|
|[<span data-ttu-id="0b5b9-267">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="0b5b9-267">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0b5b9-268">1.1</span><span class="sxs-lookup"><span data-stu-id="0b5b9-268">1.1</span></span>|
|[<span data-ttu-id="0b5b9-269">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="0b5b9-269">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0b5b9-270">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0b5b9-270">ReadItem</span></span>|
|[<span data-ttu-id="0b5b9-271">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="0b5b9-271">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0b5b9-272">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="0b5b9-272">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0b5b9-273">Пример</span><span class="sxs-lookup"><span data-stu-id="0b5b9-273">Example</span></span>

<span data-ttu-id="0b5b9-274">В этом примере возвращается текст сообщения в формате обычного текста.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-274">This example gets the body of the message in plain text.</span></span>

```javascript
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="0b5b9-275">Ниже приведен пример итогового параметра, переданного функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-275">The following is an example of the result parameter passed to the callback function.</span></span>

```json
{
  "value": "TEXT of whole body (including threads below)",
  "status": "succeeded",
  "asyncContext": "This is passed to the callback"
}
```

---
---

#### <a name="categories-categoriesjavascriptapioutlookofficecategories"></a><span data-ttu-id="0b5b9-276">Категории: [категории](/javascript/api/outlook/office.categories)</span><span class="sxs-lookup"><span data-stu-id="0b5b9-276">categories: [Categories](/javascript/api/outlook/office.categories)</span></span>

<span data-ttu-id="0b5b9-277">Получает объект, предоставляющий методы для управления категориями элемента.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-277">Gets an object that provides methods for managing the item's categories.</span></span>

> [!NOTE]
> <span data-ttu-id="0b5b9-278">Этот элемент не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-278">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="0b5b9-279">Тип</span><span class="sxs-lookup"><span data-stu-id="0b5b9-279">Type</span></span>

*   [<span data-ttu-id="0b5b9-280">Categories</span><span class="sxs-lookup"><span data-stu-id="0b5b9-280">Categories</span></span>](/javascript/api/outlook/office.categories)

##### <a name="requirements"></a><span data-ttu-id="0b5b9-281">Требования</span><span class="sxs-lookup"><span data-stu-id="0b5b9-281">Requirements</span></span>

|<span data-ttu-id="0b5b9-282">Требование</span><span class="sxs-lookup"><span data-stu-id="0b5b9-282">Requirement</span></span>|<span data-ttu-id="0b5b9-283">Значение</span><span class="sxs-lookup"><span data-stu-id="0b5b9-283">Value</span></span>|
|---|---|
|[<span data-ttu-id="0b5b9-284">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="0b5b9-284">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0b5b9-285">Предварительная версия</span><span class="sxs-lookup"><span data-stu-id="0b5b9-285">Preview</span></span>|
|[<span data-ttu-id="0b5b9-286">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="0b5b9-286">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0b5b9-287">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0b5b9-287">ReadItem</span></span>|
|[<span data-ttu-id="0b5b9-288">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="0b5b9-288">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0b5b9-289">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="0b5b9-289">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0b5b9-290">Пример</span><span class="sxs-lookup"><span data-stu-id="0b5b9-290">Example</span></span>

<span data-ttu-id="0b5b9-291">В этом примере возвращаются категории элемента.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-291">This example gets the item's categories.</span></span>

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

#### <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="0b5b9-292">CC: Array. <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[получатели](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="0b5b9-292">cc: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="0b5b9-293">Предоставляет доступ к получателям копии сообщения.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-293">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="0b5b9-294">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-294">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="0b5b9-295">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="0b5b9-295">Read mode</span></span>

<span data-ttu-id="0b5b9-p106">Свойство `cc` возвращает массив, который содержит объект `EmailAddressDetails` для каждого получателя, указанного в строке **Копия** сообщения. Коллекция может включать не более 100 элементов.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="0b5b9-298">Режим создания</span><span class="sxs-lookup"><span data-stu-id="0b5b9-298">Compose mode</span></span>

<span data-ttu-id="0b5b9-299">Свойство `cc` возвращает объект `Recipients`, предоставляющий методы для получения или обновления получателей, которые указаны в строке **Копия** сообщения.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-299">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

```javascript
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="0b5b9-300">Тип</span><span class="sxs-lookup"><span data-stu-id="0b5b9-300">Type</span></span>

*   <span data-ttu-id="0b5b9-301">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="0b5b9-301">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="0b5b9-302">Требования</span><span class="sxs-lookup"><span data-stu-id="0b5b9-302">Requirements</span></span>

|<span data-ttu-id="0b5b9-303">Требование</span><span class="sxs-lookup"><span data-stu-id="0b5b9-303">Requirement</span></span>|<span data-ttu-id="0b5b9-304">Значение</span><span class="sxs-lookup"><span data-stu-id="0b5b9-304">Value</span></span>|
|---|---|
|[<span data-ttu-id="0b5b9-305">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="0b5b9-305">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0b5b9-306">1.0</span><span class="sxs-lookup"><span data-stu-id="0b5b9-306">1.0</span></span>|
|[<span data-ttu-id="0b5b9-307">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="0b5b9-307">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0b5b9-308">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0b5b9-308">ReadItem</span></span>|
|[<span data-ttu-id="0b5b9-309">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="0b5b9-309">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0b5b9-310">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="0b5b9-310">Compose or Read</span></span>|

---
---

#### <a name="nullable-conversationid-string"></a><span data-ttu-id="0b5b9-311">(Nullable) conversationId: строка</span><span class="sxs-lookup"><span data-stu-id="0b5b9-311">(nullable) conversationId: String</span></span>

<span data-ttu-id="0b5b9-312">Получает идентификатор разговора по электронной почте, содержащего конкретное сообщение.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-312">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="0b5b9-p107">Вы можете получить целочисленное значение этого свойства, если ваше почтовое приложение активируется в формах просмотра или формах создания ответов. Если пользователь изменит тему ответа, после его отправки идентификатор беседы будет изменен, и полученное ранее значение будет недействительным.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="0b5b9-p108">Это свойство имеет значение NULL для нового элемента в форме создания. Свойство `conversationId` вернет значение, если пользователь задаст тему и сохранит элемент.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="0b5b9-317">Тип</span><span class="sxs-lookup"><span data-stu-id="0b5b9-317">Type</span></span>

*   <span data-ttu-id="0b5b9-318">String</span><span class="sxs-lookup"><span data-stu-id="0b5b9-318">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="0b5b9-319">Требования</span><span class="sxs-lookup"><span data-stu-id="0b5b9-319">Requirements</span></span>

|<span data-ttu-id="0b5b9-320">Требование</span><span class="sxs-lookup"><span data-stu-id="0b5b9-320">Requirement</span></span>|<span data-ttu-id="0b5b9-321">Значение</span><span class="sxs-lookup"><span data-stu-id="0b5b9-321">Value</span></span>|
|---|---|
|[<span data-ttu-id="0b5b9-322">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="0b5b9-322">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0b5b9-323">1.0</span><span class="sxs-lookup"><span data-stu-id="0b5b9-323">1.0</span></span>|
|[<span data-ttu-id="0b5b9-324">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="0b5b9-324">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0b5b9-325">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0b5b9-325">ReadItem</span></span>|
|[<span data-ttu-id="0b5b9-326">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="0b5b9-326">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0b5b9-327">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="0b5b9-327">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0b5b9-328">Пример</span><span class="sxs-lookup"><span data-stu-id="0b5b9-328">Example</span></span>

```javascript
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

---
---

#### <a name="datetimecreated-date"></a><span data-ttu-id="0b5b9-329">dateTimeCreated: Дата</span><span class="sxs-lookup"><span data-stu-id="0b5b9-329">dateTimeCreated: Date</span></span>

<span data-ttu-id="0b5b9-p109">Получает дату и время создания элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="0b5b9-332">Тип</span><span class="sxs-lookup"><span data-stu-id="0b5b9-332">Type</span></span>

*   <span data-ttu-id="0b5b9-333">Дата</span><span class="sxs-lookup"><span data-stu-id="0b5b9-333">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="0b5b9-334">Требования</span><span class="sxs-lookup"><span data-stu-id="0b5b9-334">Requirements</span></span>

|<span data-ttu-id="0b5b9-335">Требование</span><span class="sxs-lookup"><span data-stu-id="0b5b9-335">Requirement</span></span>|<span data-ttu-id="0b5b9-336">Значение</span><span class="sxs-lookup"><span data-stu-id="0b5b9-336">Value</span></span>|
|---|---|
|[<span data-ttu-id="0b5b9-337">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="0b5b9-337">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0b5b9-338">1.0</span><span class="sxs-lookup"><span data-stu-id="0b5b9-338">1.0</span></span>|
|[<span data-ttu-id="0b5b9-339">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="0b5b9-339">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0b5b9-340">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0b5b9-340">ReadItem</span></span>|
|[<span data-ttu-id="0b5b9-341">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="0b5b9-341">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0b5b9-342">Чтение</span><span class="sxs-lookup"><span data-stu-id="0b5b9-342">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0b5b9-343">Пример</span><span class="sxs-lookup"><span data-stu-id="0b5b9-343">Example</span></span>

```javascript
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

---
---

#### <a name="datetimemodified-date"></a><span data-ttu-id="0b5b9-344">dateTimeModified: Дата</span><span class="sxs-lookup"><span data-stu-id="0b5b9-344">dateTimeModified: Date</span></span>

<span data-ttu-id="0b5b9-p110">Получает дату и время последнего изменения элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="0b5b9-347">Этот элемент не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-347">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="0b5b9-348">Тип</span><span class="sxs-lookup"><span data-stu-id="0b5b9-348">Type</span></span>

*   <span data-ttu-id="0b5b9-349">Дата</span><span class="sxs-lookup"><span data-stu-id="0b5b9-349">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="0b5b9-350">Требования</span><span class="sxs-lookup"><span data-stu-id="0b5b9-350">Requirements</span></span>

|<span data-ttu-id="0b5b9-351">Требование</span><span class="sxs-lookup"><span data-stu-id="0b5b9-351">Requirement</span></span>|<span data-ttu-id="0b5b9-352">Значение</span><span class="sxs-lookup"><span data-stu-id="0b5b9-352">Value</span></span>|
|---|---|
|[<span data-ttu-id="0b5b9-353">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="0b5b9-353">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0b5b9-354">1.0</span><span class="sxs-lookup"><span data-stu-id="0b5b9-354">1.0</span></span>|
|[<span data-ttu-id="0b5b9-355">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="0b5b9-355">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0b5b9-356">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0b5b9-356">ReadItem</span></span>|
|[<span data-ttu-id="0b5b9-357">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="0b5b9-357">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0b5b9-358">Чтение</span><span class="sxs-lookup"><span data-stu-id="0b5b9-358">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0b5b9-359">Пример</span><span class="sxs-lookup"><span data-stu-id="0b5b9-359">Example</span></span>

```javascript
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

---
---

#### <a name="end-datetimejavascriptapioutlookofficetime"></a><span data-ttu-id="0b5b9-360">конец: Дата | [Time (время](/javascript/api/outlook/office.time) )</span><span class="sxs-lookup"><span data-stu-id="0b5b9-360">end: Date|[Time](/javascript/api/outlook/office.time)</span></span>

<span data-ttu-id="0b5b9-361">Получает или задает дату и время окончания встречи.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-361">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="0b5b9-p111">Свойство `end` представлено в виде значения даты и времени в формате UTC. Преобразовать значение свойства end в местные значения даты и времени клиента можно с помощью метода [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime).</span><span class="sxs-lookup"><span data-stu-id="0b5b9-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="0b5b9-364">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="0b5b9-364">Read mode</span></span>

<span data-ttu-id="0b5b9-365">Свойство `end` возвращает объект `Date`.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-365">The `end` property returns a `Date` object.</span></span>

```javascript
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="0b5b9-366">Режим создания</span><span class="sxs-lookup"><span data-stu-id="0b5b9-366">Compose mode</span></span>

<span data-ttu-id="0b5b9-367">Свойство `end` возвращает объект `Time`.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-367">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="0b5b9-368">Если вы задаете время окончания с помощью метода [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-), необходимо использовать метод [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) для преобразования местного времени на клиенте в формат UTC для сервера.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-368">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="0b5b9-369">В примере ниже показано, как с помощью метода [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) объекта `Time` задать время окончания встречи.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-369">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="0b5b9-370">Тип</span><span class="sxs-lookup"><span data-stu-id="0b5b9-370">Type</span></span>

*   <span data-ttu-id="0b5b9-371">Date | [Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="0b5b9-371">Date | [Time](/javascript/api/outlook/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="0b5b9-372">Требования</span><span class="sxs-lookup"><span data-stu-id="0b5b9-372">Requirements</span></span>

|<span data-ttu-id="0b5b9-373">Требование</span><span class="sxs-lookup"><span data-stu-id="0b5b9-373">Requirement</span></span>|<span data-ttu-id="0b5b9-374">Значение</span><span class="sxs-lookup"><span data-stu-id="0b5b9-374">Value</span></span>|
|---|---|
|[<span data-ttu-id="0b5b9-375">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="0b5b9-375">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0b5b9-376">1.0</span><span class="sxs-lookup"><span data-stu-id="0b5b9-376">1.0</span></span>|
|[<span data-ttu-id="0b5b9-377">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="0b5b9-377">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0b5b9-378">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0b5b9-378">ReadItem</span></span>|
|[<span data-ttu-id="0b5b9-379">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="0b5b9-379">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0b5b9-380">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="0b5b9-380">Compose or Read</span></span>|

---
---

#### <a name="enhancedlocation-enhancedlocationjavascriptapioutlookofficeenhancedlocation"></a><span data-ttu-id="0b5b9-381">Енханцедлокатион: [енханцедлокатион](/javascript/api/outlook/office.enhancedlocation)</span><span class="sxs-lookup"><span data-stu-id="0b5b9-381">enhancedLocation: [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation)</span></span>

<span data-ttu-id="0b5b9-382">Получает или задает расположение встречи.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-382">Gets or sets the locations of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="0b5b9-383">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="0b5b9-383">Read mode</span></span>

<span data-ttu-id="0b5b9-384">Свойство возвращает объект [енханцедлокатион](/javascript/api/outlook/office.enhancedlocation) , который позволяет получить набор расположений (каждый, представленный объектом локатиондетаилс), связанный с встречей. [](/javascript/api/outlook/office.locationdetails) `enhancedLocation`</span><span class="sxs-lookup"><span data-stu-id="0b5b9-384">The `enhancedLocation` property returns an [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation) object that allows you to get the set of locations (each represented by a [LocationDetails](/javascript/api/outlook/office.locationdetails) object) associated with the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="0b5b9-385">Режим создания</span><span class="sxs-lookup"><span data-stu-id="0b5b9-385">Compose mode</span></span>

<span data-ttu-id="0b5b9-386">`enhancedLocation` Свойство возвращает объект [енханцедлокатион](/javascript/api/outlook/office.enhancedlocation) , который предоставляет методы для получения, удаления или добавления расположений для встречи.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-386">The `enhancedLocation` property returns an [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation) object that provides methods to get, remove, or add locations on an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="0b5b9-387">Тип</span><span class="sxs-lookup"><span data-stu-id="0b5b9-387">Type</span></span>

*   [<span data-ttu-id="0b5b9-388">EnhancedLocation</span><span class="sxs-lookup"><span data-stu-id="0b5b9-388">EnhancedLocation</span></span>](/javascript/api/outlook/office.enhancedlocation)

##### <a name="requirements"></a><span data-ttu-id="0b5b9-389">Требования</span><span class="sxs-lookup"><span data-stu-id="0b5b9-389">Requirements</span></span>

|<span data-ttu-id="0b5b9-390">Требование</span><span class="sxs-lookup"><span data-stu-id="0b5b9-390">Requirement</span></span>|<span data-ttu-id="0b5b9-391">Значение</span><span class="sxs-lookup"><span data-stu-id="0b5b9-391">Value</span></span>|
|---|---|
|[<span data-ttu-id="0b5b9-392">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="0b5b9-392">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0b5b9-393">Предварительная версия</span><span class="sxs-lookup"><span data-stu-id="0b5b9-393">Preview</span></span>|
|[<span data-ttu-id="0b5b9-394">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="0b5b9-394">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0b5b9-395">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0b5b9-395">ReadItem</span></span>|
|[<span data-ttu-id="0b5b9-396">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="0b5b9-396">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0b5b9-397">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="0b5b9-397">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0b5b9-398">Пример</span><span class="sxs-lookup"><span data-stu-id="0b5b9-398">Example</span></span>

<span data-ttu-id="0b5b9-399">В следующем примере показано получение текущих расположений, связанных с встречей.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-399">The following example gets the current locations associated with the appointment.</span></span>

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

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom"></a><span data-ttu-id="0b5b9-400">от: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[](/javascript/api/outlook/office.from)</span><span class="sxs-lookup"><span data-stu-id="0b5b9-400">from: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[From](/javascript/api/outlook/office.from)</span></span>

<span data-ttu-id="0b5b9-401">Получает электронный адрес отправителя сообщения.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-401">Gets the email address of the sender of a message.</span></span>

<span data-ttu-id="0b5b9-p112">Свойства `from` и [`sender`](#sender-emailaddressdetails) представляют одно лицо, если сообщение не отправлено представителем. В противном случае свойство `from` представляет лицо, делегировавшее полномочия, а свойство sender — представителя.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-p112">The `from` and [`sender`](#sender-emailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="0b5b9-404">Свойству `recipientType`, принадлежащему объекту `EmailAddressDetails` в свойстве `from`, задано значение `undefined`.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-404">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="0b5b9-405">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="0b5b9-405">Read mode</span></span>

<span data-ttu-id="0b5b9-406">`from` Свойство возвращает `EmailAddressDetails` объект.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-406">The `from` property returns an `EmailAddressDetails` object.</span></span>

```javascript
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

##### <a name="compose-mode"></a><span data-ttu-id="0b5b9-407">Режим создания</span><span class="sxs-lookup"><span data-stu-id="0b5b9-407">Compose mode</span></span>

<span data-ttu-id="0b5b9-408">`from` Свойство возвращает `From` объект, который предоставляет метод для получения значения From.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-408">The `from` property returns a `From` object that provides a method to get the from value.</span></span>

```javascript
Office.context.mailbox.item.from.getAsync(callback);

function callback(asyncResult) {
  var from = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="0b5b9-409">Тип</span><span class="sxs-lookup"><span data-stu-id="0b5b9-409">Type</span></span>

*   <span data-ttu-id="0b5b9-410">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [из](/javascript/api/outlook/office.from)</span><span class="sxs-lookup"><span data-stu-id="0b5b9-410">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [From](/javascript/api/outlook/office.from)</span></span>

##### <a name="requirements"></a><span data-ttu-id="0b5b9-411">Требования</span><span class="sxs-lookup"><span data-stu-id="0b5b9-411">Requirements</span></span>

|<span data-ttu-id="0b5b9-412">Требование</span><span class="sxs-lookup"><span data-stu-id="0b5b9-412">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="0b5b9-413">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="0b5b9-413">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0b5b9-414">1.0</span><span class="sxs-lookup"><span data-stu-id="0b5b9-414">1.0</span></span>|<span data-ttu-id="0b5b9-415">1.7</span><span class="sxs-lookup"><span data-stu-id="0b5b9-415">1.7</span></span>|
|[<span data-ttu-id="0b5b9-416">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="0b5b9-416">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0b5b9-417">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0b5b9-417">ReadItem</span></span>|<span data-ttu-id="0b5b9-418">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="0b5b9-418">ReadWriteItem</span></span>|
|[<span data-ttu-id="0b5b9-419">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="0b5b9-419">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0b5b9-420">Чтение</span><span class="sxs-lookup"><span data-stu-id="0b5b9-420">Read</span></span>|<span data-ttu-id="0b5b9-421">Создание</span><span class="sxs-lookup"><span data-stu-id="0b5b9-421">Compose</span></span>|

---
---

#### <a name="internetheaders-internetheadersjavascriptapioutlookofficeinternetheaders"></a><span data-ttu-id="0b5b9-422">Internetheaders:: [internetheaders:](/javascript/api/outlook/office.internetheaders)</span><span class="sxs-lookup"><span data-stu-id="0b5b9-422">internetHeaders: [InternetHeaders](/javascript/api/outlook/office.internetheaders)</span></span>

<span data-ttu-id="0b5b9-423">Возвращает или задает заголовки Интернета сообщения.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-423">Gets or sets the internet headers of a message.</span></span>

##### <a name="type"></a><span data-ttu-id="0b5b9-424">Тип</span><span class="sxs-lookup"><span data-stu-id="0b5b9-424">Type</span></span>

*   [<span data-ttu-id="0b5b9-425">InternetHeaders</span><span class="sxs-lookup"><span data-stu-id="0b5b9-425">InternetHeaders</span></span>](/javascript/api/outlook/office.internetheaders)

##### <a name="requirements"></a><span data-ttu-id="0b5b9-426">Требования</span><span class="sxs-lookup"><span data-stu-id="0b5b9-426">Requirements</span></span>

|<span data-ttu-id="0b5b9-427">Требование</span><span class="sxs-lookup"><span data-stu-id="0b5b9-427">Requirement</span></span>|<span data-ttu-id="0b5b9-428">Значение</span><span class="sxs-lookup"><span data-stu-id="0b5b9-428">Value</span></span>|
|---|---|
|[<span data-ttu-id="0b5b9-429">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="0b5b9-429">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0b5b9-430">Предварительная версия</span><span class="sxs-lookup"><span data-stu-id="0b5b9-430">Preview</span></span>|
|[<span data-ttu-id="0b5b9-431">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="0b5b9-431">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0b5b9-432">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0b5b9-432">ReadItem</span></span>|
|[<span data-ttu-id="0b5b9-433">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="0b5b9-433">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0b5b9-434">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="0b5b9-434">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0b5b9-435">Пример</span><span class="sxs-lookup"><span data-stu-id="0b5b9-435">Example</span></span>

```javascript
Office.context.mailbox.item.internetHeaders.getAsync(["header1", "header2"], callback);

function callback(asyncResult) {
  var dictionary = asyncResult.value;
  var header1_value = dictionary["header1"];
}
```

---
---

#### <a name="internetmessageid-string"></a><span data-ttu-id="0b5b9-436">internetMessageId: строка</span><span class="sxs-lookup"><span data-stu-id="0b5b9-436">internetMessageId: String</span></span>

<span data-ttu-id="0b5b9-p113">Получает идентификатор интернет-сообщения для электронного сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-p113">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="0b5b9-439">Тип</span><span class="sxs-lookup"><span data-stu-id="0b5b9-439">Type</span></span>

*   <span data-ttu-id="0b5b9-440">String</span><span class="sxs-lookup"><span data-stu-id="0b5b9-440">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="0b5b9-441">Требования</span><span class="sxs-lookup"><span data-stu-id="0b5b9-441">Requirements</span></span>

|<span data-ttu-id="0b5b9-442">Требование</span><span class="sxs-lookup"><span data-stu-id="0b5b9-442">Requirement</span></span>|<span data-ttu-id="0b5b9-443">Значение</span><span class="sxs-lookup"><span data-stu-id="0b5b9-443">Value</span></span>|
|---|---|
|[<span data-ttu-id="0b5b9-444">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="0b5b9-444">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0b5b9-445">1.0</span><span class="sxs-lookup"><span data-stu-id="0b5b9-445">1.0</span></span>|
|[<span data-ttu-id="0b5b9-446">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="0b5b9-446">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0b5b9-447">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0b5b9-447">ReadItem</span></span>|
|[<span data-ttu-id="0b5b9-448">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="0b5b9-448">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0b5b9-449">Чтение</span><span class="sxs-lookup"><span data-stu-id="0b5b9-449">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0b5b9-450">Пример</span><span class="sxs-lookup"><span data-stu-id="0b5b9-450">Example</span></span>

```javascript
var internetMessageId = Office.context.mailbox.item.internetMessageId;
console.log("internetMessageId: " + internetMessageId);
```

---
---

#### <a name="itemclass-string"></a><span data-ttu-id="0b5b9-451">itemClass: строка</span><span class="sxs-lookup"><span data-stu-id="0b5b9-451">itemClass: String</span></span>

<span data-ttu-id="0b5b9-p114">Получает класс элемента веб-служб Exchange для выбранного элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-p114">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="0b5b9-p115">Свойство `itemClass` указывает класс сообщения выбранного элемента. Ниже приводятся классы сообщения по умолчанию для элемента сообщения или встречи.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-p115">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

|<span data-ttu-id="0b5b9-456">Тип</span><span class="sxs-lookup"><span data-stu-id="0b5b9-456">Type</span></span>|<span data-ttu-id="0b5b9-457">Описание</span><span class="sxs-lookup"><span data-stu-id="0b5b9-457">Description</span></span>|<span data-ttu-id="0b5b9-458">Класс элемента</span><span class="sxs-lookup"><span data-stu-id="0b5b9-458">item class</span></span>|
|---|---|---|
|<span data-ttu-id="0b5b9-459">Элементы встречи</span><span class="sxs-lookup"><span data-stu-id="0b5b9-459">Appointment items</span></span>|<span data-ttu-id="0b5b9-460">Это элементы календаря для класса элемента `IPM.Appointment` или `IPM.Appointment.Occurrence`.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-460">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span>|`IPM.Appointment`<br />`IPM.Appointment.Occurrence`|
|<span data-ttu-id="0b5b9-461">Элементы сообщения</span><span class="sxs-lookup"><span data-stu-id="0b5b9-461">Message items</span></span>|<span data-ttu-id="0b5b9-462">Сюда входят электронные сообщения, для которых по умолчанию задан класс сообщения `IPM.Note`, а также приглашения на собрания, ответы на них и уведомления об их отмене, использующие `IPM.Schedule.Meeting` в качестве базового класса сообщения.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-462">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span>|`IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled`|

<span data-ttu-id="0b5b9-463">Можно создавать настраиваемые классы сообщения, расширяющие классы сообщения по умолчанию, например настраиваемый класс сообщения о встрече `IPM.Appointment.Contoso`.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-463">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="0b5b9-464">Тип</span><span class="sxs-lookup"><span data-stu-id="0b5b9-464">Type</span></span>

*   <span data-ttu-id="0b5b9-465">String</span><span class="sxs-lookup"><span data-stu-id="0b5b9-465">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="0b5b9-466">Требования</span><span class="sxs-lookup"><span data-stu-id="0b5b9-466">Requirements</span></span>

|<span data-ttu-id="0b5b9-467">Требование</span><span class="sxs-lookup"><span data-stu-id="0b5b9-467">Requirement</span></span>|<span data-ttu-id="0b5b9-468">Значение</span><span class="sxs-lookup"><span data-stu-id="0b5b9-468">Value</span></span>|
|---|---|
|[<span data-ttu-id="0b5b9-469">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="0b5b9-469">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0b5b9-470">1.0</span><span class="sxs-lookup"><span data-stu-id="0b5b9-470">1.0</span></span>|
|[<span data-ttu-id="0b5b9-471">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="0b5b9-471">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0b5b9-472">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0b5b9-472">ReadItem</span></span>|
|[<span data-ttu-id="0b5b9-473">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="0b5b9-473">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0b5b9-474">Чтение</span><span class="sxs-lookup"><span data-stu-id="0b5b9-474">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0b5b9-475">Пример</span><span class="sxs-lookup"><span data-stu-id="0b5b9-475">Example</span></span>

```javascript
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

---
---

#### <a name="nullable-itemid-string"></a><span data-ttu-id="0b5b9-476">(Nullable) itemId: строка</span><span class="sxs-lookup"><span data-stu-id="0b5b9-476">(nullable) itemId: String</span></span>

<span data-ttu-id="0b5b9-p116">Получает идентификатор элемента веб-служб Exchange для текущего элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-p116">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="0b5b9-479">Идентификатор, возвращаемый свойством `itemId`, совпадает с идентификатором элемента веб-служб Exchange.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-479">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="0b5b9-480">Свойство `itemId` не совпадает с идентификатором записи Outlook, а также идентификатором, который используется REST API Outlook.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-480">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="0b5b9-481">Прежде чем совершать вызовы REST API, используя это значение, его необходимо преобразовать с помощью [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="0b5b9-481">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="0b5b9-482">Дополнительные сведения см. в статье [Использование REST API Outlook из надстройки Outlook](/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="0b5b9-482">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="0b5b9-p118">Свойство `itemId` недоступно в режиме создания. Если требуется идентификатор элемента, с помощью метода [`saveAsync`](#saveasyncoptions-callback) можно сохранить элемент в хранилище. При этом в параметре [`AsyncResult.value`](/javascript/api/office/office.asyncresult) функции обратного вызова возвращается идентификатор элемента.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-p118">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="0b5b9-485">Тип</span><span class="sxs-lookup"><span data-stu-id="0b5b9-485">Type</span></span>

*   <span data-ttu-id="0b5b9-486">String</span><span class="sxs-lookup"><span data-stu-id="0b5b9-486">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="0b5b9-487">Требования</span><span class="sxs-lookup"><span data-stu-id="0b5b9-487">Requirements</span></span>

|<span data-ttu-id="0b5b9-488">Требование</span><span class="sxs-lookup"><span data-stu-id="0b5b9-488">Requirement</span></span>|<span data-ttu-id="0b5b9-489">Значение</span><span class="sxs-lookup"><span data-stu-id="0b5b9-489">Value</span></span>|
|---|---|
|[<span data-ttu-id="0b5b9-490">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="0b5b9-490">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0b5b9-491">1.0</span><span class="sxs-lookup"><span data-stu-id="0b5b9-491">1.0</span></span>|
|[<span data-ttu-id="0b5b9-492">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="0b5b9-492">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0b5b9-493">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0b5b9-493">ReadItem</span></span>|
|[<span data-ttu-id="0b5b9-494">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="0b5b9-494">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0b5b9-495">Чтение</span><span class="sxs-lookup"><span data-stu-id="0b5b9-495">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0b5b9-496">Пример</span><span class="sxs-lookup"><span data-stu-id="0b5b9-496">Example</span></span>

<span data-ttu-id="0b5b9-p119">Указанный ниже код проверяет наличие идентификатора элемента. Если свойство `itemId` возвращает значение `null` или `undefined`, элемент будет сохранен в хранилище, а из асинхронного результата будет получен идентификатор элемента.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-p119">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

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

#### <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype"></a><span data-ttu-id="0b5b9-499">itemType: [Office. MailboxEnums. ItemType](/javascript/api/outlook/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="0b5b9-499">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="0b5b9-500">Получает тип элемента, который представляет экземпляр.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-500">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="0b5b9-501">Свойство `itemType` возвращает одно из значений перечисления `ItemType`, которое указывает, является ли экземпляр объекта `item` сообщением или встречей.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-501">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="0b5b9-502">Тип</span><span class="sxs-lookup"><span data-stu-id="0b5b9-502">Type</span></span>

*   [<span data-ttu-id="0b5b9-503">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="0b5b9-503">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="0b5b9-504">Требования</span><span class="sxs-lookup"><span data-stu-id="0b5b9-504">Requirements</span></span>

|<span data-ttu-id="0b5b9-505">Требование</span><span class="sxs-lookup"><span data-stu-id="0b5b9-505">Requirement</span></span>|<span data-ttu-id="0b5b9-506">Значение</span><span class="sxs-lookup"><span data-stu-id="0b5b9-506">Value</span></span>|
|---|---|
|[<span data-ttu-id="0b5b9-507">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="0b5b9-507">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0b5b9-508">1.0</span><span class="sxs-lookup"><span data-stu-id="0b5b9-508">1.0</span></span>|
|[<span data-ttu-id="0b5b9-509">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="0b5b9-509">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0b5b9-510">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0b5b9-510">ReadItem</span></span>|
|[<span data-ttu-id="0b5b9-511">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="0b5b9-511">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0b5b9-512">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="0b5b9-512">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0b5b9-513">Пример</span><span class="sxs-lookup"><span data-stu-id="0b5b9-513">Example</span></span>

```javascript
if (Office.context.mailbox.item.itemType === Office.MailboxEnums.ItemType.Message) {
  // Do something.
} else {
  // Do something else.
}
```

---
---

#### <a name="location-stringlocationjavascriptapioutlookofficelocation"></a><span data-ttu-id="0b5b9-514">Местоположение: строка | [Location (расположение](/javascript/api/outlook/office.location) )</span><span class="sxs-lookup"><span data-stu-id="0b5b9-514">location: String|[Location](/javascript/api/outlook/office.location)</span></span>

<span data-ttu-id="0b5b9-515">Получает или задает место встречи.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-515">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="0b5b9-516">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="0b5b9-516">Read mode</span></span>

<span data-ttu-id="0b5b9-517">Свойство `location` возвращает строку, содержащую сведения о месте встречи.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-517">The `location` property returns a string that contains the location of the appointment.</span></span>

```javascript
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="0b5b9-518">Режим создания</span><span class="sxs-lookup"><span data-stu-id="0b5b9-518">Compose mode</span></span>

<span data-ttu-id="0b5b9-519">Свойство `location` возвращает объект `Location`, предоставляющий методы, которые используются для получения и задания места встречи.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-519">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```javascript
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="0b5b9-520">Тип</span><span class="sxs-lookup"><span data-stu-id="0b5b9-520">Type</span></span>

*   <span data-ttu-id="0b5b9-521">String | [Location](/javascript/api/outlook/office.location)</span><span class="sxs-lookup"><span data-stu-id="0b5b9-521">String | [Location](/javascript/api/outlook/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="0b5b9-522">Требования</span><span class="sxs-lookup"><span data-stu-id="0b5b9-522">Requirements</span></span>

|<span data-ttu-id="0b5b9-523">Требование</span><span class="sxs-lookup"><span data-stu-id="0b5b9-523">Requirement</span></span>|<span data-ttu-id="0b5b9-524">Значение</span><span class="sxs-lookup"><span data-stu-id="0b5b9-524">Value</span></span>|
|---|---|
|[<span data-ttu-id="0b5b9-525">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="0b5b9-525">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0b5b9-526">1.0</span><span class="sxs-lookup"><span data-stu-id="0b5b9-526">1.0</span></span>|
|[<span data-ttu-id="0b5b9-527">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="0b5b9-527">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0b5b9-528">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0b5b9-528">ReadItem</span></span>|
|[<span data-ttu-id="0b5b9-529">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="0b5b9-529">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0b5b9-530">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="0b5b9-530">Compose or Read</span></span>|

---
---

#### <a name="normalizedsubject-string"></a><span data-ttu-id="0b5b9-531">normalizedSubject: строка</span><span class="sxs-lookup"><span data-stu-id="0b5b9-531">normalizedSubject: String</span></span>

<span data-ttu-id="0b5b9-p120">Получает тему элемента со всеми удаленными префиксами (включая `RE:` и `FWD:`). Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-p120">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="0b5b9-p121">Свойство normalizedSubject получает тему элемента со стандартными префиксами (такими как `RE:` и `FW:`), добавляемыми почтовыми программами. Для получения темы элемента с неизмененными префиксами используйте свойство [`subject`](#subject-stringsubject).</span><span class="sxs-lookup"><span data-stu-id="0b5b9-p121">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="0b5b9-536">Тип</span><span class="sxs-lookup"><span data-stu-id="0b5b9-536">Type</span></span>

*   <span data-ttu-id="0b5b9-537">String</span><span class="sxs-lookup"><span data-stu-id="0b5b9-537">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="0b5b9-538">Требования</span><span class="sxs-lookup"><span data-stu-id="0b5b9-538">Requirements</span></span>

|<span data-ttu-id="0b5b9-539">Требование</span><span class="sxs-lookup"><span data-stu-id="0b5b9-539">Requirement</span></span>|<span data-ttu-id="0b5b9-540">Значение</span><span class="sxs-lookup"><span data-stu-id="0b5b9-540">Value</span></span>|
|---|---|
|[<span data-ttu-id="0b5b9-541">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="0b5b9-541">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0b5b9-542">1.0</span><span class="sxs-lookup"><span data-stu-id="0b5b9-542">1.0</span></span>|
|[<span data-ttu-id="0b5b9-543">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="0b5b9-543">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0b5b9-544">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0b5b9-544">ReadItem</span></span>|
|[<span data-ttu-id="0b5b9-545">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="0b5b9-545">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0b5b9-546">Чтение</span><span class="sxs-lookup"><span data-stu-id="0b5b9-546">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0b5b9-547">Пример</span><span class="sxs-lookup"><span data-stu-id="0b5b9-547">Example</span></span>

```javascript
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

---
---

#### <a name="notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessages"></a><span data-ttu-id="0b5b9-548">notificationMessages: [notificationMessages](/javascript/api/outlook/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="0b5b9-548">notificationMessages: [NotificationMessages](/javascript/api/outlook/office.notificationmessages)</span></span>

<span data-ttu-id="0b5b9-549">Получает сообщения уведомления для элемента.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-549">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="0b5b9-550">Тип</span><span class="sxs-lookup"><span data-stu-id="0b5b9-550">Type</span></span>

*   [<span data-ttu-id="0b5b9-551">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="0b5b9-551">NotificationMessages</span></span>](/javascript/api/outlook/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="0b5b9-552">Требования</span><span class="sxs-lookup"><span data-stu-id="0b5b9-552">Requirements</span></span>

|<span data-ttu-id="0b5b9-553">Требование</span><span class="sxs-lookup"><span data-stu-id="0b5b9-553">Requirement</span></span>|<span data-ttu-id="0b5b9-554">Значение</span><span class="sxs-lookup"><span data-stu-id="0b5b9-554">Value</span></span>|
|---|---|
|[<span data-ttu-id="0b5b9-555">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="0b5b9-555">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0b5b9-556">1.3</span><span class="sxs-lookup"><span data-stu-id="0b5b9-556">1.3</span></span>|
|[<span data-ttu-id="0b5b9-557">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="0b5b9-557">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0b5b9-558">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0b5b9-558">ReadItem</span></span>|
|[<span data-ttu-id="0b5b9-559">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="0b5b9-559">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0b5b9-560">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="0b5b9-560">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0b5b9-561">Пример</span><span class="sxs-lookup"><span data-stu-id="0b5b9-561">Example</span></span>

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

#### <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="0b5b9-562">optionalAttendees: Array. <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[получатели](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="0b5b9-562">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="0b5b9-563">Предоставляет доступ к необязательным участникам события.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-563">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="0b5b9-564">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-564">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="0b5b9-565">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="0b5b9-565">Read mode</span></span>

<span data-ttu-id="0b5b9-566">Свойство `optionalAttendees` возвращает массив, содержащий объект `EmailAddressDetails` для каждого необязательного участника собрания.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-566">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

```javascript
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="0b5b9-567">Режим создания</span><span class="sxs-lookup"><span data-stu-id="0b5b9-567">Compose mode</span></span>

<span data-ttu-id="0b5b9-568">Свойство `optionalAttendees` возвращает объект `Recipients`, который предоставляет методы для получения или обновления необязательных участников собрания.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-568">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="0b5b9-569">Тип</span><span class="sxs-lookup"><span data-stu-id="0b5b9-569">Type</span></span>

*   <span data-ttu-id="0b5b9-570">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="0b5b9-570">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="0b5b9-571">Требования</span><span class="sxs-lookup"><span data-stu-id="0b5b9-571">Requirements</span></span>

|<span data-ttu-id="0b5b9-572">Требование</span><span class="sxs-lookup"><span data-stu-id="0b5b9-572">Requirement</span></span>|<span data-ttu-id="0b5b9-573">Значение</span><span class="sxs-lookup"><span data-stu-id="0b5b9-573">Value</span></span>|
|---|---|
|[<span data-ttu-id="0b5b9-574">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="0b5b9-574">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0b5b9-575">1.0</span><span class="sxs-lookup"><span data-stu-id="0b5b9-575">1.0</span></span>|
|[<span data-ttu-id="0b5b9-576">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="0b5b9-576">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0b5b9-577">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0b5b9-577">ReadItem</span></span>|
|[<span data-ttu-id="0b5b9-578">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="0b5b9-578">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0b5b9-579">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="0b5b9-579">Compose or Read</span></span>|

---
---

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsorganizerjavascriptapioutlookofficeorganizer"></a><span data-ttu-id="0b5b9-580">Организатор: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[Организатор](/javascript/api/outlook/office.organizer)</span><span class="sxs-lookup"><span data-stu-id="0b5b9-580">organizer: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[Organizer](/javascript/api/outlook/office.organizer)</span></span>

<span data-ttu-id="0b5b9-581">Получает адрес электронной почты организатора для указанного собрания.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-581">Gets the email address of the organizer for a specified meeting.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="0b5b9-582">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="0b5b9-582">Read mode</span></span>

<span data-ttu-id="0b5b9-583">`organizer` Свойство возвращает объект [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) , представляющий организатора собрания.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-583">The `organizer` property returns an [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) object that represents the meeting organizer.</span></span>

```javascript
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

##### <a name="compose-mode"></a><span data-ttu-id="0b5b9-584">Режим создания</span><span class="sxs-lookup"><span data-stu-id="0b5b9-584">Compose mode</span></span>

<span data-ttu-id="0b5b9-585">Свойство возвращает объект организатора, который предоставляет метод для получения значения организатора. [](/javascript/api/outlook/office.organizer) `organizer`</span><span class="sxs-lookup"><span data-stu-id="0b5b9-585">The `organizer` property returns an [Organizer](/javascript/api/outlook/office.organizer) object that provides a method to get the organizer value.</span></span>

```javascript
Office.context.mailbox.item.organizer.getAsync(
  function(asyncResult) {
    console.log(JSON.stringify(asyncResult));
  }
);
```

##### <a name="type"></a><span data-ttu-id="0b5b9-586">Тип</span><span class="sxs-lookup"><span data-stu-id="0b5b9-586">Type</span></span>

*   <span data-ttu-id="0b5b9-587">[](/javascript/api/outlook/office.emailaddressdetails) | [Организатор](/javascript/api/outlook/office.organizer) EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="0b5b9-587">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [Organizer](/javascript/api/outlook/office.organizer)</span></span>

##### <a name="requirements"></a><span data-ttu-id="0b5b9-588">Требования</span><span class="sxs-lookup"><span data-stu-id="0b5b9-588">Requirements</span></span>

|<span data-ttu-id="0b5b9-589">Требование</span><span class="sxs-lookup"><span data-stu-id="0b5b9-589">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="0b5b9-590">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="0b5b9-590">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0b5b9-591">1.0</span><span class="sxs-lookup"><span data-stu-id="0b5b9-591">1.0</span></span>|<span data-ttu-id="0b5b9-592">1.7</span><span class="sxs-lookup"><span data-stu-id="0b5b9-592">1.7</span></span>|
|[<span data-ttu-id="0b5b9-593">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="0b5b9-593">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0b5b9-594">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0b5b9-594">ReadItem</span></span>|<span data-ttu-id="0b5b9-595">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="0b5b9-595">ReadWriteItem</span></span>|
|[<span data-ttu-id="0b5b9-596">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="0b5b9-596">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0b5b9-597">Чтение</span><span class="sxs-lookup"><span data-stu-id="0b5b9-597">Read</span></span>|<span data-ttu-id="0b5b9-598">Создание</span><span class="sxs-lookup"><span data-stu-id="0b5b9-598">Compose</span></span>|

---
---

#### <a name="nullable-recurrence-recurrencejavascriptapioutlookofficerecurrence"></a><span data-ttu-id="0b5b9-599">(Nullable) повторение [](/javascript/api/outlook/office.recurrence) : повторение</span><span class="sxs-lookup"><span data-stu-id="0b5b9-599">(nullable) recurrence: [Recurrence](/javascript/api/outlook/office.recurrence)</span></span>

<span data-ttu-id="0b5b9-600">Получает или задает шаблон повторения встречи.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-600">Gets or sets the recurrence pattern of an appointment.</span></span> <span data-ttu-id="0b5b9-601">Получает шаблон повторения приглашения на собрание.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-601">Gets the recurrence pattern of a meeting request.</span></span> <span data-ttu-id="0b5b9-602">Режимы чтения и создания для элементов встречи.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-602">Read and compose modes for appointment items.</span></span> <span data-ttu-id="0b5b9-603">Режим чтения для элементов приглашения на собрания.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-603">Read mode for meeting request items.</span></span>

<span data-ttu-id="0b5b9-604">`recurrence` Свойство возвращает объект [повторения](/javascript/api/outlook/office.recurrence) для повторяющихся встреч или приглашений на собрания, если элемент представляет собой серию или экземпляр в ряду.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-604">The `recurrence` property returns a [recurrence](/javascript/api/outlook/office.recurrence) object for recurring appointments or meetings requests if an item is a series or an instance in a series.</span></span> <span data-ttu-id="0b5b9-605">`null`возвращается для отдельных встреч и приглашений на собрание для отдельных встреч.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-605">`null` is returned for single appointments and meeting requests of single appointments.</span></span> <span data-ttu-id="0b5b9-606">`undefined`возвращается для сообщений, которые не являются приглашениями на собрания.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-606">`undefined` is returned for messages that are not meeting requests.</span></span>

> <span data-ttu-id="0b5b9-607">Note: приглашения на `itemClass` собрания имеют значение IPM. Schedule. Meeting. Request.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-607">Note: Meeting requests have an `itemClass` value of IPM.Schedule.Meeting.Request.</span></span>

> <span data-ttu-id="0b5b9-608">Note: при наличии объекта `null`повторения это указывает на то, что объект является одной встречей или приглашением на собрание одной встречи, а не частью ряда.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-608">Note: If the recurrence object is `null`, this indicates that the object is a single appointment or a meeting request of a single appointment and NOT a part of a series.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="0b5b9-609">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="0b5b9-609">Read mode</span></span>

<span data-ttu-id="0b5b9-610">`recurrence` Свойство возвращает объект [повторения](/javascript/api/outlook/office.recurrence) , представляющий повторение встречи.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-610">The `recurrence` property returns a [Recurrence](/javascript/api/outlook/office.recurrence) object that represents the appointment recurrence.</span></span> <span data-ttu-id="0b5b9-611">Оно доступно для встреч и приглашений на собрания.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-611">This is available for appointments and meeting requests.</span></span>

```javascript
var recurrence = Office.context.mailbox.item.recurrence;
console.log("Recurrence: " + JSON.stringify(recurrence));
```

##### <a name="compose-mode"></a><span data-ttu-id="0b5b9-612">Режим создания</span><span class="sxs-lookup"><span data-stu-id="0b5b9-612">Compose mode</span></span>

<span data-ttu-id="0b5b9-613">`recurrence` Свойство возвращает объект [повторения](/javascript/api/outlook/office.recurrence) , который предоставляет методы для управления повторением встречи.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-613">The `recurrence` property returns a [Recurrence](/javascript/api/outlook/office.recurrence) object that provides methods to manage the appointment recurrence.</span></span> <span data-ttu-id="0b5b9-614">Оно доступно для встреч.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-614">This is available for appointments.</span></span>

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

##### <a name="type"></a><span data-ttu-id="0b5b9-615">Тип</span><span class="sxs-lookup"><span data-stu-id="0b5b9-615">Type</span></span>

* [<span data-ttu-id="0b5b9-616">Повторения</span><span class="sxs-lookup"><span data-stu-id="0b5b9-616">Recurrence</span></span>](/javascript/api/outlook/office.recurrence)

|<span data-ttu-id="0b5b9-617">Требование</span><span class="sxs-lookup"><span data-stu-id="0b5b9-617">Requirement</span></span>|<span data-ttu-id="0b5b9-618">Значение</span><span class="sxs-lookup"><span data-stu-id="0b5b9-618">Value</span></span>|
|---|---|
|[<span data-ttu-id="0b5b9-619">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="0b5b9-619">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0b5b9-620">1.7</span><span class="sxs-lookup"><span data-stu-id="0b5b9-620">1.7</span></span>|
|[<span data-ttu-id="0b5b9-621">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="0b5b9-621">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0b5b9-622">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0b5b9-622">ReadItem</span></span>|
|[<span data-ttu-id="0b5b9-623">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="0b5b9-623">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0b5b9-624">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="0b5b9-624">Compose or Read</span></span>|

---
---

#### <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="0b5b9-625">requiredAttendees: Array. <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[получатели](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="0b5b9-625">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="0b5b9-626">Предоставляет доступ к обязательным участникам события.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-626">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="0b5b9-627">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-627">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="0b5b9-628">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="0b5b9-628">Read mode</span></span>

<span data-ttu-id="0b5b9-629">Свойство `requiredAttendees` возвращает массив, содержащий объект `EmailAddressDetails` для каждого обязательного участника собрания.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-629">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

```javascript
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="0b5b9-630">Режим создания</span><span class="sxs-lookup"><span data-stu-id="0b5b9-630">Compose mode</span></span>

<span data-ttu-id="0b5b9-631">Свойство `requiredAttendees` возвращает объект `Recipients`, предоставляющий методы, с помощью которых можно получить или обновить сведения об обязательных участниках собрания.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-631">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="0b5b9-632">Тип</span><span class="sxs-lookup"><span data-stu-id="0b5b9-632">Type</span></span>

*   <span data-ttu-id="0b5b9-633">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="0b5b9-633">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="0b5b9-634">Требования</span><span class="sxs-lookup"><span data-stu-id="0b5b9-634">Requirements</span></span>

|<span data-ttu-id="0b5b9-635">Требование</span><span class="sxs-lookup"><span data-stu-id="0b5b9-635">Requirement</span></span>|<span data-ttu-id="0b5b9-636">Значение</span><span class="sxs-lookup"><span data-stu-id="0b5b9-636">Value</span></span>|
|---|---|
|[<span data-ttu-id="0b5b9-637">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="0b5b9-637">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0b5b9-638">1.0</span><span class="sxs-lookup"><span data-stu-id="0b5b9-638">1.0</span></span>|
|[<span data-ttu-id="0b5b9-639">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="0b5b9-639">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0b5b9-640">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0b5b9-640">ReadItem</span></span>|
|[<span data-ttu-id="0b5b9-641">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="0b5b9-641">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0b5b9-642">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="0b5b9-642">Compose or Read</span></span>|

---
---

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails"></a><span data-ttu-id="0b5b9-643">Отправитель: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="0b5b9-643">sender: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)</span></span>

<span data-ttu-id="0b5b9-p128">Получает электронный адрес отправителя электронного сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-p128">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="0b5b9-p129">Свойства [`from`](#from-emailaddressdetailsfrom) и `sender` представляют одно лицо, если сообщение не отправлено представителем. В противном случае свойство `from` представляет лицо, делегировавшее полномочия, а свойство sender — представителя.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-p129">The [`from`](#from-emailaddressdetailsfrom) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="0b5b9-648">Свойству `recipientType`, принадлежащему объекту `EmailAddressDetails` в свойстве `sender`, задано значение `undefined`.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-648">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="0b5b9-649">Тип</span><span class="sxs-lookup"><span data-stu-id="0b5b9-649">Type</span></span>

*   [<span data-ttu-id="0b5b9-650">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="0b5b9-650">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="0b5b9-651">Требования</span><span class="sxs-lookup"><span data-stu-id="0b5b9-651">Requirements</span></span>

|<span data-ttu-id="0b5b9-652">Требование</span><span class="sxs-lookup"><span data-stu-id="0b5b9-652">Requirement</span></span>|<span data-ttu-id="0b5b9-653">Значение</span><span class="sxs-lookup"><span data-stu-id="0b5b9-653">Value</span></span>|
|---|---|
|[<span data-ttu-id="0b5b9-654">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="0b5b9-654">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0b5b9-655">1.0</span><span class="sxs-lookup"><span data-stu-id="0b5b9-655">1.0</span></span>|
|[<span data-ttu-id="0b5b9-656">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="0b5b9-656">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0b5b9-657">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0b5b9-657">ReadItem</span></span>|
|[<span data-ttu-id="0b5b9-658">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="0b5b9-658">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0b5b9-659">Чтение</span><span class="sxs-lookup"><span data-stu-id="0b5b9-659">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0b5b9-660">Пример</span><span class="sxs-lookup"><span data-stu-id="0b5b9-660">Example</span></span>

```javascript
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

---
---

#### <a name="nullable-seriesid-string"></a><span data-ttu-id="0b5b9-661">(Nullable) seriesId: строка</span><span class="sxs-lookup"><span data-stu-id="0b5b9-661">(nullable) seriesId: String</span></span>

<span data-ttu-id="0b5b9-662">Получает идентификатор ряда, к которому принадлежит экземпляр.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-662">Gets the id of the series that an instance belongs to.</span></span>

<span data-ttu-id="0b5b9-663">В OWA и Outlook `seriesId` возвращается идентификатор веб-служб Exchange (EWS) родительского элемента (ряда), к которому принадлежит этот элемент.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-663">In OWA and Outlook, the `seriesId` returns the Exchange Web Services (EWS) ID of the parent (series) item that this item belongs to.</span></span> <span data-ttu-id="0b5b9-664">Однако в iOS и Android `seriesId` возвращается идентификатор REST родительского элемента.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-664">However, in iOS and Android, the `seriesId` returns the REST ID of the parent item.</span></span>

> [!NOTE]
> <span data-ttu-id="0b5b9-665">Идентификатор, возвращаемый свойством `seriesId`, совпадает с идентификатором элемента веб-служб Exchange.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-665">The identifier returned by the `seriesId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="0b5b9-666">`seriesId` Свойство не совпадает с идентификаторами Outlook, используемыми в REST API Outlook.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-666">The `seriesId` property is not identical to the Outlook IDs used by the Outlook REST API.</span></span> <span data-ttu-id="0b5b9-667">Прежде чем совершать вызовы REST API, используя это значение, его необходимо преобразовать с помощью [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="0b5b9-667">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="0b5b9-668">Дополнительные сведения см. в статье [Использование REST API Outlook из надстройки Outlook](/outlook/add-ins/use-rest-api).</span><span class="sxs-lookup"><span data-stu-id="0b5b9-668">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api).</span></span>

<span data-ttu-id="0b5b9-669">`seriesId` Свойство возвращает `null` элементы, у которых нет родительских элементов, таких как одиночные встречи, элементы ряда или приглашения на собрание, `undefined` и возвращаемые для других элементов, не являющиеся приглашениями на собрания.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-669">The `seriesId` property returns `null` for items that do not have parent items such as single appointments, series items, or meeting requests and returns `undefined` for any other items that are not meeting requests.</span></span>

##### <a name="type"></a><span data-ttu-id="0b5b9-670">Тип</span><span class="sxs-lookup"><span data-stu-id="0b5b9-670">Type</span></span>

* <span data-ttu-id="0b5b9-671">String</span><span class="sxs-lookup"><span data-stu-id="0b5b9-671">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="0b5b9-672">Требования</span><span class="sxs-lookup"><span data-stu-id="0b5b9-672">Requirements</span></span>

|<span data-ttu-id="0b5b9-673">Требование</span><span class="sxs-lookup"><span data-stu-id="0b5b9-673">Requirement</span></span>|<span data-ttu-id="0b5b9-674">Значение</span><span class="sxs-lookup"><span data-stu-id="0b5b9-674">Value</span></span>|
|---|---|
|[<span data-ttu-id="0b5b9-675">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="0b5b9-675">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0b5b9-676">1.7</span><span class="sxs-lookup"><span data-stu-id="0b5b9-676">1.7</span></span>|
|[<span data-ttu-id="0b5b9-677">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="0b5b9-677">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0b5b9-678">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0b5b9-678">ReadItem</span></span>|
|[<span data-ttu-id="0b5b9-679">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="0b5b9-679">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0b5b9-680">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="0b5b9-680">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0b5b9-681">Пример</span><span class="sxs-lookup"><span data-stu-id="0b5b9-681">Example</span></span>

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

#### <a name="start-datetimejavascriptapioutlookofficetime"></a><span data-ttu-id="0b5b9-682">Начало: Дата | [Time (время](/javascript/api/outlook/office.time) )</span><span class="sxs-lookup"><span data-stu-id="0b5b9-682">start: Date|[Time](/javascript/api/outlook/office.time)</span></span>

<span data-ttu-id="0b5b9-683">Получает или задает дату и время начала встречи.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-683">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="0b5b9-p132">Свойство `start` представлено в виде значения даты и времени в формате UTC. Это значение можно преобразовать в местные значения даты и времени клиента с помощью метода [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime).</span><span class="sxs-lookup"><span data-stu-id="0b5b9-p132">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="0b5b9-686">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="0b5b9-686">Read mode</span></span>

<span data-ttu-id="0b5b9-687">Свойство `start` возвращает объект `Date`.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-687">The `start` property returns a `Date` object.</span></span>

```javascript
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="0b5b9-688">Режим создания</span><span class="sxs-lookup"><span data-stu-id="0b5b9-688">Compose mode</span></span>

<span data-ttu-id="0b5b9-689">Свойство `start` возвращает объект `Time`.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-689">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="0b5b9-690">Если вы задаете время начала с помощью метода [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-), необходимо использовать метод [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) для преобразования местного времени на клиенте в формат UTC для сервера.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-690">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="0b5b9-691">В примере ниже с помощью метода [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) объекта `Time` задается время начала встречи в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-691">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="0b5b9-692">Тип</span><span class="sxs-lookup"><span data-stu-id="0b5b9-692">Type</span></span>

*   <span data-ttu-id="0b5b9-693">Date | [Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="0b5b9-693">Date | [Time](/javascript/api/outlook/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="0b5b9-694">Требования</span><span class="sxs-lookup"><span data-stu-id="0b5b9-694">Requirements</span></span>

|<span data-ttu-id="0b5b9-695">Требование</span><span class="sxs-lookup"><span data-stu-id="0b5b9-695">Requirement</span></span>|<span data-ttu-id="0b5b9-696">Значение</span><span class="sxs-lookup"><span data-stu-id="0b5b9-696">Value</span></span>|
|---|---|
|[<span data-ttu-id="0b5b9-697">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="0b5b9-697">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0b5b9-698">1.0</span><span class="sxs-lookup"><span data-stu-id="0b5b9-698">1.0</span></span>|
|[<span data-ttu-id="0b5b9-699">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="0b5b9-699">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0b5b9-700">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0b5b9-700">ReadItem</span></span>|
|[<span data-ttu-id="0b5b9-701">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="0b5b9-701">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0b5b9-702">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="0b5b9-702">Compose or Read</span></span>|

---
---

#### <a name="subject-stringsubjectjavascriptapioutlookofficesubject"></a><span data-ttu-id="0b5b9-703">Тема: строка | [Subject (тема](/javascript/api/outlook/office.subject) )</span><span class="sxs-lookup"><span data-stu-id="0b5b9-703">subject: String|[Subject](/javascript/api/outlook/office.subject)</span></span>

<span data-ttu-id="0b5b9-704">Получает или задает описание, которое отображается в поле темы элемента.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-704">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="0b5b9-705">Свойство `subject` получает или задает всю тему элемента для отправки с почтового сервера.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-705">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="0b5b9-706">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="0b5b9-706">Read mode</span></span>

<span data-ttu-id="0b5b9-p133">Свойство `subject` возвращает строку. С помощью свойства [`normalizedSubject`](#normalizedsubject-string) можно получить тему без начальных префиксов, таких как `RE:` и `FW:`.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-p133">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

<span data-ttu-id="0b5b9-709">В примере кода JavaScript, приведенном ниже, показано, как получить доступ к свойству `subject` текущего элемента в Outlook.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-709">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

```javascript
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="0b5b9-710">Режим создания</span><span class="sxs-lookup"><span data-stu-id="0b5b9-710">Compose mode</span></span>
<span data-ttu-id="0b5b9-711">Свойство `subject` возвращает объект `Subject`, который предоставляет методы для получения и задания темы.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-711">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```javascript
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="0b5b9-712">Тип</span><span class="sxs-lookup"><span data-stu-id="0b5b9-712">Type</span></span>

*   <span data-ttu-id="0b5b9-713">String | [Subject](/javascript/api/outlook/office.subject)</span><span class="sxs-lookup"><span data-stu-id="0b5b9-713">String | [Subject](/javascript/api/outlook/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="0b5b9-714">Требования</span><span class="sxs-lookup"><span data-stu-id="0b5b9-714">Requirements</span></span>

|<span data-ttu-id="0b5b9-715">Требование</span><span class="sxs-lookup"><span data-stu-id="0b5b9-715">Requirement</span></span>|<span data-ttu-id="0b5b9-716">Значение</span><span class="sxs-lookup"><span data-stu-id="0b5b9-716">Value</span></span>|
|---|---|
|[<span data-ttu-id="0b5b9-717">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="0b5b9-717">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0b5b9-718">1.0</span><span class="sxs-lookup"><span data-stu-id="0b5b9-718">1.0</span></span>|
|[<span data-ttu-id="0b5b9-719">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="0b5b9-719">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0b5b9-720">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0b5b9-720">ReadItem</span></span>|
|[<span data-ttu-id="0b5b9-721">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="0b5b9-721">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0b5b9-722">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="0b5b9-722">Compose or Read</span></span>|

---
---

#### <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="0b5b9-723">Кому: Array. <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[получатели](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="0b5b9-723">to: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="0b5b9-724">Предоставляет доступ к получателям, указанным в строке **Кому** сообщения.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-724">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="0b5b9-725">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-725">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="0b5b9-726">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="0b5b9-726">Read mode</span></span>

<span data-ttu-id="0b5b9-p135">Свойство `to` возвращает массив, содержащий объект `EmailAddressDetails` для каждого получателя в строке **Кому** сообщения. Коллекция может включать не более 100 элементов.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-p135">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="0b5b9-729">Режим создания</span><span class="sxs-lookup"><span data-stu-id="0b5b9-729">Compose mode</span></span>

<span data-ttu-id="0b5b9-730">Свойство `to` возвращает объект `Recipients`, предоставляющий методы для получения или обновления получателей, которые указаны в строке **Кому** сообщения.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-730">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

```javascript
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="0b5b9-731">Тип</span><span class="sxs-lookup"><span data-stu-id="0b5b9-731">Type</span></span>

*   <span data-ttu-id="0b5b9-732">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="0b5b9-732">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="0b5b9-733">Требования</span><span class="sxs-lookup"><span data-stu-id="0b5b9-733">Requirements</span></span>

|<span data-ttu-id="0b5b9-734">Требование</span><span class="sxs-lookup"><span data-stu-id="0b5b9-734">Requirement</span></span>|<span data-ttu-id="0b5b9-735">Значение</span><span class="sxs-lookup"><span data-stu-id="0b5b9-735">Value</span></span>|
|---|---|
|[<span data-ttu-id="0b5b9-736">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="0b5b9-736">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0b5b9-737">1.0</span><span class="sxs-lookup"><span data-stu-id="0b5b9-737">1.0</span></span>|
|[<span data-ttu-id="0b5b9-738">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="0b5b9-738">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0b5b9-739">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0b5b9-739">ReadItem</span></span>|
|[<span data-ttu-id="0b5b9-740">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="0b5b9-740">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0b5b9-741">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="0b5b9-741">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="0b5b9-742">Методы</span><span class="sxs-lookup"><span data-stu-id="0b5b9-742">Methods</span></span>

#### <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="0b5b9-743">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="0b5b9-743">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="0b5b9-744">Добавляет файл в сообщение или встречу в качестве вложения.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-744">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="0b5b9-745">Метод `addFileAttachmentAsync` передает файл по указанному универсальному коду ресурса (URI) и вкладывает его в элемент в форме создания.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-745">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="0b5b9-746">Идентификатор можно использовать с методом [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback), чтобы удалить вложение, добавленное во время текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-746">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0b5b9-747">Параметры</span><span class="sxs-lookup"><span data-stu-id="0b5b9-747">Parameters</span></span>
|<span data-ttu-id="0b5b9-748">Имя</span><span class="sxs-lookup"><span data-stu-id="0b5b9-748">Name</span></span>|<span data-ttu-id="0b5b9-749">Тип</span><span class="sxs-lookup"><span data-stu-id="0b5b9-749">Type</span></span>|<span data-ttu-id="0b5b9-750">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="0b5b9-750">Attributes</span></span>|<span data-ttu-id="0b5b9-751">Описание</span><span class="sxs-lookup"><span data-stu-id="0b5b9-751">Description</span></span>|
|---|---|---|---|
|`uri`|<span data-ttu-id="0b5b9-752">String</span><span class="sxs-lookup"><span data-stu-id="0b5b9-752">String</span></span>||<span data-ttu-id="0b5b9-p136">Универсальный код ресурса (URI), представляющий расположение файла, который нужно вложить в сообщение или встречу. Максимальная длина — 2048 символов.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-p136">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="0b5b9-755">String</span><span class="sxs-lookup"><span data-stu-id="0b5b9-755">String</span></span>||<span data-ttu-id="0b5b9-p137">Имя вложения, которое отображается при передаче вложения. Максимальная длина — 255 символов.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-p137">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="0b5b9-758">Объект</span><span class="sxs-lookup"><span data-stu-id="0b5b9-758">Object</span></span>|<span data-ttu-id="0b5b9-759">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="0b5b9-759">&lt;optional&gt;</span></span>|<span data-ttu-id="0b5b9-760">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-760">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="0b5b9-761">Object</span><span class="sxs-lookup"><span data-stu-id="0b5b9-761">Object</span></span>|<span data-ttu-id="0b5b9-762">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="0b5b9-762">&lt;optional&gt;</span></span>|<span data-ttu-id="0b5b9-763">В методе обратного вызова разработчики могут указать любой объект, к которому необходимо получить доступ.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-763">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="0b5b9-764">Boolean</span><span class="sxs-lookup"><span data-stu-id="0b5b9-764">Boolean</span></span>|<span data-ttu-id="0b5b9-765">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="0b5b9-765">&lt;optional&gt;</span></span>|<span data-ttu-id="0b5b9-766">Значение `true` указывает на то, что вложение будет встроено в текст сообщения и не должно отображаться в списке вложений.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-766">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="0b5b9-767">function</span><span class="sxs-lookup"><span data-stu-id="0b5b9-767">function</span></span>|<span data-ttu-id="0b5b9-768">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="0b5b9-768">&lt;optional&gt;</span></span>|<span data-ttu-id="0b5b9-769">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="0b5b9-769">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="0b5b9-770">После успешного выполнения идентификатор вложения будет представлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-770">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="0b5b9-771">Если добавить вложение не удастся, объект `asyncResult` будет содержать объект `Error` с описанием ошибки.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-771">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="0b5b9-772">Ошибки</span><span class="sxs-lookup"><span data-stu-id="0b5b9-772">Errors</span></span>

|<span data-ttu-id="0b5b9-773">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="0b5b9-773">Error code</span></span>|<span data-ttu-id="0b5b9-774">Описание</span><span class="sxs-lookup"><span data-stu-id="0b5b9-774">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="0b5b9-775">Вложение превышает максимальный размер.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-775">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="0b5b9-776">Расширение вложения не поддерживается.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-776">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="0b5b9-777">Сообщение или встреча содержат слишком много вложений.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-777">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0b5b9-778">Требования</span><span class="sxs-lookup"><span data-stu-id="0b5b9-778">Requirements</span></span>

|<span data-ttu-id="0b5b9-779">Требование</span><span class="sxs-lookup"><span data-stu-id="0b5b9-779">Requirement</span></span>|<span data-ttu-id="0b5b9-780">Значение</span><span class="sxs-lookup"><span data-stu-id="0b5b9-780">Value</span></span>|
|---|---|
|[<span data-ttu-id="0b5b9-781">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="0b5b9-781">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0b5b9-782">1.1</span><span class="sxs-lookup"><span data-stu-id="0b5b9-782">1.1</span></span>|
|[<span data-ttu-id="0b5b9-783">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="0b5b9-783">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0b5b9-784">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="0b5b9-784">ReadWriteItem</span></span>|
|[<span data-ttu-id="0b5b9-785">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="0b5b9-785">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0b5b9-786">Создание</span><span class="sxs-lookup"><span data-stu-id="0b5b9-786">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="0b5b9-787">Примеры</span><span class="sxs-lookup"><span data-stu-id="0b5b9-787">Examples</span></span>

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

<span data-ttu-id="0b5b9-788">В приведенном ниже примере файл изображения добавляется в качестве встроенного вложения, а ссылка на вложение добавляется в текст сообщения.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-788">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

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

#### <a name="addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback"></a><span data-ttu-id="0b5b9-789">addFileAttachmentFromBase64Async (base64File, Аттачментнаме, [параметры], [обратный вызов])</span><span class="sxs-lookup"><span data-stu-id="0b5b9-789">addFileAttachmentFromBase64Async(base64File, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="0b5b9-790">Добавляет файл из кодировки Base64 в сообщение или встречу в качестве вложения.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-790">Adds a file from the base64 encoding to a message or appointment as an attachment.</span></span>

<span data-ttu-id="0b5b9-791">`addFileAttachmentFromBase64Async` Метод передает файл из кодировки Base64 и вкладывает его в элемент в форме создания.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-791">The `addFileAttachmentFromBase64Async` method uploads the file from the base64 encoding and attaches it to the item in the compose form.</span></span> <span data-ttu-id="0b5b9-792">Этот метод возвращает идентификатор вложения в объекте AsyncResult. Value.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-792">This method returns the attachment identifier in the AsyncResult.value object.</span></span>

<span data-ttu-id="0b5b9-793">Идентификатор можно использовать с методом [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback), чтобы удалить вложение, добавленное во время текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-793">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0b5b9-794">Параметры</span><span class="sxs-lookup"><span data-stu-id="0b5b9-794">Parameters</span></span>

|<span data-ttu-id="0b5b9-795">Имя</span><span class="sxs-lookup"><span data-stu-id="0b5b9-795">Name</span></span>|<span data-ttu-id="0b5b9-796">Тип</span><span class="sxs-lookup"><span data-stu-id="0b5b9-796">Type</span></span>|<span data-ttu-id="0b5b9-797">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="0b5b9-797">Attributes</span></span>|<span data-ttu-id="0b5b9-798">Описание</span><span class="sxs-lookup"><span data-stu-id="0b5b9-798">Description</span></span>|
|---|---|---|---|
|`base64File`|<span data-ttu-id="0b5b9-799">String</span><span class="sxs-lookup"><span data-stu-id="0b5b9-799">String</span></span>||<span data-ttu-id="0b5b9-800">Содержимое изображения или файла в кодировке Base64, которое добавляется в сообщение электронной почты или событие.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-800">The base64 encoded content of an image or file to be added to an email or event.</span></span>|
|`attachmentName`|<span data-ttu-id="0b5b9-801">String</span><span class="sxs-lookup"><span data-stu-id="0b5b9-801">String</span></span>||<span data-ttu-id="0b5b9-p139">Имя вложения, которое отображается при передаче вложения. Максимальная длина — 255 символов.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-p139">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="0b5b9-804">Объект</span><span class="sxs-lookup"><span data-stu-id="0b5b9-804">Object</span></span>|<span data-ttu-id="0b5b9-805">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="0b5b9-805">&lt;optional&gt;</span></span>|<span data-ttu-id="0b5b9-806">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-806">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="0b5b9-807">Объект</span><span class="sxs-lookup"><span data-stu-id="0b5b9-807">Object</span></span>|<span data-ttu-id="0b5b9-808">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="0b5b9-808">&lt;optional&gt;</span></span>|<span data-ttu-id="0b5b9-809">В методе обратного вызова разработчики могут указать любой объект, к которому необходимо получить доступ.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-809">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="0b5b9-810">Boolean</span><span class="sxs-lookup"><span data-stu-id="0b5b9-810">Boolean</span></span>|<span data-ttu-id="0b5b9-811">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="0b5b9-811">&lt;optional&gt;</span></span>|<span data-ttu-id="0b5b9-812">Значение `true` указывает на то, что вложение будет встроено в текст сообщения и не должно отображаться в списке вложений.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-812">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="0b5b9-813">function</span><span class="sxs-lookup"><span data-stu-id="0b5b9-813">function</span></span>|<span data-ttu-id="0b5b9-814">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="0b5b9-814">&lt;optional&gt;</span></span>|<span data-ttu-id="0b5b9-815">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="0b5b9-815">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="0b5b9-816">После успешного выполнения идентификатор вложения будет представлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-816">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="0b5b9-817">Если добавить вложение не удастся, объект `asyncResult` будет содержать объект `Error` с описанием ошибки.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-817">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="0b5b9-818">Ошибки</span><span class="sxs-lookup"><span data-stu-id="0b5b9-818">Errors</span></span>

|<span data-ttu-id="0b5b9-819">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="0b5b9-819">Error code</span></span>|<span data-ttu-id="0b5b9-820">Описание</span><span class="sxs-lookup"><span data-stu-id="0b5b9-820">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="0b5b9-821">Вложение превышает максимальный размер.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-821">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="0b5b9-822">Расширение вложения не поддерживается.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-822">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="0b5b9-823">Сообщение или встреча содержат слишком много вложений.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-823">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0b5b9-824">Требования</span><span class="sxs-lookup"><span data-stu-id="0b5b9-824">Requirements</span></span>

|<span data-ttu-id="0b5b9-825">Требование</span><span class="sxs-lookup"><span data-stu-id="0b5b9-825">Requirement</span></span>|<span data-ttu-id="0b5b9-826">Значение</span><span class="sxs-lookup"><span data-stu-id="0b5b9-826">Value</span></span>|
|---|---|
|[<span data-ttu-id="0b5b9-827">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="0b5b9-827">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0b5b9-828">Предварительная версия</span><span class="sxs-lookup"><span data-stu-id="0b5b9-828">Preview</span></span>|
|[<span data-ttu-id="0b5b9-829">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="0b5b9-829">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0b5b9-830">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="0b5b9-830">ReadWriteItem</span></span>|
|[<span data-ttu-id="0b5b9-831">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="0b5b9-831">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0b5b9-832">Создание</span><span class="sxs-lookup"><span data-stu-id="0b5b9-832">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="0b5b9-833">Примеры</span><span class="sxs-lookup"><span data-stu-id="0b5b9-833">Examples</span></span>

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

#### <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="0b5b9-834">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="0b5b9-834">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="0b5b9-835">Добавляет обработчик для поддерживаемого события.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-835">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="0b5b9-836">В настоящее время поддерживаются типы `Office.EventType.AttachmentsChanged`событий `Office.EventType.AppointmentTimeChanged`, `Office.EventType.EnhancedLocationsChanged` `Office.EventType.RecipientsChanged`,, и `Office.EventType.RecurrenceChanged`.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-836">Currently the supported event types are `Office.EventType.AttachmentsChanged`, `Office.EventType.AppointmentTimeChanged`, `Office.EventType.EnhancedLocationsChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0b5b9-837">Параметры</span><span class="sxs-lookup"><span data-stu-id="0b5b9-837">Parameters</span></span>

| <span data-ttu-id="0b5b9-838">Имя</span><span class="sxs-lookup"><span data-stu-id="0b5b9-838">Name</span></span> | <span data-ttu-id="0b5b9-839">Тип</span><span class="sxs-lookup"><span data-stu-id="0b5b9-839">Type</span></span> | <span data-ttu-id="0b5b9-840">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="0b5b9-840">Attributes</span></span> | <span data-ttu-id="0b5b9-841">Описание</span><span class="sxs-lookup"><span data-stu-id="0b5b9-841">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="0b5b9-842">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="0b5b9-842">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="0b5b9-843">Событие, которое должно вызвать обработчик.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-843">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="0b5b9-844">Function</span><span class="sxs-lookup"><span data-stu-id="0b5b9-844">Function</span></span> || <span data-ttu-id="0b5b9-p140">Функция для обработки события. Функция должна принимать один параметр, представляющий собой объектный литерал. Значение свойства `type` параметра совпадет со значением параметра `eventType`, переданного методу `addHandlerAsync`.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-p140">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="0b5b9-848">Объект</span><span class="sxs-lookup"><span data-stu-id="0b5b9-848">Object</span></span> | <span data-ttu-id="0b5b9-849">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="0b5b9-849">&lt;optional&gt;</span></span> | <span data-ttu-id="0b5b9-850">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-850">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="0b5b9-851">Объект</span><span class="sxs-lookup"><span data-stu-id="0b5b9-851">Object</span></span> | <span data-ttu-id="0b5b9-852">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="0b5b9-852">&lt;optional&gt;</span></span> | <span data-ttu-id="0b5b9-853">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-853">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="0b5b9-854">функция</span><span class="sxs-lookup"><span data-stu-id="0b5b9-854">function</span></span>| <span data-ttu-id="0b5b9-855">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="0b5b9-855">&lt;optional&gt;</span></span>|<span data-ttu-id="0b5b9-856">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="0b5b9-856">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0b5b9-857">Требования</span><span class="sxs-lookup"><span data-stu-id="0b5b9-857">Requirements</span></span>

|<span data-ttu-id="0b5b9-858">Требование</span><span class="sxs-lookup"><span data-stu-id="0b5b9-858">Requirement</span></span>| <span data-ttu-id="0b5b9-859">Значение</span><span class="sxs-lookup"><span data-stu-id="0b5b9-859">Value</span></span>|
|---|---|
|[<span data-ttu-id="0b5b9-860">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="0b5b9-860">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0b5b9-861">1.7</span><span class="sxs-lookup"><span data-stu-id="0b5b9-861">1.7</span></span> |
|[<span data-ttu-id="0b5b9-862">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="0b5b9-862">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0b5b9-863">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0b5b9-863">ReadItem</span></span> |
|[<span data-ttu-id="0b5b9-864">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="0b5b9-864">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0b5b9-865">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="0b5b9-865">Compose or Read</span></span> |

##### <a name="example"></a><span data-ttu-id="0b5b9-866">Пример</span><span class="sxs-lookup"><span data-stu-id="0b5b9-866">Example</span></span>

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

#### <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="0b5b9-867">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="0b5b9-867">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="0b5b9-868">Добавляет к сообщению элемент Exchange, например сообщение, в виде вложения.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-868">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="0b5b9-p141">С помощью метода `addItemAttachmentAsync` можно в элемент формы создания вложить элемент с указанным идентификатором Exchange. Если указать метод обратного вызова, то этот метод вызывается с помощью параметра `asyncResult`, который содержит идентификатор вложения или код, указывающий на ошибки, которые произошли при вложении элемента. При необходимости можно использовать параметр `options` для передачи сведений о состоянии методу обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-p141">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="0b5b9-872">Идентификатор можно использовать с методом [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback), чтобы удалить вложение, добавленное во время текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-872">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="0b5b9-873">Если ваша надстройка Office выполняется в Outlook Web App, метод `addItemAttachmentAsync` обеспечивает вложение элементов в элементы, отличные от редактируемого. Однако это действие не рекомендуем выполнять, так как оно не поддерживается.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-873">If your Office Add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0b5b9-874">Параметры</span><span class="sxs-lookup"><span data-stu-id="0b5b9-874">Parameters</span></span>

|<span data-ttu-id="0b5b9-875">Имя</span><span class="sxs-lookup"><span data-stu-id="0b5b9-875">Name</span></span>|<span data-ttu-id="0b5b9-876">Тип</span><span class="sxs-lookup"><span data-stu-id="0b5b9-876">Type</span></span>|<span data-ttu-id="0b5b9-877">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="0b5b9-877">Attributes</span></span>|<span data-ttu-id="0b5b9-878">Описание</span><span class="sxs-lookup"><span data-stu-id="0b5b9-878">Description</span></span>|
|---|---|---|---|
|`itemId`|<span data-ttu-id="0b5b9-879">String</span><span class="sxs-lookup"><span data-stu-id="0b5b9-879">String</span></span>||<span data-ttu-id="0b5b9-p142">Идентификатор Exchange для вкладываемого элемента. Максимальная длина — 100 символов.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-p142">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="0b5b9-882">String</span><span class="sxs-lookup"><span data-stu-id="0b5b9-882">String</span></span>||<span data-ttu-id="0b5b9-883">Тема вкладываемого элемента.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-883">The subject of the item to be attached.</span></span> <span data-ttu-id="0b5b9-884">Максимальная длина: 255 символов.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-884">The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="0b5b9-885">Object</span><span class="sxs-lookup"><span data-stu-id="0b5b9-885">Object</span></span>|<span data-ttu-id="0b5b9-886">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="0b5b9-886">&lt;optional&gt;</span></span>|<span data-ttu-id="0b5b9-887">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-887">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="0b5b9-888">Объект</span><span class="sxs-lookup"><span data-stu-id="0b5b9-888">Object</span></span>|<span data-ttu-id="0b5b9-889">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="0b5b9-889">&lt;optional&gt;</span></span>|<span data-ttu-id="0b5b9-890">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-890">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="0b5b9-891">функция</span><span class="sxs-lookup"><span data-stu-id="0b5b9-891">function</span></span>|<span data-ttu-id="0b5b9-892">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="0b5b9-892">&lt;optional&gt;</span></span>|<span data-ttu-id="0b5b9-893">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="0b5b9-893">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="0b5b9-894">После успешного выполнения идентификатор вложения будет представлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-894">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="0b5b9-895">Если добавить вложение не удастся, объект `asyncResult` будет содержать объект `Error` с описанием ошибки.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-895">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="0b5b9-896">Ошибки</span><span class="sxs-lookup"><span data-stu-id="0b5b9-896">Errors</span></span>

|<span data-ttu-id="0b5b9-897">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="0b5b9-897">Error code</span></span>|<span data-ttu-id="0b5b9-898">Описание</span><span class="sxs-lookup"><span data-stu-id="0b5b9-898">Description</span></span>|
|------------|-------------|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="0b5b9-899">Сообщение или встреча содержат слишком много вложений.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-899">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0b5b9-900">Требования</span><span class="sxs-lookup"><span data-stu-id="0b5b9-900">Requirements</span></span>

|<span data-ttu-id="0b5b9-901">Требование</span><span class="sxs-lookup"><span data-stu-id="0b5b9-901">Requirement</span></span>|<span data-ttu-id="0b5b9-902">Значение</span><span class="sxs-lookup"><span data-stu-id="0b5b9-902">Value</span></span>|
|---|---|
|[<span data-ttu-id="0b5b9-903">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="0b5b9-903">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0b5b9-904">1.1</span><span class="sxs-lookup"><span data-stu-id="0b5b9-904">1.1</span></span>|
|[<span data-ttu-id="0b5b9-905">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="0b5b9-905">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0b5b9-906">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="0b5b9-906">ReadWriteItem</span></span>|
|[<span data-ttu-id="0b5b9-907">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="0b5b9-907">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0b5b9-908">Создание</span><span class="sxs-lookup"><span data-stu-id="0b5b9-908">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="0b5b9-909">Пример</span><span class="sxs-lookup"><span data-stu-id="0b5b9-909">Example</span></span>

<span data-ttu-id="0b5b9-910">В следующем примере существующий элемент Outlook добавляется в виде вложения с именем `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-910">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

#### <a name="close"></a><span data-ttu-id="0b5b9-911">close()</span><span class="sxs-lookup"><span data-stu-id="0b5b9-911">close()</span></span>

<span data-ttu-id="0b5b9-912">Закрывает текущий создаваемый элемент.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-912">Closes the current item that is being composed.</span></span>

<span data-ttu-id="0b5b9-p144">Работа метода `close` зависит от текущего состояния создаваемого элемента. Если элемент содержит несохраненные изменения, клиент предложит пользователю сохранить или отклонить их либо отменить действие закрытия.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-p144">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="0b5b9-915">Если в Outlook в Интернете элемент представляет собой встречу, ранее сохраненную с помощью метода `saveAsync`, пользователю предлагается сохранить, отклонить или отменить действие, даже если с момента последнего сохранения элемента не вносились какие-либо изменения.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-915">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="0b5b9-916">Если в клиенте Outlook для настольных ПК сообщение представляет собой ответ в тексте, метод `close` не работает.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-916">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="0b5b9-917">Требования</span><span class="sxs-lookup"><span data-stu-id="0b5b9-917">Requirements</span></span>

|<span data-ttu-id="0b5b9-918">Требование</span><span class="sxs-lookup"><span data-stu-id="0b5b9-918">Requirement</span></span>|<span data-ttu-id="0b5b9-919">Значение</span><span class="sxs-lookup"><span data-stu-id="0b5b9-919">Value</span></span>|
|---|---|
|[<span data-ttu-id="0b5b9-920">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="0b5b9-920">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0b5b9-921">1.3</span><span class="sxs-lookup"><span data-stu-id="0b5b9-921">1.3</span></span>|
|[<span data-ttu-id="0b5b9-922">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="0b5b9-922">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0b5b9-923">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="0b5b9-923">Restricted</span></span>|
|[<span data-ttu-id="0b5b9-924">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="0b5b9-924">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0b5b9-925">Создание</span><span class="sxs-lookup"><span data-stu-id="0b5b9-925">Compose</span></span>|

---
---

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="0b5b9-926">displayReplyAllForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="0b5b9-926">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="0b5b9-927">Отображает форму ответа, включающую отправителя и всех получателей выбранного сообщения или организатора и всех участников выбранной встречи.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-927">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="0b5b9-928">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-928">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="0b5b9-929">В Outlook Web App форма ответа отображается в виде всплывающей формы в представлении с 3 либо 1 или 2 колонками.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-929">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="0b5b9-930">Если любой строковый параметр превышает указанные для него ограничения, `displayReplyAllForm` возвращает исключение.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-930">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="0b5b9-p145">Если в параметре `formData.attachments` указаны вложения, Outlook и Outlook Web App пытаются скачать их и вложить в форму ответа. Если какие-либо вложения добавить не удается, в форме отображается сообщение об ошибке. Если сообщения об ошибках не предусмотрены, то они не отображаются.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-p145">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0b5b9-934">Параметры</span><span class="sxs-lookup"><span data-stu-id="0b5b9-934">Parameters</span></span>

|<span data-ttu-id="0b5b9-935">Имя</span><span class="sxs-lookup"><span data-stu-id="0b5b9-935">Name</span></span>|<span data-ttu-id="0b5b9-936">Тип</span><span class="sxs-lookup"><span data-stu-id="0b5b9-936">Type</span></span>|<span data-ttu-id="0b5b9-937">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="0b5b9-937">Attributes</span></span>|<span data-ttu-id="0b5b9-938">Описание</span><span class="sxs-lookup"><span data-stu-id="0b5b9-938">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="0b5b9-939">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="0b5b9-939">String &#124; Object</span></span>||<span data-ttu-id="0b5b9-p146">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-p146">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="0b5b9-942">**ИЛИ**</span><span class="sxs-lookup"><span data-stu-id="0b5b9-942">**OR**</span></span><br/><span data-ttu-id="0b5b9-p147">Объект, который содержит текст или данные вложения и функцию обратного вызова. Ниже представлено определение этого объекта.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-p147">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="0b5b9-945">String</span><span class="sxs-lookup"><span data-stu-id="0b5b9-945">String</span></span>|<span data-ttu-id="0b5b9-946">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="0b5b9-946">&lt;optional&gt;</span></span>|<span data-ttu-id="0b5b9-p148">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-p148">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="0b5b9-949">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="0b5b9-949">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="0b5b9-950">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="0b5b9-950">&lt;optional&gt;</span></span>|<span data-ttu-id="0b5b9-951">Массив объектов JSON, представляющих собой вложенные файлы или элементы.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-951">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="0b5b9-952">String</span><span class="sxs-lookup"><span data-stu-id="0b5b9-952">String</span></span>||<span data-ttu-id="0b5b9-p149">Указывает тип вложения. Допустимые значения: `file` для вложенного файла и `item` для вложенного элемента.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-p149">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="0b5b9-955">Строка</span><span class="sxs-lookup"><span data-stu-id="0b5b9-955">String</span></span>||<span data-ttu-id="0b5b9-956">Строка, содержащая имя вложения, длиной до 255 символов.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-956">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="0b5b9-957">String</span><span class="sxs-lookup"><span data-stu-id="0b5b9-957">String</span></span>||<span data-ttu-id="0b5b9-p150">Используется, только если свойству `type` задано значение `file`. URI расположения файла.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-p150">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="0b5b9-960">Логический</span><span class="sxs-lookup"><span data-stu-id="0b5b9-960">Boolean</span></span>||<span data-ttu-id="0b5b9-p151">Используется, только если свойству `type` задано значение `file`. Значение `true` указывает на то, что вложение будет встроено в текст сообщения и не должно отображаться в списке вложений.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-p151">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="0b5b9-963">String</span><span class="sxs-lookup"><span data-stu-id="0b5b9-963">String</span></span>||<span data-ttu-id="0b5b9-p152">Используется, только если свойству `type` задано значение `item`. Идентификатор вложения EWS. Это строка длиной до 100 символов.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-p152">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="0b5b9-967">function</span><span class="sxs-lookup"><span data-stu-id="0b5b9-967">function</span></span>|<span data-ttu-id="0b5b9-968">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="0b5b9-968">&lt;optional&gt;</span></span>|<span data-ttu-id="0b5b9-969">По завершении работы метода функция, переданная параметру `callback`, вызывается с помощью одного параметра `asyncResult`, представляющего собой объект [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="0b5b9-969">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0b5b9-970">Требования</span><span class="sxs-lookup"><span data-stu-id="0b5b9-970">Requirements</span></span>

|<span data-ttu-id="0b5b9-971">Требование</span><span class="sxs-lookup"><span data-stu-id="0b5b9-971">Requirement</span></span>|<span data-ttu-id="0b5b9-972">Значение</span><span class="sxs-lookup"><span data-stu-id="0b5b9-972">Value</span></span>|
|---|---|
|[<span data-ttu-id="0b5b9-973">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="0b5b9-973">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0b5b9-974">1.0</span><span class="sxs-lookup"><span data-stu-id="0b5b9-974">1.0</span></span>|
|[<span data-ttu-id="0b5b9-975">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="0b5b9-975">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0b5b9-976">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0b5b9-976">ReadItem</span></span>|
|[<span data-ttu-id="0b5b9-977">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="0b5b9-977">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0b5b9-978">Чтение</span><span class="sxs-lookup"><span data-stu-id="0b5b9-978">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="0b5b9-979">Примеры</span><span class="sxs-lookup"><span data-stu-id="0b5b9-979">Examples</span></span>

<span data-ttu-id="0b5b9-980">Приведенный ниже код передает строку в функцию `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-980">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="0b5b9-981">Ответ с пустым текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-981">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="0b5b9-982">Ответ только с текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-982">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="0b5b9-983">Ответ с текстом сообщения и вложенным файлом.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-983">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="0b5b9-984">Ответ с текстом сообщения и вложенным элементом.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-984">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="0b5b9-985">Ответ с текстом сообщения, вложенным файлом, вложенным элементом и обратным вызовом.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-985">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="0b5b9-986">displayReplyForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="0b5b9-986">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="0b5b9-987">Отображает форму ответа, включающую только отправителя выбранного сообщения или организатора выбранной встречи.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-987">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="0b5b9-988">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-988">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="0b5b9-989">В Outlook Web App форма ответа отображается в виде всплывающей формы в представлении с 3 либо 1 или 2 колонками.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-989">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="0b5b9-990">Если любой строковый параметр превышает указанные для него ограничения, `displayReplyForm` возвращает исключение.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-990">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="0b5b9-p153">Если в параметре `formData.attachments` указаны вложения, Outlook и Outlook Web App пытаются скачать их и вложить в форму ответа. Если какие-либо вложения добавить не удается, в форме отображается сообщение об ошибке. Если сообщения об ошибках не предусмотрены, то они не отображаются.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-p153">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0b5b9-994">Параметры</span><span class="sxs-lookup"><span data-stu-id="0b5b9-994">Parameters</span></span>

|<span data-ttu-id="0b5b9-995">Имя</span><span class="sxs-lookup"><span data-stu-id="0b5b9-995">Name</span></span>|<span data-ttu-id="0b5b9-996">Тип</span><span class="sxs-lookup"><span data-stu-id="0b5b9-996">Type</span></span>|<span data-ttu-id="0b5b9-997">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="0b5b9-997">Attributes</span></span>|<span data-ttu-id="0b5b9-998">Описание</span><span class="sxs-lookup"><span data-stu-id="0b5b9-998">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="0b5b9-999">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="0b5b9-999">String &#124; Object</span></span>||<span data-ttu-id="0b5b9-p154">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-p154">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="0b5b9-1002">**ИЛИ**</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1002">**OR**</span></span><br/><span data-ttu-id="0b5b9-p155">Объект, который содержит текст или данные вложения и функцию обратного вызова. Ниже представлено определение этого объекта.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-p155">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="0b5b9-1005">String</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1005">String</span></span>|<span data-ttu-id="0b5b9-1006">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1006">&lt;optional&gt;</span></span>|<span data-ttu-id="0b5b9-p156">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-p156">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="0b5b9-1009">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1009">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="0b5b9-1010">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1010">&lt;optional&gt;</span></span>|<span data-ttu-id="0b5b9-1011">Массив объектов JSON, представляющих собой вложенные файлы или элементы.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1011">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="0b5b9-1012">String</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1012">String</span></span>||<span data-ttu-id="0b5b9-p157">Указывает тип вложения. Допустимые значения: `file` для вложенного файла и `item` для вложенного элемента.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-p157">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="0b5b9-1015">Строка</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1015">String</span></span>||<span data-ttu-id="0b5b9-1016">Строка, содержащая имя вложения, длиной до 255 символов.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1016">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="0b5b9-1017">String</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1017">String</span></span>||<span data-ttu-id="0b5b9-p158">Используется, только если свойству `type` задано значение `file`. URI расположения файла.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-p158">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="0b5b9-1020">Логический</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1020">Boolean</span></span>||<span data-ttu-id="0b5b9-p159">Используется, только если свойству `type` задано значение `file`. Значение `true` указывает на то, что вложение будет встроено в текст сообщения и не должно отображаться в списке вложений.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-p159">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="0b5b9-1023">String</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1023">String</span></span>||<span data-ttu-id="0b5b9-p160">Используется, только если свойству `type` задано значение `item`. Идентификатор вложения EWS. Это строка длиной до 100 символов.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-p160">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="0b5b9-1027">function</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1027">function</span></span>|<span data-ttu-id="0b5b9-1028">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1028">&lt;optional&gt;</span></span>|<span data-ttu-id="0b5b9-1029">По завершении работы метода функция, переданная параметру `callback`, вызывается с помощью одного параметра `asyncResult`, представляющего собой объект [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1029">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0b5b9-1030">Требования</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1030">Requirements</span></span>

|<span data-ttu-id="0b5b9-1031">Требование</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1031">Requirement</span></span>|<span data-ttu-id="0b5b9-1032">Значение</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1032">Value</span></span>|
|---|---|
|[<span data-ttu-id="0b5b9-1033">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1033">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0b5b9-1034">1.0</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1034">1.0</span></span>|
|[<span data-ttu-id="0b5b9-1035">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1035">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0b5b9-1036">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1036">ReadItem</span></span>|
|[<span data-ttu-id="0b5b9-1037">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1037">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0b5b9-1038">Чтение</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1038">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="0b5b9-1039">Примеры</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1039">Examples</span></span>

<span data-ttu-id="0b5b9-1040">Приведенный ниже код передает строку в функцию `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1040">The following code passes a string to the `displayReplyForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="0b5b9-1041">Ответ с пустым текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1041">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="0b5b9-1042">Ответ только с текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1042">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="0b5b9-1043">Ответ с текстом сообщения и вложенным файлом.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1043">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="0b5b9-1044">Ответ с текстом сообщения и вложенным элементом.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1044">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="0b5b9-1045">Ответ с текстом сообщения, вложенным файлом, вложенным элементом и обратным вызовом.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1045">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getattachmentcontentasyncattachmentid-options-callback--attachmentcontentjavascriptapioutlookofficeattachmentcontent"></a><span data-ttu-id="0b5b9-1046">Жетаттачментконтентасинк (attachmentId, [параметры], [callback]) → [вложениеимеет содержимое](/javascript/api/outlook/office.attachmentcontent)</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1046">getAttachmentContentAsync(attachmentId, [options], [callback]) → [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span></span>

<span data-ttu-id="0b5b9-1047">Получает указанное вложение из сообщения или встречи и возвращает его в виде `AttachmentContent` объекта.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1047">Gets the specified attachment from a message or appointment and returns it as an `AttachmentContent` object.</span></span>

<span data-ttu-id="0b5b9-1048">`getAttachmentContentAsync` Метод получает вложение с указанным идентификатором из элемента.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1048">The `getAttachmentContentAsync` method gets the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="0b5b9-1049">Рекомендуется использовать идентификатор для получения вложения в том же сеансе, когда Аттачментидс был получен с помощью вызова `getAttachmentsAsync` или. `item.attachments`</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1049">As a best practice, you should use the identifier to retrieve an attachment in the same session that the attachmentIds were retrieved with the `getAttachmentsAsync` or `item.attachments` call.</span></span> <span data-ttu-id="0b5b9-1050">В Outlook Web App и Outlook Web App для устройств идентификатор вложения действителен только в рамках одного сеанса.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1050">In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="0b5b9-1051">Сеанс переходит к моменту, когда пользователь закрывает приложение, или если пользователь начинает создание встроенной формы, затем извлекает форму, чтобы продолжить работу в отдельном окне.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1051">A session is over when the user closes the app, or if the user starts composing an inline form then subsequently pops out the form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0b5b9-1052">Параметры</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1052">Parameters</span></span>

|<span data-ttu-id="0b5b9-1053">Имя</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1053">Name</span></span>|<span data-ttu-id="0b5b9-1054">Тип</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1054">Type</span></span>|<span data-ttu-id="0b5b9-1055">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1055">Attributes</span></span>|<span data-ttu-id="0b5b9-1056">Описание</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1056">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="0b5b9-1057">String</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1057">String</span></span>||<span data-ttu-id="0b5b9-1058">Идентификатор вложения, которое требуется получить.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1058">The identifier of the attachment you want to get.</span></span>|
|`options`|<span data-ttu-id="0b5b9-1059">Объект</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1059">Object</span></span>|<span data-ttu-id="0b5b9-1060">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1060">&lt;optional&gt;</span></span>|<span data-ttu-id="0b5b9-1061">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1061">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="0b5b9-1062">Объект</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1062">Object</span></span>|<span data-ttu-id="0b5b9-1063">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1063">&lt;optional&gt;</span></span>|<span data-ttu-id="0b5b9-1064">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1064">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="0b5b9-1065">функция</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1065">function</span></span>|<span data-ttu-id="0b5b9-1066">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1066">&lt;optional&gt;</span></span>|<span data-ttu-id="0b5b9-1067">По завершении работы метода функция, переданная параметру `callback`, вызывается с помощью одного параметра `asyncResult`, представляющего собой объект [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1067">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0b5b9-1068">Требования</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1068">Requirements</span></span>

|<span data-ttu-id="0b5b9-1069">Требование</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1069">Requirement</span></span>|<span data-ttu-id="0b5b9-1070">Значение</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1070">Value</span></span>|
|---|---|
|[<span data-ttu-id="0b5b9-1071">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1071">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0b5b9-1072">Предварительная версия</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1072">Preview</span></span>|
|[<span data-ttu-id="0b5b9-1073">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1073">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0b5b9-1074">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1074">ReadItem</span></span>|
|[<span data-ttu-id="0b5b9-1075">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1075">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0b5b9-1076">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1076">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="0b5b9-1077">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1077">Returns:</span></span>

<span data-ttu-id="0b5b9-1078">Тип: [вложениеимеет содержимое](/javascript/api/outlook/office.attachmentcontent)</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1078">Type: [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span></span>

##### <a name="example"></a><span data-ttu-id="0b5b9-1079">Пример</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1079">Example</span></span>

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

#### <a name="getattachmentsasyncoptions-callback--arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails"></a><span data-ttu-id="0b5b9-1080">Жетаттачментсасинк ([параметры], [обратный вызов]) → массив. <[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="0b5b9-1080">getAttachmentsAsync([options], [callback]) → Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

<span data-ttu-id="0b5b9-1081">Получает вложения элемента в виде массива.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1081">Gets the item's attachments as an array.</span></span> <span data-ttu-id="0b5b9-1082">Только в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1082">Compose mode only.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0b5b9-1083">Параметры</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1083">Parameters</span></span>

|<span data-ttu-id="0b5b9-1084">Имя</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1084">Name</span></span>|<span data-ttu-id="0b5b9-1085">Тип</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1085">Type</span></span>|<span data-ttu-id="0b5b9-1086">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1086">Attributes</span></span>|<span data-ttu-id="0b5b9-1087">Описание</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1087">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="0b5b9-1088">Object</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1088">Object</span></span>|<span data-ttu-id="0b5b9-1089">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1089">&lt;optional&gt;</span></span>|<span data-ttu-id="0b5b9-1090">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1090">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="0b5b9-1091">Объект</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1091">Object</span></span>|<span data-ttu-id="0b5b9-1092">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1092">&lt;optional&gt;</span></span>|<span data-ttu-id="0b5b9-1093">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1093">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="0b5b9-1094">функция</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1094">function</span></span>|<span data-ttu-id="0b5b9-1095">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1095">&lt;optional&gt;</span></span>|<span data-ttu-id="0b5b9-1096">По завершении работы метода функция, переданная параметру `callback`, вызывается с помощью одного параметра `asyncResult`, представляющего собой объект [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1096">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0b5b9-1097">Требования</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1097">Requirements</span></span>

|<span data-ttu-id="0b5b9-1098">Требование</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1098">Requirement</span></span>|<span data-ttu-id="0b5b9-1099">Значение</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1099">Value</span></span>|
|---|---|
|[<span data-ttu-id="0b5b9-1100">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1100">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0b5b9-1101">Предварительная версия</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1101">Preview</span></span>|
|[<span data-ttu-id="0b5b9-1102">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1102">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0b5b9-1103">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1103">ReadItem</span></span>|
|[<span data-ttu-id="0b5b9-1104">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1104">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0b5b9-1105">Создание</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1105">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="0b5b9-1106">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1106">Returns:</span></span>

<span data-ttu-id="0b5b9-1107">Тип: Array. <[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="0b5b9-1107">Type: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

##### <a name="example"></a><span data-ttu-id="0b5b9-1108">Пример</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1108">Example</span></span>

<span data-ttu-id="0b5b9-1109">В приведенном ниже примере создается строка HTML со сведениями обо всех вложениях в текущем элементе.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1109">The following example builds an HTML string with details of all attachments on the current item.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlookofficeentities"></a><span data-ttu-id="0b5b9-1110">getEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1110">getEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span></span>

<span data-ttu-id="0b5b9-1111">Получает сущности, обнаруженные в теле выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1111">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="0b5b9-1112">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1112">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="0b5b9-1113">Требования</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1113">Requirements</span></span>

|<span data-ttu-id="0b5b9-1114">Требование</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1114">Requirement</span></span>|<span data-ttu-id="0b5b9-1115">Значение</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1115">Value</span></span>|
|---|---|
|[<span data-ttu-id="0b5b9-1116">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1116">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0b5b9-1117">1.0</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1117">1.0</span></span>|
|[<span data-ttu-id="0b5b9-1118">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1118">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0b5b9-1119">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1119">ReadItem</span></span>|
|[<span data-ttu-id="0b5b9-1120">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1120">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0b5b9-1121">Чтение</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1121">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="0b5b9-1122">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1122">Returns:</span></span>

<span data-ttu-id="0b5b9-1123">Тип: [Entities](/javascript/api/outlook/office.entities)</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1123">Type: [Entities](/javascript/api/outlook/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="0b5b9-1124">Пример</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1124">Example</span></span>

<span data-ttu-id="0b5b9-1125">Ниже приведен пример получения доступа к сущностям контактов в теле текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1125">The following example accesses the contacts entities in the current item's body.</span></span>

```javascript
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

---
---

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion"></a><span data-ttu-id="0b5b9-1126">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1126">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span></span>

<span data-ttu-id="0b5b9-1127">Получает массив всех сущностей указанного типа, обнаруженных в теле выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1127">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="0b5b9-1128">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1128">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0b5b9-1129">Параметры</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1129">Parameters</span></span>

|<span data-ttu-id="0b5b9-1130">Имя</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1130">Name</span></span>|<span data-ttu-id="0b5b9-1131">Тип</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1131">Type</span></span>|<span data-ttu-id="0b5b9-1132">Описание</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1132">Description</span></span>|
|---|---|---|
|`entityType`|[<span data-ttu-id="0b5b9-1133">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1133">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook/office.mailboxenums.entitytype)|<span data-ttu-id="0b5b9-1134">Одно из значений перечисления EntityType.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1134">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0b5b9-1135">Требования</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1135">Requirements</span></span>

|<span data-ttu-id="0b5b9-1136">Требование</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1136">Requirement</span></span>|<span data-ttu-id="0b5b9-1137">Значение</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1137">Value</span></span>|
|---|---|
|[<span data-ttu-id="0b5b9-1138">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1138">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0b5b9-1139">1.0</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1139">1.0</span></span>|
|[<span data-ttu-id="0b5b9-1140">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1140">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0b5b9-1141">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1141">Restricted</span></span>|
|[<span data-ttu-id="0b5b9-1142">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1142">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0b5b9-1143">Чтение</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1143">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="0b5b9-1144">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1144">Returns:</span></span>

<span data-ttu-id="0b5b9-1145">Если значение, переданное в `entityType`, не является допустимым членом перечисления `EntityType`, метод возвращает значение NULL.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1145">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="0b5b9-1146">Если в теле элемента отсутствуют сущности указанного типа, метод возвращает пустой массив.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1146">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="0b5b9-1147">В противном случае тип объектов в возвращаемом массиве зависит от типа сущности, запрошенной в параметре `entityType`.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1147">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="0b5b9-1148">Хотя минимальный уровень разрешений для использования этого метода — **Restricted**, для некоторых типов сущностей требуется доступ на уровне **ReadItem**, как указано в приведенной ниже таблице.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1148">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

|<span data-ttu-id="0b5b9-1149">Значение параметра `entityType`</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1149">Value of `entityType`</span></span>|<span data-ttu-id="0b5b9-1150">Тип объектов в возвращаемом массиве</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1150">Type of objects in returned array</span></span>|<span data-ttu-id="0b5b9-1151">Необходимый уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1151">Required Permission Level</span></span>|
|---|---|---|
|`Address`|<span data-ttu-id="0b5b9-1152">String</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1152">String</span></span>|<span data-ttu-id="0b5b9-1153">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1153">**Restricted**</span></span>|
|`Contact`|<span data-ttu-id="0b5b9-1154">Contact</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1154">Contact</span></span>|<span data-ttu-id="0b5b9-1155">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1155">**ReadItem**</span></span>|
|`EmailAddress`|<span data-ttu-id="0b5b9-1156">String</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1156">String</span></span>|<span data-ttu-id="0b5b9-1157">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1157">**ReadItem**</span></span>|
|`MeetingSuggestion`|<span data-ttu-id="0b5b9-1158">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1158">MeetingSuggestion</span></span>|<span data-ttu-id="0b5b9-1159">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1159">**ReadItem**</span></span>|
|`PhoneNumber`|<span data-ttu-id="0b5b9-1160">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1160">PhoneNumber</span></span>|<span data-ttu-id="0b5b9-1161">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1161">**Restricted**</span></span>|
|`TaskSuggestion`|<span data-ttu-id="0b5b9-1162">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1162">TaskSuggestion</span></span>|<span data-ttu-id="0b5b9-1163">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1163">**ReadItem**</span></span>|
|`URL`|<span data-ttu-id="0b5b9-1164">String</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1164">String</span></span>|<span data-ttu-id="0b5b9-1165">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1165">**Restricted**</span></span>|

<span data-ttu-id="0b5b9-1166">Тип: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="0b5b9-1166">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="0b5b9-1167">Пример</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1167">Example</span></span>

<span data-ttu-id="0b5b9-1168">В примере ниже показано, как получить доступ к массиву строк, которые представляют собой почтовые адреса в теле текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1168">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion"></a><span data-ttu-id="0b5b9-1169">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1169">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span></span>

<span data-ttu-id="0b5b9-1170">Возвращает известные сущности в выбранном элементе, которые проходят через именованный фильтр, определяемый в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1170">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="0b5b9-1171">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1171">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="0b5b9-1172">Метод `getFilteredEntitiesByName` возвращает сущности, соответствующие регулярному выражению, которое определяется в элементе правила [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) в XML-файле манифеста, с использованием указанного значения элемента `FilterName`.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1172">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0b5b9-1173">Параметры</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1173">Parameters</span></span>

|<span data-ttu-id="0b5b9-1174">Имя</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1174">Name</span></span>|<span data-ttu-id="0b5b9-1175">Тип</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1175">Type</span></span>|<span data-ttu-id="0b5b9-1176">Описание</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1176">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="0b5b9-1177">String</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1177">String</span></span>|<span data-ttu-id="0b5b9-1178">Имя элемента правила `ItemHasKnownEntity`, определяющее соответствующий фильтр.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1178">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0b5b9-1179">Требования</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1179">Requirements</span></span>

|<span data-ttu-id="0b5b9-1180">Требование</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1180">Requirement</span></span>|<span data-ttu-id="0b5b9-1181">Значение</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1181">Value</span></span>|
|---|---|
|[<span data-ttu-id="0b5b9-1182">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1182">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0b5b9-1183">1.0</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1183">1.0</span></span>|
|[<span data-ttu-id="0b5b9-1184">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1184">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0b5b9-1185">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1185">ReadItem</span></span>|
|[<span data-ttu-id="0b5b9-1186">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1186">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0b5b9-1187">Чтение</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1187">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="0b5b9-1188">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1188">Returns:</span></span>

<span data-ttu-id="0b5b9-p164">Если в манифесте нет элемента `ItemHasKnownEntity` со значением `FilterName`, соответствующим параметру `name`, метод возвращает `null`. Если параметр `name` соответствует элементу `ItemHasKnownEntity` в манифесте, но при этом в текущем элементе нет соответствующих сущностей, метод возвращает пустой массив.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-p164">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="0b5b9-1191">Тип: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="0b5b9-1191">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span></span>

---
---

#### <a name="getinitializationcontextasyncoptions-callback"></a><span data-ttu-id="0b5b9-1192">getInitializationContextAsync ([параметры], [обратный вызов])</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1192">getInitializationContextAsync([options], [callback])</span></span>

<span data-ttu-id="0b5b9-1193">Получает данные инициализации, передаваемые при активации надстройки [сообщением с действиями](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1193">Gets initialization data passed when the add-in is [activated by an actionable message](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span></span>

> [!NOTE]
> <span data-ttu-id="0b5b9-1194">Этот метод поддерживается только в Outlook 2016 или более поздней версии для Windows ("нажми и работай" более поздней версии, чем 16.0.8413.1000) и Outlook в Интернете для Office 365.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1194">This method is only supported by Outlook 2016 or later on Windows (Click-to-Run versions later than 16.0.8413.1000) and Outlook on the web for Office 365.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0b5b9-1195">Параметры</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1195">Parameters</span></span>

|<span data-ttu-id="0b5b9-1196">Имя</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1196">Name</span></span>|<span data-ttu-id="0b5b9-1197">Тип</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1197">Type</span></span>|<span data-ttu-id="0b5b9-1198">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1198">Attributes</span></span>|<span data-ttu-id="0b5b9-1199">Описание</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1199">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="0b5b9-1200">Object</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1200">Object</span></span>|<span data-ttu-id="0b5b9-1201">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1201">&lt;optional&gt;</span></span>|<span data-ttu-id="0b5b9-1202">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1202">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="0b5b9-1203">Объект</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1203">Object</span></span>|<span data-ttu-id="0b5b9-1204">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1204">&lt;optional&gt;</span></span>|<span data-ttu-id="0b5b9-1205">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1205">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="0b5b9-1206">функция</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1206">function</span></span>|<span data-ttu-id="0b5b9-1207">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1207">&lt;optional&gt;</span></span>|<span data-ttu-id="0b5b9-1208">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1208">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="0b5b9-1209">При успешном выполнении данные инициализации предоставляются в `asyncResult.value` свойстве в виде строки.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1209">On success, the initialization data is provided in the `asyncResult.value` property as a string.</span></span><br/><span data-ttu-id="0b5b9-1210">Если `asyncResult` контекст инициализации отсутствует, объект будет содержать `Error` объект со `code` свойством, `9020` `name` для свойства которого задано значение. `GenericResponseError`</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1210">If there is no initialization context, the `asyncResult` object will contain an `Error` object with its `code` property set to `9020` and its `name` property set to `GenericResponseError`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0b5b9-1211">Требования</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1211">Requirements</span></span>

|<span data-ttu-id="0b5b9-1212">Требование</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1212">Requirement</span></span>|<span data-ttu-id="0b5b9-1213">Значение</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1213">Value</span></span>|
|---|---|
|[<span data-ttu-id="0b5b9-1214">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1214">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0b5b9-1215">Предварительная версия</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1215">Preview</span></span>|
|[<span data-ttu-id="0b5b9-1216">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1216">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0b5b9-1217">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1217">ReadItem</span></span>|
|[<span data-ttu-id="0b5b9-1218">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1218">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0b5b9-1219">Чтение</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1219">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0b5b9-1220">Пример</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1220">Example</span></span>

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

#### <a name="getitemidasyncoptions-callback"></a><span data-ttu-id="0b5b9-1221">Жетитемидасинк ([параметры], обратный вызов)</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1221">getItemIdAsync([options], callback)</span></span>

<span data-ttu-id="0b5b9-1222">Асинхронно получает идентификатор сохраненного элемента.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1222">Asynchronously gets the ID of a saved item.</span></span> <span data-ttu-id="0b5b9-1223">Только в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1223">Compose mode only.</span></span>

<span data-ttu-id="0b5b9-1224">При вызове этот метод возвращает идентификатор элемента с помощью метода обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1224">When invoked, this method returns the item ID via the callback method.</span></span>

> [!NOTE]
> <span data-ttu-id="0b5b9-1225">Если надстройка вызывает `getItemIdAsync` элемент в режиме создания (например, чтобы получить доступ `itemId` к использованию с помощью EWS или REST API), имейте в виду, что если Outlook находится в режиме кэширования, может потребоваться некоторое время до синхронизации элемента с сервером.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1225">If your add-in calls `getItemIdAsync` on an item in compose mode (e.g., to get an `itemId` to use with EWS or the REST API), be aware that when Outlook is in cached mode, it may take some time before the item is synced to the server.</span></span> <span data-ttu-id="0b5b9-1226">Пока элемент не будет синхронизирован, он не `itemId` распознается и не будет использоваться, возвращается сообщение об ошибке.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1226">Until the item is synced, the `itemId` is not recognized and using it returns an error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0b5b9-1227">Параметры</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1227">Parameters</span></span>

|<span data-ttu-id="0b5b9-1228">Имя</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1228">Name</span></span>|<span data-ttu-id="0b5b9-1229">Тип</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1229">Type</span></span>|<span data-ttu-id="0b5b9-1230">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1230">Attributes</span></span>|<span data-ttu-id="0b5b9-1231">Описание</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1231">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="0b5b9-1232">Object</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1232">Object</span></span>|<span data-ttu-id="0b5b9-1233">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1233">&lt;optional&gt;</span></span>|<span data-ttu-id="0b5b9-1234">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1234">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="0b5b9-1235">Объект</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1235">Object</span></span>|<span data-ttu-id="0b5b9-1236">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1236">&lt;optional&gt;</span></span>|<span data-ttu-id="0b5b9-1237">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1237">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="0b5b9-1238">функция</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1238">function</span></span>||<span data-ttu-id="0b5b9-1239">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1239">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="0b5b9-1240">После успешного выполнения идентификатор элемента будет предоставлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1240">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="errors"></a><span data-ttu-id="0b5b9-1241">Ошибки</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1241">Errors</span></span>

|<span data-ttu-id="0b5b9-1242">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1242">Error code</span></span>|<span data-ttu-id="0b5b9-1243">Описание</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1243">Description</span></span>|
|------------|-------------|
|`ItemNotSaved`|<span data-ttu-id="0b5b9-1244">Идентификатор невозможно извлечь, пока не будет сохранен элемент.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1244">The id can't be retrieved until the item is saved.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0b5b9-1245">Требования</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1245">Requirements</span></span>

|<span data-ttu-id="0b5b9-1246">Требование</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1246">Requirement</span></span>|<span data-ttu-id="0b5b9-1247">Значение</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1247">Value</span></span>|
|---|---|
|[<span data-ttu-id="0b5b9-1248">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1248">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0b5b9-1249">Предварительная версия</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1249">Preview</span></span>|
|[<span data-ttu-id="0b5b9-1250">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1250">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0b5b9-1251">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1251">ReadItem</span></span>|
|[<span data-ttu-id="0b5b9-1252">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1252">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0b5b9-1253">Создание</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1253">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="0b5b9-1254">Примеры</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1254">Examples</span></span>

```javascript
Office.context.mailbox.item.getItemIdAsync(
  function callback(result) {
    // Process the result.
  });
```

<span data-ttu-id="0b5b9-1255">В следующем примере показана структура `result` параметра, переданного функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1255">The following example shows the structure of the `result` parameter that's passed to the callback function.</span></span> <span data-ttu-id="0b5b9-1256">`value` Свойство содержит идентификатор элемента.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1256">The `value` property contains the item ID.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

---
---

#### <a name="getregexmatches--object"></a><span data-ttu-id="0b5b9-1257">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1257">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="0b5b9-1258">Возвращает строковые значения в выбранном элементе, которые соответствуют регулярным выражениям, определенным в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1258">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="0b5b9-1259">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1259">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="0b5b9-p168">Метод `getRegExMatches` возвращает строки, соответствующие регулярному выражению, которое определяется в каждом элементе правила `ItemHasRegularExpressionMatch` или `ItemHasKnownEntity` в XML-файле манифеста. Для правила `ItemHasRegularExpressionMatch` соответствующую строку должно содержать свойство элемента, указанного этим правилом. Простой тип `PropertyName` определяет поддерживаемые свойства.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-p168">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="0b5b9-1263">Например, рассмотрим манифест надстройки, который содержит указанный ниже элемент `Rule`.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1263">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="0b5b9-1264">Объект, возвращаемый методом `getRegExMatches`, будет содержать два свойства: `fruits` и `veggies`.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1264">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="0b5b9-p169">Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты. Лучше используйте метод [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) для этого.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-p169">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="0b5b9-1268">Requirements</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1268">Requirements</span></span>

|<span data-ttu-id="0b5b9-1269">Требование</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1269">Requirement</span></span>|<span data-ttu-id="0b5b9-1270">Значение</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1270">Value</span></span>|
|---|---|
|[<span data-ttu-id="0b5b9-1271">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1271">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0b5b9-1272">1.0</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1272">1.0</span></span>|
|[<span data-ttu-id="0b5b9-1273">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1273">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0b5b9-1274">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1274">ReadItem</span></span>|
|[<span data-ttu-id="0b5b9-1275">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1275">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0b5b9-1276">Чтение</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1276">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="0b5b9-1277">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1277">Returns:</span></span>

<span data-ttu-id="0b5b9-p170">Объект, содержащий массив строк, которые соответствуют регулярным выражениям, определяемым в XML-файле манифеста. Имя каждого массива равно соответствующему значению атрибута `RegExName` подходящего правила `ItemHasRegularExpressionMatch` или атрибута `FilterName` соответствующего правила `ItemHasKnownEntity`.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-p170">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="0b5b9-1280">Тип:</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1280">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="0b5b9-1281">Object</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1281">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="0b5b9-1282">Пример</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1282">Example</span></span>

<span data-ttu-id="0b5b9-1283">В приведенном ниже примере показано, как получить доступ к массиву совпадений с элементами `fruits` и `veggies` правил активации регулярных выражений, указанными в манифесте.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1283">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```javascript
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

---
---

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="0b5b9-1284">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1284">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="0b5b9-1285">Возвращает строковые значения в выбранном элементе, которые соответствуют именованному регулярному выражению, определенному в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1285">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="0b5b9-1286">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1286">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="0b5b9-1287">Метод `getRegExMatchesByName` возвращает строки, соответствующие регулярному выражению, которое определяется в элементе правила `ItemHasRegularExpressionMatch` в XML-файле манифеста, с использованием указанного значения элемента `RegExName`.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1287">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="0b5b9-p171">Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-p171">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0b5b9-1290">Параметры</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1290">Parameters</span></span>

|<span data-ttu-id="0b5b9-1291">Имя</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1291">Name</span></span>|<span data-ttu-id="0b5b9-1292">Тип</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1292">Type</span></span>|<span data-ttu-id="0b5b9-1293">Описание</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1293">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="0b5b9-1294">String</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1294">String</span></span>|<span data-ttu-id="0b5b9-1295">Имя элемента правила `ItemHasRegularExpressionMatch`, определяющее соответствующий фильтр.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1295">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0b5b9-1296">Требования</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1296">Requirements</span></span>

|<span data-ttu-id="0b5b9-1297">Требование</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1297">Requirement</span></span>|<span data-ttu-id="0b5b9-1298">Значение</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1298">Value</span></span>|
|---|---|
|[<span data-ttu-id="0b5b9-1299">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1299">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0b5b9-1300">1.0</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1300">1.0</span></span>|
|[<span data-ttu-id="0b5b9-1301">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1301">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0b5b9-1302">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1302">ReadItem</span></span>|
|[<span data-ttu-id="0b5b9-1303">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1303">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0b5b9-1304">Чтение</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1304">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="0b5b9-1305">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1305">Returns:</span></span>

<span data-ttu-id="0b5b9-1306">Массив строк, соответствующих регулярному выражению, определяемому в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1306">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="0b5b9-1307">

<dt>Тип</dt>

</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1307">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="0b5b9-1308">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="0b5b9-1308">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="0b5b9-1309">Пример</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1309">Example</span></span>

```javascript
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

---
---

#### <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="0b5b9-1310">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1310">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="0b5b9-1311">Асинхронно возвращает данные, выбранные в теме или тексте сообщения.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1311">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="0b5b9-p172">Если выделенный фрагмент отсутствует, но курсор находится в тексте или теме, метод возвращает значение NULL для выбранных данных. Если выбраны не текст и не тема, метод возвращает ошибку `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-p172">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0b5b9-1314">Параметры</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1314">Parameters</span></span>

|<span data-ttu-id="0b5b9-1315">Имя</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1315">Name</span></span>|<span data-ttu-id="0b5b9-1316">Тип</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1316">Type</span></span>|<span data-ttu-id="0b5b9-1317">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1317">Attributes</span></span>|<span data-ttu-id="0b5b9-1318">Описание</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1318">Description</span></span>|
|---|---|---|---|
|`coercionType`|[<span data-ttu-id="0b5b9-1319">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1319">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="0b5b9-p173">Запрашивает формат данных. Если задано значение Text, метод возвращает обычный текст как строку, удаляя все имеющиеся HTML-теги. Если задано значение HTML, метод возвращает выделенный текст (обычный текст или HTML).</span><span class="sxs-lookup"><span data-stu-id="0b5b9-p173">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`|<span data-ttu-id="0b5b9-1323">Объект</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1323">Object</span></span>|<span data-ttu-id="0b5b9-1324">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1324">&lt;optional&gt;</span></span>|<span data-ttu-id="0b5b9-1325">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1325">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="0b5b9-1326">Объект</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1326">Object</span></span>|<span data-ttu-id="0b5b9-1327">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1327">&lt;optional&gt;</span></span>|<span data-ttu-id="0b5b9-1328">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1328">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="0b5b9-1329">функция</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1329">function</span></span>||<span data-ttu-id="0b5b9-1330">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1330">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="0b5b9-1331">Чтобы получить доступ к выбранным данным из метода обратного вызова, вызовите `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1331">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="0b5b9-1332">Для доступа к исходному свойству, представляющему собой источник выбранных данных, вызовите параметр `asyncResult.value.sourceProperty`, который может иметь значение `body` или `subject`.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1332">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0b5b9-1333">Требования</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1333">Requirements</span></span>

|<span data-ttu-id="0b5b9-1334">Требование</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1334">Requirement</span></span>|<span data-ttu-id="0b5b9-1335">Значение</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1335">Value</span></span>|
|---|---|
|[<span data-ttu-id="0b5b9-1336">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1336">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0b5b9-1337">1.2</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1337">1.2</span></span>|
|[<span data-ttu-id="0b5b9-1338">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1338">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0b5b9-1339">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1339">ReadWriteItem</span></span>|
|[<span data-ttu-id="0b5b9-1340">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1340">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0b5b9-1341">Создание</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1341">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="0b5b9-1342">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1342">Returns:</span></span>

<span data-ttu-id="0b5b9-1343">Выбранные данные в виде строки с форматом, определенным в параметре `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1343">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="0b5b9-1344">

<dt>Тип</dt>

</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1344">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="0b5b9-1345">String</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1345">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="0b5b9-1346">Пример</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1346">Example</span></span>

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

#### <a name="getselectedentities--entitiesjavascriptapioutlookofficeentities"></a><span data-ttu-id="0b5b9-1347">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1347">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span></span>

<span data-ttu-id="0b5b9-1348">Возвращает сущности, найденные в выделенном совпадении, выбранном пользователем.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1348">Gets the entities found in a highlighted match a user has selected.</span></span> <span data-ttu-id="0b5b9-1349">Выделенные совпадения применяются к [контекстным надстройкам](/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1349">Highlighted matches apply to [contextual add-ins](/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="0b5b9-1350">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1350">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="0b5b9-1351">Требования</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1351">Requirements</span></span>

|<span data-ttu-id="0b5b9-1352">Требование</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1352">Requirement</span></span>|<span data-ttu-id="0b5b9-1353">Значение</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1353">Value</span></span>|
|---|---|
|[<span data-ttu-id="0b5b9-1354">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1354">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0b5b9-1355">1.6</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1355">1.6</span></span>|
|[<span data-ttu-id="0b5b9-1356">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1356">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0b5b9-1357">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1357">ReadItem</span></span>|
|[<span data-ttu-id="0b5b9-1358">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1358">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0b5b9-1359">Чтение</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1359">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="0b5b9-1360">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1360">Returns:</span></span>

<span data-ttu-id="0b5b9-1361">Тип: [Entities](/javascript/api/outlook/office.entities)</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1361">Type: [Entities](/javascript/api/outlook/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="0b5b9-1362">Пример</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1362">Example</span></span>

<span data-ttu-id="0b5b9-1363">В приведенном ниже примере показано, как получить доступ к сущностям адресов в выделенном совпадении, выбранном пользователем.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1363">The following example accesses the addresses entities in the highlighted match selected by the user.</span></span>

```javascript
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

---
---

#### <a name="getselectedregexmatches--object"></a><span data-ttu-id="0b5b9-1364">getSelectedRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1364">getSelectedRegExMatches() → {Object}</span></span>

<span data-ttu-id="0b5b9-p176">Возвращает строковые значения в выделенном совпадении, которые соответствуют регулярным выражениям, определенным в XML-файле манифеста. Выделенные совпадения применяются к [контекстным надстройкам](/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="0b5b9-p176">Returns string values in a highlighted match that match the regular expressions defined in the manifest XML file. Highlighted matches apply to [contextual add-ins](/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="0b5b9-1367">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1367">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="0b5b9-p177">Метод `getSelectedRegExMatches` возвращает строки, соответствующие регулярному выражению, которое определяется в каждом элементе правила `ItemHasRegularExpressionMatch` или `ItemHasKnownEntity` в XML-файле манифеста. Для правила `ItemHasRegularExpressionMatch` соответствующую строку должно содержать свойство элемента, указанного этим правилом. Простой тип `PropertyName` определяет поддерживаемые свойства.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-p177">The `getSelectedRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="0b5b9-1371">Например, рассмотрим манифест надстройки, который содержит указанный ниже элемент `Rule`.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1371">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="0b5b9-1372">Объект, возвращаемый методом `getRegExMatches`, будет содержать два свойства: `fruits` и `veggies`.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1372">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="0b5b9-p178">Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты. Лучше используйте метод [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) для этого.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-p178">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="0b5b9-1376">Requirements</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1376">Requirements</span></span>

|<span data-ttu-id="0b5b9-1377">Требование</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1377">Requirement</span></span>|<span data-ttu-id="0b5b9-1378">Значение</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1378">Value</span></span>|
|---|---|
|[<span data-ttu-id="0b5b9-1379">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1379">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0b5b9-1380">1.6</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1380">1.6</span></span>|
|[<span data-ttu-id="0b5b9-1381">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1381">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0b5b9-1382">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1382">ReadItem</span></span>|
|[<span data-ttu-id="0b5b9-1383">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1383">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0b5b9-1384">Чтение</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1384">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="0b5b9-1385">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1385">Returns:</span></span>

<span data-ttu-id="0b5b9-p179">Объект, содержащий массив строк, которые соответствуют регулярным выражениям, определяемым в XML-файле манифеста. Имя каждого массива равно соответствующему значению атрибута `RegExName` подходящего правила `ItemHasRegularExpressionMatch` или атрибута `FilterName` соответствующего правила `ItemHasKnownEntity`.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-p179">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

##### <a name="example"></a><span data-ttu-id="0b5b9-1388">Пример</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1388">Example</span></span>

<span data-ttu-id="0b5b9-1389">В приведенном ниже примере показано, как получить доступ к массиву совпадений с элементами `fruits` и `veggies` правил активации регулярных выражений, указанными в манифесте.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1389">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```javascript
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

---
---

#### <a name="getsharedpropertiesasyncoptions-callback"></a><span data-ttu-id="0b5b9-1390">Жетшаредпропертиесасинк ([параметры], обратный вызов)</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1390">getSharedPropertiesAsync([options], callback)</span></span>

<span data-ttu-id="0b5b9-1391">Получает свойства выбранной встречи или сообщения в общей папке, календаре или почтовом ящике.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1391">Gets the properties of the selected appointment or message in a shared folder, calendar, or mailbox.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0b5b9-1392">Параметры</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1392">Parameters</span></span>

|<span data-ttu-id="0b5b9-1393">Имя</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1393">Name</span></span>|<span data-ttu-id="0b5b9-1394">Тип</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1394">Type</span></span>|<span data-ttu-id="0b5b9-1395">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1395">Attributes</span></span>|<span data-ttu-id="0b5b9-1396">Описание</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1396">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="0b5b9-1397">Object</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1397">Object</span></span>|<span data-ttu-id="0b5b9-1398">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1398">&lt;optional&gt;</span></span>|<span data-ttu-id="0b5b9-1399">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1399">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="0b5b9-1400">Объект</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1400">Object</span></span>|<span data-ttu-id="0b5b9-1401">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1401">&lt;optional&gt;</span></span>|<span data-ttu-id="0b5b9-1402">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1402">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="0b5b9-1403">функция</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1403">function</span></span>||<span data-ttu-id="0b5b9-1404">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1404">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="0b5b9-1405">Общие свойства предоставляются в виде [`SharedProperties`](/javascript/api/outlook/office.sharedproperties) объекта в `asyncResult.value` свойстве.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1405">The shared properties are provided as a [`SharedProperties`](/javascript/api/outlook/office.sharedproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="0b5b9-1406">Этот объект можно использовать для получения общих свойств элемента.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1406">This object can be used to get the item's shared properties.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0b5b9-1407">Требования</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1407">Requirements</span></span>

|<span data-ttu-id="0b5b9-1408">Требование</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1408">Requirement</span></span>|<span data-ttu-id="0b5b9-1409">Значение</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1409">Value</span></span>|
|---|---|
|[<span data-ttu-id="0b5b9-1410">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1410">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0b5b9-1411">Предварительная версия</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1411">Preview</span></span>|
|[<span data-ttu-id="0b5b9-1412">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1412">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0b5b9-1413">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1413">ReadItem</span></span>|
|[<span data-ttu-id="0b5b9-1414">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1414">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0b5b9-1415">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1415">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0b5b9-1416">Пример</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1416">Example</span></span>

```javascript
Office.context.mailbox.item.getSharedPropertiesAsync(callback);

function callback (asyncResult) {
  var context = asyncResult.context;
  var sharedProperties = asyncResult.value;
}
```

---
---

#### <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="0b5b9-1417">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1417">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="0b5b9-1418">Асинхронно загружает настраиваемые свойства для надстройки для выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1418">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="0b5b9-p181">Настраиваемые свойства сохраняются в виде пар "ключ-значение" для каждого приложения и каждого элемента. Этот метод возвращает объект `CustomProperties` при обратном вызове, который предоставляет методы для доступа к настраиваемым свойствам, характерным для текущего элемента и текущей надстройки. Настраиваемые свойства не шифруются для элемента, поэтому этот способ хранения не является безопасным.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-p181">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0b5b9-1422">Параметры</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1422">Parameters</span></span>

|<span data-ttu-id="0b5b9-1423">Имя</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1423">Name</span></span>|<span data-ttu-id="0b5b9-1424">Тип</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1424">Type</span></span>|<span data-ttu-id="0b5b9-1425">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1425">Attributes</span></span>|<span data-ttu-id="0b5b9-1426">Описание</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1426">Description</span></span>|
|---|---|---|---|
|`callback`|<span data-ttu-id="0b5b9-1427">function</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1427">function</span></span>||<span data-ttu-id="0b5b9-1428">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1428">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="0b5b9-1429">Настраиваемые свойства предоставляются в виде объекта [`CustomProperties`](/javascript/api/outlook/office.customproperties) в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1429">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="0b5b9-1430">Этот объект позволяет получить, задать и удалить настраиваемые свойства для элемента, а также сохранить изменения, внесенные в набор настраиваемых свойств, на сервере.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1430">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`|<span data-ttu-id="0b5b9-1431">Объект</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1431">Object</span></span>|<span data-ttu-id="0b5b9-1432">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1432">&lt;optional&gt;</span></span>|<span data-ttu-id="0b5b9-1433">Разработчики могут указать любой объект, к которому необходимо получить доступ, в функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1433">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="0b5b9-1434">Доступ к этому объекту можно получить с помощью свойства `asyncResult.asyncContext` в функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1434">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0b5b9-1435">Требования</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1435">Requirements</span></span>

|<span data-ttu-id="0b5b9-1436">Требование</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1436">Requirement</span></span>|<span data-ttu-id="0b5b9-1437">Значение</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1437">Value</span></span>|
|---|---|
|[<span data-ttu-id="0b5b9-1438">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1438">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0b5b9-1439">1.0</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1439">1.0</span></span>|
|[<span data-ttu-id="0b5b9-1440">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1440">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0b5b9-1441">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1441">ReadItem</span></span>|
|[<span data-ttu-id="0b5b9-1442">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1442">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0b5b9-1443">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1443">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0b5b9-1444">Пример</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1444">Example</span></span>

<span data-ttu-id="0b5b9-p184">Приведенный ниже пример кода показывает, как асинхронно загружать настраиваемые свойства, характерные для текущего элемента, с помощью метода `loadCustomPropertiesAsync`. Этот пример также показывает, как сохранять эти свойства на сервере с помощью метода `CustomProperties.saveAsync`. После загрузки настраиваемых свойств в этом примере кода метод `CustomProperties.get` используется для считывания настраиваемого свойства `myProp`, метод `CustomProperties.set` — для записи настраиваемого свойства `otherProp`, а метод `saveAsync` — для сохранения настраиваемых свойств.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-p184">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

#### <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="0b5b9-1448">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1448">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="0b5b9-1449">Удаляет вложение из сообщения или встречи.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1449">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="0b5b9-1450">Метод `removeAttachmentAsync` удаляет из элемента вложение с указанным идентификатором.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1450">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="0b5b9-1451">Идентификатор вложения рекомендуется использовать для удаления вложения, только если оно добавлено тем же почтовым приложением в ходе текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1451">As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session.</span></span> <span data-ttu-id="0b5b9-1452">В Outlook Web App и Outlook Web App для устройств идентификатор вложения действителен только в рамках одного сеанса.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1452">In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="0b5b9-1453">Сеанс переходит к моменту, когда пользователь закрывает приложение, или если пользователь начинает создание встроенной формы, затем извлекает форму, чтобы продолжить работу в отдельном окне.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1453">A session is over when the user closes the app, or if the user starts composing an inline form then subsequently pops out the form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0b5b9-1454">Параметры</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1454">Parameters</span></span>

|<span data-ttu-id="0b5b9-1455">Имя</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1455">Name</span></span>|<span data-ttu-id="0b5b9-1456">Тип</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1456">Type</span></span>|<span data-ttu-id="0b5b9-1457">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1457">Attributes</span></span>|<span data-ttu-id="0b5b9-1458">Описание</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1458">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="0b5b9-1459">String</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1459">String</span></span>||<span data-ttu-id="0b5b9-1460">Идентификатор удаляемого вложения.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1460">The identifier of the attachment to remove.</span></span>|
|`options`|<span data-ttu-id="0b5b9-1461">Объект</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1461">Object</span></span>|<span data-ttu-id="0b5b9-1462">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1462">&lt;optional&gt;</span></span>|<span data-ttu-id="0b5b9-1463">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1463">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="0b5b9-1464">Объект</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1464">Object</span></span>|<span data-ttu-id="0b5b9-1465">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1465">&lt;optional&gt;</span></span>|<span data-ttu-id="0b5b9-1466">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1466">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="0b5b9-1467">функция</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1467">function</span></span>|<span data-ttu-id="0b5b9-1468">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1468">&lt;optional&gt;</span></span>|<span data-ttu-id="0b5b9-1469">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1469">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="0b5b9-1470">Если удалить вложение не удается, свойство `asyncResult.error` содержит код ошибки с указанием ее причины.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1470">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="0b5b9-1471">Ошибки</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1471">Errors</span></span>

|<span data-ttu-id="0b5b9-1472">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1472">Error code</span></span>|<span data-ttu-id="0b5b9-1473">Описание</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1473">Description</span></span>|
|------------|-------------|
|`InvalidAttachmentId`|<span data-ttu-id="0b5b9-1474">Идентификатор вложения не существует.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1474">The attachment identifier does not exist.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0b5b9-1475">Требования</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1475">Requirements</span></span>

|<span data-ttu-id="0b5b9-1476">Требование</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1476">Requirement</span></span>|<span data-ttu-id="0b5b9-1477">Значение</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1477">Value</span></span>|
|---|---|
|[<span data-ttu-id="0b5b9-1478">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1478">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0b5b9-1479">1.1</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1479">1.1</span></span>|
|[<span data-ttu-id="0b5b9-1480">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1480">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0b5b9-1481">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1481">ReadWriteItem</span></span>|
|[<span data-ttu-id="0b5b9-1482">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1482">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0b5b9-1483">Создание</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1483">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="0b5b9-1484">Пример</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1484">Example</span></span>

<span data-ttu-id="0b5b9-1485">Указанный ниже код удаляет вложение с идентификатором "0".</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1485">The following code removes an attachment with an identifier of '0'.</span></span>

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

#### <a name="removehandlerasynceventtype-options-callback"></a><span data-ttu-id="0b5b9-1486">removeHandlerAsync(eventType, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1486">removeHandlerAsync(eventType, [options], [callback])</span></span>

<span data-ttu-id="0b5b9-1487">Удаляет обработчиков для поддерживаемого типа события.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1487">Removes the event handlers for a supported event type.</span></span>

<span data-ttu-id="0b5b9-1488">В настоящее время поддерживаются типы `Office.EventType.AttachmentsChanged`событий `Office.EventType.AppointmentTimeChanged`, `Office.EventType.EnhancedLocationsChanged` `Office.EventType.RecipientsChanged`,, и `Office.EventType.RecurrenceChanged`.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1488">Currently the supported event types are `Office.EventType.AttachmentsChanged`, `Office.EventType.AppointmentTimeChanged`, `Office.EventType.EnhancedLocationsChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0b5b9-1489">Параметры</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1489">Parameters</span></span>

| <span data-ttu-id="0b5b9-1490">Имя</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1490">Name</span></span> | <span data-ttu-id="0b5b9-1491">Тип</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1491">Type</span></span> | <span data-ttu-id="0b5b9-1492">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1492">Attributes</span></span> | <span data-ttu-id="0b5b9-1493">Описание</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1493">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="0b5b9-1494">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1494">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="0b5b9-1495">Событие, которое должно отменить обработчик.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1495">The event that should revoke the handler.</span></span> |
| `options` | <span data-ttu-id="0b5b9-1496">Объект</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1496">Object</span></span> | <span data-ttu-id="0b5b9-1497">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1497">&lt;optional&gt;</span></span> | <span data-ttu-id="0b5b9-1498">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1498">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="0b5b9-1499">Объект</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1499">Object</span></span> | <span data-ttu-id="0b5b9-1500">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1500">&lt;optional&gt;</span></span> | <span data-ttu-id="0b5b9-1501">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1501">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="0b5b9-1502">функция</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1502">function</span></span>| <span data-ttu-id="0b5b9-1503">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1503">&lt;optional&gt;</span></span>|<span data-ttu-id="0b5b9-1504">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1504">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0b5b9-1505">Требования</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1505">Requirements</span></span>

|<span data-ttu-id="0b5b9-1506">Требование</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1506">Requirement</span></span>| <span data-ttu-id="0b5b9-1507">Значение</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1507">Value</span></span>|
|---|---|
|[<span data-ttu-id="0b5b9-1508">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1508">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0b5b9-1509">1.7</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1509">1.7</span></span> |
|[<span data-ttu-id="0b5b9-1510">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1510">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0b5b9-1511">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1511">ReadItem</span></span> |
|[<span data-ttu-id="0b5b9-1512">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1512">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0b5b9-1513">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1513">Compose or Read</span></span> |

---
---

#### <a name="saveasyncoptions-callback"></a><span data-ttu-id="0b5b9-1514">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1514">saveAsync([options], callback)</span></span>

<span data-ttu-id="0b5b9-1515">Асинхронно сохраняет элемент.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1515">Asynchronously saves an item.</span></span>

<span data-ttu-id="0b5b9-p186">При вызове этот метод сохраняет текущее сообщение в виде черновика и возвращает идентификатор элемента с помощью метода обратного вызова. В Outlook Web App или интерактивном режиме Outlook этот элемент сохраняется на сервере. В Outlook в режиме кэширования этот элемент сохраняется в локальном кэше.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-p186">When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="0b5b9-1519">Если в режиме создания надстройка вызывает для элемента метод `saveAsync`, чтобы получить параметр `itemId` для использования с EWS или REST API, необходимо помнить, что синхронизация элемента с сервером может занять много времени, если Outlook работает в режиме кэширования данных.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1519">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="0b5b9-1520">До окончания синхронизации элемента применение параметра `itemId` будет приводить к ошибке.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1520">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="0b5b9-p188">Если метод `saveAsync` вызывается для встречи в режиме создания, она сохраняется как обычная встреча в календаре пользователя, а не как черновик. При сохранении новой встречи приглашения не отправляются. При сохранении существующей встречи уведомления отправляются добавленным или удаленным участникам.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-p188">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="0b5b9-1524">Следующие клиенты отличаются другим поведением `saveAsync` в отношении встреч в режиме создания:</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1524">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="0b5b9-1525">Outlook для Mac не поддерживает сохранение собраний.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1525">Outlook for Mac does not support saving a meeting.</span></span> <span data-ttu-id="0b5b9-1526">`saveAsync` Метод завершается с ошибкой при вызове из собрания в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1526">The `saveAsync` method fails when called from a meeting in compose mode.</span></span> <span data-ttu-id="0b5b9-1527">Просмотреть [не удается сохранить собрание в виде черновика в Outlook для Mac с помощью API Office JS](https://support.microsoft.com/help/4505745) для обхода.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1527">See [Cannot save a meeting as a draft in Outlook for Mac by using Office JS API](https://support.microsoft.com/help/4505745) for a workaround.</span></span>
> - <span data-ttu-id="0b5b9-1528">Outlook в Интернете всегда отправляет приглашение или обновление при вызове `saveAsync` для встречи в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1528">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0b5b9-1529">Параметры</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1529">Parameters</span></span>

|<span data-ttu-id="0b5b9-1530">Имя</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1530">Name</span></span>|<span data-ttu-id="0b5b9-1531">Тип</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1531">Type</span></span>|<span data-ttu-id="0b5b9-1532">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1532">Attributes</span></span>|<span data-ttu-id="0b5b9-1533">Описание</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1533">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="0b5b9-1534">Объект</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1534">Object</span></span>|<span data-ttu-id="0b5b9-1535">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1535">&lt;optional&gt;</span></span>|<span data-ttu-id="0b5b9-1536">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1536">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="0b5b9-1537">Объект</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1537">Object</span></span>|<span data-ttu-id="0b5b9-1538">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1538">&lt;optional&gt;</span></span>|<span data-ttu-id="0b5b9-1539">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1539">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="0b5b9-1540">функция</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1540">function</span></span>||<span data-ttu-id="0b5b9-1541">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1541">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="0b5b9-1542">После успешного выполнения идентификатор элемента будет предоставлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1542">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0b5b9-1543">Требования</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1543">Requirements</span></span>

|<span data-ttu-id="0b5b9-1544">Требование</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1544">Requirement</span></span>|<span data-ttu-id="0b5b9-1545">Значение</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1545">Value</span></span>|
|---|---|
|[<span data-ttu-id="0b5b9-1546">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1546">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0b5b9-1547">1.3</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1547">1.3</span></span>|
|[<span data-ttu-id="0b5b9-1548">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1548">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0b5b9-1549">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1549">ReadWriteItem</span></span>|
|[<span data-ttu-id="0b5b9-1550">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1550">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0b5b9-1551">Создание</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1551">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="0b5b9-1552">Примеры</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1552">Examples</span></span>

```javascript
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result.
  });
```

<span data-ttu-id="0b5b9-p190">Ниже приведен пример параметра `result`, переданного функции обратного вызова. Свойство `value` содержит идентификатор элемента.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-p190">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

---
---

#### <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="0b5b9-1555">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1555">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="0b5b9-1556">Асинхронно вставляет данные в текст или тему сообщения.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1556">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="0b5b9-p191">Метод `setSelectedDataAsync` вставляет указанную строку в местоположение курсора в теме или тексте элемента либо, если текст выделен в редакторе, он заменяет выделенный текст. Если курсор находится вне текста или темы элемента, возвращается ошибка. После вставки курсор помещается в конец вставленного содержимого.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-p191">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0b5b9-1560">Параметры</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1560">Parameters</span></span>

|<span data-ttu-id="0b5b9-1561">Имя</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1561">Name</span></span>|<span data-ttu-id="0b5b9-1562">Тип</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1562">Type</span></span>|<span data-ttu-id="0b5b9-1563">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1563">Attributes</span></span>|<span data-ttu-id="0b5b9-1564">Описание</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1564">Description</span></span>|
|---|---|---|---|
|`data`|<span data-ttu-id="0b5b9-1565">String</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1565">String</span></span>||<span data-ttu-id="0b5b9-p192">Вставляемые данные. Объем данных не должен превышать 1 000 000 символов. Если передано больше 1 000 000 символов, возвращается исключение `ArgumentOutOfRange`.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-p192">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`|<span data-ttu-id="0b5b9-1569">Object</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1569">Object</span></span>|<span data-ttu-id="0b5b9-1570">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1570">&lt;optional&gt;</span></span>|<span data-ttu-id="0b5b9-1571">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1571">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="0b5b9-1572">Объект</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1572">Object</span></span>|<span data-ttu-id="0b5b9-1573">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1573">&lt;optional&gt;</span></span>|<span data-ttu-id="0b5b9-1574">В методе обратного вызова разработчики могут указать любой объект, к которому необходимо получить доступ.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1574">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="0b5b9-1575">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1575">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="0b5b9-1576">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1576">&lt;optional&gt;</span></span>|<span data-ttu-id="0b5b9-p193">Если задано значение `text`, текущий стиль применяется в Outlook Web App и Outlook. Если поле представляет собой редактор HTML, вставляются только текстовые данные, даже если они имеют формат HTML.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-p193">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="0b5b9-p194">Если задано значение `html` и поле (не тема) поддерживает HTML, в Outlook Web App применяется текущий стиль, а в Outlook — стиль по умолчанию. Если поле является текстовым, возвращается ошибка `InvalidDataFormat`.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-p194">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="0b5b9-1581">Если свойство `coercionType` не задано, результат зависит от поля: если поле имеет формат HTML, используется текст в формате HTML, а если поле текстовое, применяется обычный текст.</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1581">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`|<span data-ttu-id="0b5b9-1582">функция</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1582">function</span></span>||<span data-ttu-id="0b5b9-1583">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1583">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0b5b9-1584">Требования</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1584">Requirements</span></span>

|<span data-ttu-id="0b5b9-1585">Требование</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1585">Requirement</span></span>|<span data-ttu-id="0b5b9-1586">Значение</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1586">Value</span></span>|
|---|---|
|[<span data-ttu-id="0b5b9-1587">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1587">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0b5b9-1588">1.2</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1588">1.2</span></span>|
|[<span data-ttu-id="0b5b9-1589">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1589">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0b5b9-1590">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1590">ReadWriteItem</span></span>|
|[<span data-ttu-id="0b5b9-1591">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1591">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0b5b9-1592">Создание</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1592">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="0b5b9-1593">Пример</span><span class="sxs-lookup"><span data-stu-id="0b5b9-1593">Example</span></span>

```javascript
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
