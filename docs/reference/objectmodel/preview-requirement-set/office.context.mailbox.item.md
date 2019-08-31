---
title: Office. Context. Mailbox. Item — Предварительная версия набора требований
description: ''
ms.date: 08/30/2019
localization_priority: Normal
ms.openlocfilehash: 9939d939e7b1de7af71d7b5532dcf306330e5b6e
ms.sourcegitcommit: 1fb99b1b4e63868a0e81a928c69a34c42bf7e209
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/30/2019
ms.locfileid: "36696500"
---
# <a name="item"></a><span data-ttu-id="e5c37-102">item</span><span class="sxs-lookup"><span data-stu-id="e5c37-102">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="e5c37-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="e5c37-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="e5c37-p101">Пространство имен `item` используется для доступа к выбранному в данный момент сообщению, приглашению на собрание или описанию встречи. Вы можете определить тип пространства имен `item` с помощью свойства [itemType](#itemtype-officemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="e5c37-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="e5c37-106">Requirements</span><span class="sxs-lookup"><span data-stu-id="e5c37-106">Requirements</span></span>

|<span data-ttu-id="e5c37-107">Требование</span><span class="sxs-lookup"><span data-stu-id="e5c37-107">Requirement</span></span>|<span data-ttu-id="e5c37-108">Значение</span><span class="sxs-lookup"><span data-stu-id="e5c37-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="e5c37-109">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="e5c37-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e5c37-110">1.0</span><span class="sxs-lookup"><span data-stu-id="e5c37-110">1.0</span></span>|
|[<span data-ttu-id="e5c37-111">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e5c37-111">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e5c37-112">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="e5c37-112">Restricted</span></span>|
|[<span data-ttu-id="e5c37-113">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e5c37-113">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="e5c37-114">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="e5c37-114">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="e5c37-115">Элементы и методы</span><span class="sxs-lookup"><span data-stu-id="e5c37-115">Members and methods</span></span>

| <span data-ttu-id="e5c37-116">Элемент</span><span class="sxs-lookup"><span data-stu-id="e5c37-116">Member</span></span> | <span data-ttu-id="e5c37-117">Тип</span><span class="sxs-lookup"><span data-stu-id="e5c37-117">Type</span></span> |
|--------|------|
| [<span data-ttu-id="e5c37-118">attachments</span><span class="sxs-lookup"><span data-stu-id="e5c37-118">attachments</span></span>](#attachments-arrayattachmentdetails) | <span data-ttu-id="e5c37-119">Элемент</span><span class="sxs-lookup"><span data-stu-id="e5c37-119">Member</span></span> |
| [<span data-ttu-id="e5c37-120">bcc</span><span class="sxs-lookup"><span data-stu-id="e5c37-120">bcc</span></span>](#bcc-recipients) | <span data-ttu-id="e5c37-121">Элемент</span><span class="sxs-lookup"><span data-stu-id="e5c37-121">Member</span></span> |
| [<span data-ttu-id="e5c37-122">body</span><span class="sxs-lookup"><span data-stu-id="e5c37-122">body</span></span>](#body-body) | <span data-ttu-id="e5c37-123">Элемент</span><span class="sxs-lookup"><span data-stu-id="e5c37-123">Member</span></span> |
| [<span data-ttu-id="e5c37-124">разделов</span><span class="sxs-lookup"><span data-stu-id="e5c37-124">categories</span></span>](#categories-categories) | <span data-ttu-id="e5c37-125">Элемент</span><span class="sxs-lookup"><span data-stu-id="e5c37-125">Member</span></span> |
| [<span data-ttu-id="e5c37-126">cc</span><span class="sxs-lookup"><span data-stu-id="e5c37-126">cc</span></span>](#cc-arrayemailaddressdetailsrecipients) | <span data-ttu-id="e5c37-127">Элемент</span><span class="sxs-lookup"><span data-stu-id="e5c37-127">Member</span></span> |
| [<span data-ttu-id="e5c37-128">conversationId</span><span class="sxs-lookup"><span data-stu-id="e5c37-128">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="e5c37-129">Элемент</span><span class="sxs-lookup"><span data-stu-id="e5c37-129">Member</span></span> |
| [<span data-ttu-id="e5c37-130">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="e5c37-130">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="e5c37-131">Элемент</span><span class="sxs-lookup"><span data-stu-id="e5c37-131">Member</span></span> |
| [<span data-ttu-id="e5c37-132">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="e5c37-132">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="e5c37-133">Элемент</span><span class="sxs-lookup"><span data-stu-id="e5c37-133">Member</span></span> |
| [<span data-ttu-id="e5c37-134">end</span><span class="sxs-lookup"><span data-stu-id="e5c37-134">end</span></span>](#end-datetime) | <span data-ttu-id="e5c37-135">Элемент</span><span class="sxs-lookup"><span data-stu-id="e5c37-135">Member</span></span> |
| [<span data-ttu-id="e5c37-136">енханцедлокатион</span><span class="sxs-lookup"><span data-stu-id="e5c37-136">enhancedLocation</span></span>](#enhancedlocation-enhancedlocation) | <span data-ttu-id="e5c37-137">Элемент</span><span class="sxs-lookup"><span data-stu-id="e5c37-137">Member</span></span> |
| [<span data-ttu-id="e5c37-138">from</span><span class="sxs-lookup"><span data-stu-id="e5c37-138">from</span></span>](#from-emailaddressdetailsfrom) | <span data-ttu-id="e5c37-139">Элемент</span><span class="sxs-lookup"><span data-stu-id="e5c37-139">Member</span></span> |
| [<span data-ttu-id="e5c37-140">Internetheaders:</span><span class="sxs-lookup"><span data-stu-id="e5c37-140">internetHeaders</span></span>](#internetheaders-internetheaders) | <span data-ttu-id="e5c37-141">Элемент</span><span class="sxs-lookup"><span data-stu-id="e5c37-141">Member</span></span> |
| [<span data-ttu-id="e5c37-142">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="e5c37-142">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="e5c37-143">Элемент</span><span class="sxs-lookup"><span data-stu-id="e5c37-143">Member</span></span> |
| [<span data-ttu-id="e5c37-144">itemClass</span><span class="sxs-lookup"><span data-stu-id="e5c37-144">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="e5c37-145">Элемент</span><span class="sxs-lookup"><span data-stu-id="e5c37-145">Member</span></span> |
| [<span data-ttu-id="e5c37-146">itemId</span><span class="sxs-lookup"><span data-stu-id="e5c37-146">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="e5c37-147">Элемент</span><span class="sxs-lookup"><span data-stu-id="e5c37-147">Member</span></span> |
| [<span data-ttu-id="e5c37-148">itemType</span><span class="sxs-lookup"><span data-stu-id="e5c37-148">itemType</span></span>](#itemtype-officemailboxenumsitemtype) | <span data-ttu-id="e5c37-149">Элемент</span><span class="sxs-lookup"><span data-stu-id="e5c37-149">Member</span></span> |
| [<span data-ttu-id="e5c37-150">location</span><span class="sxs-lookup"><span data-stu-id="e5c37-150">location</span></span>](#location-stringlocation) | <span data-ttu-id="e5c37-151">Элемент</span><span class="sxs-lookup"><span data-stu-id="e5c37-151">Member</span></span> |
| [<span data-ttu-id="e5c37-152">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="e5c37-152">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="e5c37-153">Элемент</span><span class="sxs-lookup"><span data-stu-id="e5c37-153">Member</span></span> |
| [<span data-ttu-id="e5c37-154">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="e5c37-154">notificationMessages</span></span>](#notificationmessages-notificationmessages) | <span data-ttu-id="e5c37-155">Member</span><span class="sxs-lookup"><span data-stu-id="e5c37-155">Member</span></span> |
| [<span data-ttu-id="e5c37-156">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="e5c37-156">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="e5c37-157">Элемент</span><span class="sxs-lookup"><span data-stu-id="e5c37-157">Member</span></span> |
| [<span data-ttu-id="e5c37-158">organizer</span><span class="sxs-lookup"><span data-stu-id="e5c37-158">organizer</span></span>](#organizer-emailaddressdetailsorganizer) | <span data-ttu-id="e5c37-159">Элемент</span><span class="sxs-lookup"><span data-stu-id="e5c37-159">Member</span></span> |
| [<span data-ttu-id="e5c37-160">recurrence</span><span class="sxs-lookup"><span data-stu-id="e5c37-160">recurrence</span></span>](#nullable-recurrence-recurrence) | <span data-ttu-id="e5c37-161">Элемент</span><span class="sxs-lookup"><span data-stu-id="e5c37-161">Member</span></span> |
| [<span data-ttu-id="e5c37-162">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="e5c37-162">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="e5c37-163">Элемент</span><span class="sxs-lookup"><span data-stu-id="e5c37-163">Member</span></span> |
| [<span data-ttu-id="e5c37-164">sender</span><span class="sxs-lookup"><span data-stu-id="e5c37-164">sender</span></span>](#sender-emailaddressdetails) | <span data-ttu-id="e5c37-165">Member</span><span class="sxs-lookup"><span data-stu-id="e5c37-165">Member</span></span> |
| [<span data-ttu-id="e5c37-166">seriesId</span><span class="sxs-lookup"><span data-stu-id="e5c37-166">seriesId</span></span>](#nullable-seriesid-string) | <span data-ttu-id="e5c37-167">Member</span><span class="sxs-lookup"><span data-stu-id="e5c37-167">Member</span></span> |
| [<span data-ttu-id="e5c37-168">start</span><span class="sxs-lookup"><span data-stu-id="e5c37-168">start</span></span>](#start-datetime) | <span data-ttu-id="e5c37-169">Member</span><span class="sxs-lookup"><span data-stu-id="e5c37-169">Member</span></span> |
| [<span data-ttu-id="e5c37-170">subject</span><span class="sxs-lookup"><span data-stu-id="e5c37-170">subject</span></span>](#subject-stringsubject) | <span data-ttu-id="e5c37-171">Элемент</span><span class="sxs-lookup"><span data-stu-id="e5c37-171">Member</span></span> |
| [<span data-ttu-id="e5c37-172">to</span><span class="sxs-lookup"><span data-stu-id="e5c37-172">to</span></span>](#to-arrayemailaddressdetailsrecipients) | <span data-ttu-id="e5c37-173">Элемент</span><span class="sxs-lookup"><span data-stu-id="e5c37-173">Member</span></span> |
| [<span data-ttu-id="e5c37-174">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="e5c37-174">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="e5c37-175">Метод</span><span class="sxs-lookup"><span data-stu-id="e5c37-175">Method</span></span> |
| [<span data-ttu-id="e5c37-176">addFileAttachmentFromBase64Async</span><span class="sxs-lookup"><span data-stu-id="e5c37-176">addFileAttachmentFromBase64Async</span></span>](#addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback) | <span data-ttu-id="e5c37-177">Метод</span><span class="sxs-lookup"><span data-stu-id="e5c37-177">Method</span></span> |
| [<span data-ttu-id="e5c37-178">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="e5c37-178">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="e5c37-179">Метод</span><span class="sxs-lookup"><span data-stu-id="e5c37-179">Method</span></span> |
| [<span data-ttu-id="e5c37-180">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="e5c37-180">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="e5c37-181">Метод</span><span class="sxs-lookup"><span data-stu-id="e5c37-181">Method</span></span> |
| [<span data-ttu-id="e5c37-182">close</span><span class="sxs-lookup"><span data-stu-id="e5c37-182">close</span></span>](#close) | <span data-ttu-id="e5c37-183">Метод</span><span class="sxs-lookup"><span data-stu-id="e5c37-183">Method</span></span> |
| [<span data-ttu-id="e5c37-184">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="e5c37-184">displayReplyAllForm</span></span>](#displayreplyallformformdata-callback) | <span data-ttu-id="e5c37-185">Метод</span><span class="sxs-lookup"><span data-stu-id="e5c37-185">Method</span></span> |
| [<span data-ttu-id="e5c37-186">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="e5c37-186">displayReplyForm</span></span>](#displayreplyformformdata-callback) | <span data-ttu-id="e5c37-187">Метод</span><span class="sxs-lookup"><span data-stu-id="e5c37-187">Method</span></span> |
| [<span data-ttu-id="e5c37-188">жетаттачментконтентасинк</span><span class="sxs-lookup"><span data-stu-id="e5c37-188">getAttachmentContentAsync</span></span>](#getattachmentcontentasyncattachmentid-options-callback--attachmentcontent) | <span data-ttu-id="e5c37-189">Метод</span><span class="sxs-lookup"><span data-stu-id="e5c37-189">Method</span></span> |
| [<span data-ttu-id="e5c37-190">жетаттачментсасинк</span><span class="sxs-lookup"><span data-stu-id="e5c37-190">getAttachmentsAsync</span></span>](#getattachmentsasyncoptions-callback--arrayattachmentdetails) | <span data-ttu-id="e5c37-191">Метод</span><span class="sxs-lookup"><span data-stu-id="e5c37-191">Method</span></span> |
| [<span data-ttu-id="e5c37-192">getEntities</span><span class="sxs-lookup"><span data-stu-id="e5c37-192">getEntities</span></span>](#getentities--entities) | <span data-ttu-id="e5c37-193">Метод</span><span class="sxs-lookup"><span data-stu-id="e5c37-193">Method</span></span> |
| [<span data-ttu-id="e5c37-194">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="e5c37-194">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="e5c37-195">Метод</span><span class="sxs-lookup"><span data-stu-id="e5c37-195">Method</span></span> |
| [<span data-ttu-id="e5c37-196">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="e5c37-196">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="e5c37-197">Метод</span><span class="sxs-lookup"><span data-stu-id="e5c37-197">Method</span></span> |
| [<span data-ttu-id="e5c37-198">getInitializationContextAsync</span><span class="sxs-lookup"><span data-stu-id="e5c37-198">getInitializationContextAsync</span></span>](#getinitializationcontextasyncoptions-callback) | <span data-ttu-id="e5c37-199">Метод</span><span class="sxs-lookup"><span data-stu-id="e5c37-199">Method</span></span> |
| [<span data-ttu-id="e5c37-200">жетитемидасинк</span><span class="sxs-lookup"><span data-stu-id="e5c37-200">getItemIdAsync</span></span>](#getitemidasyncoptions-callback) | <span data-ttu-id="e5c37-201">Метод</span><span class="sxs-lookup"><span data-stu-id="e5c37-201">Method</span></span> |
| [<span data-ttu-id="e5c37-202">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="e5c37-202">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="e5c37-203">Метод</span><span class="sxs-lookup"><span data-stu-id="e5c37-203">Method</span></span> |
| [<span data-ttu-id="e5c37-204">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="e5c37-204">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="e5c37-205">Метод</span><span class="sxs-lookup"><span data-stu-id="e5c37-205">Method</span></span> |
| [<span data-ttu-id="e5c37-206">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="e5c37-206">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="e5c37-207">Метод</span><span class="sxs-lookup"><span data-stu-id="e5c37-207">Method</span></span> |
| [<span data-ttu-id="e5c37-208">жетселектедентитиес</span><span class="sxs-lookup"><span data-stu-id="e5c37-208">getSelectedEntities</span></span>](#getselectedentities--entities) | <span data-ttu-id="e5c37-209">Метод</span><span class="sxs-lookup"><span data-stu-id="e5c37-209">Method</span></span> |
| [<span data-ttu-id="e5c37-210">жетселектедрежексматчес</span><span class="sxs-lookup"><span data-stu-id="e5c37-210">getSelectedRegExMatches</span></span>](#getselectedregexmatches--object) | <span data-ttu-id="e5c37-211">Метод</span><span class="sxs-lookup"><span data-stu-id="e5c37-211">Method</span></span> |
| [<span data-ttu-id="e5c37-212">жетшаредпропертиесасинк</span><span class="sxs-lookup"><span data-stu-id="e5c37-212">getSharedPropertiesAsync</span></span>](#getsharedpropertiesasyncoptions-callback) | <span data-ttu-id="e5c37-213">Метод</span><span class="sxs-lookup"><span data-stu-id="e5c37-213">Method</span></span> |
| [<span data-ttu-id="e5c37-214">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="e5c37-214">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="e5c37-215">Метод</span><span class="sxs-lookup"><span data-stu-id="e5c37-215">Method</span></span> |
| [<span data-ttu-id="e5c37-216">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="e5c37-216">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="e5c37-217">Метод</span><span class="sxs-lookup"><span data-stu-id="e5c37-217">Method</span></span> |
| [<span data-ttu-id="e5c37-218">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="e5c37-218">removeHandlerAsync</span></span>](#removehandlerasynceventtype-options-callback) | <span data-ttu-id="e5c37-219">Метод</span><span class="sxs-lookup"><span data-stu-id="e5c37-219">Method</span></span> |
| [<span data-ttu-id="e5c37-220">saveAsync</span><span class="sxs-lookup"><span data-stu-id="e5c37-220">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="e5c37-221">Метод</span><span class="sxs-lookup"><span data-stu-id="e5c37-221">Method</span></span> |
| [<span data-ttu-id="e5c37-222">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="e5c37-222">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="e5c37-223">Метод</span><span class="sxs-lookup"><span data-stu-id="e5c37-223">Method</span></span> |

### <a name="example"></a><span data-ttu-id="e5c37-224">Пример</span><span class="sxs-lookup"><span data-stu-id="e5c37-224">Example</span></span>

<span data-ttu-id="e5c37-225">В примере кода JavaScript, приведенном ниже, показано, как получить доступ к свойству `subject` текущего элемента в Outlook.</span><span class="sxs-lookup"><span data-stu-id="e5c37-225">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="e5c37-226">Элементы</span><span class="sxs-lookup"><span data-stu-id="e5c37-226">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails"></a><span data-ttu-id="e5c37-227">вложения: Array. <[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="e5c37-227">attachments: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

<span data-ttu-id="e5c37-228">Получает вложения элемента в виде массива.</span><span class="sxs-lookup"><span data-stu-id="e5c37-228">Gets the item's attachments as an array.</span></span> <span data-ttu-id="e5c37-229">Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="e5c37-229">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="e5c37-230">Outlook блокирует определенные типы файлов из-за потенциальных проблем с безопасностью, поэтому они не возвращаются.</span><span class="sxs-lookup"><span data-stu-id="e5c37-230">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="e5c37-231">Дополнительные сведения см. в статье [Блокировка вложений в Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="e5c37-231">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="e5c37-232">Тип</span><span class="sxs-lookup"><span data-stu-id="e5c37-232">Type</span></span>

*   <span data-ttu-id="e5c37-233">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="e5c37-233">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="e5c37-234">Требования</span><span class="sxs-lookup"><span data-stu-id="e5c37-234">Requirements</span></span>

|<span data-ttu-id="e5c37-235">Требование</span><span class="sxs-lookup"><span data-stu-id="e5c37-235">Requirement</span></span>|<span data-ttu-id="e5c37-236">Значение</span><span class="sxs-lookup"><span data-stu-id="e5c37-236">Value</span></span>|
|---|---|
|[<span data-ttu-id="e5c37-237">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="e5c37-237">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e5c37-238">1.0</span><span class="sxs-lookup"><span data-stu-id="e5c37-238">1.0</span></span>|
|[<span data-ttu-id="e5c37-239">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e5c37-239">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e5c37-240">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e5c37-240">ReadItem</span></span>|
|[<span data-ttu-id="e5c37-241">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e5c37-241">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="e5c37-242">Чтение</span><span class="sxs-lookup"><span data-stu-id="e5c37-242">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e5c37-243">Пример</span><span class="sxs-lookup"><span data-stu-id="e5c37-243">Example</span></span>

<span data-ttu-id="e5c37-244">С помощью приведенного ниже кода можно создать HTML-строку с подробными сведениями обо всех вложениях для текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="e5c37-244">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

#### <a name="bcc-recipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="e5c37-245">СК: [получатели](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="e5c37-245">bcc: [Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="e5c37-246">Получает объект, который предоставляет методы для получения или обновления получателей скрытой копии сообщения.</span><span class="sxs-lookup"><span data-stu-id="e5c37-246">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="e5c37-247">Только в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="e5c37-247">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="e5c37-248">Тип</span><span class="sxs-lookup"><span data-stu-id="e5c37-248">Type</span></span>

*   [<span data-ttu-id="e5c37-249">Получатели</span><span class="sxs-lookup"><span data-stu-id="e5c37-249">Recipients</span></span>](/javascript/api/outlook/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="e5c37-250">Требования</span><span class="sxs-lookup"><span data-stu-id="e5c37-250">Requirements</span></span>

|<span data-ttu-id="e5c37-251">Требование</span><span class="sxs-lookup"><span data-stu-id="e5c37-251">Requirement</span></span>|<span data-ttu-id="e5c37-252">Значение</span><span class="sxs-lookup"><span data-stu-id="e5c37-252">Value</span></span>|
|---|---|
|[<span data-ttu-id="e5c37-253">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="e5c37-253">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e5c37-254">1.1</span><span class="sxs-lookup"><span data-stu-id="e5c37-254">1.1</span></span>|
|[<span data-ttu-id="e5c37-255">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e5c37-255">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e5c37-256">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e5c37-256">ReadItem</span></span>|
|[<span data-ttu-id="e5c37-257">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e5c37-257">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="e5c37-258">Создание</span><span class="sxs-lookup"><span data-stu-id="e5c37-258">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="e5c37-259">Пример</span><span class="sxs-lookup"><span data-stu-id="e5c37-259">Example</span></span>

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

#### <a name="body-bodyjavascriptapioutlookofficebody"></a><span data-ttu-id="e5c37-260">основной текст: [Body](/javascript/api/outlook/office.body)</span><span class="sxs-lookup"><span data-stu-id="e5c37-260">body: [Body](/javascript/api/outlook/office.body)</span></span>

<span data-ttu-id="e5c37-261">Получает объект, предоставляющий методы для работы с основным текстом элемента.</span><span class="sxs-lookup"><span data-stu-id="e5c37-261">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="e5c37-262">Тип</span><span class="sxs-lookup"><span data-stu-id="e5c37-262">Type</span></span>

*   [<span data-ttu-id="e5c37-263">Body</span><span class="sxs-lookup"><span data-stu-id="e5c37-263">Body</span></span>](/javascript/api/outlook/office.body)

##### <a name="requirements"></a><span data-ttu-id="e5c37-264">Требования</span><span class="sxs-lookup"><span data-stu-id="e5c37-264">Requirements</span></span>

|<span data-ttu-id="e5c37-265">Требование</span><span class="sxs-lookup"><span data-stu-id="e5c37-265">Requirement</span></span>|<span data-ttu-id="e5c37-266">Значение</span><span class="sxs-lookup"><span data-stu-id="e5c37-266">Value</span></span>|
|---|---|
|[<span data-ttu-id="e5c37-267">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="e5c37-267">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e5c37-268">1.1</span><span class="sxs-lookup"><span data-stu-id="e5c37-268">1.1</span></span>|
|[<span data-ttu-id="e5c37-269">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e5c37-269">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e5c37-270">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e5c37-270">ReadItem</span></span>|
|[<span data-ttu-id="e5c37-271">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e5c37-271">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="e5c37-272">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="e5c37-272">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e5c37-273">Пример</span><span class="sxs-lookup"><span data-stu-id="e5c37-273">Example</span></span>

<span data-ttu-id="e5c37-274">В этом примере возвращается текст сообщения в формате обычного текста.</span><span class="sxs-lookup"><span data-stu-id="e5c37-274">This example gets the body of the message in plain text.</span></span>

```js
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="e5c37-275">Ниже приведен пример итогового параметра, переданного функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="e5c37-275">The following is an example of the result parameter passed to the callback function.</span></span>

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

#### <a name="categories-categoriesjavascriptapioutlookofficecategories"></a><span data-ttu-id="e5c37-276">Категории: [категории](/javascript/api/outlook/office.categories)</span><span class="sxs-lookup"><span data-stu-id="e5c37-276">categories: [Categories](/javascript/api/outlook/office.categories)</span></span>

<span data-ttu-id="e5c37-277">Получает объект, предоставляющий методы для управления категориями элемента.</span><span class="sxs-lookup"><span data-stu-id="e5c37-277">Gets an object that provides methods for managing the item's categories.</span></span>

> [!NOTE]
> <span data-ttu-id="e5c37-278">Этот элемент не поддерживается в Outlook на iOS или Android.</span><span class="sxs-lookup"><span data-stu-id="e5c37-278">This member is not supported in Outlook on iOS or Android.</span></span>

##### <a name="type"></a><span data-ttu-id="e5c37-279">Тип</span><span class="sxs-lookup"><span data-stu-id="e5c37-279">Type</span></span>

*   [<span data-ttu-id="e5c37-280">Categories</span><span class="sxs-lookup"><span data-stu-id="e5c37-280">Categories</span></span>](/javascript/api/outlook/office.categories)

##### <a name="requirements"></a><span data-ttu-id="e5c37-281">Требования</span><span class="sxs-lookup"><span data-stu-id="e5c37-281">Requirements</span></span>

|<span data-ttu-id="e5c37-282">Требование</span><span class="sxs-lookup"><span data-stu-id="e5c37-282">Requirement</span></span>|<span data-ttu-id="e5c37-283">Значение</span><span class="sxs-lookup"><span data-stu-id="e5c37-283">Value</span></span>|
|---|---|
|[<span data-ttu-id="e5c37-284">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="e5c37-284">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e5c37-285">Предварительная версия</span><span class="sxs-lookup"><span data-stu-id="e5c37-285">Preview</span></span>|
|[<span data-ttu-id="e5c37-286">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e5c37-286">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e5c37-287">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e5c37-287">ReadItem</span></span>|
|[<span data-ttu-id="e5c37-288">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e5c37-288">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="e5c37-289">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="e5c37-289">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e5c37-290">Пример</span><span class="sxs-lookup"><span data-stu-id="e5c37-290">Example</span></span>

<span data-ttu-id="e5c37-291">В этом примере возвращаются категории элемента.</span><span class="sxs-lookup"><span data-stu-id="e5c37-291">This example gets the item's categories.</span></span>

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

#### <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="e5c37-292">CC: Array. <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[получатели](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="e5c37-292">cc: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="e5c37-293">Предоставляет доступ к получателям копии сообщения.</span><span class="sxs-lookup"><span data-stu-id="e5c37-293">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="e5c37-294">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="e5c37-294">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="e5c37-295">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="e5c37-295">Read mode</span></span>

<span data-ttu-id="e5c37-p106">Свойство `cc` возвращает массив, который содержит объект `EmailAddressDetails` для каждого получателя, указанного в строке **Копия** сообщения. Коллекция может включать не более 100 элементов.</span><span class="sxs-lookup"><span data-stu-id="e5c37-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="e5c37-298">Режим создания</span><span class="sxs-lookup"><span data-stu-id="e5c37-298">Compose mode</span></span>

<span data-ttu-id="e5c37-299">Свойство `cc` возвращает объект `Recipients`, предоставляющий методы для получения или обновления получателей, которые указаны в строке **Копия** сообщения.</span><span class="sxs-lookup"><span data-stu-id="e5c37-299">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

```js
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="e5c37-300">Тип</span><span class="sxs-lookup"><span data-stu-id="e5c37-300">Type</span></span>

*   <span data-ttu-id="e5c37-301">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="e5c37-301">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="e5c37-302">Требования</span><span class="sxs-lookup"><span data-stu-id="e5c37-302">Requirements</span></span>

|<span data-ttu-id="e5c37-303">Требование</span><span class="sxs-lookup"><span data-stu-id="e5c37-303">Requirement</span></span>|<span data-ttu-id="e5c37-304">Значение</span><span class="sxs-lookup"><span data-stu-id="e5c37-304">Value</span></span>|
|---|---|
|[<span data-ttu-id="e5c37-305">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="e5c37-305">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e5c37-306">1.0</span><span class="sxs-lookup"><span data-stu-id="e5c37-306">1.0</span></span>|
|[<span data-ttu-id="e5c37-307">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e5c37-307">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e5c37-308">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e5c37-308">ReadItem</span></span>|
|[<span data-ttu-id="e5c37-309">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e5c37-309">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="e5c37-310">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="e5c37-310">Compose or Read</span></span>|

<br>

---
---

#### <a name="nullable-conversationid-string"></a><span data-ttu-id="e5c37-311">(Nullable) conversationId: строка</span><span class="sxs-lookup"><span data-stu-id="e5c37-311">(nullable) conversationId: String</span></span>

<span data-ttu-id="e5c37-312">Получает идентификатор разговора по электронной почте, содержащего конкретное сообщение.</span><span class="sxs-lookup"><span data-stu-id="e5c37-312">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="e5c37-p107">Вы можете получить целочисленное значение этого свойства, если ваше почтовое приложение активируется в формах просмотра или формах создания ответов. Если пользователь изменит тему ответа, после его отправки идентификатор беседы будет изменен, и полученное ранее значение будет недействительным.</span><span class="sxs-lookup"><span data-stu-id="e5c37-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="e5c37-p108">Это свойство имеет значение NULL для нового элемента в форме создания. Свойство `conversationId` вернет значение, если пользователь задаст тему и сохранит элемент.</span><span class="sxs-lookup"><span data-stu-id="e5c37-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="e5c37-317">Тип</span><span class="sxs-lookup"><span data-stu-id="e5c37-317">Type</span></span>

*   <span data-ttu-id="e5c37-318">String</span><span class="sxs-lookup"><span data-stu-id="e5c37-318">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="e5c37-319">Требования</span><span class="sxs-lookup"><span data-stu-id="e5c37-319">Requirements</span></span>

|<span data-ttu-id="e5c37-320">Требование</span><span class="sxs-lookup"><span data-stu-id="e5c37-320">Requirement</span></span>|<span data-ttu-id="e5c37-321">Значение</span><span class="sxs-lookup"><span data-stu-id="e5c37-321">Value</span></span>|
|---|---|
|[<span data-ttu-id="e5c37-322">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="e5c37-322">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e5c37-323">1.0</span><span class="sxs-lookup"><span data-stu-id="e5c37-323">1.0</span></span>|
|[<span data-ttu-id="e5c37-324">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e5c37-324">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e5c37-325">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e5c37-325">ReadItem</span></span>|
|[<span data-ttu-id="e5c37-326">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e5c37-326">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="e5c37-327">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="e5c37-327">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e5c37-328">Пример</span><span class="sxs-lookup"><span data-stu-id="e5c37-328">Example</span></span>

```js
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

<br>

---
---

#### <a name="datetimecreated-date"></a><span data-ttu-id="e5c37-329">dateTimeCreated: Дата</span><span class="sxs-lookup"><span data-stu-id="e5c37-329">dateTimeCreated: Date</span></span>

<span data-ttu-id="e5c37-p109">Получает дату и время создания элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="e5c37-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="e5c37-332">Тип</span><span class="sxs-lookup"><span data-stu-id="e5c37-332">Type</span></span>

*   <span data-ttu-id="e5c37-333">Дата</span><span class="sxs-lookup"><span data-stu-id="e5c37-333">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="e5c37-334">Требования</span><span class="sxs-lookup"><span data-stu-id="e5c37-334">Requirements</span></span>

|<span data-ttu-id="e5c37-335">Требование</span><span class="sxs-lookup"><span data-stu-id="e5c37-335">Requirement</span></span>|<span data-ttu-id="e5c37-336">Значение</span><span class="sxs-lookup"><span data-stu-id="e5c37-336">Value</span></span>|
|---|---|
|[<span data-ttu-id="e5c37-337">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="e5c37-337">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e5c37-338">1.0</span><span class="sxs-lookup"><span data-stu-id="e5c37-338">1.0</span></span>|
|[<span data-ttu-id="e5c37-339">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e5c37-339">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e5c37-340">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e5c37-340">ReadItem</span></span>|
|[<span data-ttu-id="e5c37-341">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e5c37-341">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="e5c37-342">Чтение</span><span class="sxs-lookup"><span data-stu-id="e5c37-342">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e5c37-343">Пример</span><span class="sxs-lookup"><span data-stu-id="e5c37-343">Example</span></span>

```js
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

<br>

---
---

#### <a name="datetimemodified-date"></a><span data-ttu-id="e5c37-344">dateTimeModified: Дата</span><span class="sxs-lookup"><span data-stu-id="e5c37-344">dateTimeModified: Date</span></span>

<span data-ttu-id="e5c37-345">Получает дату и время последнего изменения элемента.</span><span class="sxs-lookup"><span data-stu-id="e5c37-345">Gets the date and time that an item was last modified.</span></span> <span data-ttu-id="e5c37-346">Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="e5c37-346">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="e5c37-347">Этот элемент не поддерживается в Outlook на iOS или Android.</span><span class="sxs-lookup"><span data-stu-id="e5c37-347">This member is not supported in Outlook on iOS or Android.</span></span>

##### <a name="type"></a><span data-ttu-id="e5c37-348">Тип</span><span class="sxs-lookup"><span data-stu-id="e5c37-348">Type</span></span>

*   <span data-ttu-id="e5c37-349">Дата</span><span class="sxs-lookup"><span data-stu-id="e5c37-349">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="e5c37-350">Требования</span><span class="sxs-lookup"><span data-stu-id="e5c37-350">Requirements</span></span>

|<span data-ttu-id="e5c37-351">Требование</span><span class="sxs-lookup"><span data-stu-id="e5c37-351">Requirement</span></span>|<span data-ttu-id="e5c37-352">Значение</span><span class="sxs-lookup"><span data-stu-id="e5c37-352">Value</span></span>|
|---|---|
|[<span data-ttu-id="e5c37-353">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="e5c37-353">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e5c37-354">1.0</span><span class="sxs-lookup"><span data-stu-id="e5c37-354">1.0</span></span>|
|[<span data-ttu-id="e5c37-355">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e5c37-355">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e5c37-356">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e5c37-356">ReadItem</span></span>|
|[<span data-ttu-id="e5c37-357">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e5c37-357">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="e5c37-358">Чтение</span><span class="sxs-lookup"><span data-stu-id="e5c37-358">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e5c37-359">Пример</span><span class="sxs-lookup"><span data-stu-id="e5c37-359">Example</span></span>

```js
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

<br>

---
---

#### <a name="end-datetimejavascriptapioutlookofficetime"></a><span data-ttu-id="e5c37-360">конец: Дата | [Time (время](/javascript/api/outlook/office.time) )</span><span class="sxs-lookup"><span data-stu-id="e5c37-360">end: Date|[Time](/javascript/api/outlook/office.time)</span></span>

<span data-ttu-id="e5c37-361">Получает или задает дату и время окончания встречи.</span><span class="sxs-lookup"><span data-stu-id="e5c37-361">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="e5c37-p111">Свойство `end` представлено в виде значения даты и времени в формате UTC. Преобразовать значение свойства end в местные значения даты и времени клиента можно с помощью метода [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime).</span><span class="sxs-lookup"><span data-stu-id="e5c37-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="e5c37-364">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="e5c37-364">Read mode</span></span>

<span data-ttu-id="e5c37-365">Свойство `end` возвращает объект `Date`.</span><span class="sxs-lookup"><span data-stu-id="e5c37-365">The `end` property returns a `Date` object.</span></span>

```js
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="e5c37-366">Режим создания</span><span class="sxs-lookup"><span data-stu-id="e5c37-366">Compose mode</span></span>

<span data-ttu-id="e5c37-367">Свойство `end` возвращает объект `Time`.</span><span class="sxs-lookup"><span data-stu-id="e5c37-367">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="e5c37-368">Если вы задаете время окончания с помощью метода [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-), необходимо использовать метод [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) для преобразования местного времени на клиенте в формат UTC для сервера.</span><span class="sxs-lookup"><span data-stu-id="e5c37-368">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="e5c37-369">В примере ниже показано, как с помощью метода [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) объекта `Time` задать время окончания встречи.</span><span class="sxs-lookup"><span data-stu-id="e5c37-369">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="e5c37-370">Тип</span><span class="sxs-lookup"><span data-stu-id="e5c37-370">Type</span></span>

*   <span data-ttu-id="e5c37-371">Date | [Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="e5c37-371">Date | [Time](/javascript/api/outlook/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="e5c37-372">Требования</span><span class="sxs-lookup"><span data-stu-id="e5c37-372">Requirements</span></span>

|<span data-ttu-id="e5c37-373">Требование</span><span class="sxs-lookup"><span data-stu-id="e5c37-373">Requirement</span></span>|<span data-ttu-id="e5c37-374">Значение</span><span class="sxs-lookup"><span data-stu-id="e5c37-374">Value</span></span>|
|---|---|
|[<span data-ttu-id="e5c37-375">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="e5c37-375">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e5c37-376">1.0</span><span class="sxs-lookup"><span data-stu-id="e5c37-376">1.0</span></span>|
|[<span data-ttu-id="e5c37-377">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e5c37-377">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e5c37-378">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e5c37-378">ReadItem</span></span>|
|[<span data-ttu-id="e5c37-379">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e5c37-379">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="e5c37-380">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="e5c37-380">Compose or Read</span></span>|

<br>

---
---

#### <a name="enhancedlocation-enhancedlocationjavascriptapioutlookofficeenhancedlocation"></a><span data-ttu-id="e5c37-381">Енханцедлокатион: [енханцедлокатион](/javascript/api/outlook/office.enhancedlocation)</span><span class="sxs-lookup"><span data-stu-id="e5c37-381">enhancedLocation: [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation)</span></span>

<span data-ttu-id="e5c37-382">Получает или задает расположение встречи.</span><span class="sxs-lookup"><span data-stu-id="e5c37-382">Gets or sets the locations of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="e5c37-383">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="e5c37-383">Read mode</span></span>

<span data-ttu-id="e5c37-384">Свойство возвращает объект [енханцедлокатион](/javascript/api/outlook/office.enhancedlocation) , который позволяет получить набор расположений (каждый, представленный объектом локатиондетаилс), связанный с встречей. [](/javascript/api/outlook/office.locationdetails) `enhancedLocation`</span><span class="sxs-lookup"><span data-stu-id="e5c37-384">The `enhancedLocation` property returns an [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation) object that allows you to get the set of locations (each represented by a [LocationDetails](/javascript/api/outlook/office.locationdetails) object) associated with the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="e5c37-385">Режим создания</span><span class="sxs-lookup"><span data-stu-id="e5c37-385">Compose mode</span></span>

<span data-ttu-id="e5c37-386">`enhancedLocation` Свойство возвращает объект [енханцедлокатион](/javascript/api/outlook/office.enhancedlocation) , который предоставляет методы для получения, удаления или добавления расположений для встречи.</span><span class="sxs-lookup"><span data-stu-id="e5c37-386">The `enhancedLocation` property returns an [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation) object that provides methods to get, remove, or add locations on an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="e5c37-387">Тип</span><span class="sxs-lookup"><span data-stu-id="e5c37-387">Type</span></span>

*   [<span data-ttu-id="e5c37-388">EnhancedLocation</span><span class="sxs-lookup"><span data-stu-id="e5c37-388">EnhancedLocation</span></span>](/javascript/api/outlook/office.enhancedlocation)

##### <a name="requirements"></a><span data-ttu-id="e5c37-389">Требования</span><span class="sxs-lookup"><span data-stu-id="e5c37-389">Requirements</span></span>

|<span data-ttu-id="e5c37-390">Требование</span><span class="sxs-lookup"><span data-stu-id="e5c37-390">Requirement</span></span>|<span data-ttu-id="e5c37-391">Значение</span><span class="sxs-lookup"><span data-stu-id="e5c37-391">Value</span></span>|
|---|---|
|[<span data-ttu-id="e5c37-392">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="e5c37-392">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e5c37-393">Предварительная версия</span><span class="sxs-lookup"><span data-stu-id="e5c37-393">Preview</span></span>|
|[<span data-ttu-id="e5c37-394">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e5c37-394">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e5c37-395">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e5c37-395">ReadItem</span></span>|
|[<span data-ttu-id="e5c37-396">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e5c37-396">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="e5c37-397">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="e5c37-397">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e5c37-398">Пример</span><span class="sxs-lookup"><span data-stu-id="e5c37-398">Example</span></span>

<span data-ttu-id="e5c37-399">В следующем примере показано получение текущих расположений, связанных с встречей.</span><span class="sxs-lookup"><span data-stu-id="e5c37-399">The following example gets the current locations associated with the appointment.</span></span>

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

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom"></a><span data-ttu-id="e5c37-400">от: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[](/javascript/api/outlook/office.from)</span><span class="sxs-lookup"><span data-stu-id="e5c37-400">from: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[From](/javascript/api/outlook/office.from)</span></span>

<span data-ttu-id="e5c37-401">Получает электронный адрес отправителя сообщения.</span><span class="sxs-lookup"><span data-stu-id="e5c37-401">Gets the email address of the sender of a message.</span></span>

<span data-ttu-id="e5c37-p112">Свойства `from` и [`sender`](#sender-emailaddressdetails) представляют одно лицо, если сообщение не отправлено представителем. В противном случае свойство `from` представляет лицо, делегировавшее полномочия, а свойство sender — представителя.</span><span class="sxs-lookup"><span data-stu-id="e5c37-p112">The `from` and [`sender`](#sender-emailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="e5c37-404">Свойству `recipientType`, принадлежащему объекту `EmailAddressDetails` в свойстве `from`, задано значение `undefined`.</span><span class="sxs-lookup"><span data-stu-id="e5c37-404">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="e5c37-405">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="e5c37-405">Read mode</span></span>

<span data-ttu-id="e5c37-406">`from` Свойство возвращает `EmailAddressDetails` объект.</span><span class="sxs-lookup"><span data-stu-id="e5c37-406">The `from` property returns an `EmailAddressDetails` object.</span></span>

```js
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

##### <a name="compose-mode"></a><span data-ttu-id="e5c37-407">Режим создания</span><span class="sxs-lookup"><span data-stu-id="e5c37-407">Compose mode</span></span>

<span data-ttu-id="e5c37-408">`from` Свойство возвращает `From` объект, который предоставляет метод для получения значения From.</span><span class="sxs-lookup"><span data-stu-id="e5c37-408">The `from` property returns a `From` object that provides a method to get the from value.</span></span>

```js
Office.context.mailbox.item.from.getAsync(callback);

function callback(asyncResult) {
  var from = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="e5c37-409">Тип</span><span class="sxs-lookup"><span data-stu-id="e5c37-409">Type</span></span>

*   <span data-ttu-id="e5c37-410">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [из](/javascript/api/outlook/office.from)</span><span class="sxs-lookup"><span data-stu-id="e5c37-410">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [From](/javascript/api/outlook/office.from)</span></span>

##### <a name="requirements"></a><span data-ttu-id="e5c37-411">Требования</span><span class="sxs-lookup"><span data-stu-id="e5c37-411">Requirements</span></span>

|<span data-ttu-id="e5c37-412">Требование</span><span class="sxs-lookup"><span data-stu-id="e5c37-412">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="e5c37-413">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="e5c37-413">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e5c37-414">1.0</span><span class="sxs-lookup"><span data-stu-id="e5c37-414">1.0</span></span>|<span data-ttu-id="e5c37-415">1.7</span><span class="sxs-lookup"><span data-stu-id="e5c37-415">1.7</span></span>|
|[<span data-ttu-id="e5c37-416">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e5c37-416">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e5c37-417">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e5c37-417">ReadItem</span></span>|<span data-ttu-id="e5c37-418">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="e5c37-418">ReadWriteItem</span></span>|
|[<span data-ttu-id="e5c37-419">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e5c37-419">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="e5c37-420">Чтение</span><span class="sxs-lookup"><span data-stu-id="e5c37-420">Read</span></span>|<span data-ttu-id="e5c37-421">Создание</span><span class="sxs-lookup"><span data-stu-id="e5c37-421">Compose</span></span>|

<br>

---
---

#### <a name="internetheaders-internetheadersjavascriptapioutlookofficeinternetheaders"></a><span data-ttu-id="e5c37-422">Internetheaders:: [internetheaders:](/javascript/api/outlook/office.internetheaders)</span><span class="sxs-lookup"><span data-stu-id="e5c37-422">internetHeaders: [InternetHeaders](/javascript/api/outlook/office.internetheaders)</span></span>

<span data-ttu-id="e5c37-423">Возвращает или задает настраиваемые заголовки Интернета для сообщения.</span><span class="sxs-lookup"><span data-stu-id="e5c37-423">Gets or sets custom internet headers on a message.</span></span>

##### <a name="type"></a><span data-ttu-id="e5c37-424">Тип</span><span class="sxs-lookup"><span data-stu-id="e5c37-424">Type</span></span>

*   [<span data-ttu-id="e5c37-425">InternetHeaders</span><span class="sxs-lookup"><span data-stu-id="e5c37-425">InternetHeaders</span></span>](/javascript/api/outlook/office.internetheaders)

##### <a name="requirements"></a><span data-ttu-id="e5c37-426">Требования</span><span class="sxs-lookup"><span data-stu-id="e5c37-426">Requirements</span></span>

|<span data-ttu-id="e5c37-427">Требование</span><span class="sxs-lookup"><span data-stu-id="e5c37-427">Requirement</span></span>|<span data-ttu-id="e5c37-428">Значение</span><span class="sxs-lookup"><span data-stu-id="e5c37-428">Value</span></span>|
|---|---|
|[<span data-ttu-id="e5c37-429">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="e5c37-429">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e5c37-430">Предварительная версия</span><span class="sxs-lookup"><span data-stu-id="e5c37-430">Preview</span></span>|
|[<span data-ttu-id="e5c37-431">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e5c37-431">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e5c37-432">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e5c37-432">ReadItem</span></span>|
|[<span data-ttu-id="e5c37-433">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e5c37-433">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="e5c37-434">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="e5c37-434">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e5c37-435">Пример</span><span class="sxs-lookup"><span data-stu-id="e5c37-435">Example</span></span>

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

#### <a name="internetmessageid-string"></a><span data-ttu-id="e5c37-436">internetMessageId: строка</span><span class="sxs-lookup"><span data-stu-id="e5c37-436">internetMessageId: String</span></span>

<span data-ttu-id="e5c37-p113">Получает идентификатор интернет-сообщения для электронного сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="e5c37-p113">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="e5c37-439">Тип</span><span class="sxs-lookup"><span data-stu-id="e5c37-439">Type</span></span>

*   <span data-ttu-id="e5c37-440">String</span><span class="sxs-lookup"><span data-stu-id="e5c37-440">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="e5c37-441">Требования</span><span class="sxs-lookup"><span data-stu-id="e5c37-441">Requirements</span></span>

|<span data-ttu-id="e5c37-442">Требование</span><span class="sxs-lookup"><span data-stu-id="e5c37-442">Requirement</span></span>|<span data-ttu-id="e5c37-443">Значение</span><span class="sxs-lookup"><span data-stu-id="e5c37-443">Value</span></span>|
|---|---|
|[<span data-ttu-id="e5c37-444">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="e5c37-444">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e5c37-445">1.0</span><span class="sxs-lookup"><span data-stu-id="e5c37-445">1.0</span></span>|
|[<span data-ttu-id="e5c37-446">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e5c37-446">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e5c37-447">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e5c37-447">ReadItem</span></span>|
|[<span data-ttu-id="e5c37-448">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e5c37-448">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="e5c37-449">Чтение</span><span class="sxs-lookup"><span data-stu-id="e5c37-449">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e5c37-450">Пример</span><span class="sxs-lookup"><span data-stu-id="e5c37-450">Example</span></span>

```js
var internetMessageId = Office.context.mailbox.item.internetMessageId;
console.log("internetMessageId: " + internetMessageId);
```

<br>

---
---

#### <a name="itemclass-string"></a><span data-ttu-id="e5c37-451">itemClass: строка</span><span class="sxs-lookup"><span data-stu-id="e5c37-451">itemClass: String</span></span>

<span data-ttu-id="e5c37-p114">Получает класс элемента веб-служб Exchange для выбранного элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="e5c37-p114">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="e5c37-p115">Свойство `itemClass` указывает класс сообщения выбранного элемента. Ниже приводятся классы сообщения по умолчанию для элемента сообщения или встречи.</span><span class="sxs-lookup"><span data-stu-id="e5c37-p115">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

|<span data-ttu-id="e5c37-456">Тип</span><span class="sxs-lookup"><span data-stu-id="e5c37-456">Type</span></span>|<span data-ttu-id="e5c37-457">Описание</span><span class="sxs-lookup"><span data-stu-id="e5c37-457">Description</span></span>|<span data-ttu-id="e5c37-458">Класс элемента</span><span class="sxs-lookup"><span data-stu-id="e5c37-458">item class</span></span>|
|---|---|---|
|<span data-ttu-id="e5c37-459">Элементы встречи</span><span class="sxs-lookup"><span data-stu-id="e5c37-459">Appointment items</span></span>|<span data-ttu-id="e5c37-460">Это элементы календаря для класса элемента `IPM.Appointment` или `IPM.Appointment.Occurrence`.</span><span class="sxs-lookup"><span data-stu-id="e5c37-460">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span>|`IPM.Appointment`<br />`IPM.Appointment.Occurrence`|
|<span data-ttu-id="e5c37-461">Элементы сообщения</span><span class="sxs-lookup"><span data-stu-id="e5c37-461">Message items</span></span>|<span data-ttu-id="e5c37-462">Сюда входят электронные сообщения, для которых по умолчанию задан класс сообщения `IPM.Note`, а также приглашения на собрания, ответы на них и уведомления об их отмене, использующие `IPM.Schedule.Meeting` в качестве базового класса сообщения.</span><span class="sxs-lookup"><span data-stu-id="e5c37-462">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span>|`IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled`|

<span data-ttu-id="e5c37-463">Можно создавать настраиваемые классы сообщения, расширяющие классы сообщения по умолчанию, например настраиваемый класс сообщения о встрече `IPM.Appointment.Contoso`.</span><span class="sxs-lookup"><span data-stu-id="e5c37-463">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="e5c37-464">Тип</span><span class="sxs-lookup"><span data-stu-id="e5c37-464">Type</span></span>

*   <span data-ttu-id="e5c37-465">String</span><span class="sxs-lookup"><span data-stu-id="e5c37-465">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="e5c37-466">Требования</span><span class="sxs-lookup"><span data-stu-id="e5c37-466">Requirements</span></span>

|<span data-ttu-id="e5c37-467">Требование</span><span class="sxs-lookup"><span data-stu-id="e5c37-467">Requirement</span></span>|<span data-ttu-id="e5c37-468">Значение</span><span class="sxs-lookup"><span data-stu-id="e5c37-468">Value</span></span>|
|---|---|
|[<span data-ttu-id="e5c37-469">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="e5c37-469">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e5c37-470">1.0</span><span class="sxs-lookup"><span data-stu-id="e5c37-470">1.0</span></span>|
|[<span data-ttu-id="e5c37-471">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e5c37-471">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e5c37-472">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e5c37-472">ReadItem</span></span>|
|[<span data-ttu-id="e5c37-473">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e5c37-473">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="e5c37-474">Чтение</span><span class="sxs-lookup"><span data-stu-id="e5c37-474">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e5c37-475">Пример</span><span class="sxs-lookup"><span data-stu-id="e5c37-475">Example</span></span>

```js
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

<br>

---
---

#### <a name="nullable-itemid-string"></a><span data-ttu-id="e5c37-476">(Nullable) itemId: строка</span><span class="sxs-lookup"><span data-stu-id="e5c37-476">(nullable) itemId: String</span></span>

<span data-ttu-id="e5c37-p116">Получает идентификатор элемента веб-служб Exchange для текущего элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="e5c37-p116">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="e5c37-479">Идентификатор, возвращаемый свойством `itemId`, совпадает с идентификатором элемента веб-служб Exchange.</span><span class="sxs-lookup"><span data-stu-id="e5c37-479">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="e5c37-480">Свойство `itemId` не совпадает с идентификатором записи Outlook, а также идентификатором, который используется REST API Outlook.</span><span class="sxs-lookup"><span data-stu-id="e5c37-480">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="e5c37-481">Прежде чем совершать вызовы REST API, используя это значение, его необходимо преобразовать с помощью [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="e5c37-481">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="e5c37-482">Дополнительные сведения см. в статье [Использование REST API Outlook из надстройки Outlook](/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="e5c37-482">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="e5c37-p118">Свойство `itemId` недоступно в режиме создания. Если требуется идентификатор элемента, с помощью метода [`saveAsync`](#saveasyncoptions-callback) можно сохранить элемент в хранилище. При этом в параметре [`AsyncResult.value`](/javascript/api/office/office.asyncresult) функции обратного вызова возвращается идентификатор элемента.</span><span class="sxs-lookup"><span data-stu-id="e5c37-p118">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="e5c37-485">Тип</span><span class="sxs-lookup"><span data-stu-id="e5c37-485">Type</span></span>

*   <span data-ttu-id="e5c37-486">String</span><span class="sxs-lookup"><span data-stu-id="e5c37-486">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="e5c37-487">Требования</span><span class="sxs-lookup"><span data-stu-id="e5c37-487">Requirements</span></span>

|<span data-ttu-id="e5c37-488">Требование</span><span class="sxs-lookup"><span data-stu-id="e5c37-488">Requirement</span></span>|<span data-ttu-id="e5c37-489">Значение</span><span class="sxs-lookup"><span data-stu-id="e5c37-489">Value</span></span>|
|---|---|
|[<span data-ttu-id="e5c37-490">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="e5c37-490">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e5c37-491">1.0</span><span class="sxs-lookup"><span data-stu-id="e5c37-491">1.0</span></span>|
|[<span data-ttu-id="e5c37-492">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e5c37-492">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e5c37-493">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e5c37-493">ReadItem</span></span>|
|[<span data-ttu-id="e5c37-494">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e5c37-494">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="e5c37-495">Чтение</span><span class="sxs-lookup"><span data-stu-id="e5c37-495">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e5c37-496">Пример</span><span class="sxs-lookup"><span data-stu-id="e5c37-496">Example</span></span>

<span data-ttu-id="e5c37-p119">Указанный ниже код проверяет наличие идентификатора элемента. Если свойство `itemId` возвращает значение `null` или `undefined`, элемент будет сохранен в хранилище, а из асинхронного результата будет получен идентификатор элемента.</span><span class="sxs-lookup"><span data-stu-id="e5c37-p119">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

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

#### <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype"></a><span data-ttu-id="e5c37-499">itemType: [Office. MailboxEnums. ItemType](/javascript/api/outlook/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="e5c37-499">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="e5c37-500">Получает тип элемента, который представляет экземпляр.</span><span class="sxs-lookup"><span data-stu-id="e5c37-500">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="e5c37-501">Свойство `itemType` возвращает одно из значений перечисления `ItemType`, которое указывает, является ли экземпляр объекта `item` сообщением или встречей.</span><span class="sxs-lookup"><span data-stu-id="e5c37-501">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="e5c37-502">Тип</span><span class="sxs-lookup"><span data-stu-id="e5c37-502">Type</span></span>

*   [<span data-ttu-id="e5c37-503">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="e5c37-503">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="e5c37-504">Требования</span><span class="sxs-lookup"><span data-stu-id="e5c37-504">Requirements</span></span>

|<span data-ttu-id="e5c37-505">Требование</span><span class="sxs-lookup"><span data-stu-id="e5c37-505">Requirement</span></span>|<span data-ttu-id="e5c37-506">Значение</span><span class="sxs-lookup"><span data-stu-id="e5c37-506">Value</span></span>|
|---|---|
|[<span data-ttu-id="e5c37-507">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="e5c37-507">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e5c37-508">1.0</span><span class="sxs-lookup"><span data-stu-id="e5c37-508">1.0</span></span>|
|[<span data-ttu-id="e5c37-509">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e5c37-509">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e5c37-510">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e5c37-510">ReadItem</span></span>|
|[<span data-ttu-id="e5c37-511">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e5c37-511">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="e5c37-512">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="e5c37-512">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e5c37-513">Пример</span><span class="sxs-lookup"><span data-stu-id="e5c37-513">Example</span></span>

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

#### <a name="location-stringlocationjavascriptapioutlookofficelocation"></a><span data-ttu-id="e5c37-514">Местоположение: строка | [Location (расположение](/javascript/api/outlook/office.location) )</span><span class="sxs-lookup"><span data-stu-id="e5c37-514">location: String|[Location](/javascript/api/outlook/office.location)</span></span>

<span data-ttu-id="e5c37-515">Получает или задает место встречи.</span><span class="sxs-lookup"><span data-stu-id="e5c37-515">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="e5c37-516">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="e5c37-516">Read mode</span></span>

<span data-ttu-id="e5c37-517">Свойство `location` возвращает строку, содержащую сведения о месте встречи.</span><span class="sxs-lookup"><span data-stu-id="e5c37-517">The `location` property returns a string that contains the location of the appointment.</span></span>

```js
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="e5c37-518">Режим создания</span><span class="sxs-lookup"><span data-stu-id="e5c37-518">Compose mode</span></span>

<span data-ttu-id="e5c37-519">Свойство `location` возвращает объект `Location`, предоставляющий методы, которые используются для получения и задания места встречи.</span><span class="sxs-lookup"><span data-stu-id="e5c37-519">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```js
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="e5c37-520">Тип</span><span class="sxs-lookup"><span data-stu-id="e5c37-520">Type</span></span>

*   <span data-ttu-id="e5c37-521">String | [Location](/javascript/api/outlook/office.location)</span><span class="sxs-lookup"><span data-stu-id="e5c37-521">String | [Location](/javascript/api/outlook/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="e5c37-522">Требования</span><span class="sxs-lookup"><span data-stu-id="e5c37-522">Requirements</span></span>

|<span data-ttu-id="e5c37-523">Требование</span><span class="sxs-lookup"><span data-stu-id="e5c37-523">Requirement</span></span>|<span data-ttu-id="e5c37-524">Значение</span><span class="sxs-lookup"><span data-stu-id="e5c37-524">Value</span></span>|
|---|---|
|[<span data-ttu-id="e5c37-525">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="e5c37-525">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e5c37-526">1.0</span><span class="sxs-lookup"><span data-stu-id="e5c37-526">1.0</span></span>|
|[<span data-ttu-id="e5c37-527">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e5c37-527">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e5c37-528">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e5c37-528">ReadItem</span></span>|
|[<span data-ttu-id="e5c37-529">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e5c37-529">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="e5c37-530">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="e5c37-530">Compose or Read</span></span>|

<br>

---
---

#### <a name="normalizedsubject-string"></a><span data-ttu-id="e5c37-531">normalizedSubject: строка</span><span class="sxs-lookup"><span data-stu-id="e5c37-531">normalizedSubject: String</span></span>

<span data-ttu-id="e5c37-p120">Получает тему элемента со всеми удаленными префиксами (включая `RE:` и `FWD:`). Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="e5c37-p120">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="e5c37-p121">Свойство normalizedSubject получает тему элемента со стандартными префиксами (такими как `RE:` и `FW:`), добавляемыми почтовыми программами. Для получения темы элемента с неизмененными префиксами используйте свойство [`subject`](#subject-stringsubject).</span><span class="sxs-lookup"><span data-stu-id="e5c37-p121">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="e5c37-536">Тип</span><span class="sxs-lookup"><span data-stu-id="e5c37-536">Type</span></span>

*   <span data-ttu-id="e5c37-537">String</span><span class="sxs-lookup"><span data-stu-id="e5c37-537">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="e5c37-538">Требования</span><span class="sxs-lookup"><span data-stu-id="e5c37-538">Requirements</span></span>

|<span data-ttu-id="e5c37-539">Требование</span><span class="sxs-lookup"><span data-stu-id="e5c37-539">Requirement</span></span>|<span data-ttu-id="e5c37-540">Значение</span><span class="sxs-lookup"><span data-stu-id="e5c37-540">Value</span></span>|
|---|---|
|[<span data-ttu-id="e5c37-541">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="e5c37-541">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e5c37-542">1.0</span><span class="sxs-lookup"><span data-stu-id="e5c37-542">1.0</span></span>|
|[<span data-ttu-id="e5c37-543">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e5c37-543">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e5c37-544">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e5c37-544">ReadItem</span></span>|
|[<span data-ttu-id="e5c37-545">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e5c37-545">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="e5c37-546">Чтение</span><span class="sxs-lookup"><span data-stu-id="e5c37-546">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e5c37-547">Пример</span><span class="sxs-lookup"><span data-stu-id="e5c37-547">Example</span></span>

```js
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

<br>

---
---

#### <a name="notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessages"></a><span data-ttu-id="e5c37-548">notificationMessages: [notificationMessages](/javascript/api/outlook/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="e5c37-548">notificationMessages: [NotificationMessages](/javascript/api/outlook/office.notificationmessages)</span></span>

<span data-ttu-id="e5c37-549">Получает сообщения уведомления для элемента.</span><span class="sxs-lookup"><span data-stu-id="e5c37-549">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="e5c37-550">Тип</span><span class="sxs-lookup"><span data-stu-id="e5c37-550">Type</span></span>

*   [<span data-ttu-id="e5c37-551">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="e5c37-551">NotificationMessages</span></span>](/javascript/api/outlook/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="e5c37-552">Требования</span><span class="sxs-lookup"><span data-stu-id="e5c37-552">Requirements</span></span>

|<span data-ttu-id="e5c37-553">Требование</span><span class="sxs-lookup"><span data-stu-id="e5c37-553">Requirement</span></span>|<span data-ttu-id="e5c37-554">Значение</span><span class="sxs-lookup"><span data-stu-id="e5c37-554">Value</span></span>|
|---|---|
|[<span data-ttu-id="e5c37-555">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="e5c37-555">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e5c37-556">1.3</span><span class="sxs-lookup"><span data-stu-id="e5c37-556">1.3</span></span>|
|[<span data-ttu-id="e5c37-557">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e5c37-557">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e5c37-558">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e5c37-558">ReadItem</span></span>|
|[<span data-ttu-id="e5c37-559">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e5c37-559">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="e5c37-560">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="e5c37-560">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e5c37-561">Пример</span><span class="sxs-lookup"><span data-stu-id="e5c37-561">Example</span></span>

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

#### <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="e5c37-562">optionalAttendees: Array. <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[получатели](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="e5c37-562">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="e5c37-563">Предоставляет доступ к необязательным участникам события.</span><span class="sxs-lookup"><span data-stu-id="e5c37-563">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="e5c37-564">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="e5c37-564">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="e5c37-565">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="e5c37-565">Read mode</span></span>

<span data-ttu-id="e5c37-566">Свойство `optionalAttendees` возвращает массив, содержащий объект `EmailAddressDetails` для каждого необязательного участника собрания.</span><span class="sxs-lookup"><span data-stu-id="e5c37-566">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

```js
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="e5c37-567">Режим создания</span><span class="sxs-lookup"><span data-stu-id="e5c37-567">Compose mode</span></span>

<span data-ttu-id="e5c37-568">Свойство `optionalAttendees` возвращает объект `Recipients`, который предоставляет методы для получения или обновления необязательных участников собрания.</span><span class="sxs-lookup"><span data-stu-id="e5c37-568">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

```js
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="e5c37-569">Тип</span><span class="sxs-lookup"><span data-stu-id="e5c37-569">Type</span></span>

*   <span data-ttu-id="e5c37-570">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="e5c37-570">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="e5c37-571">Требования</span><span class="sxs-lookup"><span data-stu-id="e5c37-571">Requirements</span></span>

|<span data-ttu-id="e5c37-572">Требование</span><span class="sxs-lookup"><span data-stu-id="e5c37-572">Requirement</span></span>|<span data-ttu-id="e5c37-573">Значение</span><span class="sxs-lookup"><span data-stu-id="e5c37-573">Value</span></span>|
|---|---|
|[<span data-ttu-id="e5c37-574">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="e5c37-574">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e5c37-575">1.0</span><span class="sxs-lookup"><span data-stu-id="e5c37-575">1.0</span></span>|
|[<span data-ttu-id="e5c37-576">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e5c37-576">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e5c37-577">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e5c37-577">ReadItem</span></span>|
|[<span data-ttu-id="e5c37-578">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e5c37-578">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="e5c37-579">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="e5c37-579">Compose or Read</span></span>|

<br>

---
---

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsorganizerjavascriptapioutlookofficeorganizer"></a><span data-ttu-id="e5c37-580">Организатор: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[Организатор](/javascript/api/outlook/office.organizer)</span><span class="sxs-lookup"><span data-stu-id="e5c37-580">organizer: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[Organizer](/javascript/api/outlook/office.organizer)</span></span>

<span data-ttu-id="e5c37-581">Получает адрес электронной почты организатора для указанного собрания.</span><span class="sxs-lookup"><span data-stu-id="e5c37-581">Gets the email address of the organizer for a specified meeting.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="e5c37-582">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="e5c37-582">Read mode</span></span>

<span data-ttu-id="e5c37-583">`organizer` Свойство возвращает объект [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) , представляющий организатора собрания.</span><span class="sxs-lookup"><span data-stu-id="e5c37-583">The `organizer` property returns an [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) object that represents the meeting organizer.</span></span>

```js
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

##### <a name="compose-mode"></a><span data-ttu-id="e5c37-584">Режим создания</span><span class="sxs-lookup"><span data-stu-id="e5c37-584">Compose mode</span></span>

<span data-ttu-id="e5c37-585">Свойство возвращает объект организатора, который предоставляет метод для получения значения организатора. [](/javascript/api/outlook/office.organizer) `organizer`</span><span class="sxs-lookup"><span data-stu-id="e5c37-585">The `organizer` property returns an [Organizer](/javascript/api/outlook/office.organizer) object that provides a method to get the organizer value.</span></span>

```js
Office.context.mailbox.item.organizer.getAsync(
  function(asyncResult) {
    console.log(JSON.stringify(asyncResult));
  }
);
```

##### <a name="type"></a><span data-ttu-id="e5c37-586">Тип</span><span class="sxs-lookup"><span data-stu-id="e5c37-586">Type</span></span>

*   <span data-ttu-id="e5c37-587">[](/javascript/api/outlook/office.emailaddressdetails) | [Организатор](/javascript/api/outlook/office.organizer) EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="e5c37-587">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [Organizer](/javascript/api/outlook/office.organizer)</span></span>

##### <a name="requirements"></a><span data-ttu-id="e5c37-588">Требования</span><span class="sxs-lookup"><span data-stu-id="e5c37-588">Requirements</span></span>

|<span data-ttu-id="e5c37-589">Требование</span><span class="sxs-lookup"><span data-stu-id="e5c37-589">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="e5c37-590">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="e5c37-590">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e5c37-591">1.0</span><span class="sxs-lookup"><span data-stu-id="e5c37-591">1.0</span></span>|<span data-ttu-id="e5c37-592">1.7</span><span class="sxs-lookup"><span data-stu-id="e5c37-592">1.7</span></span>|
|[<span data-ttu-id="e5c37-593">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e5c37-593">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e5c37-594">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e5c37-594">ReadItem</span></span>|<span data-ttu-id="e5c37-595">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="e5c37-595">ReadWriteItem</span></span>|
|[<span data-ttu-id="e5c37-596">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e5c37-596">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="e5c37-597">Чтение</span><span class="sxs-lookup"><span data-stu-id="e5c37-597">Read</span></span>|<span data-ttu-id="e5c37-598">Создание</span><span class="sxs-lookup"><span data-stu-id="e5c37-598">Compose</span></span>|

<br>

---
---

#### <a name="nullable-recurrence-recurrencejavascriptapioutlookofficerecurrence"></a><span data-ttu-id="e5c37-599">(Nullable) повторение [](/javascript/api/outlook/office.recurrence) : повторение</span><span class="sxs-lookup"><span data-stu-id="e5c37-599">(nullable) recurrence: [Recurrence](/javascript/api/outlook/office.recurrence)</span></span>

<span data-ttu-id="e5c37-600">Получает или задает шаблон повторения встречи.</span><span class="sxs-lookup"><span data-stu-id="e5c37-600">Gets or sets the recurrence pattern of an appointment.</span></span> <span data-ttu-id="e5c37-601">Получает шаблон повторения приглашения на собрание.</span><span class="sxs-lookup"><span data-stu-id="e5c37-601">Gets the recurrence pattern of a meeting request.</span></span> <span data-ttu-id="e5c37-602">Режимы чтения и создания для элементов встречи.</span><span class="sxs-lookup"><span data-stu-id="e5c37-602">Read and compose modes for appointment items.</span></span> <span data-ttu-id="e5c37-603">Режим чтения для элементов приглашения на собрания.</span><span class="sxs-lookup"><span data-stu-id="e5c37-603">Read mode for meeting request items.</span></span>

<span data-ttu-id="e5c37-604">`recurrence` Свойство возвращает объект [повторения](/javascript/api/outlook/office.recurrence) для повторяющихся встреч или приглашений на собрания, если элемент представляет собой серию или экземпляр в ряду.</span><span class="sxs-lookup"><span data-stu-id="e5c37-604">The `recurrence` property returns a [recurrence](/javascript/api/outlook/office.recurrence) object for recurring appointments or meetings requests if an item is a series or an instance in a series.</span></span> <span data-ttu-id="e5c37-605">`null`возвращается для отдельных встреч и приглашений на собрание для отдельных встреч.</span><span class="sxs-lookup"><span data-stu-id="e5c37-605">`null` is returned for single appointments and meeting requests of single appointments.</span></span> <span data-ttu-id="e5c37-606">`undefined`возвращается для сообщений, которые не являются приглашениями на собрания.</span><span class="sxs-lookup"><span data-stu-id="e5c37-606">`undefined` is returned for messages that are not meeting requests.</span></span>

> <span data-ttu-id="e5c37-607">Note: приглашения на `itemClass` собрания имеют значение IPM. Schedule. Meeting. Request.</span><span class="sxs-lookup"><span data-stu-id="e5c37-607">Note: Meeting requests have an `itemClass` value of IPM.Schedule.Meeting.Request.</span></span>

> <span data-ttu-id="e5c37-608">Note: при наличии объекта `null`повторения это указывает на то, что объект является одной встречей или приглашением на собрание одной встречи, а не частью ряда.</span><span class="sxs-lookup"><span data-stu-id="e5c37-608">Note: If the recurrence object is `null`, this indicates that the object is a single appointment or a meeting request of a single appointment and NOT a part of a series.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="e5c37-609">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="e5c37-609">Read mode</span></span>

<span data-ttu-id="e5c37-610">`recurrence` Свойство возвращает объект [повторения](/javascript/api/outlook/office.recurrence) , представляющий повторение встречи.</span><span class="sxs-lookup"><span data-stu-id="e5c37-610">The `recurrence` property returns a [Recurrence](/javascript/api/outlook/office.recurrence) object that represents the appointment recurrence.</span></span> <span data-ttu-id="e5c37-611">Оно доступно для встреч и приглашений на собрания.</span><span class="sxs-lookup"><span data-stu-id="e5c37-611">This is available for appointments and meeting requests.</span></span>

```js
var recurrence = Office.context.mailbox.item.recurrence;
console.log("Recurrence: " + JSON.stringify(recurrence));
```

##### <a name="compose-mode"></a><span data-ttu-id="e5c37-612">Режим создания</span><span class="sxs-lookup"><span data-stu-id="e5c37-612">Compose mode</span></span>

<span data-ttu-id="e5c37-613">`recurrence` Свойство возвращает объект [повторения](/javascript/api/outlook/office.recurrence) , который предоставляет методы для управления повторением встречи.</span><span class="sxs-lookup"><span data-stu-id="e5c37-613">The `recurrence` property returns a [Recurrence](/javascript/api/outlook/office.recurrence) object that provides methods to manage the appointment recurrence.</span></span> <span data-ttu-id="e5c37-614">Оно доступно для встреч.</span><span class="sxs-lookup"><span data-stu-id="e5c37-614">This is available for appointments.</span></span>

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

##### <a name="type"></a><span data-ttu-id="e5c37-615">Тип</span><span class="sxs-lookup"><span data-stu-id="e5c37-615">Type</span></span>

* [<span data-ttu-id="e5c37-616">Повторения</span><span class="sxs-lookup"><span data-stu-id="e5c37-616">Recurrence</span></span>](/javascript/api/outlook/office.recurrence)

|<span data-ttu-id="e5c37-617">Требование</span><span class="sxs-lookup"><span data-stu-id="e5c37-617">Requirement</span></span>|<span data-ttu-id="e5c37-618">Значение</span><span class="sxs-lookup"><span data-stu-id="e5c37-618">Value</span></span>|
|---|---|
|[<span data-ttu-id="e5c37-619">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="e5c37-619">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e5c37-620">1.7</span><span class="sxs-lookup"><span data-stu-id="e5c37-620">1.7</span></span>|
|[<span data-ttu-id="e5c37-621">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e5c37-621">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e5c37-622">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e5c37-622">ReadItem</span></span>|
|[<span data-ttu-id="e5c37-623">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e5c37-623">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="e5c37-624">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="e5c37-624">Compose or Read</span></span>|

<br>

---
---

#### <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="e5c37-625">requiredAttendees: Array. <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[получатели](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="e5c37-625">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="e5c37-626">Предоставляет доступ к обязательным участникам события.</span><span class="sxs-lookup"><span data-stu-id="e5c37-626">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="e5c37-627">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="e5c37-627">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="e5c37-628">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="e5c37-628">Read mode</span></span>

<span data-ttu-id="e5c37-629">Свойство `requiredAttendees` возвращает массив, содержащий объект `EmailAddressDetails` для каждого обязательного участника собрания.</span><span class="sxs-lookup"><span data-stu-id="e5c37-629">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

```js
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="e5c37-630">Режим создания</span><span class="sxs-lookup"><span data-stu-id="e5c37-630">Compose mode</span></span>

<span data-ttu-id="e5c37-631">Свойство `requiredAttendees` возвращает объект `Recipients`, предоставляющий методы, с помощью которых можно получить или обновить сведения об обязательных участниках собрания.</span><span class="sxs-lookup"><span data-stu-id="e5c37-631">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

```js
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="e5c37-632">Тип</span><span class="sxs-lookup"><span data-stu-id="e5c37-632">Type</span></span>

*   <span data-ttu-id="e5c37-633">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="e5c37-633">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="e5c37-634">Требования</span><span class="sxs-lookup"><span data-stu-id="e5c37-634">Requirements</span></span>

|<span data-ttu-id="e5c37-635">Требование</span><span class="sxs-lookup"><span data-stu-id="e5c37-635">Requirement</span></span>|<span data-ttu-id="e5c37-636">Значение</span><span class="sxs-lookup"><span data-stu-id="e5c37-636">Value</span></span>|
|---|---|
|[<span data-ttu-id="e5c37-637">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="e5c37-637">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e5c37-638">1.0</span><span class="sxs-lookup"><span data-stu-id="e5c37-638">1.0</span></span>|
|[<span data-ttu-id="e5c37-639">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e5c37-639">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e5c37-640">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e5c37-640">ReadItem</span></span>|
|[<span data-ttu-id="e5c37-641">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e5c37-641">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="e5c37-642">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="e5c37-642">Compose or Read</span></span>|

<br>

---
---

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails"></a><span data-ttu-id="e5c37-643">Отправитель: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="e5c37-643">sender: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)</span></span>

<span data-ttu-id="e5c37-p128">Получает электронный адрес отправителя электронного сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="e5c37-p128">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="e5c37-p129">Свойства [`from`](#from-emailaddressdetailsfrom) и `sender` представляют одно лицо, если сообщение не отправлено представителем. В противном случае свойство `from` представляет лицо, делегировавшее полномочия, а свойство sender — представителя.</span><span class="sxs-lookup"><span data-stu-id="e5c37-p129">The [`from`](#from-emailaddressdetailsfrom) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="e5c37-648">Свойству `recipientType`, принадлежащему объекту `EmailAddressDetails` в свойстве `sender`, задано значение `undefined`.</span><span class="sxs-lookup"><span data-stu-id="e5c37-648">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="e5c37-649">Тип</span><span class="sxs-lookup"><span data-stu-id="e5c37-649">Type</span></span>

*   [<span data-ttu-id="e5c37-650">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="e5c37-650">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="e5c37-651">Требования</span><span class="sxs-lookup"><span data-stu-id="e5c37-651">Requirements</span></span>

|<span data-ttu-id="e5c37-652">Требование</span><span class="sxs-lookup"><span data-stu-id="e5c37-652">Requirement</span></span>|<span data-ttu-id="e5c37-653">Значение</span><span class="sxs-lookup"><span data-stu-id="e5c37-653">Value</span></span>|
|---|---|
|[<span data-ttu-id="e5c37-654">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="e5c37-654">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e5c37-655">1.0</span><span class="sxs-lookup"><span data-stu-id="e5c37-655">1.0</span></span>|
|[<span data-ttu-id="e5c37-656">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e5c37-656">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e5c37-657">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e5c37-657">ReadItem</span></span>|
|[<span data-ttu-id="e5c37-658">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e5c37-658">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="e5c37-659">Чтение</span><span class="sxs-lookup"><span data-stu-id="e5c37-659">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e5c37-660">Пример</span><span class="sxs-lookup"><span data-stu-id="e5c37-660">Example</span></span>

```js
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

<br>

---
---

#### <a name="nullable-seriesid-string"></a><span data-ttu-id="e5c37-661">(Nullable) seriesId: строка</span><span class="sxs-lookup"><span data-stu-id="e5c37-661">(nullable) seriesId: String</span></span>

<span data-ttu-id="e5c37-662">Получает идентификатор ряда, к которому принадлежит экземпляр.</span><span class="sxs-lookup"><span data-stu-id="e5c37-662">Gets the id of the series that an instance belongs to.</span></span>

<span data-ttu-id="e5c37-663">В Outlook в Интернете и на настольных клиентах `seriesId` возвращается идентификатор веб-служб Exchange (EWS) родительского элемента (ряда), к которому принадлежит этот элемент.</span><span class="sxs-lookup"><span data-stu-id="e5c37-663">In Outlook on the web and desktop clients, the `seriesId` returns the Exchange Web Services (EWS) ID of the parent (series) item that this item belongs to.</span></span> <span data-ttu-id="e5c37-664">Однако в iOS и Android `seriesId` возвращается идентификатор REST родительского элемента.</span><span class="sxs-lookup"><span data-stu-id="e5c37-664">However, in iOS and Android, the `seriesId` returns the REST ID of the parent item.</span></span>

> [!NOTE]
> <span data-ttu-id="e5c37-665">Идентификатор, возвращаемый свойством `seriesId`, совпадает с идентификатором элемента веб-служб Exchange.</span><span class="sxs-lookup"><span data-stu-id="e5c37-665">The identifier returned by the `seriesId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="e5c37-666">`seriesId` Свойство не совпадает с идентификаторами Outlook, используемыми в REST API Outlook.</span><span class="sxs-lookup"><span data-stu-id="e5c37-666">The `seriesId` property is not identical to the Outlook IDs used by the Outlook REST API.</span></span> <span data-ttu-id="e5c37-667">Прежде чем совершать вызовы REST API, используя это значение, его необходимо преобразовать с помощью [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="e5c37-667">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="e5c37-668">Дополнительные сведения см. в статье [Использование REST API Outlook из надстройки Outlook](/outlook/add-ins/use-rest-api).</span><span class="sxs-lookup"><span data-stu-id="e5c37-668">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api).</span></span>

<span data-ttu-id="e5c37-669">`seriesId` Свойство возвращает `null` элементы, у которых нет родительских элементов, таких как одиночные встречи, элементы ряда или приглашения на собрание, `undefined` и возвращаемые для других элементов, не являющиеся приглашениями на собрания.</span><span class="sxs-lookup"><span data-stu-id="e5c37-669">The `seriesId` property returns `null` for items that do not have parent items such as single appointments, series items, or meeting requests and returns `undefined` for any other items that are not meeting requests.</span></span>

##### <a name="type"></a><span data-ttu-id="e5c37-670">Тип</span><span class="sxs-lookup"><span data-stu-id="e5c37-670">Type</span></span>

* <span data-ttu-id="e5c37-671">String</span><span class="sxs-lookup"><span data-stu-id="e5c37-671">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="e5c37-672">Требования</span><span class="sxs-lookup"><span data-stu-id="e5c37-672">Requirements</span></span>

|<span data-ttu-id="e5c37-673">Требование</span><span class="sxs-lookup"><span data-stu-id="e5c37-673">Requirement</span></span>|<span data-ttu-id="e5c37-674">Значение</span><span class="sxs-lookup"><span data-stu-id="e5c37-674">Value</span></span>|
|---|---|
|[<span data-ttu-id="e5c37-675">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="e5c37-675">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e5c37-676">1.7</span><span class="sxs-lookup"><span data-stu-id="e5c37-676">1.7</span></span>|
|[<span data-ttu-id="e5c37-677">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e5c37-677">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e5c37-678">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e5c37-678">ReadItem</span></span>|
|[<span data-ttu-id="e5c37-679">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e5c37-679">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="e5c37-680">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="e5c37-680">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e5c37-681">Пример</span><span class="sxs-lookup"><span data-stu-id="e5c37-681">Example</span></span>

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

#### <a name="start-datetimejavascriptapioutlookofficetime"></a><span data-ttu-id="e5c37-682">Начало: Дата | [Time (время](/javascript/api/outlook/office.time) )</span><span class="sxs-lookup"><span data-stu-id="e5c37-682">start: Date|[Time](/javascript/api/outlook/office.time)</span></span>

<span data-ttu-id="e5c37-683">Получает или задает дату и время начала встречи.</span><span class="sxs-lookup"><span data-stu-id="e5c37-683">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="e5c37-p132">Свойство `start` представлено в виде значения даты и времени в формате UTC. Это значение можно преобразовать в местные значения даты и времени клиента с помощью метода [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime).</span><span class="sxs-lookup"><span data-stu-id="e5c37-p132">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="e5c37-686">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="e5c37-686">Read mode</span></span>

<span data-ttu-id="e5c37-687">Свойство `start` возвращает объект `Date`.</span><span class="sxs-lookup"><span data-stu-id="e5c37-687">The `start` property returns a `Date` object.</span></span>

```js
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="e5c37-688">Режим создания</span><span class="sxs-lookup"><span data-stu-id="e5c37-688">Compose mode</span></span>

<span data-ttu-id="e5c37-689">Свойство `start` возвращает объект `Time`.</span><span class="sxs-lookup"><span data-stu-id="e5c37-689">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="e5c37-690">Если вы задаете время начала с помощью метода [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-), необходимо использовать метод [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) для преобразования местного времени на клиенте в формат UTC для сервера.</span><span class="sxs-lookup"><span data-stu-id="e5c37-690">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="e5c37-691">В примере ниже с помощью метода [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) объекта `Time` задается время начала встречи в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="e5c37-691">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="e5c37-692">Тип</span><span class="sxs-lookup"><span data-stu-id="e5c37-692">Type</span></span>

*   <span data-ttu-id="e5c37-693">Date | [Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="e5c37-693">Date | [Time](/javascript/api/outlook/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="e5c37-694">Требования</span><span class="sxs-lookup"><span data-stu-id="e5c37-694">Requirements</span></span>

|<span data-ttu-id="e5c37-695">Требование</span><span class="sxs-lookup"><span data-stu-id="e5c37-695">Requirement</span></span>|<span data-ttu-id="e5c37-696">Значение</span><span class="sxs-lookup"><span data-stu-id="e5c37-696">Value</span></span>|
|---|---|
|[<span data-ttu-id="e5c37-697">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="e5c37-697">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e5c37-698">1.0</span><span class="sxs-lookup"><span data-stu-id="e5c37-698">1.0</span></span>|
|[<span data-ttu-id="e5c37-699">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e5c37-699">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e5c37-700">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e5c37-700">ReadItem</span></span>|
|[<span data-ttu-id="e5c37-701">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e5c37-701">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="e5c37-702">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="e5c37-702">Compose or Read</span></span>|

<br>

---
---

#### <a name="subject-stringsubjectjavascriptapioutlookofficesubject"></a><span data-ttu-id="e5c37-703">Тема: строка | [Subject (тема](/javascript/api/outlook/office.subject) )</span><span class="sxs-lookup"><span data-stu-id="e5c37-703">subject: String|[Subject](/javascript/api/outlook/office.subject)</span></span>

<span data-ttu-id="e5c37-704">Получает или задает описание, которое отображается в поле темы элемента.</span><span class="sxs-lookup"><span data-stu-id="e5c37-704">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="e5c37-705">Свойство `subject` получает или задает всю тему элемента для отправки с почтового сервера.</span><span class="sxs-lookup"><span data-stu-id="e5c37-705">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="e5c37-706">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="e5c37-706">Read mode</span></span>

<span data-ttu-id="e5c37-p133">Свойство `subject` возвращает строку. С помощью свойства [`normalizedSubject`](#normalizedsubject-string) можно получить тему без начальных префиксов, таких как `RE:` и `FW:`.</span><span class="sxs-lookup"><span data-stu-id="e5c37-p133">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

<span data-ttu-id="e5c37-709">В примере кода JavaScript, приведенном ниже, показано, как получить доступ к свойству `subject` текущего элемента в Outlook.</span><span class="sxs-lookup"><span data-stu-id="e5c37-709">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

```js
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="e5c37-710">Режим создания</span><span class="sxs-lookup"><span data-stu-id="e5c37-710">Compose mode</span></span>
<span data-ttu-id="e5c37-711">Свойство `subject` возвращает объект `Subject`, который предоставляет методы для получения и задания темы.</span><span class="sxs-lookup"><span data-stu-id="e5c37-711">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```js
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="e5c37-712">Тип</span><span class="sxs-lookup"><span data-stu-id="e5c37-712">Type</span></span>

*   <span data-ttu-id="e5c37-713">String | [Subject](/javascript/api/outlook/office.subject)</span><span class="sxs-lookup"><span data-stu-id="e5c37-713">String | [Subject](/javascript/api/outlook/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="e5c37-714">Требования</span><span class="sxs-lookup"><span data-stu-id="e5c37-714">Requirements</span></span>

|<span data-ttu-id="e5c37-715">Требование</span><span class="sxs-lookup"><span data-stu-id="e5c37-715">Requirement</span></span>|<span data-ttu-id="e5c37-716">Значение</span><span class="sxs-lookup"><span data-stu-id="e5c37-716">Value</span></span>|
|---|---|
|[<span data-ttu-id="e5c37-717">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="e5c37-717">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e5c37-718">1.0</span><span class="sxs-lookup"><span data-stu-id="e5c37-718">1.0</span></span>|
|[<span data-ttu-id="e5c37-719">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e5c37-719">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e5c37-720">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e5c37-720">ReadItem</span></span>|
|[<span data-ttu-id="e5c37-721">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e5c37-721">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="e5c37-722">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="e5c37-722">Compose or Read</span></span>|

<br>

---
---

#### <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="e5c37-723">Кому: Array. <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[получатели](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="e5c37-723">to: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="e5c37-724">Предоставляет доступ к получателям, указанным в строке **Кому** сообщения.</span><span class="sxs-lookup"><span data-stu-id="e5c37-724">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="e5c37-725">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="e5c37-725">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="e5c37-726">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="e5c37-726">Read mode</span></span>

<span data-ttu-id="e5c37-p135">Свойство `to` возвращает массив, содержащий объект `EmailAddressDetails` для каждого получателя в строке **Кому** сообщения. Коллекция может включать не более 100 элементов.</span><span class="sxs-lookup"><span data-stu-id="e5c37-p135">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="e5c37-729">Режим создания</span><span class="sxs-lookup"><span data-stu-id="e5c37-729">Compose mode</span></span>

<span data-ttu-id="e5c37-730">Свойство `to` возвращает объект `Recipients`, предоставляющий методы для получения или обновления получателей, которые указаны в строке **Кому** сообщения.</span><span class="sxs-lookup"><span data-stu-id="e5c37-730">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

```js
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="e5c37-731">Тип</span><span class="sxs-lookup"><span data-stu-id="e5c37-731">Type</span></span>

*   <span data-ttu-id="e5c37-732">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="e5c37-732">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="e5c37-733">Требования</span><span class="sxs-lookup"><span data-stu-id="e5c37-733">Requirements</span></span>

|<span data-ttu-id="e5c37-734">Требование</span><span class="sxs-lookup"><span data-stu-id="e5c37-734">Requirement</span></span>|<span data-ttu-id="e5c37-735">Значение</span><span class="sxs-lookup"><span data-stu-id="e5c37-735">Value</span></span>|
|---|---|
|[<span data-ttu-id="e5c37-736">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="e5c37-736">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e5c37-737">1.0</span><span class="sxs-lookup"><span data-stu-id="e5c37-737">1.0</span></span>|
|[<span data-ttu-id="e5c37-738">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e5c37-738">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e5c37-739">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e5c37-739">ReadItem</span></span>|
|[<span data-ttu-id="e5c37-740">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e5c37-740">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="e5c37-741">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="e5c37-741">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="e5c37-742">Методы</span><span class="sxs-lookup"><span data-stu-id="e5c37-742">Methods</span></span>

#### <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="e5c37-743">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="e5c37-743">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="e5c37-744">Добавляет файл в сообщение или встречу в качестве вложения.</span><span class="sxs-lookup"><span data-stu-id="e5c37-744">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="e5c37-745">Метод `addFileAttachmentAsync` передает файл по указанному универсальному коду ресурса (URI) и вкладывает его в элемент в форме создания.</span><span class="sxs-lookup"><span data-stu-id="e5c37-745">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="e5c37-746">Идентификатор можно использовать с методом [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback), чтобы удалить вложение, добавленное во время текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="e5c37-746">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e5c37-747">Параметры</span><span class="sxs-lookup"><span data-stu-id="e5c37-747">Parameters</span></span>
|<span data-ttu-id="e5c37-748">Имя</span><span class="sxs-lookup"><span data-stu-id="e5c37-748">Name</span></span>|<span data-ttu-id="e5c37-749">Тип</span><span class="sxs-lookup"><span data-stu-id="e5c37-749">Type</span></span>|<span data-ttu-id="e5c37-750">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="e5c37-750">Attributes</span></span>|<span data-ttu-id="e5c37-751">Описание</span><span class="sxs-lookup"><span data-stu-id="e5c37-751">Description</span></span>|
|---|---|---|---|
|`uri`|<span data-ttu-id="e5c37-752">String</span><span class="sxs-lookup"><span data-stu-id="e5c37-752">String</span></span>||<span data-ttu-id="e5c37-p136">Универсальный код ресурса (URI), представляющий расположение файла, который нужно вложить в сообщение или встречу. Максимальная длина — 2048 символов.</span><span class="sxs-lookup"><span data-stu-id="e5c37-p136">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="e5c37-755">String</span><span class="sxs-lookup"><span data-stu-id="e5c37-755">String</span></span>||<span data-ttu-id="e5c37-p137">Имя вложения, которое отображается при передаче вложения. Максимальная длина — 255 символов.</span><span class="sxs-lookup"><span data-stu-id="e5c37-p137">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="e5c37-758">Объект</span><span class="sxs-lookup"><span data-stu-id="e5c37-758">Object</span></span>|<span data-ttu-id="e5c37-759">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="e5c37-759">&lt;optional&gt;</span></span>|<span data-ttu-id="e5c37-760">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="e5c37-760">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="e5c37-761">Object</span><span class="sxs-lookup"><span data-stu-id="e5c37-761">Object</span></span>|<span data-ttu-id="e5c37-762">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="e5c37-762">&lt;optional&gt;</span></span>|<span data-ttu-id="e5c37-763">В методе обратного вызова разработчики могут указать любой объект, к которому необходимо получить доступ.</span><span class="sxs-lookup"><span data-stu-id="e5c37-763">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="e5c37-764">Boolean</span><span class="sxs-lookup"><span data-stu-id="e5c37-764">Boolean</span></span>|<span data-ttu-id="e5c37-765">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="e5c37-765">&lt;optional&gt;</span></span>|<span data-ttu-id="e5c37-766">Значение `true` указывает на то, что вложение будет встроено в текст сообщения и не должно отображаться в списке вложений.</span><span class="sxs-lookup"><span data-stu-id="e5c37-766">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="e5c37-767">function</span><span class="sxs-lookup"><span data-stu-id="e5c37-767">function</span></span>|<span data-ttu-id="e5c37-768">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="e5c37-768">&lt;optional&gt;</span></span>|<span data-ttu-id="e5c37-769">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="e5c37-769">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="e5c37-770">После успешного выполнения идентификатор вложения будет представлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="e5c37-770">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="e5c37-771">Если добавить вложение не удастся, объект `asyncResult` будет содержать объект `Error` с описанием ошибки.</span><span class="sxs-lookup"><span data-stu-id="e5c37-771">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="e5c37-772">Ошибки</span><span class="sxs-lookup"><span data-stu-id="e5c37-772">Errors</span></span>

|<span data-ttu-id="e5c37-773">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="e5c37-773">Error code</span></span>|<span data-ttu-id="e5c37-774">Описание</span><span class="sxs-lookup"><span data-stu-id="e5c37-774">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="e5c37-775">Вложение превышает максимальный размер.</span><span class="sxs-lookup"><span data-stu-id="e5c37-775">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="e5c37-776">Расширение вложения не поддерживается.</span><span class="sxs-lookup"><span data-stu-id="e5c37-776">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="e5c37-777">Сообщение или встреча содержат слишком много вложений.</span><span class="sxs-lookup"><span data-stu-id="e5c37-777">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e5c37-778">Требования</span><span class="sxs-lookup"><span data-stu-id="e5c37-778">Requirements</span></span>

|<span data-ttu-id="e5c37-779">Требование</span><span class="sxs-lookup"><span data-stu-id="e5c37-779">Requirement</span></span>|<span data-ttu-id="e5c37-780">Значение</span><span class="sxs-lookup"><span data-stu-id="e5c37-780">Value</span></span>|
|---|---|
|[<span data-ttu-id="e5c37-781">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="e5c37-781">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e5c37-782">1.1</span><span class="sxs-lookup"><span data-stu-id="e5c37-782">1.1</span></span>|
|[<span data-ttu-id="e5c37-783">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e5c37-783">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e5c37-784">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="e5c37-784">ReadWriteItem</span></span>|
|[<span data-ttu-id="e5c37-785">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e5c37-785">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="e5c37-786">Создание</span><span class="sxs-lookup"><span data-stu-id="e5c37-786">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="e5c37-787">Примеры</span><span class="sxs-lookup"><span data-stu-id="e5c37-787">Examples</span></span>

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

<span data-ttu-id="e5c37-788">В приведенном ниже примере файл изображения добавляется в качестве встроенного вложения, а ссылка на вложение добавляется в текст сообщения.</span><span class="sxs-lookup"><span data-stu-id="e5c37-788">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

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

#### <a name="addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback"></a><span data-ttu-id="e5c37-789">addFileAttachmentFromBase64Async (base64File, Аттачментнаме, [параметры], [обратный вызов])</span><span class="sxs-lookup"><span data-stu-id="e5c37-789">addFileAttachmentFromBase64Async(base64File, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="e5c37-790">Добавляет файл из кодировки Base64 в сообщение или встречу в качестве вложения.</span><span class="sxs-lookup"><span data-stu-id="e5c37-790">Adds a file from the base64 encoding to a message or appointment as an attachment.</span></span>

<span data-ttu-id="e5c37-791">`addFileAttachmentFromBase64Async` Метод передает файл из кодировки Base64 и вкладывает его в элемент в форме создания.</span><span class="sxs-lookup"><span data-stu-id="e5c37-791">The `addFileAttachmentFromBase64Async` method uploads the file from the base64 encoding and attaches it to the item in the compose form.</span></span> <span data-ttu-id="e5c37-792">Этот метод возвращает идентификатор вложения в объекте AsyncResult. Value.</span><span class="sxs-lookup"><span data-stu-id="e5c37-792">This method returns the attachment identifier in the AsyncResult.value object.</span></span>

<span data-ttu-id="e5c37-793">Идентификатор можно использовать с методом [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback), чтобы удалить вложение, добавленное во время текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="e5c37-793">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e5c37-794">Параметры</span><span class="sxs-lookup"><span data-stu-id="e5c37-794">Parameters</span></span>

|<span data-ttu-id="e5c37-795">Имя</span><span class="sxs-lookup"><span data-stu-id="e5c37-795">Name</span></span>|<span data-ttu-id="e5c37-796">Тип</span><span class="sxs-lookup"><span data-stu-id="e5c37-796">Type</span></span>|<span data-ttu-id="e5c37-797">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="e5c37-797">Attributes</span></span>|<span data-ttu-id="e5c37-798">Описание</span><span class="sxs-lookup"><span data-stu-id="e5c37-798">Description</span></span>|
|---|---|---|---|
|`base64File`|<span data-ttu-id="e5c37-799">String</span><span class="sxs-lookup"><span data-stu-id="e5c37-799">String</span></span>||<span data-ttu-id="e5c37-800">Содержимое изображения или файла в кодировке Base64, которое добавляется в сообщение электронной почты или событие.</span><span class="sxs-lookup"><span data-stu-id="e5c37-800">The base64 encoded content of an image or file to be added to an email or event.</span></span>|
|`attachmentName`|<span data-ttu-id="e5c37-801">String.</span><span class="sxs-lookup"><span data-stu-id="e5c37-801">String</span></span>||<span data-ttu-id="e5c37-p139">Имя вложения, которое отображается при передаче вложения. Максимальная длина — 255 символов.</span><span class="sxs-lookup"><span data-stu-id="e5c37-p139">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="e5c37-804">Объект</span><span class="sxs-lookup"><span data-stu-id="e5c37-804">Object</span></span>|<span data-ttu-id="e5c37-805">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="e5c37-805">&lt;optional&gt;</span></span>|<span data-ttu-id="e5c37-806">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="e5c37-806">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="e5c37-807">Объект</span><span class="sxs-lookup"><span data-stu-id="e5c37-807">Object</span></span>|<span data-ttu-id="e5c37-808">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="e5c37-808">&lt;optional&gt;</span></span>|<span data-ttu-id="e5c37-809">В методе обратного вызова разработчики могут указать любой объект, к которому необходимо получить доступ.</span><span class="sxs-lookup"><span data-stu-id="e5c37-809">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="e5c37-810">Boolean</span><span class="sxs-lookup"><span data-stu-id="e5c37-810">Boolean</span></span>|<span data-ttu-id="e5c37-811">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="e5c37-811">&lt;optional&gt;</span></span>|<span data-ttu-id="e5c37-812">Значение `true` указывает на то, что вложение будет встроено в текст сообщения и не должно отображаться в списке вложений.</span><span class="sxs-lookup"><span data-stu-id="e5c37-812">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="e5c37-813">function</span><span class="sxs-lookup"><span data-stu-id="e5c37-813">function</span></span>|<span data-ttu-id="e5c37-814">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="e5c37-814">&lt;optional&gt;</span></span>|<span data-ttu-id="e5c37-815">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="e5c37-815">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="e5c37-816">После успешного выполнения идентификатор вложения будет представлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="e5c37-816">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="e5c37-817">Если добавить вложение не удастся, объект `asyncResult` будет содержать объект `Error` с описанием ошибки.</span><span class="sxs-lookup"><span data-stu-id="e5c37-817">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="e5c37-818">Ошибки</span><span class="sxs-lookup"><span data-stu-id="e5c37-818">Errors</span></span>

|<span data-ttu-id="e5c37-819">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="e5c37-819">Error code</span></span>|<span data-ttu-id="e5c37-820">Описание</span><span class="sxs-lookup"><span data-stu-id="e5c37-820">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="e5c37-821">Вложение превышает максимальный размер.</span><span class="sxs-lookup"><span data-stu-id="e5c37-821">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="e5c37-822">Расширение вложения не поддерживается.</span><span class="sxs-lookup"><span data-stu-id="e5c37-822">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="e5c37-823">Сообщение или встреча содержат слишком много вложений.</span><span class="sxs-lookup"><span data-stu-id="e5c37-823">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e5c37-824">Требования</span><span class="sxs-lookup"><span data-stu-id="e5c37-824">Requirements</span></span>

|<span data-ttu-id="e5c37-825">Требование</span><span class="sxs-lookup"><span data-stu-id="e5c37-825">Requirement</span></span>|<span data-ttu-id="e5c37-826">Значение</span><span class="sxs-lookup"><span data-stu-id="e5c37-826">Value</span></span>|
|---|---|
|[<span data-ttu-id="e5c37-827">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="e5c37-827">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e5c37-828">Предварительная версия</span><span class="sxs-lookup"><span data-stu-id="e5c37-828">Preview</span></span>|
|[<span data-ttu-id="e5c37-829">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e5c37-829">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e5c37-830">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="e5c37-830">ReadWriteItem</span></span>|
|[<span data-ttu-id="e5c37-831">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e5c37-831">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="e5c37-832">Создание</span><span class="sxs-lookup"><span data-stu-id="e5c37-832">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="e5c37-833">Примеры</span><span class="sxs-lookup"><span data-stu-id="e5c37-833">Examples</span></span>

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

#### <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="e5c37-834">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="e5c37-834">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="e5c37-835">Добавляет обработчик для поддерживаемого события.</span><span class="sxs-lookup"><span data-stu-id="e5c37-835">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="e5c37-836">В настоящее время поддерживаются типы `Office.EventType.AttachmentsChanged`событий `Office.EventType.AppointmentTimeChanged`, `Office.EventType.EnhancedLocationsChanged` `Office.EventType.RecipientsChanged`,, и `Office.EventType.RecurrenceChanged`.</span><span class="sxs-lookup"><span data-stu-id="e5c37-836">Currently the supported event types are `Office.EventType.AttachmentsChanged`, `Office.EventType.AppointmentTimeChanged`, `Office.EventType.EnhancedLocationsChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e5c37-837">Параметры</span><span class="sxs-lookup"><span data-stu-id="e5c37-837">Parameters</span></span>

| <span data-ttu-id="e5c37-838">Имя</span><span class="sxs-lookup"><span data-stu-id="e5c37-838">Name</span></span> | <span data-ttu-id="e5c37-839">Тип</span><span class="sxs-lookup"><span data-stu-id="e5c37-839">Type</span></span> | <span data-ttu-id="e5c37-840">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="e5c37-840">Attributes</span></span> | <span data-ttu-id="e5c37-841">Описание</span><span class="sxs-lookup"><span data-stu-id="e5c37-841">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="e5c37-842">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="e5c37-842">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="e5c37-843">Событие, которое должно вызвать обработчик.</span><span class="sxs-lookup"><span data-stu-id="e5c37-843">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="e5c37-844">Function</span><span class="sxs-lookup"><span data-stu-id="e5c37-844">Function</span></span> || <span data-ttu-id="e5c37-p140">Функция для обработки события. Функция должна принимать один параметр, представляющий собой объектный литерал. Значение свойства `type` параметра совпадет со значением параметра `eventType`, переданного методу `addHandlerAsync`.</span><span class="sxs-lookup"><span data-stu-id="e5c37-p140">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="e5c37-848">Объект</span><span class="sxs-lookup"><span data-stu-id="e5c37-848">Object</span></span> | <span data-ttu-id="e5c37-849">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="e5c37-849">&lt;optional&gt;</span></span> | <span data-ttu-id="e5c37-850">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="e5c37-850">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="e5c37-851">Объект</span><span class="sxs-lookup"><span data-stu-id="e5c37-851">Object</span></span> | <span data-ttu-id="e5c37-852">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="e5c37-852">&lt;optional&gt;</span></span> | <span data-ttu-id="e5c37-853">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="e5c37-853">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="e5c37-854">функция</span><span class="sxs-lookup"><span data-stu-id="e5c37-854">function</span></span>| <span data-ttu-id="e5c37-855">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="e5c37-855">&lt;optional&gt;</span></span>|<span data-ttu-id="e5c37-856">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="e5c37-856">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e5c37-857">Требования</span><span class="sxs-lookup"><span data-stu-id="e5c37-857">Requirements</span></span>

|<span data-ttu-id="e5c37-858">Требование</span><span class="sxs-lookup"><span data-stu-id="e5c37-858">Requirement</span></span>| <span data-ttu-id="e5c37-859">Значение</span><span class="sxs-lookup"><span data-stu-id="e5c37-859">Value</span></span>|
|---|---|
|[<span data-ttu-id="e5c37-860">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="e5c37-860">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e5c37-861">1.7</span><span class="sxs-lookup"><span data-stu-id="e5c37-861">1.7</span></span> |
|[<span data-ttu-id="e5c37-862">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e5c37-862">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e5c37-863">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e5c37-863">ReadItem</span></span> |
|[<span data-ttu-id="e5c37-864">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e5c37-864">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e5c37-865">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="e5c37-865">Compose or Read</span></span> |

##### <a name="example"></a><span data-ttu-id="e5c37-866">Пример</span><span class="sxs-lookup"><span data-stu-id="e5c37-866">Example</span></span>

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

#### <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="e5c37-867">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="e5c37-867">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="e5c37-868">Добавляет к сообщению элемент Exchange, например сообщение, в виде вложения.</span><span class="sxs-lookup"><span data-stu-id="e5c37-868">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="e5c37-p141">С помощью метода `addItemAttachmentAsync` можно в элемент формы создания вложить элемент с указанным идентификатором Exchange. Если указать метод обратного вызова, то этот метод вызывается с помощью параметра `asyncResult`, который содержит идентификатор вложения или код, указывающий на ошибки, которые произошли при вложении элемента. При необходимости можно использовать параметр `options` для передачи сведений о состоянии методу обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="e5c37-p141">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="e5c37-872">Идентификатор можно использовать с методом [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback), чтобы удалить вложение, добавленное во время текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="e5c37-872">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="e5c37-873">Если ваша надстройка Office работает в Outlook в Интернете, `addItemAttachmentAsync` метод может присоединять элементы к элементам, отличным от редактируемого элемента; Однако это не поддерживается и не рекомендуется.</span><span class="sxs-lookup"><span data-stu-id="e5c37-873">If your Office Add-in is running in Outlook on the web, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e5c37-874">Параметры</span><span class="sxs-lookup"><span data-stu-id="e5c37-874">Parameters</span></span>

|<span data-ttu-id="e5c37-875">Имя</span><span class="sxs-lookup"><span data-stu-id="e5c37-875">Name</span></span>|<span data-ttu-id="e5c37-876">Тип</span><span class="sxs-lookup"><span data-stu-id="e5c37-876">Type</span></span>|<span data-ttu-id="e5c37-877">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="e5c37-877">Attributes</span></span>|<span data-ttu-id="e5c37-878">Описание</span><span class="sxs-lookup"><span data-stu-id="e5c37-878">Description</span></span>|
|---|---|---|---|
|`itemId`|<span data-ttu-id="e5c37-879">String</span><span class="sxs-lookup"><span data-stu-id="e5c37-879">String</span></span>||<span data-ttu-id="e5c37-p142">Идентификатор Exchange для вкладываемого элемента. Максимальная длина — 100 символов.</span><span class="sxs-lookup"><span data-stu-id="e5c37-p142">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="e5c37-882">String</span><span class="sxs-lookup"><span data-stu-id="e5c37-882">String</span></span>||<span data-ttu-id="e5c37-883">Тема вкладываемого элемента.</span><span class="sxs-lookup"><span data-stu-id="e5c37-883">The subject of the item to be attached.</span></span> <span data-ttu-id="e5c37-884">Максимальная длина: 255 символов.</span><span class="sxs-lookup"><span data-stu-id="e5c37-884">The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="e5c37-885">Object</span><span class="sxs-lookup"><span data-stu-id="e5c37-885">Object</span></span>|<span data-ttu-id="e5c37-886">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="e5c37-886">&lt;optional&gt;</span></span>|<span data-ttu-id="e5c37-887">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="e5c37-887">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="e5c37-888">Объект</span><span class="sxs-lookup"><span data-stu-id="e5c37-888">Object</span></span>|<span data-ttu-id="e5c37-889">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="e5c37-889">&lt;optional&gt;</span></span>|<span data-ttu-id="e5c37-890">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="e5c37-890">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="e5c37-891">функция</span><span class="sxs-lookup"><span data-stu-id="e5c37-891">function</span></span>|<span data-ttu-id="e5c37-892">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="e5c37-892">&lt;optional&gt;</span></span>|<span data-ttu-id="e5c37-893">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="e5c37-893">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="e5c37-894">После успешного выполнения идентификатор вложения будет представлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="e5c37-894">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="e5c37-895">Если добавить вложение не удастся, объект `asyncResult` будет содержать объект `Error` с описанием ошибки.</span><span class="sxs-lookup"><span data-stu-id="e5c37-895">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="e5c37-896">Ошибки</span><span class="sxs-lookup"><span data-stu-id="e5c37-896">Errors</span></span>

|<span data-ttu-id="e5c37-897">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="e5c37-897">Error code</span></span>|<span data-ttu-id="e5c37-898">Описание</span><span class="sxs-lookup"><span data-stu-id="e5c37-898">Description</span></span>|
|------------|-------------|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="e5c37-899">Сообщение или встреча содержат слишком много вложений.</span><span class="sxs-lookup"><span data-stu-id="e5c37-899">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e5c37-900">Требования</span><span class="sxs-lookup"><span data-stu-id="e5c37-900">Requirements</span></span>

|<span data-ttu-id="e5c37-901">Требование</span><span class="sxs-lookup"><span data-stu-id="e5c37-901">Requirement</span></span>|<span data-ttu-id="e5c37-902">Значение</span><span class="sxs-lookup"><span data-stu-id="e5c37-902">Value</span></span>|
|---|---|
|[<span data-ttu-id="e5c37-903">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="e5c37-903">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e5c37-904">1.1</span><span class="sxs-lookup"><span data-stu-id="e5c37-904">1.1</span></span>|
|[<span data-ttu-id="e5c37-905">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e5c37-905">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e5c37-906">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="e5c37-906">ReadWriteItem</span></span>|
|[<span data-ttu-id="e5c37-907">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e5c37-907">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="e5c37-908">Создание</span><span class="sxs-lookup"><span data-stu-id="e5c37-908">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="e5c37-909">Пример</span><span class="sxs-lookup"><span data-stu-id="e5c37-909">Example</span></span>

<span data-ttu-id="e5c37-910">В следующем примере существующий элемент Outlook добавляется в виде вложения с именем `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="e5c37-910">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

#### <a name="close"></a><span data-ttu-id="e5c37-911">close()</span><span class="sxs-lookup"><span data-stu-id="e5c37-911">close()</span></span>

<span data-ttu-id="e5c37-912">Закрывает текущий создаваемый элемент.</span><span class="sxs-lookup"><span data-stu-id="e5c37-912">Closes the current item that is being composed.</span></span>

<span data-ttu-id="e5c37-p144">Работа метода `close` зависит от текущего состояния создаваемого элемента. Если элемент содержит несохраненные изменения, клиент предложит пользователю сохранить или отклонить их либо отменить действие закрытия.</span><span class="sxs-lookup"><span data-stu-id="e5c37-p144">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="e5c37-915">Если в Outlook в Интернете элемент представляет собой встречу, ранее сохраненную с помощью метода `saveAsync`, пользователю предлагается сохранить, отклонить или отменить действие, даже если с момента последнего сохранения элемента не вносились какие-либо изменения.</span><span class="sxs-lookup"><span data-stu-id="e5c37-915">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="e5c37-916">Если в клиенте Outlook для настольных ПК сообщение представляет собой ответ в тексте, метод `close` не работает.</span><span class="sxs-lookup"><span data-stu-id="e5c37-916">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="e5c37-917">Требования</span><span class="sxs-lookup"><span data-stu-id="e5c37-917">Requirements</span></span>

|<span data-ttu-id="e5c37-918">Требование</span><span class="sxs-lookup"><span data-stu-id="e5c37-918">Requirement</span></span>|<span data-ttu-id="e5c37-919">Значение</span><span class="sxs-lookup"><span data-stu-id="e5c37-919">Value</span></span>|
|---|---|
|[<span data-ttu-id="e5c37-920">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="e5c37-920">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e5c37-921">1.3</span><span class="sxs-lookup"><span data-stu-id="e5c37-921">1.3</span></span>|
|[<span data-ttu-id="e5c37-922">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e5c37-922">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e5c37-923">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="e5c37-923">Restricted</span></span>|
|[<span data-ttu-id="e5c37-924">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e5c37-924">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="e5c37-925">Создание</span><span class="sxs-lookup"><span data-stu-id="e5c37-925">Compose</span></span>|

<br>

---
---

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="e5c37-926">displayReplyAllForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="e5c37-926">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="e5c37-927">Отображает форму ответа, включающую отправителя и всех получателей выбранного сообщения или организатора и всех участников выбранной встречи.</span><span class="sxs-lookup"><span data-stu-id="e5c37-927">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="e5c37-928">Этот метод не поддерживается в Outlook на iOS или Android.</span><span class="sxs-lookup"><span data-stu-id="e5c37-928">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="e5c37-929">В Outlook в Интернете форма ответа отображается в виде всплывающей формы в представлении из трех столбцов и всплывающей формы в представлении с 2 или 1 столбца.</span><span class="sxs-lookup"><span data-stu-id="e5c37-929">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="e5c37-930">Если любой строковый параметр превышает указанные для него ограничения, `displayReplyAllForm` возвращает исключение.</span><span class="sxs-lookup"><span data-stu-id="e5c37-930">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="e5c37-931">Если в `formData.attachments` параметре указаны вложения, Outlook в Интернете и клиенте для настольных компьютеров пытаются скачать все вложения и присоединить их к форме ответа.</span><span class="sxs-lookup"><span data-stu-id="e5c37-931">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form.</span></span> <span data-ttu-id="e5c37-932">Если какие-либо вложения добавить не удается, в форме отображается сообщение об ошибке.</span><span class="sxs-lookup"><span data-stu-id="e5c37-932">If any attachments fail to be added, an error is shown in the form UI.</span></span> <span data-ttu-id="e5c37-933">Если сообщения об ошибках не предусмотрены, то они не отображаются.</span><span class="sxs-lookup"><span data-stu-id="e5c37-933">If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e5c37-934">Параметры</span><span class="sxs-lookup"><span data-stu-id="e5c37-934">Parameters</span></span>

|<span data-ttu-id="e5c37-935">Имя</span><span class="sxs-lookup"><span data-stu-id="e5c37-935">Name</span></span>|<span data-ttu-id="e5c37-936">Тип</span><span class="sxs-lookup"><span data-stu-id="e5c37-936">Type</span></span>|<span data-ttu-id="e5c37-937">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="e5c37-937">Attributes</span></span>|<span data-ttu-id="e5c37-938">Описание</span><span class="sxs-lookup"><span data-stu-id="e5c37-938">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="e5c37-939">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="e5c37-939">String &#124; Object</span></span>||<span data-ttu-id="e5c37-p146">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="e5c37-p146">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="e5c37-942">**ИЛИ**</span><span class="sxs-lookup"><span data-stu-id="e5c37-942">**OR**</span></span><br/><span data-ttu-id="e5c37-p147">Объект, который содержит текст или данные вложения и функцию обратного вызова. Ниже представлено определение этого объекта.</span><span class="sxs-lookup"><span data-stu-id="e5c37-p147">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="e5c37-945">String.</span><span class="sxs-lookup"><span data-stu-id="e5c37-945">String</span></span>|<span data-ttu-id="e5c37-946">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="e5c37-946">&lt;optional&gt;</span></span>|<span data-ttu-id="e5c37-p148">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="e5c37-p148">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="e5c37-949">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="e5c37-949">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="e5c37-950">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="e5c37-950">&lt;optional&gt;</span></span>|<span data-ttu-id="e5c37-951">Массив объектов JSON, представляющих собой вложенные файлы или элементы.</span><span class="sxs-lookup"><span data-stu-id="e5c37-951">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="e5c37-952">String.</span><span class="sxs-lookup"><span data-stu-id="e5c37-952">String</span></span>||<span data-ttu-id="e5c37-p149">Указывает тип вложения. Допустимые значения: `file` для вложенного файла и `item` для вложенного элемента.</span><span class="sxs-lookup"><span data-stu-id="e5c37-p149">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="e5c37-955">Строка</span><span class="sxs-lookup"><span data-stu-id="e5c37-955">String</span></span>||<span data-ttu-id="e5c37-956">Строка, содержащая имя вложения, длиной до 255 символов.</span><span class="sxs-lookup"><span data-stu-id="e5c37-956">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="e5c37-957">String</span><span class="sxs-lookup"><span data-stu-id="e5c37-957">String</span></span>||<span data-ttu-id="e5c37-p150">Используется, только если свойству `type` задано значение `file`. URI расположения файла.</span><span class="sxs-lookup"><span data-stu-id="e5c37-p150">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="e5c37-960">Логический</span><span class="sxs-lookup"><span data-stu-id="e5c37-960">Boolean</span></span>||<span data-ttu-id="e5c37-p151">Используется, только если свойству `type` задано значение `file`. Значение `true` указывает на то, что вложение будет встроено в текст сообщения и не должно отображаться в списке вложений.</span><span class="sxs-lookup"><span data-stu-id="e5c37-p151">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="e5c37-963">String</span><span class="sxs-lookup"><span data-stu-id="e5c37-963">String</span></span>||<span data-ttu-id="e5c37-p152">Используется, только если свойству `type` задано значение `item`. Идентификатор вложения EWS. Это строка длиной до 100 символов.</span><span class="sxs-lookup"><span data-stu-id="e5c37-p152">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="e5c37-967">function</span><span class="sxs-lookup"><span data-stu-id="e5c37-967">function</span></span>|<span data-ttu-id="e5c37-968">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="e5c37-968">&lt;optional&gt;</span></span>|<span data-ttu-id="e5c37-969">По завершении работы метода функция, переданная параметру `callback`, вызывается с помощью одного параметра `asyncResult`, представляющего собой объект [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="e5c37-969">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e5c37-970">Требования</span><span class="sxs-lookup"><span data-stu-id="e5c37-970">Requirements</span></span>

|<span data-ttu-id="e5c37-971">Требование</span><span class="sxs-lookup"><span data-stu-id="e5c37-971">Requirement</span></span>|<span data-ttu-id="e5c37-972">Значение</span><span class="sxs-lookup"><span data-stu-id="e5c37-972">Value</span></span>|
|---|---|
|[<span data-ttu-id="e5c37-973">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="e5c37-973">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e5c37-974">1.0</span><span class="sxs-lookup"><span data-stu-id="e5c37-974">1.0</span></span>|
|[<span data-ttu-id="e5c37-975">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e5c37-975">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e5c37-976">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e5c37-976">ReadItem</span></span>|
|[<span data-ttu-id="e5c37-977">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e5c37-977">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="e5c37-978">Чтение</span><span class="sxs-lookup"><span data-stu-id="e5c37-978">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="e5c37-979">Примеры</span><span class="sxs-lookup"><span data-stu-id="e5c37-979">Examples</span></span>

<span data-ttu-id="e5c37-980">Приведенный ниже код передает строку в функцию `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="e5c37-980">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="e5c37-981">Ответ с пустым текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="e5c37-981">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="e5c37-982">Ответ только с текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="e5c37-982">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="e5c37-983">Ответ с текстом сообщения и вложенным файлом.</span><span class="sxs-lookup"><span data-stu-id="e5c37-983">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="e5c37-984">Ответ с текстом сообщения и вложенным элементом.</span><span class="sxs-lookup"><span data-stu-id="e5c37-984">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="e5c37-985">Ответ с текстом сообщения, вложенным файлом, вложенным элементом и обратным вызовом.</span><span class="sxs-lookup"><span data-stu-id="e5c37-985">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="e5c37-986">displayReplyForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="e5c37-986">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="e5c37-987">Отображает форму ответа, включающую только отправителя выбранного сообщения или организатора выбранной встречи.</span><span class="sxs-lookup"><span data-stu-id="e5c37-987">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="e5c37-988">Этот метод не поддерживается в Outlook на iOS или Android.</span><span class="sxs-lookup"><span data-stu-id="e5c37-988">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="e5c37-989">В Outlook в Интернете форма ответа отображается в виде всплывающей формы в представлении из трех столбцов и всплывающей формы в представлении с 2 или 1 столбца.</span><span class="sxs-lookup"><span data-stu-id="e5c37-989">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="e5c37-990">Если любой строковый параметр превышает указанные для него ограничения, `displayReplyForm` возвращает исключение.</span><span class="sxs-lookup"><span data-stu-id="e5c37-990">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="e5c37-991">Если в `formData.attachments` параметре указаны вложения, Outlook в Интернете и клиенте для настольных компьютеров пытаются скачать все вложения и присоединить их к форме ответа.</span><span class="sxs-lookup"><span data-stu-id="e5c37-991">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form.</span></span> <span data-ttu-id="e5c37-992">Если какие-либо вложения добавить не удается, в форме отображается сообщение об ошибке.</span><span class="sxs-lookup"><span data-stu-id="e5c37-992">If any attachments fail to be added, an error is shown in the form UI.</span></span> <span data-ttu-id="e5c37-993">Если сообщения об ошибках не предусмотрены, то они не отображаются.</span><span class="sxs-lookup"><span data-stu-id="e5c37-993">If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e5c37-994">Параметры</span><span class="sxs-lookup"><span data-stu-id="e5c37-994">Parameters</span></span>

|<span data-ttu-id="e5c37-995">Имя</span><span class="sxs-lookup"><span data-stu-id="e5c37-995">Name</span></span>|<span data-ttu-id="e5c37-996">Тип</span><span class="sxs-lookup"><span data-stu-id="e5c37-996">Type</span></span>|<span data-ttu-id="e5c37-997">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="e5c37-997">Attributes</span></span>|<span data-ttu-id="e5c37-998">Описание</span><span class="sxs-lookup"><span data-stu-id="e5c37-998">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="e5c37-999">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="e5c37-999">String &#124; Object</span></span>||<span data-ttu-id="e5c37-p154">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="e5c37-p154">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="e5c37-1002">**ИЛИ**</span><span class="sxs-lookup"><span data-stu-id="e5c37-1002">**OR**</span></span><br/><span data-ttu-id="e5c37-p155">Объект, который содержит текст или данные вложения и функцию обратного вызова. Ниже представлено определение этого объекта.</span><span class="sxs-lookup"><span data-stu-id="e5c37-p155">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="e5c37-1005">String.</span><span class="sxs-lookup"><span data-stu-id="e5c37-1005">String</span></span>|<span data-ttu-id="e5c37-1006">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="e5c37-1006">&lt;optional&gt;</span></span>|<span data-ttu-id="e5c37-p156">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="e5c37-p156">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="e5c37-1009">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="e5c37-1009">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="e5c37-1010">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="e5c37-1010">&lt;optional&gt;</span></span>|<span data-ttu-id="e5c37-1011">Массив объектов JSON, представляющих собой вложенные файлы или элементы.</span><span class="sxs-lookup"><span data-stu-id="e5c37-1011">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="e5c37-1012">String.</span><span class="sxs-lookup"><span data-stu-id="e5c37-1012">String</span></span>||<span data-ttu-id="e5c37-p157">Указывает тип вложения. Допустимые значения: `file` для вложенного файла и `item` для вложенного элемента.</span><span class="sxs-lookup"><span data-stu-id="e5c37-p157">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="e5c37-1015">Строка</span><span class="sxs-lookup"><span data-stu-id="e5c37-1015">String</span></span>||<span data-ttu-id="e5c37-1016">Строка, содержащая имя вложения, длиной до 255 символов.</span><span class="sxs-lookup"><span data-stu-id="e5c37-1016">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="e5c37-1017">String</span><span class="sxs-lookup"><span data-stu-id="e5c37-1017">String</span></span>||<span data-ttu-id="e5c37-p158">Используется, только если свойству `type` задано значение `file`. URI расположения файла.</span><span class="sxs-lookup"><span data-stu-id="e5c37-p158">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="e5c37-1020">Логический</span><span class="sxs-lookup"><span data-stu-id="e5c37-1020">Boolean</span></span>||<span data-ttu-id="e5c37-p159">Используется, только если свойству `type` задано значение `file`. Значение `true` указывает на то, что вложение будет встроено в текст сообщения и не должно отображаться в списке вложений.</span><span class="sxs-lookup"><span data-stu-id="e5c37-p159">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="e5c37-1023">String</span><span class="sxs-lookup"><span data-stu-id="e5c37-1023">String</span></span>||<span data-ttu-id="e5c37-p160">Используется, только если свойству `type` задано значение `item`. Идентификатор вложения EWS. Это строка длиной до 100 символов.</span><span class="sxs-lookup"><span data-stu-id="e5c37-p160">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="e5c37-1027">function</span><span class="sxs-lookup"><span data-stu-id="e5c37-1027">function</span></span>|<span data-ttu-id="e5c37-1028">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="e5c37-1028">&lt;optional&gt;</span></span>|<span data-ttu-id="e5c37-1029">По завершении работы метода функция, переданная параметру `callback`, вызывается с помощью одного параметра `asyncResult`, представляющего собой объект [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="e5c37-1029">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e5c37-1030">Требования</span><span class="sxs-lookup"><span data-stu-id="e5c37-1030">Requirements</span></span>

|<span data-ttu-id="e5c37-1031">Требование</span><span class="sxs-lookup"><span data-stu-id="e5c37-1031">Requirement</span></span>|<span data-ttu-id="e5c37-1032">Значение</span><span class="sxs-lookup"><span data-stu-id="e5c37-1032">Value</span></span>|
|---|---|
|[<span data-ttu-id="e5c37-1033">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="e5c37-1033">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e5c37-1034">1.0</span><span class="sxs-lookup"><span data-stu-id="e5c37-1034">1.0</span></span>|
|[<span data-ttu-id="e5c37-1035">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e5c37-1035">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e5c37-1036">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e5c37-1036">ReadItem</span></span>|
|[<span data-ttu-id="e5c37-1037">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e5c37-1037">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="e5c37-1038">Чтение</span><span class="sxs-lookup"><span data-stu-id="e5c37-1038">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="e5c37-1039">Примеры</span><span class="sxs-lookup"><span data-stu-id="e5c37-1039">Examples</span></span>

<span data-ttu-id="e5c37-1040">Приведенный ниже код передает строку в функцию `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="e5c37-1040">The following code passes a string to the `displayReplyForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="e5c37-1041">Ответ с пустым текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="e5c37-1041">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="e5c37-1042">Ответ только с текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="e5c37-1042">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="e5c37-1043">Ответ с текстом сообщения и вложенным файлом.</span><span class="sxs-lookup"><span data-stu-id="e5c37-1043">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="e5c37-1044">Ответ с текстом сообщения и вложенным элементом.</span><span class="sxs-lookup"><span data-stu-id="e5c37-1044">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="e5c37-1045">Ответ с текстом сообщения, вложенным файлом, вложенным элементом и обратным вызовом.</span><span class="sxs-lookup"><span data-stu-id="e5c37-1045">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getattachmentcontentasyncattachmentid-options-callback--attachmentcontentjavascriptapioutlookofficeattachmentcontent"></a><span data-ttu-id="e5c37-1046">Жетаттачментконтентасинк (attachmentId, [параметры], [callback]) → [вложениеимеет содержимое](/javascript/api/outlook/office.attachmentcontent)</span><span class="sxs-lookup"><span data-stu-id="e5c37-1046">getAttachmentContentAsync(attachmentId, [options], [callback]) → [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span></span>

<span data-ttu-id="e5c37-1047">Получает указанное вложение из сообщения или встречи и возвращает его в виде `AttachmentContent` объекта.</span><span class="sxs-lookup"><span data-stu-id="e5c37-1047">Gets the specified attachment from a message or appointment and returns it as an `AttachmentContent` object.</span></span>

<span data-ttu-id="e5c37-1048">`getAttachmentContentAsync` Метод получает вложение с указанным идентификатором из элемента.</span><span class="sxs-lookup"><span data-stu-id="e5c37-1048">The `getAttachmentContentAsync` method gets the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="e5c37-1049">Рекомендуется использовать идентификатор для получения вложения в том же сеансе, когда Аттачментидс был получен с помощью вызова `getAttachmentsAsync` или. `item.attachments`</span><span class="sxs-lookup"><span data-stu-id="e5c37-1049">As a best practice, you should use the identifier to retrieve an attachment in the same session that the attachmentIds were retrieved with the `getAttachmentsAsync` or `item.attachments` call.</span></span> <span data-ttu-id="e5c37-1050">В Outlook в Интернете и мобильных устройствах идентификатор вложения действителен только в рамках одного сеанса.</span><span class="sxs-lookup"><span data-stu-id="e5c37-1050">In Outlook on the web and mobile devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="e5c37-1051">Сеанс переходит к моменту, когда пользователь закрывает приложение, или если пользователь начинает создание встроенной формы, затем извлекает форму, чтобы продолжить работу в отдельном окне.</span><span class="sxs-lookup"><span data-stu-id="e5c37-1051">A session is over when the user closes the app, or if the user starts composing an inline form then subsequently pops out the form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e5c37-1052">Параметры</span><span class="sxs-lookup"><span data-stu-id="e5c37-1052">Parameters</span></span>

|<span data-ttu-id="e5c37-1053">Имя</span><span class="sxs-lookup"><span data-stu-id="e5c37-1053">Name</span></span>|<span data-ttu-id="e5c37-1054">Тип</span><span class="sxs-lookup"><span data-stu-id="e5c37-1054">Type</span></span>|<span data-ttu-id="e5c37-1055">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="e5c37-1055">Attributes</span></span>|<span data-ttu-id="e5c37-1056">Описание</span><span class="sxs-lookup"><span data-stu-id="e5c37-1056">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="e5c37-1057">String</span><span class="sxs-lookup"><span data-stu-id="e5c37-1057">String</span></span>||<span data-ttu-id="e5c37-1058">Идентификатор вложения, которое требуется получить.</span><span class="sxs-lookup"><span data-stu-id="e5c37-1058">The identifier of the attachment you want to get.</span></span>|
|`options`|<span data-ttu-id="e5c37-1059">Объект</span><span class="sxs-lookup"><span data-stu-id="e5c37-1059">Object</span></span>|<span data-ttu-id="e5c37-1060">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="e5c37-1060">&lt;optional&gt;</span></span>|<span data-ttu-id="e5c37-1061">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="e5c37-1061">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="e5c37-1062">Объект</span><span class="sxs-lookup"><span data-stu-id="e5c37-1062">Object</span></span>|<span data-ttu-id="e5c37-1063">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="e5c37-1063">&lt;optional&gt;</span></span>|<span data-ttu-id="e5c37-1064">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="e5c37-1064">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="e5c37-1065">функция</span><span class="sxs-lookup"><span data-stu-id="e5c37-1065">function</span></span>|<span data-ttu-id="e5c37-1066">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="e5c37-1066">&lt;optional&gt;</span></span>|<span data-ttu-id="e5c37-1067">По завершении работы метода функция, переданная параметру `callback`, вызывается с помощью одного параметра `asyncResult`, представляющего собой объект [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="e5c37-1067">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e5c37-1068">Требования</span><span class="sxs-lookup"><span data-stu-id="e5c37-1068">Requirements</span></span>

|<span data-ttu-id="e5c37-1069">Требование</span><span class="sxs-lookup"><span data-stu-id="e5c37-1069">Requirement</span></span>|<span data-ttu-id="e5c37-1070">Значение</span><span class="sxs-lookup"><span data-stu-id="e5c37-1070">Value</span></span>|
|---|---|
|[<span data-ttu-id="e5c37-1071">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="e5c37-1071">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e5c37-1072">Предварительная версия</span><span class="sxs-lookup"><span data-stu-id="e5c37-1072">Preview</span></span>|
|[<span data-ttu-id="e5c37-1073">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e5c37-1073">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e5c37-1074">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e5c37-1074">ReadItem</span></span>|
|[<span data-ttu-id="e5c37-1075">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e5c37-1075">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="e5c37-1076">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="e5c37-1076">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="e5c37-1077">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="e5c37-1077">Returns:</span></span>

<span data-ttu-id="e5c37-1078">Тип: [вложениеимеет содержимое](/javascript/api/outlook/office.attachmentcontent)</span><span class="sxs-lookup"><span data-stu-id="e5c37-1078">Type: [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span></span>

##### <a name="example"></a><span data-ttu-id="e5c37-1079">Пример</span><span class="sxs-lookup"><span data-stu-id="e5c37-1079">Example</span></span>

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

#### <a name="getattachmentsasyncoptions-callback--arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails"></a><span data-ttu-id="e5c37-1080">Жетаттачментсасинк ([параметры], [обратный вызов]) → массив. <[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="e5c37-1080">getAttachmentsAsync([options], [callback]) → Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

<span data-ttu-id="e5c37-1081">Получает вложения элемента в виде массива.</span><span class="sxs-lookup"><span data-stu-id="e5c37-1081">Gets the item's attachments as an array.</span></span> <span data-ttu-id="e5c37-1082">Только в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="e5c37-1082">Compose mode only.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e5c37-1083">Параметры</span><span class="sxs-lookup"><span data-stu-id="e5c37-1083">Parameters</span></span>

|<span data-ttu-id="e5c37-1084">Имя</span><span class="sxs-lookup"><span data-stu-id="e5c37-1084">Name</span></span>|<span data-ttu-id="e5c37-1085">Тип</span><span class="sxs-lookup"><span data-stu-id="e5c37-1085">Type</span></span>|<span data-ttu-id="e5c37-1086">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="e5c37-1086">Attributes</span></span>|<span data-ttu-id="e5c37-1087">Описание</span><span class="sxs-lookup"><span data-stu-id="e5c37-1087">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="e5c37-1088">Object</span><span class="sxs-lookup"><span data-stu-id="e5c37-1088">Object</span></span>|<span data-ttu-id="e5c37-1089">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="e5c37-1089">&lt;optional&gt;</span></span>|<span data-ttu-id="e5c37-1090">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="e5c37-1090">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="e5c37-1091">Объект</span><span class="sxs-lookup"><span data-stu-id="e5c37-1091">Object</span></span>|<span data-ttu-id="e5c37-1092">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="e5c37-1092">&lt;optional&gt;</span></span>|<span data-ttu-id="e5c37-1093">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="e5c37-1093">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="e5c37-1094">функция</span><span class="sxs-lookup"><span data-stu-id="e5c37-1094">function</span></span>|<span data-ttu-id="e5c37-1095">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="e5c37-1095">&lt;optional&gt;</span></span>|<span data-ttu-id="e5c37-1096">По завершении работы метода функция, переданная параметру `callback`, вызывается с помощью одного параметра `asyncResult`, представляющего собой объект [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="e5c37-1096">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e5c37-1097">Требования</span><span class="sxs-lookup"><span data-stu-id="e5c37-1097">Requirements</span></span>

|<span data-ttu-id="e5c37-1098">Требование</span><span class="sxs-lookup"><span data-stu-id="e5c37-1098">Requirement</span></span>|<span data-ttu-id="e5c37-1099">Значение</span><span class="sxs-lookup"><span data-stu-id="e5c37-1099">Value</span></span>|
|---|---|
|[<span data-ttu-id="e5c37-1100">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="e5c37-1100">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e5c37-1101">Предварительная версия</span><span class="sxs-lookup"><span data-stu-id="e5c37-1101">Preview</span></span>|
|[<span data-ttu-id="e5c37-1102">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e5c37-1102">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e5c37-1103">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e5c37-1103">ReadItem</span></span>|
|[<span data-ttu-id="e5c37-1104">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e5c37-1104">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="e5c37-1105">Создание</span><span class="sxs-lookup"><span data-stu-id="e5c37-1105">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="e5c37-1106">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="e5c37-1106">Returns:</span></span>

<span data-ttu-id="e5c37-1107">Тип: Array. <[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="e5c37-1107">Type: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

##### <a name="example"></a><span data-ttu-id="e5c37-1108">Пример</span><span class="sxs-lookup"><span data-stu-id="e5c37-1108">Example</span></span>

<span data-ttu-id="e5c37-1109">В приведенном ниже примере создается строка HTML со сведениями обо всех вложениях в текущем элементе.</span><span class="sxs-lookup"><span data-stu-id="e5c37-1109">The following example builds an HTML string with details of all attachments on the current item.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlookofficeentities"></a><span data-ttu-id="e5c37-1110">getEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="e5c37-1110">getEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span></span>

<span data-ttu-id="e5c37-1111">Получает сущности, обнаруженные в теле выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="e5c37-1111">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="e5c37-1112">Этот метод не поддерживается в Outlook на iOS или Android.</span><span class="sxs-lookup"><span data-stu-id="e5c37-1112">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="e5c37-1113">Требования</span><span class="sxs-lookup"><span data-stu-id="e5c37-1113">Requirements</span></span>

|<span data-ttu-id="e5c37-1114">Требование</span><span class="sxs-lookup"><span data-stu-id="e5c37-1114">Requirement</span></span>|<span data-ttu-id="e5c37-1115">Значение</span><span class="sxs-lookup"><span data-stu-id="e5c37-1115">Value</span></span>|
|---|---|
|[<span data-ttu-id="e5c37-1116">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="e5c37-1116">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e5c37-1117">1.0</span><span class="sxs-lookup"><span data-stu-id="e5c37-1117">1.0</span></span>|
|[<span data-ttu-id="e5c37-1118">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e5c37-1118">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e5c37-1119">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e5c37-1119">ReadItem</span></span>|
|[<span data-ttu-id="e5c37-1120">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e5c37-1120">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="e5c37-1121">Чтение</span><span class="sxs-lookup"><span data-stu-id="e5c37-1121">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="e5c37-1122">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="e5c37-1122">Returns:</span></span>

<span data-ttu-id="e5c37-1123">Тип: [Entities](/javascript/api/outlook/office.entities)</span><span class="sxs-lookup"><span data-stu-id="e5c37-1123">Type: [Entities](/javascript/api/outlook/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="e5c37-1124">Пример</span><span class="sxs-lookup"><span data-stu-id="e5c37-1124">Example</span></span>

<span data-ttu-id="e5c37-1125">Ниже приведен пример получения доступа к сущностям контактов в теле текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="e5c37-1125">The following example accesses the contacts entities in the current item's body.</span></span>

```js
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

<br>

---
---

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion"></a><span data-ttu-id="e5c37-1126">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="e5c37-1126">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span></span>

<span data-ttu-id="e5c37-1127">Получает массив всех сущностей указанного типа, обнаруженных в теле выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="e5c37-1127">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="e5c37-1128">Этот метод не поддерживается в Outlook на iOS или Android.</span><span class="sxs-lookup"><span data-stu-id="e5c37-1128">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e5c37-1129">Параметры</span><span class="sxs-lookup"><span data-stu-id="e5c37-1129">Parameters</span></span>

|<span data-ttu-id="e5c37-1130">Имя</span><span class="sxs-lookup"><span data-stu-id="e5c37-1130">Name</span></span>|<span data-ttu-id="e5c37-1131">Тип</span><span class="sxs-lookup"><span data-stu-id="e5c37-1131">Type</span></span>|<span data-ttu-id="e5c37-1132">Описание</span><span class="sxs-lookup"><span data-stu-id="e5c37-1132">Description</span></span>|
|---|---|---|
|`entityType`|[<span data-ttu-id="e5c37-1133">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="e5c37-1133">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook/office.mailboxenums.entitytype)|<span data-ttu-id="e5c37-1134">Одно из значений перечисления EntityType.</span><span class="sxs-lookup"><span data-stu-id="e5c37-1134">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e5c37-1135">Требования</span><span class="sxs-lookup"><span data-stu-id="e5c37-1135">Requirements</span></span>

|<span data-ttu-id="e5c37-1136">Требование</span><span class="sxs-lookup"><span data-stu-id="e5c37-1136">Requirement</span></span>|<span data-ttu-id="e5c37-1137">Значение</span><span class="sxs-lookup"><span data-stu-id="e5c37-1137">Value</span></span>|
|---|---|
|[<span data-ttu-id="e5c37-1138">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="e5c37-1138">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e5c37-1139">1.0</span><span class="sxs-lookup"><span data-stu-id="e5c37-1139">1.0</span></span>|
|[<span data-ttu-id="e5c37-1140">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e5c37-1140">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e5c37-1141">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="e5c37-1141">Restricted</span></span>|
|[<span data-ttu-id="e5c37-1142">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e5c37-1142">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="e5c37-1143">Чтение</span><span class="sxs-lookup"><span data-stu-id="e5c37-1143">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="e5c37-1144">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="e5c37-1144">Returns:</span></span>

<span data-ttu-id="e5c37-1145">Если значение, переданное в `entityType`, не является допустимым членом перечисления `EntityType`, метод возвращает значение NULL.</span><span class="sxs-lookup"><span data-stu-id="e5c37-1145">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="e5c37-1146">Если в теле элемента отсутствуют сущности указанного типа, метод возвращает пустой массив.</span><span class="sxs-lookup"><span data-stu-id="e5c37-1146">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="e5c37-1147">В противном случае тип объектов в возвращаемом массиве зависит от типа сущности, запрошенной в параметре `entityType`.</span><span class="sxs-lookup"><span data-stu-id="e5c37-1147">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="e5c37-1148">Хотя минимальный уровень разрешений для использования этого метода — **Restricted**, для некоторых типов сущностей требуется доступ на уровне **ReadItem**, как указано в приведенной ниже таблице.</span><span class="sxs-lookup"><span data-stu-id="e5c37-1148">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

|<span data-ttu-id="e5c37-1149">Значение параметра `entityType`</span><span class="sxs-lookup"><span data-stu-id="e5c37-1149">Value of `entityType`</span></span>|<span data-ttu-id="e5c37-1150">Тип объектов в возвращаемом массиве</span><span class="sxs-lookup"><span data-stu-id="e5c37-1150">Type of objects in returned array</span></span>|<span data-ttu-id="e5c37-1151">Необходимый уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e5c37-1151">Required Permission Level</span></span>|
|---|---|---|
|`Address`|<span data-ttu-id="e5c37-1152">String</span><span class="sxs-lookup"><span data-stu-id="e5c37-1152">String</span></span>|<span data-ttu-id="e5c37-1153">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="e5c37-1153">**Restricted**</span></span>|
|`Contact`|<span data-ttu-id="e5c37-1154">Contact</span><span class="sxs-lookup"><span data-stu-id="e5c37-1154">Contact</span></span>|<span data-ttu-id="e5c37-1155">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="e5c37-1155">**ReadItem**</span></span>|
|`EmailAddress`|<span data-ttu-id="e5c37-1156">String</span><span class="sxs-lookup"><span data-stu-id="e5c37-1156">String</span></span>|<span data-ttu-id="e5c37-1157">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="e5c37-1157">**ReadItem**</span></span>|
|`MeetingSuggestion`|<span data-ttu-id="e5c37-1158">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="e5c37-1158">MeetingSuggestion</span></span>|<span data-ttu-id="e5c37-1159">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="e5c37-1159">**ReadItem**</span></span>|
|`PhoneNumber`|<span data-ttu-id="e5c37-1160">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="e5c37-1160">PhoneNumber</span></span>|<span data-ttu-id="e5c37-1161">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="e5c37-1161">**Restricted**</span></span>|
|`TaskSuggestion`|<span data-ttu-id="e5c37-1162">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="e5c37-1162">TaskSuggestion</span></span>|<span data-ttu-id="e5c37-1163">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="e5c37-1163">**ReadItem**</span></span>|
|`URL`|<span data-ttu-id="e5c37-1164">String</span><span class="sxs-lookup"><span data-stu-id="e5c37-1164">String</span></span>|<span data-ttu-id="e5c37-1165">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="e5c37-1165">**Restricted**</span></span>|

<span data-ttu-id="e5c37-1166">Тип: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="e5c37-1166">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="e5c37-1167">Пример</span><span class="sxs-lookup"><span data-stu-id="e5c37-1167">Example</span></span>

<span data-ttu-id="e5c37-1168">В примере ниже показано, как получить доступ к массиву строк, которые представляют собой почтовые адреса в теле текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="e5c37-1168">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion"></a><span data-ttu-id="e5c37-1169">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="e5c37-1169">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span></span>

<span data-ttu-id="e5c37-1170">Возвращает известные сущности в выбранном элементе, которые проходят через именованный фильтр, определяемый в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="e5c37-1170">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="e5c37-1171">Этот метод не поддерживается в Outlook на iOS или Android.</span><span class="sxs-lookup"><span data-stu-id="e5c37-1171">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="e5c37-1172">Метод `getFilteredEntitiesByName` возвращает сущности, соответствующие регулярному выражению, которое определяется в элементе правила [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) в XML-файле манифеста, с использованием указанного значения элемента `FilterName`.</span><span class="sxs-lookup"><span data-stu-id="e5c37-1172">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e5c37-1173">Параметры</span><span class="sxs-lookup"><span data-stu-id="e5c37-1173">Parameters</span></span>

|<span data-ttu-id="e5c37-1174">Имя</span><span class="sxs-lookup"><span data-stu-id="e5c37-1174">Name</span></span>|<span data-ttu-id="e5c37-1175">Тип</span><span class="sxs-lookup"><span data-stu-id="e5c37-1175">Type</span></span>|<span data-ttu-id="e5c37-1176">Описание</span><span class="sxs-lookup"><span data-stu-id="e5c37-1176">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="e5c37-1177">String</span><span class="sxs-lookup"><span data-stu-id="e5c37-1177">String</span></span>|<span data-ttu-id="e5c37-1178">Имя элемента правила `ItemHasKnownEntity`, определяющее соответствующий фильтр.</span><span class="sxs-lookup"><span data-stu-id="e5c37-1178">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e5c37-1179">Требования</span><span class="sxs-lookup"><span data-stu-id="e5c37-1179">Requirements</span></span>

|<span data-ttu-id="e5c37-1180">Требование</span><span class="sxs-lookup"><span data-stu-id="e5c37-1180">Requirement</span></span>|<span data-ttu-id="e5c37-1181">Значение</span><span class="sxs-lookup"><span data-stu-id="e5c37-1181">Value</span></span>|
|---|---|
|[<span data-ttu-id="e5c37-1182">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="e5c37-1182">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e5c37-1183">1.0</span><span class="sxs-lookup"><span data-stu-id="e5c37-1183">1.0</span></span>|
|[<span data-ttu-id="e5c37-1184">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e5c37-1184">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e5c37-1185">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e5c37-1185">ReadItem</span></span>|
|[<span data-ttu-id="e5c37-1186">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e5c37-1186">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="e5c37-1187">Чтение</span><span class="sxs-lookup"><span data-stu-id="e5c37-1187">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="e5c37-1188">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="e5c37-1188">Returns:</span></span>

<span data-ttu-id="e5c37-p164">Если в манифесте нет элемента `ItemHasKnownEntity` со значением `FilterName`, соответствующим параметру `name`, метод возвращает `null`. Если параметр `name` соответствует элементу `ItemHasKnownEntity` в манифесте, но при этом в текущем элементе нет соответствующих сущностей, метод возвращает пустой массив.</span><span class="sxs-lookup"><span data-stu-id="e5c37-p164">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="e5c37-1191">Тип: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="e5c37-1191">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span></span>

<br>

---
---

#### <a name="getinitializationcontextasyncoptions-callback"></a><span data-ttu-id="e5c37-1192">getInitializationContextAsync ([параметры], [обратный вызов])</span><span class="sxs-lookup"><span data-stu-id="e5c37-1192">getInitializationContextAsync([options], [callback])</span></span>

<span data-ttu-id="e5c37-1193">Получает данные инициализации, передаваемые при активации надстройки [сообщением с действиями](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span><span class="sxs-lookup"><span data-stu-id="e5c37-1193">Gets initialization data passed when the add-in is [activated by an actionable message](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span></span>

> [!NOTE]
> <span data-ttu-id="e5c37-1194">Этот метод поддерживается только в Outlook 2016 или более поздней версии для Windows ("нажми и работай" более поздней версии, чем 16.0.8413.1000) и Outlook в Интернете для Office 365.</span><span class="sxs-lookup"><span data-stu-id="e5c37-1194">This method is only supported by Outlook 2016 or later on Windows (Click-to-Run versions later than 16.0.8413.1000) and Outlook on the web for Office 365.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e5c37-1195">Параметры</span><span class="sxs-lookup"><span data-stu-id="e5c37-1195">Parameters</span></span>

|<span data-ttu-id="e5c37-1196">Имя</span><span class="sxs-lookup"><span data-stu-id="e5c37-1196">Name</span></span>|<span data-ttu-id="e5c37-1197">Тип</span><span class="sxs-lookup"><span data-stu-id="e5c37-1197">Type</span></span>|<span data-ttu-id="e5c37-1198">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="e5c37-1198">Attributes</span></span>|<span data-ttu-id="e5c37-1199">Описание</span><span class="sxs-lookup"><span data-stu-id="e5c37-1199">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="e5c37-1200">Object</span><span class="sxs-lookup"><span data-stu-id="e5c37-1200">Object</span></span>|<span data-ttu-id="e5c37-1201">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="e5c37-1201">&lt;optional&gt;</span></span>|<span data-ttu-id="e5c37-1202">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="e5c37-1202">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="e5c37-1203">Объект</span><span class="sxs-lookup"><span data-stu-id="e5c37-1203">Object</span></span>|<span data-ttu-id="e5c37-1204">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="e5c37-1204">&lt;optional&gt;</span></span>|<span data-ttu-id="e5c37-1205">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="e5c37-1205">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="e5c37-1206">функция</span><span class="sxs-lookup"><span data-stu-id="e5c37-1206">function</span></span>|<span data-ttu-id="e5c37-1207">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="e5c37-1207">&lt;optional&gt;</span></span>|<span data-ttu-id="e5c37-1208">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="e5c37-1208">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="e5c37-1209">При успешном выполнении данные инициализации предоставляются в `asyncResult.value` свойстве в виде строки.</span><span class="sxs-lookup"><span data-stu-id="e5c37-1209">On success, the initialization data is provided in the `asyncResult.value` property as a string.</span></span><br/><span data-ttu-id="e5c37-1210">Если `asyncResult` контекст инициализации отсутствует, объект будет содержать `Error` объект со `code` свойством, `9020` `name` для свойства которого задано значение. `GenericResponseError`</span><span class="sxs-lookup"><span data-stu-id="e5c37-1210">If there is no initialization context, the `asyncResult` object will contain an `Error` object with its `code` property set to `9020` and its `name` property set to `GenericResponseError`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e5c37-1211">Требования</span><span class="sxs-lookup"><span data-stu-id="e5c37-1211">Requirements</span></span>

|<span data-ttu-id="e5c37-1212">Требование</span><span class="sxs-lookup"><span data-stu-id="e5c37-1212">Requirement</span></span>|<span data-ttu-id="e5c37-1213">Значение</span><span class="sxs-lookup"><span data-stu-id="e5c37-1213">Value</span></span>|
|---|---|
|[<span data-ttu-id="e5c37-1214">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="e5c37-1214">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e5c37-1215">Предварительная версия</span><span class="sxs-lookup"><span data-stu-id="e5c37-1215">Preview</span></span>|
|[<span data-ttu-id="e5c37-1216">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e5c37-1216">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e5c37-1217">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e5c37-1217">ReadItem</span></span>|
|[<span data-ttu-id="e5c37-1218">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e5c37-1218">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="e5c37-1219">Чтение</span><span class="sxs-lookup"><span data-stu-id="e5c37-1219">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e5c37-1220">Пример</span><span class="sxs-lookup"><span data-stu-id="e5c37-1220">Example</span></span>

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

#### <a name="getitemidasyncoptions-callback"></a><span data-ttu-id="e5c37-1221">Жетитемидасинк ([параметры], обратный вызов)</span><span class="sxs-lookup"><span data-stu-id="e5c37-1221">getItemIdAsync([options], callback)</span></span>

<span data-ttu-id="e5c37-1222">Асинхронно получает идентификатор сохраненного элемента.</span><span class="sxs-lookup"><span data-stu-id="e5c37-1222">Asynchronously gets the ID of a saved item.</span></span> <span data-ttu-id="e5c37-1223">Только в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="e5c37-1223">Compose mode only.</span></span>

<span data-ttu-id="e5c37-1224">При вызове этот метод возвращает идентификатор элемента с помощью метода обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="e5c37-1224">When invoked, this method returns the item ID via the callback method.</span></span>

> [!NOTE]
> <span data-ttu-id="e5c37-1225">Если надстройка вызывает `getItemIdAsync` элемент в режиме создания (например, чтобы получить доступ `itemId` к использованию с помощью EWS или REST API), имейте в виду, что если Outlook находится в режиме кэширования, может потребоваться некоторое время до синхронизации элемента с сервером.</span><span class="sxs-lookup"><span data-stu-id="e5c37-1225">If your add-in calls `getItemIdAsync` on an item in compose mode (e.g., to get an `itemId` to use with EWS or the REST API), be aware that when Outlook is in cached mode, it may take some time before the item is synced to the server.</span></span> <span data-ttu-id="e5c37-1226">Пока элемент не будет синхронизирован, он не `itemId` распознается и не будет использоваться, возвращается сообщение об ошибке.</span><span class="sxs-lookup"><span data-stu-id="e5c37-1226">Until the item is synced, the `itemId` is not recognized and using it returns an error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e5c37-1227">Параметры</span><span class="sxs-lookup"><span data-stu-id="e5c37-1227">Parameters</span></span>

|<span data-ttu-id="e5c37-1228">Имя</span><span class="sxs-lookup"><span data-stu-id="e5c37-1228">Name</span></span>|<span data-ttu-id="e5c37-1229">Тип</span><span class="sxs-lookup"><span data-stu-id="e5c37-1229">Type</span></span>|<span data-ttu-id="e5c37-1230">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="e5c37-1230">Attributes</span></span>|<span data-ttu-id="e5c37-1231">Описание</span><span class="sxs-lookup"><span data-stu-id="e5c37-1231">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="e5c37-1232">Object</span><span class="sxs-lookup"><span data-stu-id="e5c37-1232">Object</span></span>|<span data-ttu-id="e5c37-1233">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="e5c37-1233">&lt;optional&gt;</span></span>|<span data-ttu-id="e5c37-1234">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="e5c37-1234">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="e5c37-1235">Объект</span><span class="sxs-lookup"><span data-stu-id="e5c37-1235">Object</span></span>|<span data-ttu-id="e5c37-1236">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="e5c37-1236">&lt;optional&gt;</span></span>|<span data-ttu-id="e5c37-1237">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="e5c37-1237">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="e5c37-1238">функция</span><span class="sxs-lookup"><span data-stu-id="e5c37-1238">function</span></span>||<span data-ttu-id="e5c37-1239">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="e5c37-1239">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="e5c37-1240">После успешного выполнения идентификатор элемента будет предоставлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="e5c37-1240">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="errors"></a><span data-ttu-id="e5c37-1241">Ошибки</span><span class="sxs-lookup"><span data-stu-id="e5c37-1241">Errors</span></span>

|<span data-ttu-id="e5c37-1242">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="e5c37-1242">Error code</span></span>|<span data-ttu-id="e5c37-1243">Описание</span><span class="sxs-lookup"><span data-stu-id="e5c37-1243">Description</span></span>|
|------------|-------------|
|`ItemNotSaved`|<span data-ttu-id="e5c37-1244">Идентификатор невозможно извлечь, пока не будет сохранен элемент.</span><span class="sxs-lookup"><span data-stu-id="e5c37-1244">The id can't be retrieved until the item is saved.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e5c37-1245">Требования</span><span class="sxs-lookup"><span data-stu-id="e5c37-1245">Requirements</span></span>

|<span data-ttu-id="e5c37-1246">Требование</span><span class="sxs-lookup"><span data-stu-id="e5c37-1246">Requirement</span></span>|<span data-ttu-id="e5c37-1247">Значение</span><span class="sxs-lookup"><span data-stu-id="e5c37-1247">Value</span></span>|
|---|---|
|[<span data-ttu-id="e5c37-1248">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="e5c37-1248">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e5c37-1249">Предварительная версия</span><span class="sxs-lookup"><span data-stu-id="e5c37-1249">Preview</span></span>|
|[<span data-ttu-id="e5c37-1250">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e5c37-1250">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e5c37-1251">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e5c37-1251">ReadItem</span></span>|
|[<span data-ttu-id="e5c37-1252">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e5c37-1252">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="e5c37-1253">Создание</span><span class="sxs-lookup"><span data-stu-id="e5c37-1253">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="e5c37-1254">Примеры</span><span class="sxs-lookup"><span data-stu-id="e5c37-1254">Examples</span></span>

```js
Office.context.mailbox.item.getItemIdAsync(
  function callback(result) {
    // Process the result.
  });
```

<span data-ttu-id="e5c37-1255">В следующем примере показана структура `result` параметра, переданного функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="e5c37-1255">The following example shows the structure of the `result` parameter that's passed to the callback function.</span></span> <span data-ttu-id="e5c37-1256">`value` Свойство содержит идентификатор элемента.</span><span class="sxs-lookup"><span data-stu-id="e5c37-1256">The `value` property contains the item ID.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

<br>

---
---

#### <a name="getregexmatches--object"></a><span data-ttu-id="e5c37-1257">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="e5c37-1257">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="e5c37-1258">Возвращает строковые значения в выбранном элементе, которые соответствуют регулярным выражениям, определенным в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="e5c37-1258">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="e5c37-1259">Этот метод не поддерживается в Outlook на iOS или Android.</span><span class="sxs-lookup"><span data-stu-id="e5c37-1259">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="e5c37-p168">Метод `getRegExMatches` возвращает строки, соответствующие регулярному выражению, которое определяется в каждом элементе правила `ItemHasRegularExpressionMatch` или `ItemHasKnownEntity` в XML-файле манифеста. Для правила `ItemHasRegularExpressionMatch` соответствующую строку должно содержать свойство элемента, указанного этим правилом. Простой тип `PropertyName` определяет поддерживаемые свойства.</span><span class="sxs-lookup"><span data-stu-id="e5c37-p168">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="e5c37-1263">Например, рассмотрим манифест надстройки, который содержит указанный ниже элемент `Rule`.</span><span class="sxs-lookup"><span data-stu-id="e5c37-1263">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="e5c37-1264">Объект, возвращаемый методом `getRegExMatches`, будет содержать два свойства: `fruits` и `veggies`.</span><span class="sxs-lookup"><span data-stu-id="e5c37-1264">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="e5c37-p169">Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты. Лучше используйте метод [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) для этого.</span><span class="sxs-lookup"><span data-stu-id="e5c37-p169">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="e5c37-1268">Requirements</span><span class="sxs-lookup"><span data-stu-id="e5c37-1268">Requirements</span></span>

|<span data-ttu-id="e5c37-1269">Требование</span><span class="sxs-lookup"><span data-stu-id="e5c37-1269">Requirement</span></span>|<span data-ttu-id="e5c37-1270">Значение</span><span class="sxs-lookup"><span data-stu-id="e5c37-1270">Value</span></span>|
|---|---|
|[<span data-ttu-id="e5c37-1271">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="e5c37-1271">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e5c37-1272">1.0</span><span class="sxs-lookup"><span data-stu-id="e5c37-1272">1.0</span></span>|
|[<span data-ttu-id="e5c37-1273">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e5c37-1273">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e5c37-1274">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e5c37-1274">ReadItem</span></span>|
|[<span data-ttu-id="e5c37-1275">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e5c37-1275">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="e5c37-1276">Чтение</span><span class="sxs-lookup"><span data-stu-id="e5c37-1276">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="e5c37-1277">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="e5c37-1277">Returns:</span></span>

<span data-ttu-id="e5c37-p170">Объект, содержащий массив строк, которые соответствуют регулярным выражениям, определяемым в XML-файле манифеста. Имя каждого массива равно соответствующему значению атрибута `RegExName` подходящего правила `ItemHasRegularExpressionMatch` или атрибута `FilterName` соответствующего правила `ItemHasKnownEntity`.</span><span class="sxs-lookup"><span data-stu-id="e5c37-p170">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="e5c37-1280">

<dt>Тип</dt>

</span><span class="sxs-lookup"><span data-stu-id="e5c37-1280">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="e5c37-1281">Object</span><span class="sxs-lookup"><span data-stu-id="e5c37-1281">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="e5c37-1282">Пример</span><span class="sxs-lookup"><span data-stu-id="e5c37-1282">Example</span></span>

<span data-ttu-id="e5c37-1283">В приведенном ниже примере показано, как получить доступ к массиву совпадений с элементами `fruits` и `veggies` правил активации регулярных выражений, указанными в манифесте.</span><span class="sxs-lookup"><span data-stu-id="e5c37-1283">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```js
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

<br>

---
---

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="e5c37-1284">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="e5c37-1284">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="e5c37-1285">Возвращает строковые значения в выбранном элементе, которые соответствуют именованному регулярному выражению, определенному в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="e5c37-1285">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="e5c37-1286">Этот метод не поддерживается в Outlook на iOS или Android.</span><span class="sxs-lookup"><span data-stu-id="e5c37-1286">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="e5c37-1287">Метод `getRegExMatchesByName` возвращает строки, соответствующие регулярному выражению, которое определяется в элементе правила `ItemHasRegularExpressionMatch` в XML-файле манифеста, с использованием указанного значения элемента `RegExName`.</span><span class="sxs-lookup"><span data-stu-id="e5c37-1287">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="e5c37-p171">Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты.</span><span class="sxs-lookup"><span data-stu-id="e5c37-p171">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e5c37-1290">Параметры</span><span class="sxs-lookup"><span data-stu-id="e5c37-1290">Parameters</span></span>

|<span data-ttu-id="e5c37-1291">Имя</span><span class="sxs-lookup"><span data-stu-id="e5c37-1291">Name</span></span>|<span data-ttu-id="e5c37-1292">Тип</span><span class="sxs-lookup"><span data-stu-id="e5c37-1292">Type</span></span>|<span data-ttu-id="e5c37-1293">Описание</span><span class="sxs-lookup"><span data-stu-id="e5c37-1293">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="e5c37-1294">String</span><span class="sxs-lookup"><span data-stu-id="e5c37-1294">String</span></span>|<span data-ttu-id="e5c37-1295">Имя элемента правила `ItemHasRegularExpressionMatch`, определяющее соответствующий фильтр.</span><span class="sxs-lookup"><span data-stu-id="e5c37-1295">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e5c37-1296">Требования</span><span class="sxs-lookup"><span data-stu-id="e5c37-1296">Requirements</span></span>

|<span data-ttu-id="e5c37-1297">Требование</span><span class="sxs-lookup"><span data-stu-id="e5c37-1297">Requirement</span></span>|<span data-ttu-id="e5c37-1298">Значение</span><span class="sxs-lookup"><span data-stu-id="e5c37-1298">Value</span></span>|
|---|---|
|[<span data-ttu-id="e5c37-1299">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="e5c37-1299">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e5c37-1300">1.0</span><span class="sxs-lookup"><span data-stu-id="e5c37-1300">1.0</span></span>|
|[<span data-ttu-id="e5c37-1301">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e5c37-1301">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e5c37-1302">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e5c37-1302">ReadItem</span></span>|
|[<span data-ttu-id="e5c37-1303">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e5c37-1303">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="e5c37-1304">Чтение</span><span class="sxs-lookup"><span data-stu-id="e5c37-1304">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="e5c37-1305">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="e5c37-1305">Returns:</span></span>

<span data-ttu-id="e5c37-1306">Массив строк, соответствующих регулярному выражению, определяемому в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="e5c37-1306">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<span data-ttu-id="e5c37-1307">Тип: Array. < String ></span><span class="sxs-lookup"><span data-stu-id="e5c37-1307">Type: Array.< String ></span></span>

##### <a name="example"></a><span data-ttu-id="e5c37-1308">Пример</span><span class="sxs-lookup"><span data-stu-id="e5c37-1308">Example</span></span>

```js
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

<br>

---
---

#### <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="e5c37-1309">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="e5c37-1309">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="e5c37-1310">Асинхронно возвращает данные, выбранные в теме или тексте сообщения.</span><span class="sxs-lookup"><span data-stu-id="e5c37-1310">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="e5c37-p172">Если выделенный фрагмент отсутствует, но курсор находится в тексте или теме, метод возвращает значение NULL для выбранных данных. Если выбраны не текст и не тема, метод возвращает ошибку `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="e5c37-p172">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e5c37-1313">Параметры</span><span class="sxs-lookup"><span data-stu-id="e5c37-1313">Parameters</span></span>

|<span data-ttu-id="e5c37-1314">Имя</span><span class="sxs-lookup"><span data-stu-id="e5c37-1314">Name</span></span>|<span data-ttu-id="e5c37-1315">Тип</span><span class="sxs-lookup"><span data-stu-id="e5c37-1315">Type</span></span>|<span data-ttu-id="e5c37-1316">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="e5c37-1316">Attributes</span></span>|<span data-ttu-id="e5c37-1317">Описание</span><span class="sxs-lookup"><span data-stu-id="e5c37-1317">Description</span></span>|
|---|---|---|---|
|`coercionType`|[<span data-ttu-id="e5c37-1318">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="e5c37-1318">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="e5c37-p173">Запрашивает формат данных. Если задано значение Text, метод возвращает обычный текст как строку, удаляя все имеющиеся HTML-теги. Если задано значение HTML, метод возвращает выделенный текст (обычный текст или HTML).</span><span class="sxs-lookup"><span data-stu-id="e5c37-p173">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`|<span data-ttu-id="e5c37-1322">Объект</span><span class="sxs-lookup"><span data-stu-id="e5c37-1322">Object</span></span>|<span data-ttu-id="e5c37-1323">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="e5c37-1323">&lt;optional&gt;</span></span>|<span data-ttu-id="e5c37-1324">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="e5c37-1324">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="e5c37-1325">Объект</span><span class="sxs-lookup"><span data-stu-id="e5c37-1325">Object</span></span>|<span data-ttu-id="e5c37-1326">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="e5c37-1326">&lt;optional&gt;</span></span>|<span data-ttu-id="e5c37-1327">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="e5c37-1327">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="e5c37-1328">функция</span><span class="sxs-lookup"><span data-stu-id="e5c37-1328">function</span></span>||<span data-ttu-id="e5c37-1329">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="e5c37-1329">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="e5c37-1330">Чтобы получить доступ к выбранным данным из метода обратного вызова, вызовите `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="e5c37-1330">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="e5c37-1331">Для доступа к исходному свойству, представляющему собой источник выбранных данных, вызовите параметр `asyncResult.value.sourceProperty`, который может иметь значение `body` или `subject`.</span><span class="sxs-lookup"><span data-stu-id="e5c37-1331">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e5c37-1332">Требования</span><span class="sxs-lookup"><span data-stu-id="e5c37-1332">Requirements</span></span>

|<span data-ttu-id="e5c37-1333">Требование</span><span class="sxs-lookup"><span data-stu-id="e5c37-1333">Requirement</span></span>|<span data-ttu-id="e5c37-1334">Значение</span><span class="sxs-lookup"><span data-stu-id="e5c37-1334">Value</span></span>|
|---|---|
|[<span data-ttu-id="e5c37-1335">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="e5c37-1335">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e5c37-1336">1.2</span><span class="sxs-lookup"><span data-stu-id="e5c37-1336">1.2</span></span>|
|[<span data-ttu-id="e5c37-1337">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e5c37-1337">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e5c37-1338">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="e5c37-1338">ReadWriteItem</span></span>|
|[<span data-ttu-id="e5c37-1339">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e5c37-1339">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="e5c37-1340">Создание</span><span class="sxs-lookup"><span data-stu-id="e5c37-1340">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="e5c37-1341">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="e5c37-1341">Returns:</span></span>

<span data-ttu-id="e5c37-1342">Выбранные данные в виде строки с форматом, определенным в параметре `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="e5c37-1342">The selected data as a string with format determined by `coercionType`.</span></span>

<span data-ttu-id="e5c37-1343">Тип: String</span><span class="sxs-lookup"><span data-stu-id="e5c37-1343">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="e5c37-1344">Пример</span><span class="sxs-lookup"><span data-stu-id="e5c37-1344">Example</span></span>

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

#### <a name="getselectedentities--entitiesjavascriptapioutlookofficeentities"></a><span data-ttu-id="e5c37-1345">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="e5c37-1345">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span></span>

<span data-ttu-id="e5c37-1346">Возвращает сущности, найденные в выделенном совпадении, выбранном пользователем.</span><span class="sxs-lookup"><span data-stu-id="e5c37-1346">Gets the entities found in a highlighted match a user has selected.</span></span> <span data-ttu-id="e5c37-1347">Выделенные совпадения применяются к [контекстным надстройкам](/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="e5c37-1347">Highlighted matches apply to [contextual add-ins](/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="e5c37-1348">Этот метод не поддерживается в Outlook на iOS или Android.</span><span class="sxs-lookup"><span data-stu-id="e5c37-1348">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="e5c37-1349">Требования</span><span class="sxs-lookup"><span data-stu-id="e5c37-1349">Requirements</span></span>

|<span data-ttu-id="e5c37-1350">Требование</span><span class="sxs-lookup"><span data-stu-id="e5c37-1350">Requirement</span></span>|<span data-ttu-id="e5c37-1351">Значение</span><span class="sxs-lookup"><span data-stu-id="e5c37-1351">Value</span></span>|
|---|---|
|[<span data-ttu-id="e5c37-1352">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="e5c37-1352">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e5c37-1353">1.6</span><span class="sxs-lookup"><span data-stu-id="e5c37-1353">1.6</span></span>|
|[<span data-ttu-id="e5c37-1354">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e5c37-1354">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e5c37-1355">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e5c37-1355">ReadItem</span></span>|
|[<span data-ttu-id="e5c37-1356">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e5c37-1356">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="e5c37-1357">Чтение</span><span class="sxs-lookup"><span data-stu-id="e5c37-1357">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="e5c37-1358">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="e5c37-1358">Returns:</span></span>

<span data-ttu-id="e5c37-1359">Тип: [Entities](/javascript/api/outlook/office.entities)</span><span class="sxs-lookup"><span data-stu-id="e5c37-1359">Type: [Entities](/javascript/api/outlook/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="e5c37-1360">Пример</span><span class="sxs-lookup"><span data-stu-id="e5c37-1360">Example</span></span>

<span data-ttu-id="e5c37-1361">В приведенном ниже примере показано, как получить доступ к сущностям адресов в выделенном совпадении, выбранном пользователем.</span><span class="sxs-lookup"><span data-stu-id="e5c37-1361">The following example accesses the addresses entities in the highlighted match selected by the user.</span></span>

```js
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

<br>

---
---

#### <a name="getselectedregexmatches--object"></a><span data-ttu-id="e5c37-1362">getSelectedRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="e5c37-1362">getSelectedRegExMatches() → {Object}</span></span>

<span data-ttu-id="e5c37-p176">Возвращает строковые значения в выделенном совпадении, которые соответствуют регулярным выражениям, определенным в XML-файле манифеста. Выделенные совпадения применяются к [контекстным надстройкам](/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="e5c37-p176">Returns string values in a highlighted match that match the regular expressions defined in the manifest XML file. Highlighted matches apply to [contextual add-ins](/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="e5c37-1365">Этот метод не поддерживается в Outlook на iOS или Android.</span><span class="sxs-lookup"><span data-stu-id="e5c37-1365">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="e5c37-p177">Метод `getSelectedRegExMatches` возвращает строки, соответствующие регулярному выражению, которое определяется в каждом элементе правила `ItemHasRegularExpressionMatch` или `ItemHasKnownEntity` в XML-файле манифеста. Для правила `ItemHasRegularExpressionMatch` соответствующую строку должно содержать свойство элемента, указанного этим правилом. Простой тип `PropertyName` определяет поддерживаемые свойства.</span><span class="sxs-lookup"><span data-stu-id="e5c37-p177">The `getSelectedRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="e5c37-1369">Например, рассмотрим манифест надстройки, который содержит указанный ниже элемент `Rule`.</span><span class="sxs-lookup"><span data-stu-id="e5c37-1369">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="e5c37-1370">Объект, возвращаемый методом `getRegExMatches`, будет содержать два свойства: `fruits` и `veggies`.</span><span class="sxs-lookup"><span data-stu-id="e5c37-1370">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="e5c37-p178">Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты. Лучше используйте метод [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) для этого.</span><span class="sxs-lookup"><span data-stu-id="e5c37-p178">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="e5c37-1374">Requirements</span><span class="sxs-lookup"><span data-stu-id="e5c37-1374">Requirements</span></span>

|<span data-ttu-id="e5c37-1375">Требование</span><span class="sxs-lookup"><span data-stu-id="e5c37-1375">Requirement</span></span>|<span data-ttu-id="e5c37-1376">Значение</span><span class="sxs-lookup"><span data-stu-id="e5c37-1376">Value</span></span>|
|---|---|
|[<span data-ttu-id="e5c37-1377">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="e5c37-1377">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e5c37-1378">1.6</span><span class="sxs-lookup"><span data-stu-id="e5c37-1378">1.6</span></span>|
|[<span data-ttu-id="e5c37-1379">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e5c37-1379">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e5c37-1380">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e5c37-1380">ReadItem</span></span>|
|[<span data-ttu-id="e5c37-1381">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e5c37-1381">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="e5c37-1382">Чтение</span><span class="sxs-lookup"><span data-stu-id="e5c37-1382">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="e5c37-1383">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="e5c37-1383">Returns:</span></span>

<span data-ttu-id="e5c37-p179">Объект, содержащий массив строк, которые соответствуют регулярным выражениям, определяемым в XML-файле манифеста. Имя каждого массива равно соответствующему значению атрибута `RegExName` подходящего правила `ItemHasRegularExpressionMatch` или атрибута `FilterName` соответствующего правила `ItemHasKnownEntity`.</span><span class="sxs-lookup"><span data-stu-id="e5c37-p179">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

##### <a name="example"></a><span data-ttu-id="e5c37-1386">Пример</span><span class="sxs-lookup"><span data-stu-id="e5c37-1386">Example</span></span>

<span data-ttu-id="e5c37-1387">В приведенном ниже примере показано, как получить доступ к массиву совпадений с элементами `fruits` и `veggies` правил активации регулярных выражений, указанными в манифесте.</span><span class="sxs-lookup"><span data-stu-id="e5c37-1387">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```js
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

<br>

---
---

#### <a name="getsharedpropertiesasyncoptions-callback"></a><span data-ttu-id="e5c37-1388">Жетшаредпропертиесасинк ([параметры], обратный вызов)</span><span class="sxs-lookup"><span data-stu-id="e5c37-1388">getSharedPropertiesAsync([options], callback)</span></span>

<span data-ttu-id="e5c37-1389">Получает свойства выбранной встречи или сообщения в общей папке, календаре или почтовом ящике.</span><span class="sxs-lookup"><span data-stu-id="e5c37-1389">Gets the properties of the selected appointment or message in a shared folder, calendar, or mailbox.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e5c37-1390">Параметры</span><span class="sxs-lookup"><span data-stu-id="e5c37-1390">Parameters</span></span>

|<span data-ttu-id="e5c37-1391">Имя</span><span class="sxs-lookup"><span data-stu-id="e5c37-1391">Name</span></span>|<span data-ttu-id="e5c37-1392">Тип</span><span class="sxs-lookup"><span data-stu-id="e5c37-1392">Type</span></span>|<span data-ttu-id="e5c37-1393">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="e5c37-1393">Attributes</span></span>|<span data-ttu-id="e5c37-1394">Описание</span><span class="sxs-lookup"><span data-stu-id="e5c37-1394">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="e5c37-1395">Object</span><span class="sxs-lookup"><span data-stu-id="e5c37-1395">Object</span></span>|<span data-ttu-id="e5c37-1396">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="e5c37-1396">&lt;optional&gt;</span></span>|<span data-ttu-id="e5c37-1397">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="e5c37-1397">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="e5c37-1398">Объект</span><span class="sxs-lookup"><span data-stu-id="e5c37-1398">Object</span></span>|<span data-ttu-id="e5c37-1399">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="e5c37-1399">&lt;optional&gt;</span></span>|<span data-ttu-id="e5c37-1400">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="e5c37-1400">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="e5c37-1401">функция</span><span class="sxs-lookup"><span data-stu-id="e5c37-1401">function</span></span>||<span data-ttu-id="e5c37-1402">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="e5c37-1402">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="e5c37-1403">Общие свойства предоставляются в виде [`SharedProperties`](/javascript/api/outlook/office.sharedproperties) объекта в `asyncResult.value` свойстве.</span><span class="sxs-lookup"><span data-stu-id="e5c37-1403">The shared properties are provided as a [`SharedProperties`](/javascript/api/outlook/office.sharedproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="e5c37-1404">Этот объект можно использовать для получения общих свойств элемента.</span><span class="sxs-lookup"><span data-stu-id="e5c37-1404">This object can be used to get the item's shared properties.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e5c37-1405">Требования</span><span class="sxs-lookup"><span data-stu-id="e5c37-1405">Requirements</span></span>

|<span data-ttu-id="e5c37-1406">Требование</span><span class="sxs-lookup"><span data-stu-id="e5c37-1406">Requirement</span></span>|<span data-ttu-id="e5c37-1407">Значение</span><span class="sxs-lookup"><span data-stu-id="e5c37-1407">Value</span></span>|
|---|---|
|[<span data-ttu-id="e5c37-1408">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="e5c37-1408">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e5c37-1409">Предварительная версия</span><span class="sxs-lookup"><span data-stu-id="e5c37-1409">Preview</span></span>|
|[<span data-ttu-id="e5c37-1410">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e5c37-1410">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e5c37-1411">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e5c37-1411">ReadItem</span></span>|
|[<span data-ttu-id="e5c37-1412">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e5c37-1412">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="e5c37-1413">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="e5c37-1413">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e5c37-1414">Пример</span><span class="sxs-lookup"><span data-stu-id="e5c37-1414">Example</span></span>

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

#### <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="e5c37-1415">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="e5c37-1415">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="e5c37-1416">Асинхронно загружает настраиваемые свойства для надстройки для выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="e5c37-1416">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="e5c37-p181">Настраиваемые свойства сохраняются в виде пар "ключ-значение" для каждого приложения и каждого элемента. Этот метод возвращает объект `CustomProperties` при обратном вызове, который предоставляет методы для доступа к настраиваемым свойствам, характерным для текущего элемента и текущей надстройки. Настраиваемые свойства не шифруются для элемента, поэтому этот способ хранения не является безопасным.</span><span class="sxs-lookup"><span data-stu-id="e5c37-p181">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e5c37-1420">Параметры</span><span class="sxs-lookup"><span data-stu-id="e5c37-1420">Parameters</span></span>

|<span data-ttu-id="e5c37-1421">Имя</span><span class="sxs-lookup"><span data-stu-id="e5c37-1421">Name</span></span>|<span data-ttu-id="e5c37-1422">Тип</span><span class="sxs-lookup"><span data-stu-id="e5c37-1422">Type</span></span>|<span data-ttu-id="e5c37-1423">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="e5c37-1423">Attributes</span></span>|<span data-ttu-id="e5c37-1424">Описание</span><span class="sxs-lookup"><span data-stu-id="e5c37-1424">Description</span></span>|
|---|---|---|---|
|`callback`|<span data-ttu-id="e5c37-1425">function</span><span class="sxs-lookup"><span data-stu-id="e5c37-1425">function</span></span>||<span data-ttu-id="e5c37-1426">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="e5c37-1426">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="e5c37-1427">Настраиваемые свойства предоставляются в виде объекта [`CustomProperties`](/javascript/api/outlook/office.customproperties) в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="e5c37-1427">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="e5c37-1428">Этот объект позволяет получить, задать и удалить настраиваемые свойства для элемента, а также сохранить изменения, внесенные в набор настраиваемых свойств, на сервере.</span><span class="sxs-lookup"><span data-stu-id="e5c37-1428">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`|<span data-ttu-id="e5c37-1429">Объект</span><span class="sxs-lookup"><span data-stu-id="e5c37-1429">Object</span></span>|<span data-ttu-id="e5c37-1430">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="e5c37-1430">&lt;optional&gt;</span></span>|<span data-ttu-id="e5c37-1431">Разработчики могут указать любой объект, к которому необходимо получить доступ, в функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="e5c37-1431">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="e5c37-1432">Доступ к этому объекту можно получить с помощью свойства `asyncResult.asyncContext` в функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="e5c37-1432">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e5c37-1433">Требования</span><span class="sxs-lookup"><span data-stu-id="e5c37-1433">Requirements</span></span>

|<span data-ttu-id="e5c37-1434">Требование</span><span class="sxs-lookup"><span data-stu-id="e5c37-1434">Requirement</span></span>|<span data-ttu-id="e5c37-1435">Значение</span><span class="sxs-lookup"><span data-stu-id="e5c37-1435">Value</span></span>|
|---|---|
|[<span data-ttu-id="e5c37-1436">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="e5c37-1436">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e5c37-1437">1.0</span><span class="sxs-lookup"><span data-stu-id="e5c37-1437">1.0</span></span>|
|[<span data-ttu-id="e5c37-1438">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e5c37-1438">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e5c37-1439">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e5c37-1439">ReadItem</span></span>|
|[<span data-ttu-id="e5c37-1440">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e5c37-1440">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="e5c37-1441">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="e5c37-1441">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e5c37-1442">Пример</span><span class="sxs-lookup"><span data-stu-id="e5c37-1442">Example</span></span>

<span data-ttu-id="e5c37-p184">Приведенный ниже пример кода показывает, как асинхронно загружать настраиваемые свойства, характерные для текущего элемента, с помощью метода `loadCustomPropertiesAsync`. Этот пример также показывает, как сохранять эти свойства на сервере с помощью метода `CustomProperties.saveAsync`. После загрузки настраиваемых свойств в этом примере кода метод `CustomProperties.get` используется для считывания настраиваемого свойства `myProp`, метод `CustomProperties.set` — для записи настраиваемого свойства `otherProp`, а метод `saveAsync` — для сохранения настраиваемых свойств.</span><span class="sxs-lookup"><span data-stu-id="e5c37-p184">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

#### <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="e5c37-1446">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="e5c37-1446">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="e5c37-1447">Удаляет вложение из сообщения или встречи.</span><span class="sxs-lookup"><span data-stu-id="e5c37-1447">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="e5c37-1448">Метод `removeAttachmentAsync` удаляет из элемента вложение с указанным идентификатором.</span><span class="sxs-lookup"><span data-stu-id="e5c37-1448">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="e5c37-1449">Идентификатор вложения рекомендуется использовать для удаления вложения, только если оно добавлено тем же почтовым приложением в ходе текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="e5c37-1449">As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session.</span></span> <span data-ttu-id="e5c37-1450">В Outlook в Интернете и мобильных устройствах идентификатор вложения действителен только в рамках одного сеанса.</span><span class="sxs-lookup"><span data-stu-id="e5c37-1450">In Outlook on the web and mobile devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="e5c37-1451">Сеанс переходит к моменту, когда пользователь закрывает приложение, или если пользователь начинает создание встроенной формы, затем извлекает форму, чтобы продолжить работу в отдельном окне.</span><span class="sxs-lookup"><span data-stu-id="e5c37-1451">A session is over when the user closes the app, or if the user starts composing an inline form then subsequently pops out the form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e5c37-1452">Параметры</span><span class="sxs-lookup"><span data-stu-id="e5c37-1452">Parameters</span></span>

|<span data-ttu-id="e5c37-1453">Имя</span><span class="sxs-lookup"><span data-stu-id="e5c37-1453">Name</span></span>|<span data-ttu-id="e5c37-1454">Тип</span><span class="sxs-lookup"><span data-stu-id="e5c37-1454">Type</span></span>|<span data-ttu-id="e5c37-1455">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="e5c37-1455">Attributes</span></span>|<span data-ttu-id="e5c37-1456">Описание</span><span class="sxs-lookup"><span data-stu-id="e5c37-1456">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="e5c37-1457">String</span><span class="sxs-lookup"><span data-stu-id="e5c37-1457">String</span></span>||<span data-ttu-id="e5c37-1458">Идентификатор удаляемого вложения.</span><span class="sxs-lookup"><span data-stu-id="e5c37-1458">The identifier of the attachment to remove.</span></span>|
|`options`|<span data-ttu-id="e5c37-1459">Объект</span><span class="sxs-lookup"><span data-stu-id="e5c37-1459">Object</span></span>|<span data-ttu-id="e5c37-1460">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="e5c37-1460">&lt;optional&gt;</span></span>|<span data-ttu-id="e5c37-1461">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="e5c37-1461">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="e5c37-1462">Объект</span><span class="sxs-lookup"><span data-stu-id="e5c37-1462">Object</span></span>|<span data-ttu-id="e5c37-1463">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="e5c37-1463">&lt;optional&gt;</span></span>|<span data-ttu-id="e5c37-1464">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="e5c37-1464">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="e5c37-1465">функция</span><span class="sxs-lookup"><span data-stu-id="e5c37-1465">function</span></span>|<span data-ttu-id="e5c37-1466">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="e5c37-1466">&lt;optional&gt;</span></span>|<span data-ttu-id="e5c37-1467">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="e5c37-1467">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="e5c37-1468">Если удалить вложение не удается, свойство `asyncResult.error` содержит код ошибки с указанием ее причины.</span><span class="sxs-lookup"><span data-stu-id="e5c37-1468">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="e5c37-1469">Ошибки</span><span class="sxs-lookup"><span data-stu-id="e5c37-1469">Errors</span></span>

|<span data-ttu-id="e5c37-1470">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="e5c37-1470">Error code</span></span>|<span data-ttu-id="e5c37-1471">Описание</span><span class="sxs-lookup"><span data-stu-id="e5c37-1471">Description</span></span>|
|------------|-------------|
|`InvalidAttachmentId`|<span data-ttu-id="e5c37-1472">Идентификатор вложения не существует.</span><span class="sxs-lookup"><span data-stu-id="e5c37-1472">The attachment identifier does not exist.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e5c37-1473">Требования</span><span class="sxs-lookup"><span data-stu-id="e5c37-1473">Requirements</span></span>

|<span data-ttu-id="e5c37-1474">Требование</span><span class="sxs-lookup"><span data-stu-id="e5c37-1474">Requirement</span></span>|<span data-ttu-id="e5c37-1475">Значение</span><span class="sxs-lookup"><span data-stu-id="e5c37-1475">Value</span></span>|
|---|---|
|[<span data-ttu-id="e5c37-1476">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="e5c37-1476">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e5c37-1477">1.1</span><span class="sxs-lookup"><span data-stu-id="e5c37-1477">1.1</span></span>|
|[<span data-ttu-id="e5c37-1478">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e5c37-1478">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e5c37-1479">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="e5c37-1479">ReadWriteItem</span></span>|
|[<span data-ttu-id="e5c37-1480">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e5c37-1480">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="e5c37-1481">Создание</span><span class="sxs-lookup"><span data-stu-id="e5c37-1481">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="e5c37-1482">Пример</span><span class="sxs-lookup"><span data-stu-id="e5c37-1482">Example</span></span>

<span data-ttu-id="e5c37-1483">Указанный ниже код удаляет вложение с идентификатором "0".</span><span class="sxs-lookup"><span data-stu-id="e5c37-1483">The following code removes an attachment with an identifier of '0'.</span></span>

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

#### <a name="removehandlerasynceventtype-options-callback"></a><span data-ttu-id="e5c37-1484">removeHandlerAsync(eventType, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="e5c37-1484">removeHandlerAsync(eventType, [options], [callback])</span></span>

<span data-ttu-id="e5c37-1485">Удаляет обработчиков для поддерживаемого типа события.</span><span class="sxs-lookup"><span data-stu-id="e5c37-1485">Removes the event handlers for a supported event type.</span></span>

<span data-ttu-id="e5c37-1486">В настоящее время поддерживаются типы `Office.EventType.AttachmentsChanged`событий `Office.EventType.AppointmentTimeChanged`, `Office.EventType.EnhancedLocationsChanged` `Office.EventType.RecipientsChanged`,, и `Office.EventType.RecurrenceChanged`.</span><span class="sxs-lookup"><span data-stu-id="e5c37-1486">Currently the supported event types are `Office.EventType.AttachmentsChanged`, `Office.EventType.AppointmentTimeChanged`, `Office.EventType.EnhancedLocationsChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e5c37-1487">Параметры</span><span class="sxs-lookup"><span data-stu-id="e5c37-1487">Parameters</span></span>

| <span data-ttu-id="e5c37-1488">Имя</span><span class="sxs-lookup"><span data-stu-id="e5c37-1488">Name</span></span> | <span data-ttu-id="e5c37-1489">Тип</span><span class="sxs-lookup"><span data-stu-id="e5c37-1489">Type</span></span> | <span data-ttu-id="e5c37-1490">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="e5c37-1490">Attributes</span></span> | <span data-ttu-id="e5c37-1491">Описание</span><span class="sxs-lookup"><span data-stu-id="e5c37-1491">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="e5c37-1492">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="e5c37-1492">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="e5c37-1493">Событие, которое должно отменить обработчик.</span><span class="sxs-lookup"><span data-stu-id="e5c37-1493">The event that should revoke the handler.</span></span> |
| `options` | <span data-ttu-id="e5c37-1494">Объект</span><span class="sxs-lookup"><span data-stu-id="e5c37-1494">Object</span></span> | <span data-ttu-id="e5c37-1495">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="e5c37-1495">&lt;optional&gt;</span></span> | <span data-ttu-id="e5c37-1496">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="e5c37-1496">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="e5c37-1497">Объект</span><span class="sxs-lookup"><span data-stu-id="e5c37-1497">Object</span></span> | <span data-ttu-id="e5c37-1498">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="e5c37-1498">&lt;optional&gt;</span></span> | <span data-ttu-id="e5c37-1499">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="e5c37-1499">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="e5c37-1500">функция</span><span class="sxs-lookup"><span data-stu-id="e5c37-1500">function</span></span>| <span data-ttu-id="e5c37-1501">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="e5c37-1501">&lt;optional&gt;</span></span>|<span data-ttu-id="e5c37-1502">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="e5c37-1502">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e5c37-1503">Требования</span><span class="sxs-lookup"><span data-stu-id="e5c37-1503">Requirements</span></span>

|<span data-ttu-id="e5c37-1504">Требование</span><span class="sxs-lookup"><span data-stu-id="e5c37-1504">Requirement</span></span>| <span data-ttu-id="e5c37-1505">Значение</span><span class="sxs-lookup"><span data-stu-id="e5c37-1505">Value</span></span>|
|---|---|
|[<span data-ttu-id="e5c37-1506">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="e5c37-1506">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e5c37-1507">1.7</span><span class="sxs-lookup"><span data-stu-id="e5c37-1507">1.7</span></span> |
|[<span data-ttu-id="e5c37-1508">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e5c37-1508">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e5c37-1509">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e5c37-1509">ReadItem</span></span> |
|[<span data-ttu-id="e5c37-1510">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e5c37-1510">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e5c37-1511">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="e5c37-1511">Compose or Read</span></span> |

<br>

---
---

#### <a name="saveasyncoptions-callback"></a><span data-ttu-id="e5c37-1512">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="e5c37-1512">saveAsync([options], callback)</span></span>

<span data-ttu-id="e5c37-1513">Асинхронно сохраняет элемент.</span><span class="sxs-lookup"><span data-stu-id="e5c37-1513">Asynchronously saves an item.</span></span>

<span data-ttu-id="e5c37-1514">При вызове этот метод сохраняет текущее сообщение в виде черновика и возвращает идентификатор элемента с помощью метода обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="e5c37-1514">When invoked, this method saves the current message as a draft and returns the item id via the callback method.</span></span> <span data-ttu-id="e5c37-1515">В Outlook в Интернете или Outlook в интерактивном режиме элемент сохраняется на сервере.</span><span class="sxs-lookup"><span data-stu-id="e5c37-1515">In Outlook on the web or Outlook in online mode, the item is saved to the server.</span></span> <span data-ttu-id="e5c37-1516">В Outlook в режиме кэширования этот элемент сохраняется в локальном кэше.</span><span class="sxs-lookup"><span data-stu-id="e5c37-1516">In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="e5c37-1517">Если в режиме создания надстройка вызывает для элемента метод `saveAsync`, чтобы получить параметр `itemId` для использования с EWS или REST API, необходимо помнить, что синхронизация элемента с сервером может занять много времени, если Outlook работает в режиме кэширования данных.</span><span class="sxs-lookup"><span data-stu-id="e5c37-1517">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="e5c37-1518">До окончания синхронизации элемента применение параметра `itemId` будет приводить к ошибке.</span><span class="sxs-lookup"><span data-stu-id="e5c37-1518">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="e5c37-p188">Если метод `saveAsync` вызывается для встречи в режиме создания, она сохраняется как обычная встреча в календаре пользователя, а не как черновик. При сохранении новой встречи приглашения не отправляются. При сохранении существующей встречи уведомления отправляются добавленным или удаленным участникам.</span><span class="sxs-lookup"><span data-stu-id="e5c37-p188">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="e5c37-1522">Следующие клиенты отличаются другим поведением `saveAsync` в отношении встреч в режиме создания:</span><span class="sxs-lookup"><span data-stu-id="e5c37-1522">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="e5c37-1523">Outlook в Mac не поддерживает сохранение собраний.</span><span class="sxs-lookup"><span data-stu-id="e5c37-1523">Outlook on Mac does not support saving a meeting.</span></span> <span data-ttu-id="e5c37-1524">`saveAsync` Метод завершается с ошибкой при вызове из собрания в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="e5c37-1524">The `saveAsync` method fails when called from a meeting in compose mode.</span></span> <span data-ttu-id="e5c37-1525">Просмотреть [не удается сохранить собрание в виде черновика в Outlook для Mac с помощью API Office JS](https://support.microsoft.com/help/4505745) для обхода.</span><span class="sxs-lookup"><span data-stu-id="e5c37-1525">See [Cannot save a meeting as a draft in Outlook for Mac by using Office JS API](https://support.microsoft.com/help/4505745) for a workaround.</span></span>
> - <span data-ttu-id="e5c37-1526">Outlook в Интернете всегда отправляет приглашение или обновление при вызове `saveAsync` для встречи в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="e5c37-1526">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e5c37-1527">Параметры</span><span class="sxs-lookup"><span data-stu-id="e5c37-1527">Parameters</span></span>

|<span data-ttu-id="e5c37-1528">Имя</span><span class="sxs-lookup"><span data-stu-id="e5c37-1528">Name</span></span>|<span data-ttu-id="e5c37-1529">Тип</span><span class="sxs-lookup"><span data-stu-id="e5c37-1529">Type</span></span>|<span data-ttu-id="e5c37-1530">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="e5c37-1530">Attributes</span></span>|<span data-ttu-id="e5c37-1531">Описание</span><span class="sxs-lookup"><span data-stu-id="e5c37-1531">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="e5c37-1532">Объект</span><span class="sxs-lookup"><span data-stu-id="e5c37-1532">Object</span></span>|<span data-ttu-id="e5c37-1533">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="e5c37-1533">&lt;optional&gt;</span></span>|<span data-ttu-id="e5c37-1534">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="e5c37-1534">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="e5c37-1535">Объект</span><span class="sxs-lookup"><span data-stu-id="e5c37-1535">Object</span></span>|<span data-ttu-id="e5c37-1536">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="e5c37-1536">&lt;optional&gt;</span></span>|<span data-ttu-id="e5c37-1537">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="e5c37-1537">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="e5c37-1538">функция</span><span class="sxs-lookup"><span data-stu-id="e5c37-1538">function</span></span>||<span data-ttu-id="e5c37-1539">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="e5c37-1539">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="e5c37-1540">После успешного выполнения идентификатор элемента будет предоставлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="e5c37-1540">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e5c37-1541">Требования</span><span class="sxs-lookup"><span data-stu-id="e5c37-1541">Requirements</span></span>

|<span data-ttu-id="e5c37-1542">Требование</span><span class="sxs-lookup"><span data-stu-id="e5c37-1542">Requirement</span></span>|<span data-ttu-id="e5c37-1543">Значение</span><span class="sxs-lookup"><span data-stu-id="e5c37-1543">Value</span></span>|
|---|---|
|[<span data-ttu-id="e5c37-1544">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="e5c37-1544">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e5c37-1545">1.3</span><span class="sxs-lookup"><span data-stu-id="e5c37-1545">1.3</span></span>|
|[<span data-ttu-id="e5c37-1546">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e5c37-1546">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e5c37-1547">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="e5c37-1547">ReadWriteItem</span></span>|
|[<span data-ttu-id="e5c37-1548">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e5c37-1548">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="e5c37-1549">Создание</span><span class="sxs-lookup"><span data-stu-id="e5c37-1549">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="e5c37-1550">Примеры</span><span class="sxs-lookup"><span data-stu-id="e5c37-1550">Examples</span></span>

```js
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result.
  });
```

<span data-ttu-id="e5c37-p190">Ниже приведен пример параметра `result`, переданного функции обратного вызова. Свойство `value` содержит идентификатор элемента.</span><span class="sxs-lookup"><span data-stu-id="e5c37-p190">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

<br>

---
---

#### <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="e5c37-1553">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="e5c37-1553">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="e5c37-1554">Асинхронно вставляет данные в текст или тему сообщения.</span><span class="sxs-lookup"><span data-stu-id="e5c37-1554">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="e5c37-p191">Метод `setSelectedDataAsync` вставляет указанную строку в местоположение курсора в теме или тексте элемента либо, если текст выделен в редакторе, он заменяет выделенный текст. Если курсор находится вне текста или темы элемента, возвращается ошибка. После вставки курсор помещается в конец вставленного содержимого.</span><span class="sxs-lookup"><span data-stu-id="e5c37-p191">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e5c37-1558">Параметры</span><span class="sxs-lookup"><span data-stu-id="e5c37-1558">Parameters</span></span>

|<span data-ttu-id="e5c37-1559">Имя</span><span class="sxs-lookup"><span data-stu-id="e5c37-1559">Name</span></span>|<span data-ttu-id="e5c37-1560">Тип</span><span class="sxs-lookup"><span data-stu-id="e5c37-1560">Type</span></span>|<span data-ttu-id="e5c37-1561">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="e5c37-1561">Attributes</span></span>|<span data-ttu-id="e5c37-1562">Описание</span><span class="sxs-lookup"><span data-stu-id="e5c37-1562">Description</span></span>|
|---|---|---|---|
|`data`|<span data-ttu-id="e5c37-1563">String</span><span class="sxs-lookup"><span data-stu-id="e5c37-1563">String</span></span>||<span data-ttu-id="e5c37-p192">Вставляемые данные. Объем данных не должен превышать 1 000 000 символов. Если передано больше 1 000 000 символов, возвращается исключение `ArgumentOutOfRange`.</span><span class="sxs-lookup"><span data-stu-id="e5c37-p192">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`|<span data-ttu-id="e5c37-1567">Object</span><span class="sxs-lookup"><span data-stu-id="e5c37-1567">Object</span></span>|<span data-ttu-id="e5c37-1568">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="e5c37-1568">&lt;optional&gt;</span></span>|<span data-ttu-id="e5c37-1569">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="e5c37-1569">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="e5c37-1570">Объект</span><span class="sxs-lookup"><span data-stu-id="e5c37-1570">Object</span></span>|<span data-ttu-id="e5c37-1571">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="e5c37-1571">&lt;optional&gt;</span></span>|<span data-ttu-id="e5c37-1572">В методе обратного вызова разработчики могут указать любой объект, к которому необходимо получить доступ.</span><span class="sxs-lookup"><span data-stu-id="e5c37-1572">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="e5c37-1573">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="e5c37-1573">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="e5c37-1574">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="e5c37-1574">&lt;optional&gt;</span></span>|<span data-ttu-id="e5c37-1575">Если `text`текущий стиль применяется в Outlook для веб-клиентов и клиентов для настольных ПК.</span><span class="sxs-lookup"><span data-stu-id="e5c37-1575">If `text`, the current style is applied in Outlook on the web and desktop clients.</span></span> <span data-ttu-id="e5c37-1576">Если поле представляет собой редактор HTML, вставляются только текстовые данные, даже если они имеют формат HTML.</span><span class="sxs-lookup"><span data-stu-id="e5c37-1576">If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="e5c37-1577">Если `html` и поле поддерживает HTML (тема не используется), текущий стиль применяется в Outlook в Интернете, а в настольных клиентах Outlook применяется стиль по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="e5c37-1577">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook on the web and the default style is applied in Outlook desktop clients.</span></span> <span data-ttu-id="e5c37-1578">Если поле является текстовым, возвращается ошибка `InvalidDataFormat`.</span><span class="sxs-lookup"><span data-stu-id="e5c37-1578">If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="e5c37-1579">Если свойство `coercionType` не задано, результат зависит от поля: если поле имеет формат HTML, используется текст в формате HTML, а если поле текстовое, применяется обычный текст.</span><span class="sxs-lookup"><span data-stu-id="e5c37-1579">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`|<span data-ttu-id="e5c37-1580">функция</span><span class="sxs-lookup"><span data-stu-id="e5c37-1580">function</span></span>||<span data-ttu-id="e5c37-1581">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="e5c37-1581">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e5c37-1582">Требования</span><span class="sxs-lookup"><span data-stu-id="e5c37-1582">Requirements</span></span>

|<span data-ttu-id="e5c37-1583">Требование</span><span class="sxs-lookup"><span data-stu-id="e5c37-1583">Requirement</span></span>|<span data-ttu-id="e5c37-1584">Значение</span><span class="sxs-lookup"><span data-stu-id="e5c37-1584">Value</span></span>|
|---|---|
|[<span data-ttu-id="e5c37-1585">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="e5c37-1585">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e5c37-1586">1.2</span><span class="sxs-lookup"><span data-stu-id="e5c37-1586">1.2</span></span>|
|[<span data-ttu-id="e5c37-1587">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e5c37-1587">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e5c37-1588">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="e5c37-1588">ReadWriteItem</span></span>|
|[<span data-ttu-id="e5c37-1589">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e5c37-1589">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="e5c37-1590">Создание</span><span class="sxs-lookup"><span data-stu-id="e5c37-1590">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="e5c37-1591">Пример</span><span class="sxs-lookup"><span data-stu-id="e5c37-1591">Example</span></span>

```js
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
