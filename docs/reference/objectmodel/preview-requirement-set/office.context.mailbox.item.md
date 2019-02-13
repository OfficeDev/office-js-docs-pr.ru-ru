---
title: Office.Context.Mailbox.Item - наборы требований предварительного просмотра
description: ''
ms.date: 01/30/2019
localization_priority: Normal
ms.openlocfilehash: a660f8bafdd2587f97d704e42c47abbe6c7d533d
ms.sourcegitcommit: a59f4e322238efa187f388a75b7709462c71e668
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/13/2019
ms.locfileid: "29982050"
---
# <a name="item"></a><span data-ttu-id="b789e-102">item</span><span class="sxs-lookup"><span data-stu-id="b789e-102">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="b789e-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="b789e-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="b789e-p101">Пространство имен `item` используется для доступа к выбранному в данный момент сообщению, приглашению на собрание или описанию встречи. Вы можете определить тип пространства имен `item` с помощью свойства [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="b789e-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="b789e-106">Requirements</span><span class="sxs-lookup"><span data-stu-id="b789e-106">Requirements</span></span>

|<span data-ttu-id="b789e-107">Требование</span><span class="sxs-lookup"><span data-stu-id="b789e-107">Requirement</span></span>|<span data-ttu-id="b789e-108">Значение</span><span class="sxs-lookup"><span data-stu-id="b789e-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="b789e-109">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="b789e-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="b789e-110">1.0</span><span class="sxs-lookup"><span data-stu-id="b789e-110">1.0</span></span>|
|[<span data-ttu-id="b789e-111">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b789e-111">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="b789e-112">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="b789e-112">Restricted</span></span>|
|[<span data-ttu-id="b789e-113">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b789e-113">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="b789e-114">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="b789e-114">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="b789e-115">Элементы и методы</span><span class="sxs-lookup"><span data-stu-id="b789e-115">Members and methods</span></span>

| <span data-ttu-id="b789e-116">Элемент</span><span class="sxs-lookup"><span data-stu-id="b789e-116">Member</span></span> | <span data-ttu-id="b789e-117">Тип</span><span class="sxs-lookup"><span data-stu-id="b789e-117">Type</span></span> |
|--------|------|
| [<span data-ttu-id="b789e-118">attachments</span><span class="sxs-lookup"><span data-stu-id="b789e-118">attachments</span></span>](#attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails) | <span data-ttu-id="b789e-119">Элемент</span><span class="sxs-lookup"><span data-stu-id="b789e-119">Member</span></span> |
| [<span data-ttu-id="b789e-120">bcc</span><span class="sxs-lookup"><span data-stu-id="b789e-120">bcc</span></span>](#bcc-recipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="b789e-121">Элемент</span><span class="sxs-lookup"><span data-stu-id="b789e-121">Member</span></span> |
| [<span data-ttu-id="b789e-122">body</span><span class="sxs-lookup"><span data-stu-id="b789e-122">body</span></span>](#body-bodyjavascriptapioutlookofficebody) | <span data-ttu-id="b789e-123">Элемент</span><span class="sxs-lookup"><span data-stu-id="b789e-123">Member</span></span> |
| [<span data-ttu-id="b789e-124">cc</span><span class="sxs-lookup"><span data-stu-id="b789e-124">cc</span></span>](#cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="b789e-125">Элемент</span><span class="sxs-lookup"><span data-stu-id="b789e-125">Member</span></span> |
| [<span data-ttu-id="b789e-126">conversationId</span><span class="sxs-lookup"><span data-stu-id="b789e-126">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="b789e-127">Элемент</span><span class="sxs-lookup"><span data-stu-id="b789e-127">Member</span></span> |
| [<span data-ttu-id="b789e-128">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="b789e-128">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="b789e-129">Элемент</span><span class="sxs-lookup"><span data-stu-id="b789e-129">Member</span></span> |
| [<span data-ttu-id="b789e-130">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="b789e-130">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="b789e-131">Элемент</span><span class="sxs-lookup"><span data-stu-id="b789e-131">Member</span></span> |
| [<span data-ttu-id="b789e-132">end</span><span class="sxs-lookup"><span data-stu-id="b789e-132">end</span></span>](#end-datetimejavascriptapioutlookofficetime) | <span data-ttu-id="b789e-133">Элемент</span><span class="sxs-lookup"><span data-stu-id="b789e-133">Member</span></span> |
| [<span data-ttu-id="b789e-134">enhancedLocation</span><span class="sxs-lookup"><span data-stu-id="b789e-134">enhancedLocation</span></span>](#enhancedlocation-enhancedlocationjavascriptapioutlookofficeenhancedlocation) | <span data-ttu-id="b789e-135">Элемент</span><span class="sxs-lookup"><span data-stu-id="b789e-135">Member</span></span> |
| [<span data-ttu-id="b789e-136">from</span><span class="sxs-lookup"><span data-stu-id="b789e-136">from</span></span>](#from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom) | <span data-ttu-id="b789e-137">Элемент</span><span class="sxs-lookup"><span data-stu-id="b789e-137">Member</span></span> |
| [<span data-ttu-id="b789e-138">internetHeaders</span><span class="sxs-lookup"><span data-stu-id="b789e-138">internetHeaders</span></span>](#internetheaders-internetheadersjavascriptapioutlookofficeinternetheaders) | <span data-ttu-id="b789e-139">Элемент</span><span class="sxs-lookup"><span data-stu-id="b789e-139">Member</span></span> |
| [<span data-ttu-id="b789e-140">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="b789e-140">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="b789e-141">Элемент</span><span class="sxs-lookup"><span data-stu-id="b789e-141">Member</span></span> |
| [<span data-ttu-id="b789e-142">itemClass</span><span class="sxs-lookup"><span data-stu-id="b789e-142">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="b789e-143">Элемент</span><span class="sxs-lookup"><span data-stu-id="b789e-143">Member</span></span> |
| [<span data-ttu-id="b789e-144">itemId</span><span class="sxs-lookup"><span data-stu-id="b789e-144">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="b789e-145">Элемент</span><span class="sxs-lookup"><span data-stu-id="b789e-145">Member</span></span> |
| [<span data-ttu-id="b789e-146">itemType</span><span class="sxs-lookup"><span data-stu-id="b789e-146">itemType</span></span>](#itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype) | <span data-ttu-id="b789e-147">Элемент</span><span class="sxs-lookup"><span data-stu-id="b789e-147">Member</span></span> |
| [<span data-ttu-id="b789e-148">location</span><span class="sxs-lookup"><span data-stu-id="b789e-148">location</span></span>](#location-stringlocationjavascriptapioutlookofficelocation) | <span data-ttu-id="b789e-149">Элемент</span><span class="sxs-lookup"><span data-stu-id="b789e-149">Member</span></span> |
| [<span data-ttu-id="b789e-150">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="b789e-150">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="b789e-151">Элемент</span><span class="sxs-lookup"><span data-stu-id="b789e-151">Member</span></span> |
| [<span data-ttu-id="b789e-152">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="b789e-152">notificationMessages</span></span>](#notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessages) | <span data-ttu-id="b789e-153">Элемент</span><span class="sxs-lookup"><span data-stu-id="b789e-153">Member</span></span> |
| [<span data-ttu-id="b789e-154">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="b789e-154">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="b789e-155">Member</span><span class="sxs-lookup"><span data-stu-id="b789e-155">Member</span></span> |
| [<span data-ttu-id="b789e-156">organizer</span><span class="sxs-lookup"><span data-stu-id="b789e-156">organizer</span></span>](#organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsorganizerjavascriptapioutlookofficeorganizer) | <span data-ttu-id="b789e-157">Элемент</span><span class="sxs-lookup"><span data-stu-id="b789e-157">Member</span></span> |
| [<span data-ttu-id="b789e-158">recurrence</span><span class="sxs-lookup"><span data-stu-id="b789e-158">recurrence</span></span>](#nullable-recurrence-recurrencejavascriptapioutlookofficerecurrence) | <span data-ttu-id="b789e-159">Элемент</span><span class="sxs-lookup"><span data-stu-id="b789e-159">Member</span></span> |
| [<span data-ttu-id="b789e-160">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="b789e-160">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="b789e-161">Элемент</span><span class="sxs-lookup"><span data-stu-id="b789e-161">Member</span></span> |
| [<span data-ttu-id="b789e-162">sender</span><span class="sxs-lookup"><span data-stu-id="b789e-162">sender</span></span>](#sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails) | <span data-ttu-id="b789e-163">Элемент</span><span class="sxs-lookup"><span data-stu-id="b789e-163">Member</span></span> |
| [<span data-ttu-id="b789e-164">seriesId</span><span class="sxs-lookup"><span data-stu-id="b789e-164">seriesId</span></span>](#nullable-seriesid-string) | <span data-ttu-id="b789e-165">Элемент</span><span class="sxs-lookup"><span data-stu-id="b789e-165">Member</span></span> |
| [<span data-ttu-id="b789e-166">start</span><span class="sxs-lookup"><span data-stu-id="b789e-166">start</span></span>](#start-datetimejavascriptapioutlookofficetime) | <span data-ttu-id="b789e-167">Member</span><span class="sxs-lookup"><span data-stu-id="b789e-167">Member</span></span> |
| [<span data-ttu-id="b789e-168">subject</span><span class="sxs-lookup"><span data-stu-id="b789e-168">subject</span></span>](#subject-stringsubjectjavascriptapioutlookofficesubject) | <span data-ttu-id="b789e-169">Member</span><span class="sxs-lookup"><span data-stu-id="b789e-169">Member</span></span> |
| [<span data-ttu-id="b789e-170">to</span><span class="sxs-lookup"><span data-stu-id="b789e-170">to</span></span>](#to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="b789e-171">Элемент</span><span class="sxs-lookup"><span data-stu-id="b789e-171">Member</span></span> |
| [<span data-ttu-id="b789e-172">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="b789e-172">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="b789e-173">Метод</span><span class="sxs-lookup"><span data-stu-id="b789e-173">Method</span></span> |
| [<span data-ttu-id="b789e-174">addFileAttachmentFromBase64Async</span><span class="sxs-lookup"><span data-stu-id="b789e-174">addFileAttachmentFromBase64Async</span></span>](#addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback) | <span data-ttu-id="b789e-175">Метод</span><span class="sxs-lookup"><span data-stu-id="b789e-175">Method</span></span> |
| [<span data-ttu-id="b789e-176">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="b789e-176">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="b789e-177">Метод</span><span class="sxs-lookup"><span data-stu-id="b789e-177">Method</span></span> |
| [<span data-ttu-id="b789e-178">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="b789e-178">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="b789e-179">Метод</span><span class="sxs-lookup"><span data-stu-id="b789e-179">Method</span></span> |
| [<span data-ttu-id="b789e-180">close</span><span class="sxs-lookup"><span data-stu-id="b789e-180">close</span></span>](#close) | <span data-ttu-id="b789e-181">Метод</span><span class="sxs-lookup"><span data-stu-id="b789e-181">Method</span></span> |
| [<span data-ttu-id="b789e-182">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="b789e-182">displayReplyAllForm</span></span>](#displayreplyallformformdata) | <span data-ttu-id="b789e-183">Метод</span><span class="sxs-lookup"><span data-stu-id="b789e-183">Method</span></span> |
| [<span data-ttu-id="b789e-184">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="b789e-184">displayReplyForm</span></span>](#displayreplyformformdata) | <span data-ttu-id="b789e-185">Метод</span><span class="sxs-lookup"><span data-stu-id="b789e-185">Method</span></span> |
| [<span data-ttu-id="b789e-186">getAttachmentContentAsync</span><span class="sxs-lookup"><span data-stu-id="b789e-186">getAttachmentContentAsync</span></span>](#getattachmentcontentasyncattachmentid-options-callback--attachmentcontentjavascriptapioutlookofficeattachmentcontent) | <span data-ttu-id="b789e-187">Метод</span><span class="sxs-lookup"><span data-stu-id="b789e-187">Method</span></span> |
| [<span data-ttu-id="b789e-188">getAttachmentsAsync</span><span class="sxs-lookup"><span data-stu-id="b789e-188">getAttachmentsAsync</span></span>](#getattachmentsasyncoptions-callback--arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails) | <span data-ttu-id="b789e-189">Метод</span><span class="sxs-lookup"><span data-stu-id="b789e-189">Method</span></span> |
| [<span data-ttu-id="b789e-190">getEntities</span><span class="sxs-lookup"><span data-stu-id="b789e-190">getEntities</span></span>](#getentities--entitiesjavascriptapioutlookofficeentities) | <span data-ttu-id="b789e-191">Метод</span><span class="sxs-lookup"><span data-stu-id="b789e-191">Method</span></span> |
| [<span data-ttu-id="b789e-192">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="b789e-192">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion) | <span data-ttu-id="b789e-193">Метод</span><span class="sxs-lookup"><span data-stu-id="b789e-193">Method</span></span> |
| [<span data-ttu-id="b789e-194">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="b789e-194">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion) | <span data-ttu-id="b789e-195">Метод</span><span class="sxs-lookup"><span data-stu-id="b789e-195">Method</span></span> |
| [<span data-ttu-id="b789e-196">getInitializationContextAsync</span><span class="sxs-lookup"><span data-stu-id="b789e-196">getInitializationContextAsync</span></span>](#getinitializationcontextasyncoptions-callback) | <span data-ttu-id="b789e-197">Метод</span><span class="sxs-lookup"><span data-stu-id="b789e-197">Method</span></span> |
| [<span data-ttu-id="b789e-198">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="b789e-198">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="b789e-199">Метод</span><span class="sxs-lookup"><span data-stu-id="b789e-199">Method</span></span> |
| [<span data-ttu-id="b789e-200">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="b789e-200">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="b789e-201">Метод</span><span class="sxs-lookup"><span data-stu-id="b789e-201">Method</span></span> |
| [<span data-ttu-id="b789e-202">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="b789e-202">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="b789e-203">Метод</span><span class="sxs-lookup"><span data-stu-id="b789e-203">Method</span></span> |
| [<span data-ttu-id="b789e-204">getSelectedEntities</span><span class="sxs-lookup"><span data-stu-id="b789e-204">getSelectedEntities</span></span>](#getselectedentities--entitiesjavascriptapioutlookofficeentities) | <span data-ttu-id="b789e-205">Метод</span><span class="sxs-lookup"><span data-stu-id="b789e-205">Method</span></span> |
| [<span data-ttu-id="b789e-206">getSelectedRegExMatches</span><span class="sxs-lookup"><span data-stu-id="b789e-206">getSelectedRegExMatches</span></span>](#getselectedregexmatches--object) | <span data-ttu-id="b789e-207">Метод</span><span class="sxs-lookup"><span data-stu-id="b789e-207">Method</span></span> |
| [<span data-ttu-id="b789e-208">getSharedPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="b789e-208">getSharedPropertiesAsync</span></span>](#getsharedpropertiesasyncoptions-callback) | <span data-ttu-id="b789e-209">Метод</span><span class="sxs-lookup"><span data-stu-id="b789e-209">Method</span></span> |
| [<span data-ttu-id="b789e-210">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="b789e-210">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="b789e-211">Метод</span><span class="sxs-lookup"><span data-stu-id="b789e-211">Method</span></span> |
| [<span data-ttu-id="b789e-212">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="b789e-212">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="b789e-213">Метод</span><span class="sxs-lookup"><span data-stu-id="b789e-213">Method</span></span> |
| [<span data-ttu-id="b789e-214">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="b789e-214">removeHandlerAsync</span></span>](#removehandlerasynceventtype-options-callback) | <span data-ttu-id="b789e-215">Метод</span><span class="sxs-lookup"><span data-stu-id="b789e-215">Method</span></span> |
| [<span data-ttu-id="b789e-216">saveAsync</span><span class="sxs-lookup"><span data-stu-id="b789e-216">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="b789e-217">Метод</span><span class="sxs-lookup"><span data-stu-id="b789e-217">Method</span></span> |
| [<span data-ttu-id="b789e-218">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="b789e-218">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="b789e-219">Метод</span><span class="sxs-lookup"><span data-stu-id="b789e-219">Method</span></span> |

### <a name="example"></a><span data-ttu-id="b789e-220">Пример</span><span class="sxs-lookup"><span data-stu-id="b789e-220">Example</span></span>

<span data-ttu-id="b789e-221">В примере кода JavaScript, приведенном ниже, показано, как получить доступ к свойству `subject` текущего элемента в Outlook.</span><span class="sxs-lookup"><span data-stu-id="b789e-221">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="b789e-222">Элементы</span><span class="sxs-lookup"><span data-stu-id="b789e-222">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails"></a><span data-ttu-id="b789e-223">attachments :Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="b789e-223">attachments :Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

<span data-ttu-id="b789e-224">Получает вложения элемента в качестве массива.</span><span class="sxs-lookup"><span data-stu-id="b789e-224">Gets the item's attachments as an array.</span></span> <span data-ttu-id="b789e-225">Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="b789e-225">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="b789e-226">Outlook блокирует определенные типы файлов из-за потенциальных проблем с безопасностью, поэтому они не возвращаются.</span><span class="sxs-lookup"><span data-stu-id="b789e-226">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="b789e-227">Дополнительные сведения см. в статье [Блокировка вложений в Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="b789e-227">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="b789e-228">Тип:</span><span class="sxs-lookup"><span data-stu-id="b789e-228">Type:</span></span>

*   <span data-ttu-id="b789e-229">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="b789e-229">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="b789e-230">Требования</span><span class="sxs-lookup"><span data-stu-id="b789e-230">Requirements</span></span>

|<span data-ttu-id="b789e-231">Требование</span><span class="sxs-lookup"><span data-stu-id="b789e-231">Requirement</span></span>|<span data-ttu-id="b789e-232">Значение</span><span class="sxs-lookup"><span data-stu-id="b789e-232">Value</span></span>|
|---|---|
|[<span data-ttu-id="b789e-233">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="b789e-233">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="b789e-234">1.0</span><span class="sxs-lookup"><span data-stu-id="b789e-234">1.0</span></span>|
|[<span data-ttu-id="b789e-235">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b789e-235">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="b789e-236">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b789e-236">ReadItem</span></span>|
|[<span data-ttu-id="b789e-237">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b789e-237">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="b789e-238">Чтение</span><span class="sxs-lookup"><span data-stu-id="b789e-238">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b789e-239">Пример</span><span class="sxs-lookup"><span data-stu-id="b789e-239">Example</span></span>

<span data-ttu-id="b789e-240">С помощью приведенного ниже кода можно создать HTML-строку с подробными сведениями обо всех вложениях для текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="b789e-240">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

####  <a name="bcc-recipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="b789e-241">bcc :[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="b789e-241">bcc :[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="b789e-242">Получает объект, который предоставляет методы для получения или обновления получателей скрытой копии сообщения.</span><span class="sxs-lookup"><span data-stu-id="b789e-242">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="b789e-243">Только в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="b789e-243">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="b789e-244">Тип:</span><span class="sxs-lookup"><span data-stu-id="b789e-244">Type:</span></span>

*   [<span data-ttu-id="b789e-245">Recipients</span><span class="sxs-lookup"><span data-stu-id="b789e-245">Recipients</span></span>](/javascript/api/outlook/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="b789e-246">Требования</span><span class="sxs-lookup"><span data-stu-id="b789e-246">Requirements</span></span>

|<span data-ttu-id="b789e-247">Требование</span><span class="sxs-lookup"><span data-stu-id="b789e-247">Requirement</span></span>|<span data-ttu-id="b789e-248">Значение</span><span class="sxs-lookup"><span data-stu-id="b789e-248">Value</span></span>|
|---|---|
|[<span data-ttu-id="b789e-249">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="b789e-249">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="b789e-250">1.1</span><span class="sxs-lookup"><span data-stu-id="b789e-250">1.1</span></span>|
|[<span data-ttu-id="b789e-251">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b789e-251">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="b789e-252">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b789e-252">ReadItem</span></span>|
|[<span data-ttu-id="b789e-253">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b789e-253">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="b789e-254">Создание</span><span class="sxs-lookup"><span data-stu-id="b789e-254">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="b789e-255">Пример</span><span class="sxs-lookup"><span data-stu-id="b789e-255">Example</span></span>

```javascript
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlookofficebody"></a><span data-ttu-id="b789e-256">body :[Body](/javascript/api/outlook/office.body)</span><span class="sxs-lookup"><span data-stu-id="b789e-256">body :[Body](/javascript/api/outlook/office.body)</span></span>

<span data-ttu-id="b789e-257">Получает объект, предоставляющий методы для работы с основным текстом элемента.</span><span class="sxs-lookup"><span data-stu-id="b789e-257">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="b789e-258">Тип:</span><span class="sxs-lookup"><span data-stu-id="b789e-258">Type:</span></span>

*   [<span data-ttu-id="b789e-259">Body</span><span class="sxs-lookup"><span data-stu-id="b789e-259">Body</span></span>](/javascript/api/outlook/office.body)

##### <a name="requirements"></a><span data-ttu-id="b789e-260">Требования</span><span class="sxs-lookup"><span data-stu-id="b789e-260">Requirements</span></span>

|<span data-ttu-id="b789e-261">Требование</span><span class="sxs-lookup"><span data-stu-id="b789e-261">Requirement</span></span>|<span data-ttu-id="b789e-262">Значение</span><span class="sxs-lookup"><span data-stu-id="b789e-262">Value</span></span>|
|---|---|
|[<span data-ttu-id="b789e-263">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="b789e-263">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="b789e-264">1.1</span><span class="sxs-lookup"><span data-stu-id="b789e-264">1.1</span></span>|
|[<span data-ttu-id="b789e-265">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b789e-265">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="b789e-266">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b789e-266">ReadItem</span></span>|
|[<span data-ttu-id="b789e-267">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b789e-267">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="b789e-268">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="b789e-268">Compose or read</span></span>|

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="b789e-269">cc :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="b789e-269">cc :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="b789e-270">Предоставляет доступ к получателям копии сообщения.</span><span class="sxs-lookup"><span data-stu-id="b789e-270">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="b789e-271">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="b789e-271">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="b789e-272">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="b789e-272">Read mode</span></span>

<span data-ttu-id="b789e-p106">Свойство `cc` возвращает массив, который содержит объект `EmailAddressDetails` для каждого получателя, указанного в строке **Копия** сообщения. Коллекция может включать не более 100 элементов.</span><span class="sxs-lookup"><span data-stu-id="b789e-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="b789e-275">Режим создания</span><span class="sxs-lookup"><span data-stu-id="b789e-275">Compose mode</span></span>

<span data-ttu-id="b789e-276">Свойство `cc` возвращает объект `Recipients`, предоставляющий методы для получения или обновления получателей, которые указаны в строке **Копия** сообщения.</span><span class="sxs-lookup"><span data-stu-id="b789e-276">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="b789e-277">Тип:</span><span class="sxs-lookup"><span data-stu-id="b789e-277">Type:</span></span>

*   <span data-ttu-id="b789e-278">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="b789e-278">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="b789e-279">Требования</span><span class="sxs-lookup"><span data-stu-id="b789e-279">Requirements</span></span>

|<span data-ttu-id="b789e-280">Требование</span><span class="sxs-lookup"><span data-stu-id="b789e-280">Requirement</span></span>|<span data-ttu-id="b789e-281">Значение</span><span class="sxs-lookup"><span data-stu-id="b789e-281">Value</span></span>|
|---|---|
|[<span data-ttu-id="b789e-282">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="b789e-282">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="b789e-283">1.0</span><span class="sxs-lookup"><span data-stu-id="b789e-283">1.0</span></span>|
|[<span data-ttu-id="b789e-284">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b789e-284">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="b789e-285">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b789e-285">ReadItem</span></span>|
|[<span data-ttu-id="b789e-286">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b789e-286">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="b789e-287">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="b789e-287">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="b789e-288">Пример</span><span class="sxs-lookup"><span data-stu-id="b789e-288">Example</span></span>

```javascript
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="b789e-289">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="b789e-289">(nullable) conversationId :String</span></span>

<span data-ttu-id="b789e-290">Получает идентификатор разговора по электронной почте, содержащего конкретное сообщение.</span><span class="sxs-lookup"><span data-stu-id="b789e-290">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="b789e-p107">Вы можете получить целочисленное значение этого свойства, если ваше почтовое приложение активируется в формах просмотра или формах создания ответов. Если пользователь изменит тему ответа, после его отправки идентификатор беседы будет изменен, и полученное ранее значение будет недействительным.</span><span class="sxs-lookup"><span data-stu-id="b789e-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="b789e-p108">Это свойство имеет значение NULL для нового элемента в форме создания. Свойство `conversationId` вернет значение, если пользователь задаст тему и сохранит элемент.</span><span class="sxs-lookup"><span data-stu-id="b789e-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="b789e-295">Тип:</span><span class="sxs-lookup"><span data-stu-id="b789e-295">Type:</span></span>

*   <span data-ttu-id="b789e-296">String</span><span class="sxs-lookup"><span data-stu-id="b789e-296">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="b789e-297">Требования</span><span class="sxs-lookup"><span data-stu-id="b789e-297">Requirements</span></span>

|<span data-ttu-id="b789e-298">Требование</span><span class="sxs-lookup"><span data-stu-id="b789e-298">Requirement</span></span>|<span data-ttu-id="b789e-299">Значение</span><span class="sxs-lookup"><span data-stu-id="b789e-299">Value</span></span>|
|---|---|
|[<span data-ttu-id="b789e-300">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="b789e-300">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="b789e-301">1.0</span><span class="sxs-lookup"><span data-stu-id="b789e-301">1.0</span></span>|
|[<span data-ttu-id="b789e-302">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b789e-302">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="b789e-303">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b789e-303">ReadItem</span></span>|
|[<span data-ttu-id="b789e-304">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b789e-304">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="b789e-305">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="b789e-305">Compose or read</span></span>|

#### <a name="datetimecreated-date"></a><span data-ttu-id="b789e-306">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="b789e-306">dateTimeCreated :Date</span></span>

<span data-ttu-id="b789e-p109">Получает дату и время создания элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="b789e-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="b789e-309">Тип:</span><span class="sxs-lookup"><span data-stu-id="b789e-309">Type:</span></span>

*   <span data-ttu-id="b789e-310">Date</span><span class="sxs-lookup"><span data-stu-id="b789e-310">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="b789e-311">Требования</span><span class="sxs-lookup"><span data-stu-id="b789e-311">Requirements</span></span>

|<span data-ttu-id="b789e-312">Требование</span><span class="sxs-lookup"><span data-stu-id="b789e-312">Requirement</span></span>|<span data-ttu-id="b789e-313">Значение</span><span class="sxs-lookup"><span data-stu-id="b789e-313">Value</span></span>|
|---|---|
|[<span data-ttu-id="b789e-314">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="b789e-314">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="b789e-315">1.0</span><span class="sxs-lookup"><span data-stu-id="b789e-315">1.0</span></span>|
|[<span data-ttu-id="b789e-316">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b789e-316">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="b789e-317">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b789e-317">ReadItem</span></span>|
|[<span data-ttu-id="b789e-318">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b789e-318">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="b789e-319">Чтение</span><span class="sxs-lookup"><span data-stu-id="b789e-319">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b789e-320">Пример</span><span class="sxs-lookup"><span data-stu-id="b789e-320">Example</span></span>

```javascript
var created = Office.context.mailbox.item.dateTimeCreated;
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="b789e-321">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="b789e-321">dateTimeModified :Date</span></span>

<span data-ttu-id="b789e-p110">Получает дату и время последнего изменения элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="b789e-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="b789e-324">Этот элемент не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="b789e-324">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="b789e-325">Тип:</span><span class="sxs-lookup"><span data-stu-id="b789e-325">Type:</span></span>

*   <span data-ttu-id="b789e-326">Date</span><span class="sxs-lookup"><span data-stu-id="b789e-326">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="b789e-327">Требования</span><span class="sxs-lookup"><span data-stu-id="b789e-327">Requirements</span></span>

|<span data-ttu-id="b789e-328">Требование</span><span class="sxs-lookup"><span data-stu-id="b789e-328">Requirement</span></span>|<span data-ttu-id="b789e-329">Значение</span><span class="sxs-lookup"><span data-stu-id="b789e-329">Value</span></span>|
|---|---|
|[<span data-ttu-id="b789e-330">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="b789e-330">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="b789e-331">1.0</span><span class="sxs-lookup"><span data-stu-id="b789e-331">1.0</span></span>|
|[<span data-ttu-id="b789e-332">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b789e-332">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="b789e-333">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b789e-333">ReadItem</span></span>|
|[<span data-ttu-id="b789e-334">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b789e-334">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="b789e-335">Чтение</span><span class="sxs-lookup"><span data-stu-id="b789e-335">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b789e-336">Пример</span><span class="sxs-lookup"><span data-stu-id="b789e-336">Example</span></span>

```javascript
var modified = Office.context.mailbox.item.dateTimeModified;
```

####  <a name="end-datetimejavascriptapioutlookofficetime"></a><span data-ttu-id="b789e-337">end :Date|[Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="b789e-337">end :Date|[Time](/javascript/api/outlook/office.time)</span></span>

<span data-ttu-id="b789e-338">Получает или задает дату и время окончания встречи.</span><span class="sxs-lookup"><span data-stu-id="b789e-338">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="b789e-p111">Свойство `end` представлено в виде значения даты и времени в формате UTC. Преобразовать значение свойства end в местные значения даты и времени клиента можно с помощью метода [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime).</span><span class="sxs-lookup"><span data-stu-id="b789e-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="b789e-341">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="b789e-341">Read mode</span></span>

<span data-ttu-id="b789e-342">Свойство `end` возвращает объект `Date`.</span><span class="sxs-lookup"><span data-stu-id="b789e-342">The `end` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="b789e-343">Режим создания</span><span class="sxs-lookup"><span data-stu-id="b789e-343">Compose mode</span></span>

<span data-ttu-id="b789e-344">Свойство `end` возвращает объект `Time`.</span><span class="sxs-lookup"><span data-stu-id="b789e-344">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="b789e-345">Если вы задаете время окончания с помощью метода [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-), необходимо использовать метод [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) для преобразования местного времени на клиенте в формат UTC для сервера.</span><span class="sxs-lookup"><span data-stu-id="b789e-345">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="b789e-346">Тип:</span><span class="sxs-lookup"><span data-stu-id="b789e-346">Type:</span></span>

*   <span data-ttu-id="b789e-347">Date | [Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="b789e-347">Date | [Time](/javascript/api/outlook/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="b789e-348">Требования</span><span class="sxs-lookup"><span data-stu-id="b789e-348">Requirements</span></span>

|<span data-ttu-id="b789e-349">Требование</span><span class="sxs-lookup"><span data-stu-id="b789e-349">Requirement</span></span>|<span data-ttu-id="b789e-350">Значение</span><span class="sxs-lookup"><span data-stu-id="b789e-350">Value</span></span>|
|---|---|
|[<span data-ttu-id="b789e-351">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="b789e-351">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="b789e-352">1.0</span><span class="sxs-lookup"><span data-stu-id="b789e-352">1.0</span></span>|
|[<span data-ttu-id="b789e-353">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b789e-353">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="b789e-354">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b789e-354">ReadItem</span></span>|
|[<span data-ttu-id="b789e-355">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b789e-355">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="b789e-356">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="b789e-356">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="b789e-357">Пример</span><span class="sxs-lookup"><span data-stu-id="b789e-357">Example</span></span>

<span data-ttu-id="b789e-358">В примере ниже показано, как с помощью метода [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) объекта `Time` задать время окончания встречи в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="b789e-358">The following example sets the end time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

#### <a name="enhancedlocation-enhancedlocationjavascriptapioutlookofficeenhancedlocation"></a><span data-ttu-id="b789e-359">enhancedLocation:[EnhancedLocation](/javascript/api/outlook/office.enhancedlocation)</span><span class="sxs-lookup"><span data-stu-id="b789e-359">enhancedLocation :[EnhancedLocation](/javascript/api/outlook/office.enhancedlocation)</span></span>

<span data-ttu-id="b789e-360">Получает или задает расположение встречи.</span><span class="sxs-lookup"><span data-stu-id="b789e-360">Gets or sets the locations of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="b789e-361">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="b789e-361">Read mode</span></span>

<span data-ttu-id="b789e-362">`enhancedLocation` Свойство возвращает объект [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation) , который позволяет задать расположения (каждый из которых представлен объектом [LocationDetails](/javascript/api/outlook/office.locationdetails) ), связанный с встречи.</span><span class="sxs-lookup"><span data-stu-id="b789e-362">The `enhancedLocation` property returns an [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation) object that allows you to get the set of locations (each represented by a [LocationDetails](/javascript/api/outlook/office.locationdetails) object) associated with the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="b789e-363">Режим создания</span><span class="sxs-lookup"><span data-stu-id="b789e-363">Compose mode</span></span>

<span data-ttu-id="b789e-364">`enhancedLocation` Свойство возвращает объект [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation) , который предоставляет методы для получения, удалить или добавить расположения на встречи.</span><span class="sxs-lookup"><span data-stu-id="b789e-364">The `enhancedLocation` property returns an [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation) object that provides methods to get, remove, or add locations on an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="b789e-365">Тип:</span><span class="sxs-lookup"><span data-stu-id="b789e-365">Type:</span></span>

*   [<span data-ttu-id="b789e-366">EnhancedLocation</span><span class="sxs-lookup"><span data-stu-id="b789e-366">EnhancedLocation</span></span>](/javascript/api/outlook/office.enhancedlocation)

##### <a name="requirements"></a><span data-ttu-id="b789e-367">Требования</span><span class="sxs-lookup"><span data-stu-id="b789e-367">Requirements</span></span>

|<span data-ttu-id="b789e-368">Требование</span><span class="sxs-lookup"><span data-stu-id="b789e-368">Requirement</span></span>|<span data-ttu-id="b789e-369">Значение</span><span class="sxs-lookup"><span data-stu-id="b789e-369">Value</span></span>|
|---|---|
|[<span data-ttu-id="b789e-370">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="b789e-370">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="b789e-371">Предварительная версия</span><span class="sxs-lookup"><span data-stu-id="b789e-371">Preview</span></span>|
|[<span data-ttu-id="b789e-372">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b789e-372">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="b789e-373">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b789e-373">ReadItem</span></span>|
|[<span data-ttu-id="b789e-374">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b789e-374">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="b789e-375">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="b789e-375">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="b789e-376">Пример</span><span class="sxs-lookup"><span data-stu-id="b789e-376">Example</span></span>

<span data-ttu-id="b789e-377">В следующем примере получается текущего расположения, связанной со встречей.</span><span class="sxs-lookup"><span data-stu-id="b789e-377">The following example gets the current locations associated with the appointment.</span></span>

```javascript
Office.context.mailbox.item.enhancedLocation.getAsync(callbackFunction);

function callbackFunction(asyncResult) {
  asyncResult.value.forEach(function (place) {
    console.log("Display name: " + place.displayName);
    console.log("Type: " + place.locationIdentifier.type);
    if (place.locationIdentifier.type == Office.MailboxEnums.LocationType.Room) {
      console.log("Email address: " + place.emailAddress);
    }
  });
}

// Sample output:
// Display name: Conf Room 14
// Type: room
// Email address: cr14@contoso.com
// Display name: Paris
// Type: custom
```

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom"></a><span data-ttu-id="b789e-378">from :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[From](/javascript/api/outlook/office.from)</span><span class="sxs-lookup"><span data-stu-id="b789e-378">from :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[From](/javascript/api/outlook/office.from)</span></span>

<span data-ttu-id="b789e-379">Получает адрес электронной почты отправителя сообщения.</span><span class="sxs-lookup"><span data-stu-id="b789e-379">Gets the email address of the sender of a message.</span></span>

<span data-ttu-id="b789e-p112">Свойства `from` и [`sender`](#sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails) представляют одно лицо, если сообщение не отправлено представителем. В противном случае свойство `from` представляет лицо, делегировавшее полномочия, а свойство sender — представителя.</span><span class="sxs-lookup"><span data-stu-id="b789e-p112">The `from` and [`sender`](#sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="b789e-382">Свойству `recipientType`, принадлежащему объекту `EmailAddressDetails` в свойстве `from`, задано значение `undefined`.</span><span class="sxs-lookup"><span data-stu-id="b789e-382">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="b789e-383">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="b789e-383">Read mode</span></span>

<span data-ttu-id="b789e-384">Свойство `from` возвращает объект `EmailAddressDetails`.</span><span class="sxs-lookup"><span data-stu-id="b789e-384">The `from` property returns an `EmailAddressDetails` object.</span></span>

```javascript
var from = Office.context.mailbox.item.from;
```

##### <a name="compose-mode"></a><span data-ttu-id="b789e-385">Режим создания</span><span class="sxs-lookup"><span data-stu-id="b789e-385">Compose mode</span></span>

<span data-ttu-id="b789e-386">Свойство `from` возвращает объект `From`, который предоставляет метод для получения значения отправителя.</span><span class="sxs-lookup"><span data-stu-id="b789e-386">The `from` property returns a `From` object that provides a method to get the from value.</span></span>

```javascript
Office.context.mailbox.item.from.getAsync(callback);

function callback(asyncResult) {
  var from = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="b789e-387">Тип:</span><span class="sxs-lookup"><span data-stu-id="b789e-387">Type:</span></span>

*   <span data-ttu-id="b789e-388">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [From](/javascript/api/outlook/office.from)</span><span class="sxs-lookup"><span data-stu-id="b789e-388">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [From](/javascript/api/outlook/office.from)</span></span>

##### <a name="requirements"></a><span data-ttu-id="b789e-389">Требования</span><span class="sxs-lookup"><span data-stu-id="b789e-389">Requirements</span></span>

|<span data-ttu-id="b789e-390">Требование</span><span class="sxs-lookup"><span data-stu-id="b789e-390">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="b789e-391">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="b789e-391">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="b789e-392">1.0</span><span class="sxs-lookup"><span data-stu-id="b789e-392">1.0</span></span>|<span data-ttu-id="b789e-393">1.7</span><span class="sxs-lookup"><span data-stu-id="b789e-393">1.7</span></span>|
|[<span data-ttu-id="b789e-394">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b789e-394">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="b789e-395">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b789e-395">ReadItem</span></span>|<span data-ttu-id="b789e-396">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="b789e-396">ReadWriteItem</span></span>|
|[<span data-ttu-id="b789e-397">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b789e-397">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="b789e-398">Чтение</span><span class="sxs-lookup"><span data-stu-id="b789e-398">Read</span></span>|<span data-ttu-id="b789e-399">Создание</span><span class="sxs-lookup"><span data-stu-id="b789e-399">Compose</span></span>|

#### <a name="internetheaders-internetheadersjavascriptapioutlookofficeinternetheaders"></a><span data-ttu-id="b789e-400">internetHeaders :[InternetHeaders](/javascript/api/outlook/office.internetheaders)</span><span class="sxs-lookup"><span data-stu-id="b789e-400">internetHeaders :[InternetHeaders](/javascript/api/outlook/office.internetheaders)</span></span>

<span data-ttu-id="b789e-401">Получает или задает заголовки Интернета в сообщении.</span><span class="sxs-lookup"><span data-stu-id="b789e-401">Gets or sets the internet headers of a message.</span></span>

##### <a name="type"></a><span data-ttu-id="b789e-402">Тип:</span><span class="sxs-lookup"><span data-stu-id="b789e-402">Type:</span></span>

*   [<span data-ttu-id="b789e-403">InternetHeaders</span><span class="sxs-lookup"><span data-stu-id="b789e-403">InternetHeaders</span></span>](/javascript/api/outlook/office.internetheaders)

##### <a name="requirements"></a><span data-ttu-id="b789e-404">Требования</span><span class="sxs-lookup"><span data-stu-id="b789e-404">Requirements</span></span>

|<span data-ttu-id="b789e-405">Requirement</span><span class="sxs-lookup"><span data-stu-id="b789e-405">Requirement</span></span>|<span data-ttu-id="b789e-406">Значение</span><span class="sxs-lookup"><span data-stu-id="b789e-406">Value</span></span>|
|---|---|
|[<span data-ttu-id="b789e-407">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="b789e-407">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="b789e-408">Предварительная версия</span><span class="sxs-lookup"><span data-stu-id="b789e-408">Preview</span></span>|
|[<span data-ttu-id="b789e-409">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b789e-409">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="b789e-410">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b789e-410">ReadItem</span></span>|
|[<span data-ttu-id="b789e-411">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b789e-411">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="b789e-412">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="b789e-412">Compose or read</span></span>|

#### <a name="internetmessageid-string"></a><span data-ttu-id="b789e-413">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="b789e-413">internetMessageId :String</span></span>

<span data-ttu-id="b789e-p113">Получает идентификатор интернет-сообщения для электронного сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="b789e-p113">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="b789e-416">Тип:</span><span class="sxs-lookup"><span data-stu-id="b789e-416">Type:</span></span>

*   <span data-ttu-id="b789e-417">String</span><span class="sxs-lookup"><span data-stu-id="b789e-417">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="b789e-418">Требования</span><span class="sxs-lookup"><span data-stu-id="b789e-418">Requirements</span></span>

|<span data-ttu-id="b789e-419">Requirement</span><span class="sxs-lookup"><span data-stu-id="b789e-419">Requirement</span></span>|<span data-ttu-id="b789e-420">Значение</span><span class="sxs-lookup"><span data-stu-id="b789e-420">Value</span></span>|
|---|---|
|[<span data-ttu-id="b789e-421">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="b789e-421">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="b789e-422">1.0</span><span class="sxs-lookup"><span data-stu-id="b789e-422">1.0</span></span>|
|[<span data-ttu-id="b789e-423">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b789e-423">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="b789e-424">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b789e-424">ReadItem</span></span>|
|[<span data-ttu-id="b789e-425">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b789e-425">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="b789e-426">Чтение</span><span class="sxs-lookup"><span data-stu-id="b789e-426">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b789e-427">Пример</span><span class="sxs-lookup"><span data-stu-id="b789e-427">Example</span></span>

```javascript
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="b789e-428">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="b789e-428">itemClass :String</span></span>

<span data-ttu-id="b789e-p114">Получает класс элемента веб-служб Exchange для выбранного элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="b789e-p114">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="b789e-p115">Свойство `itemClass` указывает класс сообщения выбранного элемента. Ниже приводятся классы сообщения по умолчанию для элемента сообщения или встречи.</span><span class="sxs-lookup"><span data-stu-id="b789e-p115">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

|<span data-ttu-id="b789e-433">Тип</span><span class="sxs-lookup"><span data-stu-id="b789e-433">Type</span></span>|<span data-ttu-id="b789e-434">Описание</span><span class="sxs-lookup"><span data-stu-id="b789e-434">Description</span></span>|<span data-ttu-id="b789e-435">Класс элемента</span><span class="sxs-lookup"><span data-stu-id="b789e-435">item class</span></span>|
|---|---|---|
|<span data-ttu-id="b789e-436">Элементы встречи</span><span class="sxs-lookup"><span data-stu-id="b789e-436">Appointment items</span></span>|<span data-ttu-id="b789e-437">Это элементы календаря для класса элемента `IPM.Appointment` или `IPM.Appointment.Occurrence`.</span><span class="sxs-lookup"><span data-stu-id="b789e-437">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span>|`IPM.Appointment`<br />`IPM.Appointment.Occurrence`|
|<span data-ttu-id="b789e-438">Элементы сообщения</span><span class="sxs-lookup"><span data-stu-id="b789e-438">Message items</span></span>|<span data-ttu-id="b789e-439">Сюда входят электронные сообщения, для которых по умолчанию задан класс сообщения `IPM.Note`, а также приглашения на собрания, ответы на них и уведомления об их отмене, использующие `IPM.Schedule.Meeting` в качестве базового класса сообщения.</span><span class="sxs-lookup"><span data-stu-id="b789e-439">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span>|`IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled`|

<span data-ttu-id="b789e-440">Можно создавать настраиваемые классы сообщения, расширяющие классы сообщения по умолчанию, например настраиваемый класс сообщения о встрече `IPM.Appointment.Contoso`.</span><span class="sxs-lookup"><span data-stu-id="b789e-440">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="b789e-441">Тип:</span><span class="sxs-lookup"><span data-stu-id="b789e-441">Type:</span></span>

*   <span data-ttu-id="b789e-442">String</span><span class="sxs-lookup"><span data-stu-id="b789e-442">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="b789e-443">Требования</span><span class="sxs-lookup"><span data-stu-id="b789e-443">Requirements</span></span>

|<span data-ttu-id="b789e-444">Requirement</span><span class="sxs-lookup"><span data-stu-id="b789e-444">Requirement</span></span>|<span data-ttu-id="b789e-445">Значение</span><span class="sxs-lookup"><span data-stu-id="b789e-445">Value</span></span>|
|---|---|
|[<span data-ttu-id="b789e-446">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="b789e-446">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="b789e-447">1.0</span><span class="sxs-lookup"><span data-stu-id="b789e-447">1.0</span></span>|
|[<span data-ttu-id="b789e-448">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b789e-448">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="b789e-449">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b789e-449">ReadItem</span></span>|
|[<span data-ttu-id="b789e-450">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b789e-450">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="b789e-451">Чтение</span><span class="sxs-lookup"><span data-stu-id="b789e-451">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b789e-452">Пример</span><span class="sxs-lookup"><span data-stu-id="b789e-452">Example</span></span>

```javascript
var itemClass = Office.context.mailbox.item.itemClass;
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="b789e-453">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="b789e-453">(nullable) itemId :String</span></span>

<span data-ttu-id="b789e-p116">Получает идентификатор элемента веб-служб Exchange для текущего элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="b789e-p116">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="b789e-456">Идентификатор, возвращаемый свойством `itemId`, совпадает с идентификатором элемента веб-служб Exchange.</span><span class="sxs-lookup"><span data-stu-id="b789e-456">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="b789e-457">Свойство `itemId` не совпадает с идентификатором записи Outlook, а также идентификатором, который используется REST API Outlook.</span><span class="sxs-lookup"><span data-stu-id="b789e-457">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="b789e-458">Прежде чем совершать вызовы REST API, используя это значение, его необходимо преобразовать с помощью [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="b789e-458">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="b789e-459">Дополнительные сведения см. в статье [Использование REST API Outlook из надстройки Outlook](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="b789e-459">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="b789e-p118">Свойство `itemId` недоступно в режиме создания. Если требуется идентификатор элемента, с помощью метода [`saveAsync`](#saveasyncoptions-callback) можно сохранить элемент в хранилище. При этом в параметре [`AsyncResult.value`](/javascript/api/office/office.asyncresult) функции обратного вызова возвращается идентификатор элемента.</span><span class="sxs-lookup"><span data-stu-id="b789e-p118">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="b789e-462">Тип:</span><span class="sxs-lookup"><span data-stu-id="b789e-462">Type:</span></span>

*   <span data-ttu-id="b789e-463">String</span><span class="sxs-lookup"><span data-stu-id="b789e-463">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="b789e-464">Требования</span><span class="sxs-lookup"><span data-stu-id="b789e-464">Requirements</span></span>

|<span data-ttu-id="b789e-465">Requirement</span><span class="sxs-lookup"><span data-stu-id="b789e-465">Requirement</span></span>|<span data-ttu-id="b789e-466">Значение</span><span class="sxs-lookup"><span data-stu-id="b789e-466">Value</span></span>|
|---|---|
|[<span data-ttu-id="b789e-467">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="b789e-467">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="b789e-468">1.0</span><span class="sxs-lookup"><span data-stu-id="b789e-468">1.0</span></span>|
|[<span data-ttu-id="b789e-469">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b789e-469">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="b789e-470">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b789e-470">ReadItem</span></span>|
|[<span data-ttu-id="b789e-471">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b789e-471">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="b789e-472">Чтение</span><span class="sxs-lookup"><span data-stu-id="b789e-472">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b789e-473">Пример</span><span class="sxs-lookup"><span data-stu-id="b789e-473">Example</span></span>

<span data-ttu-id="b789e-p119">Указанный ниже код проверяет наличие идентификатора элемента. Если свойство `itemId` возвращает значение `null` или `undefined`, элемент будет сохранен в хранилище, а из асинхронного результата будет получен идентификатор элемента.</span><span class="sxs-lookup"><span data-stu-id="b789e-p119">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```javascript
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result){
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype"></a><span data-ttu-id="b789e-476">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="b789e-476">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="b789e-477">Получает тип элемента, который представляет экземпляр.</span><span class="sxs-lookup"><span data-stu-id="b789e-477">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="b789e-478">Свойство `itemType` возвращает одно из значений перечисления `ItemType`, которое указывает, является ли экземпляр объекта `item` сообщением или встречей.</span><span class="sxs-lookup"><span data-stu-id="b789e-478">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="b789e-479">Тип:</span><span class="sxs-lookup"><span data-stu-id="b789e-479">Type:</span></span>

*   [<span data-ttu-id="b789e-480">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="b789e-480">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="b789e-481">Требования</span><span class="sxs-lookup"><span data-stu-id="b789e-481">Requirements</span></span>

|<span data-ttu-id="b789e-482">Requirement</span><span class="sxs-lookup"><span data-stu-id="b789e-482">Requirement</span></span>|<span data-ttu-id="b789e-483">Значение</span><span class="sxs-lookup"><span data-stu-id="b789e-483">Value</span></span>|
|---|---|
|[<span data-ttu-id="b789e-484">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="b789e-484">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="b789e-485">1.0</span><span class="sxs-lookup"><span data-stu-id="b789e-485">1.0</span></span>|
|[<span data-ttu-id="b789e-486">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b789e-486">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="b789e-487">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b789e-487">ReadItem</span></span>|
|[<span data-ttu-id="b789e-488">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b789e-488">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="b789e-489">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="b789e-489">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="b789e-490">Пример</span><span class="sxs-lookup"><span data-stu-id="b789e-490">Example</span></span>

```javascript
if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Message)
  // do something
else
  // do something else
```

####  <a name="location-stringlocationjavascriptapioutlookofficelocation"></a><span data-ttu-id="b789e-491">location :String|[Location](/javascript/api/outlook/office.location)</span><span class="sxs-lookup"><span data-stu-id="b789e-491">location :String|[Location](/javascript/api/outlook/office.location)</span></span>

<span data-ttu-id="b789e-492">Получает или задает место встречи.</span><span class="sxs-lookup"><span data-stu-id="b789e-492">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="b789e-493">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="b789e-493">Read mode</span></span>

<span data-ttu-id="b789e-494">Свойство `location` возвращает строку, содержащую сведения о месте встречи.</span><span class="sxs-lookup"><span data-stu-id="b789e-494">The `location` property returns a string that contains the location of the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="b789e-495">Режим создания</span><span class="sxs-lookup"><span data-stu-id="b789e-495">Compose mode</span></span>

<span data-ttu-id="b789e-496">Свойство `location` возвращает объект `Location`, предоставляющий методы, которые используются для получения и задания места встречи.</span><span class="sxs-lookup"><span data-stu-id="b789e-496">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="b789e-497">Тип:</span><span class="sxs-lookup"><span data-stu-id="b789e-497">Type:</span></span>

*   <span data-ttu-id="b789e-498">String | [Location](/javascript/api/outlook/office.location)</span><span class="sxs-lookup"><span data-stu-id="b789e-498">String | [Location](/javascript/api/outlook/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="b789e-499">Требования</span><span class="sxs-lookup"><span data-stu-id="b789e-499">Requirements</span></span>

|<span data-ttu-id="b789e-500">Requirement</span><span class="sxs-lookup"><span data-stu-id="b789e-500">Requirement</span></span>|<span data-ttu-id="b789e-501">Значение</span><span class="sxs-lookup"><span data-stu-id="b789e-501">Value</span></span>|
|---|---|
|[<span data-ttu-id="b789e-502">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="b789e-502">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="b789e-503">1.0</span><span class="sxs-lookup"><span data-stu-id="b789e-503">1.0</span></span>|
|[<span data-ttu-id="b789e-504">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b789e-504">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="b789e-505">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b789e-505">ReadItem</span></span>|
|[<span data-ttu-id="b789e-506">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b789e-506">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="b789e-507">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="b789e-507">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="b789e-508">Пример</span><span class="sxs-lookup"><span data-stu-id="b789e-508">Example</span></span>

```javascript
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

#### <a name="normalizedsubject-string"></a><span data-ttu-id="b789e-509">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="b789e-509">normalizedSubject :String</span></span>

<span data-ttu-id="b789e-p120">Получает тему элемента со всеми удаленными префиксами (включая `RE:` и `FWD:`). Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="b789e-p120">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="b789e-p121">Свойство normalizedSubject получает тему элемента со стандартными префиксами (такими как `RE:` и `FW:`), добавляемыми почтовыми программами. Для получения темы элемента с неизмененными префиксами используйте свойство [`subject`](#subject-stringsubjectjavascriptapioutlookofficesubject).</span><span class="sxs-lookup"><span data-stu-id="b789e-p121">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubjectjavascriptapioutlookofficesubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="b789e-514">Тип:</span><span class="sxs-lookup"><span data-stu-id="b789e-514">Type:</span></span>

*   <span data-ttu-id="b789e-515">String</span><span class="sxs-lookup"><span data-stu-id="b789e-515">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="b789e-516">Требования</span><span class="sxs-lookup"><span data-stu-id="b789e-516">Requirements</span></span>

|<span data-ttu-id="b789e-517">Requirement</span><span class="sxs-lookup"><span data-stu-id="b789e-517">Requirement</span></span>|<span data-ttu-id="b789e-518">Значение</span><span class="sxs-lookup"><span data-stu-id="b789e-518">Value</span></span>|
|---|---|
|[<span data-ttu-id="b789e-519">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="b789e-519">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="b789e-520">1.0</span><span class="sxs-lookup"><span data-stu-id="b789e-520">1.0</span></span>|
|[<span data-ttu-id="b789e-521">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b789e-521">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="b789e-522">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b789e-522">ReadItem</span></span>|
|[<span data-ttu-id="b789e-523">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b789e-523">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="b789e-524">Чтение</span><span class="sxs-lookup"><span data-stu-id="b789e-524">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b789e-525">Пример</span><span class="sxs-lookup"><span data-stu-id="b789e-525">Example</span></span>

```javascript
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
```

####  <a name="notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessages"></a><span data-ttu-id="b789e-526">notificationMessages :[NotificationMessages](/javascript/api/outlook/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="b789e-526">notificationMessages :[NotificationMessages](/javascript/api/outlook/office.notificationmessages)</span></span>

<span data-ttu-id="b789e-527">Получает сообщения уведомления для элемента.</span><span class="sxs-lookup"><span data-stu-id="b789e-527">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="b789e-528">Тип:</span><span class="sxs-lookup"><span data-stu-id="b789e-528">Type:</span></span>

*   [<span data-ttu-id="b789e-529">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="b789e-529">NotificationMessages</span></span>](/javascript/api/outlook/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="b789e-530">Требования</span><span class="sxs-lookup"><span data-stu-id="b789e-530">Requirements</span></span>

|<span data-ttu-id="b789e-531">Requirement</span><span class="sxs-lookup"><span data-stu-id="b789e-531">Requirement</span></span>|<span data-ttu-id="b789e-532">Значение</span><span class="sxs-lookup"><span data-stu-id="b789e-532">Value</span></span>|
|---|---|
|[<span data-ttu-id="b789e-533">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="b789e-533">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="b789e-534">1.3</span><span class="sxs-lookup"><span data-stu-id="b789e-534">1.3</span></span>|
|[<span data-ttu-id="b789e-535">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b789e-535">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="b789e-536">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b789e-536">ReadItem</span></span>|
|[<span data-ttu-id="b789e-537">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b789e-537">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="b789e-538">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="b789e-538">Compose or read</span></span>|

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="b789e-539">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="b789e-539">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="b789e-540">Предоставляет доступ к необязательным участникам события.</span><span class="sxs-lookup"><span data-stu-id="b789e-540">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="b789e-541">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="b789e-541">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="b789e-542">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="b789e-542">Read mode</span></span>

<span data-ttu-id="b789e-543">Свойство `optionalAttendees` возвращает массив, содержащий объект `EmailAddressDetails` для каждого необязательного участника собрания.</span><span class="sxs-lookup"><span data-stu-id="b789e-543">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="b789e-544">Режим создания</span><span class="sxs-lookup"><span data-stu-id="b789e-544">Compose mode</span></span>

<span data-ttu-id="b789e-545">Свойство `optionalAttendees` возвращает объект `Recipients`, который предоставляет методы для получения или обновления необязательных участников собрания.</span><span class="sxs-lookup"><span data-stu-id="b789e-545">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="b789e-546">Тип:</span><span class="sxs-lookup"><span data-stu-id="b789e-546">Type:</span></span>

*   <span data-ttu-id="b789e-547">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="b789e-547">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="b789e-548">Требования</span><span class="sxs-lookup"><span data-stu-id="b789e-548">Requirements</span></span>

|<span data-ttu-id="b789e-549">Requirement</span><span class="sxs-lookup"><span data-stu-id="b789e-549">Requirement</span></span>|<span data-ttu-id="b789e-550">Значение</span><span class="sxs-lookup"><span data-stu-id="b789e-550">Value</span></span>|
|---|---|
|[<span data-ttu-id="b789e-551">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="b789e-551">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="b789e-552">1.0</span><span class="sxs-lookup"><span data-stu-id="b789e-552">1.0</span></span>|
|[<span data-ttu-id="b789e-553">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b789e-553">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="b789e-554">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b789e-554">ReadItem</span></span>|
|[<span data-ttu-id="b789e-555">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b789e-555">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="b789e-556">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="b789e-556">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="b789e-557">Пример</span><span class="sxs-lookup"><span data-stu-id="b789e-557">Example</span></span>

```javascript
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsorganizerjavascriptapioutlookofficeorganizer"></a><span data-ttu-id="b789e-558">organizer :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[Organizer](/javascript/api/outlook/office.organizer)</span><span class="sxs-lookup"><span data-stu-id="b789e-558">organizer :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[Organizer](/javascript/api/outlook/office.organizer)</span></span>

<span data-ttu-id="b789e-559">Получает адрес электронной почты организатора указанного собрания.</span><span class="sxs-lookup"><span data-stu-id="b789e-559">Gets the email address of the organizer for a specified meeting.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="b789e-560">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="b789e-560">Read mode</span></span>

<span data-ttu-id="b789e-561">Свойство `organizer` возвращает объект [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails), представляющий организатора собрания.</span><span class="sxs-lookup"><span data-stu-id="b789e-561">The `organizer` property returns an [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) object that represents the meeting organizer.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="b789e-562">Режим создания</span><span class="sxs-lookup"><span data-stu-id="b789e-562">Compose mode</span></span>

<span data-ttu-id="b789e-563">Свойство `organizer` возвращает объект [Organizer](/javascript/api/outlook/office.organizer), который предоставляет метод для получения значения организатора.</span><span class="sxs-lookup"><span data-stu-id="b789e-563">The `organizer` property returns an [Organizer](/javascript/api/outlook/office.organizer) object that provides a method to get the organizer value.</span></span>

##### <a name="type"></a><span data-ttu-id="b789e-564">Тип:</span><span class="sxs-lookup"><span data-stu-id="b789e-564">Type:</span></span>

*   <span data-ttu-id="b789e-565">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [Organizer](/javascript/api/outlook/office.organizer)</span><span class="sxs-lookup"><span data-stu-id="b789e-565">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [Organizer](/javascript/api/outlook/office.organizer)</span></span>

##### <a name="requirements"></a><span data-ttu-id="b789e-566">Требования</span><span class="sxs-lookup"><span data-stu-id="b789e-566">Requirements</span></span>

|<span data-ttu-id="b789e-567">Требование</span><span class="sxs-lookup"><span data-stu-id="b789e-567">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="b789e-568">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="b789e-568">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="b789e-569">1.0</span><span class="sxs-lookup"><span data-stu-id="b789e-569">1.0</span></span>|<span data-ttu-id="b789e-570">1.7</span><span class="sxs-lookup"><span data-stu-id="b789e-570">1.7</span></span>|
|[<span data-ttu-id="b789e-571">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b789e-571">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="b789e-572">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b789e-572">ReadItem</span></span>|<span data-ttu-id="b789e-573">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="b789e-573">ReadWriteItem</span></span>|
|[<span data-ttu-id="b789e-574">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b789e-574">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="b789e-575">Чтение</span><span class="sxs-lookup"><span data-stu-id="b789e-575">Read</span></span>|<span data-ttu-id="b789e-576">Создание</span><span class="sxs-lookup"><span data-stu-id="b789e-576">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="b789e-577">Пример</span><span class="sxs-lookup"><span data-stu-id="b789e-577">Example</span></span>

```javascript
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
```

#### <a name="nullable-recurrence-recurrencejavascriptapioutlookofficerecurrence"></a><span data-ttu-id="b789e-578">(nullable) recurrence :[Recurrence](/javascript/api/outlook/office.recurrence)</span><span class="sxs-lookup"><span data-stu-id="b789e-578">(nullable) recurrence :[Recurrence](/javascript/api/outlook/office.recurrence)</span></span>

<span data-ttu-id="b789e-579">Получает или задает расписание повторения для встречи.</span><span class="sxs-lookup"><span data-stu-id="b789e-579">Gets or sets the recurrence pattern of an appointment.</span></span> <span data-ttu-id="b789e-580">Получает расписание повторения для приглашения на собрание.</span><span class="sxs-lookup"><span data-stu-id="b789e-580">Gets the recurrence pattern of a meeting request.</span></span> <span data-ttu-id="b789e-581">Доступно в режимах чтения и создания для элементов встречи.</span><span class="sxs-lookup"><span data-stu-id="b789e-581">Read and compose modes for appointment items.</span></span> <span data-ttu-id="b789e-582">Доступно в режиме чтения для элементов приглашения на собрание.</span><span class="sxs-lookup"><span data-stu-id="b789e-582">Read mode for meeting request items.</span></span>

<span data-ttu-id="b789e-583">Свойство `recurrence` возвращает объект [recurrence](/javascript/api/outlook/office.recurrence) для повторяющихся встреч или приглашений на собрание, если элемент представляет собой серию или экземпляр в пределах серии.</span><span class="sxs-lookup"><span data-stu-id="b789e-583">The `recurrence` property returns a [recurrence](/javascript/api/outlook/office.recurrence) object for recurring appointments or meetings requests if an item is a series or an instance in a series.</span></span> <span data-ttu-id="b789e-584">Значение `null` возвращается для отдельных встреч и приглашений на собрания, связанных с одной встречей.</span><span class="sxs-lookup"><span data-stu-id="b789e-584">`null` is returned for single appointments and meeting requests of single appointments.</span></span> <span data-ttu-id="b789e-585">Значение `undefined` возвращается для сообщений, которые не являются приглашениями на собрания.</span><span class="sxs-lookup"><span data-stu-id="b789e-585">`undefined` is returned for messages that are not meeting requests.</span></span>

> <span data-ttu-id="b789e-586">Примечание. Приглашения на собрания имеют значение `itemClass` для класса IPM.Schedule.Meeting.Request.</span><span class="sxs-lookup"><span data-stu-id="b789e-586">Note: Meeting requests have an `itemClass` value of IPM.Schedule.Meeting.Request.</span></span>

> <span data-ttu-id="b789e-587">Примечание. Если объект recurrence имеет значение `null`, он представляет собой отдельную встречу или приглашение на собрание, связанное с одной встречей, и НЕ входит в серию.</span><span class="sxs-lookup"><span data-stu-id="b789e-587">Note: If the recurrence object is `null`, this indicates that the object is a single appointment or a meeting request of a single appointment and NOT a part of a series.</span></span>

##### <a name="type"></a><span data-ttu-id="b789e-588">Тип:</span><span class="sxs-lookup"><span data-stu-id="b789e-588">Type:</span></span>

* [<span data-ttu-id="b789e-589">Recurrence</span><span class="sxs-lookup"><span data-stu-id="b789e-589">Recurrence</span></span>](/javascript/api/outlook/office.recurrence)

|<span data-ttu-id="b789e-590">Требование</span><span class="sxs-lookup"><span data-stu-id="b789e-590">Requirement</span></span>|<span data-ttu-id="b789e-591">Значение</span><span class="sxs-lookup"><span data-stu-id="b789e-591">Value</span></span>|
|---|---|
|[<span data-ttu-id="b789e-592">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="b789e-592">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="b789e-593">1.7</span><span class="sxs-lookup"><span data-stu-id="b789e-593">1.7</span></span>|
|[<span data-ttu-id="b789e-594">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b789e-594">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="b789e-595">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b789e-595">ReadItem</span></span>|
|[<span data-ttu-id="b789e-596">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b789e-596">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="b789e-597">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="b789e-597">Compose or read</span></span>|

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="b789e-598">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="b789e-598">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="b789e-599">Предоставляет доступ к обязательным участникам события.</span><span class="sxs-lookup"><span data-stu-id="b789e-599">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="b789e-600">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="b789e-600">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="b789e-601">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="b789e-601">Read mode</span></span>

<span data-ttu-id="b789e-602">Свойство `requiredAttendees` возвращает массив, содержащий объект `EmailAddressDetails` для каждого обязательного участника собрания.</span><span class="sxs-lookup"><span data-stu-id="b789e-602">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="b789e-603">Режим создания</span><span class="sxs-lookup"><span data-stu-id="b789e-603">Compose mode</span></span>

<span data-ttu-id="b789e-604">Свойство `requiredAttendees` возвращает объект `Recipients`, предоставляющий методы, с помощью которых можно получить или обновить сведения об обязательных участниках собрания.</span><span class="sxs-lookup"><span data-stu-id="b789e-604">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="b789e-605">Тип:</span><span class="sxs-lookup"><span data-stu-id="b789e-605">Type:</span></span>

*   <span data-ttu-id="b789e-606">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="b789e-606">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="b789e-607">Требования</span><span class="sxs-lookup"><span data-stu-id="b789e-607">Requirements</span></span>

|<span data-ttu-id="b789e-608">Requirement</span><span class="sxs-lookup"><span data-stu-id="b789e-608">Requirement</span></span>|<span data-ttu-id="b789e-609">Значение</span><span class="sxs-lookup"><span data-stu-id="b789e-609">Value</span></span>|
|---|---|
|[<span data-ttu-id="b789e-610">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="b789e-610">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="b789e-611">1.0</span><span class="sxs-lookup"><span data-stu-id="b789e-611">1.0</span></span>|
|[<span data-ttu-id="b789e-612">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b789e-612">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="b789e-613">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b789e-613">ReadItem</span></span>|
|[<span data-ttu-id="b789e-614">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b789e-614">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="b789e-615">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="b789e-615">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="b789e-616">Пример</span><span class="sxs-lookup"><span data-stu-id="b789e-616">Example</span></span>

```javascript
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
}
```

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails"></a><span data-ttu-id="b789e-617">sender :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="b789e-617">sender :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)</span></span>

<span data-ttu-id="b789e-p126">Получает электронный адрес отправителя электронного сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="b789e-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="b789e-p127">Свойства [`from`](#from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom) и `sender` представляют одно лицо, если сообщение не отправлено представителем. В противном случае свойство `from` представляет лицо, делегировавшее полномочия, а свойство sender — представителя.</span><span class="sxs-lookup"><span data-stu-id="b789e-p127">The [`from`](#from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="b789e-622">Свойству `recipientType`, принадлежащему объекту `EmailAddressDetails` в свойстве `sender`, задано значение `undefined`.</span><span class="sxs-lookup"><span data-stu-id="b789e-622">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="b789e-623">Тип:</span><span class="sxs-lookup"><span data-stu-id="b789e-623">Type:</span></span>

*   [<span data-ttu-id="b789e-624">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="b789e-624">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="b789e-625">Требования</span><span class="sxs-lookup"><span data-stu-id="b789e-625">Requirements</span></span>

|<span data-ttu-id="b789e-626">Requirement</span><span class="sxs-lookup"><span data-stu-id="b789e-626">Requirement</span></span>|<span data-ttu-id="b789e-627">Значение</span><span class="sxs-lookup"><span data-stu-id="b789e-627">Value</span></span>|
|---|---|
|[<span data-ttu-id="b789e-628">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="b789e-628">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="b789e-629">1.0</span><span class="sxs-lookup"><span data-stu-id="b789e-629">1.0</span></span>|
|[<span data-ttu-id="b789e-630">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b789e-630">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="b789e-631">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b789e-631">ReadItem</span></span>|
|[<span data-ttu-id="b789e-632">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b789e-632">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="b789e-633">Чтение</span><span class="sxs-lookup"><span data-stu-id="b789e-633">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b789e-634">Пример</span><span class="sxs-lookup"><span data-stu-id="b789e-634">Example</span></span>

```javascript
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
```

#### <a name="nullable-seriesid-string"></a><span data-ttu-id="b789e-635">(nullable) seriesId :String</span><span class="sxs-lookup"><span data-stu-id="b789e-635">(nullable) seriesId :String</span></span>

<span data-ttu-id="b789e-636">Получает идентификатор серии, к которой относится экземпляр.</span><span class="sxs-lookup"><span data-stu-id="b789e-636">Gets the id of the series that an instance belongs to.</span></span>

<span data-ttu-id="b789e-637">В Outlook Web App и Outlook свойство `seriesId` возвращает идентификатор веб-служб Exchange (EWS) родительского элемента (серии), к которому относится этот элемент.</span><span class="sxs-lookup"><span data-stu-id="b789e-637">In OWA and Outlook, the `seriesId` returns the Exchange Web Services (EWS) ID of the parent (series) item that this item belongs to.</span></span> <span data-ttu-id="b789e-638">Однако в iOS и Android свойство `seriesId` возвращает идентификатор REST родительского элемента.</span><span class="sxs-lookup"><span data-stu-id="b789e-638">However, in iOS and Android, the `seriesId` returns the REST ID of the parent item.</span></span>

> [!NOTE]
> <span data-ttu-id="b789e-639">Идентификатор, возвращаемый свойством `seriesId`, совпадает с идентификатором элемента веб-служб Exchange.</span><span class="sxs-lookup"><span data-stu-id="b789e-639">The identifier returned by the `seriesId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="b789e-640">Свойство `seriesId` не совпадает с идентификаторами Outlook, которые используются в REST API Outlook.</span><span class="sxs-lookup"><span data-stu-id="b789e-640">The `seriesId` property is not identical to the Outlook IDs used by the Outlook REST API.</span></span> <span data-ttu-id="b789e-641">Прежде чем совершать вызовы REST API, используя это значение, его необходимо преобразовать с помощью [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="b789e-641">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="b789e-642">Дополнительные сведения см. в статье [Использование REST API Outlook из надстройки Outlook](https://docs.microsoft.com/outlook/add-ins/use-rest-api).</span><span class="sxs-lookup"><span data-stu-id="b789e-642">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api).</span></span>

<span data-ttu-id="b789e-643">Свойство `seriesId` возвращает значение `null` для элементов, у которых нет родительских элементов, например отдельных встреч, элементов серий или приглашений на собрания, и возвращает значение `undefined` для всех других элементов, которые не представляют собой приглашения на собрания.</span><span class="sxs-lookup"><span data-stu-id="b789e-643">The `seriesId` property returns `null` for items that do not have parent items such as single appointments, series items, or meeting requests and returns `undefined` for any other items that are not meeting requests.</span></span>

##### <a name="type"></a><span data-ttu-id="b789e-644">Тип:</span><span class="sxs-lookup"><span data-stu-id="b789e-644">Type:</span></span>

* <span data-ttu-id="b789e-645">String</span><span class="sxs-lookup"><span data-stu-id="b789e-645">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="b789e-646">Требования</span><span class="sxs-lookup"><span data-stu-id="b789e-646">Requirements</span></span>

|<span data-ttu-id="b789e-647">Requirement</span><span class="sxs-lookup"><span data-stu-id="b789e-647">Requirement</span></span>|<span data-ttu-id="b789e-648">Значение</span><span class="sxs-lookup"><span data-stu-id="b789e-648">Value</span></span>|
|---|---|
|[<span data-ttu-id="b789e-649">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="b789e-649">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="b789e-650">1.7</span><span class="sxs-lookup"><span data-stu-id="b789e-650">1.7</span></span>|
|[<span data-ttu-id="b789e-651">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b789e-651">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="b789e-652">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b789e-652">ReadItem</span></span>|
|[<span data-ttu-id="b789e-653">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b789e-653">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="b789e-654">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="b789e-654">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="b789e-655">Пример</span><span class="sxs-lookup"><span data-stu-id="b789e-655">Example</span></span>

```javascript
var seriesId = Office.context.mailbox.item.seriesId;
var isSeries = (seriesId == null);
```

####  <a name="start-datetimejavascriptapioutlookofficetime"></a><span data-ttu-id="b789e-656">start :Date|[Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="b789e-656">start :Date|[Time](/javascript/api/outlook/office.time)</span></span>

<span data-ttu-id="b789e-657">Получает или задает дату и время начала встречи.</span><span class="sxs-lookup"><span data-stu-id="b789e-657">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="b789e-p130">Свойство `start` представлено в виде значения даты и времени в формате UTC. Это значение можно преобразовать в местные значения даты и времени клиента с помощью метода [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime).</span><span class="sxs-lookup"><span data-stu-id="b789e-p130">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="b789e-660">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="b789e-660">Read mode</span></span>

<span data-ttu-id="b789e-661">Свойство `start` возвращает объект `Date`.</span><span class="sxs-lookup"><span data-stu-id="b789e-661">The `start` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="b789e-662">Режим создания</span><span class="sxs-lookup"><span data-stu-id="b789e-662">Compose mode</span></span>

<span data-ttu-id="b789e-663">Свойство `start` возвращает объект `Time`.</span><span class="sxs-lookup"><span data-stu-id="b789e-663">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="b789e-664">Если вы задаете время начала с помощью метода [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-), необходимо использовать метод [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) для преобразования местного времени на клиенте в формат UTC для сервера.</span><span class="sxs-lookup"><span data-stu-id="b789e-664">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="b789e-665">Тип:</span><span class="sxs-lookup"><span data-stu-id="b789e-665">Type:</span></span>

*   <span data-ttu-id="b789e-666">Date | [Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="b789e-666">Date | [Time](/javascript/api/outlook/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="b789e-667">Требования</span><span class="sxs-lookup"><span data-stu-id="b789e-667">Requirements</span></span>

|<span data-ttu-id="b789e-668">Requirement</span><span class="sxs-lookup"><span data-stu-id="b789e-668">Requirement</span></span>|<span data-ttu-id="b789e-669">Значение</span><span class="sxs-lookup"><span data-stu-id="b789e-669">Value</span></span>|
|---|---|
|[<span data-ttu-id="b789e-670">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="b789e-670">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="b789e-671">1.0</span><span class="sxs-lookup"><span data-stu-id="b789e-671">1.0</span></span>|
|[<span data-ttu-id="b789e-672">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b789e-672">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="b789e-673">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b789e-673">ReadItem</span></span>|
|[<span data-ttu-id="b789e-674">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b789e-674">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="b789e-675">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="b789e-675">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="b789e-676">Пример</span><span class="sxs-lookup"><span data-stu-id="b789e-676">Example</span></span>

<span data-ttu-id="b789e-677">В примере ниже с помощью метода [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) объекта `Time` задается время начала встречи в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="b789e-677">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

####  <a name="subject-stringsubjectjavascriptapioutlookofficesubject"></a><span data-ttu-id="b789e-678">subject :String|[Subject](/javascript/api/outlook/office.subject)</span><span class="sxs-lookup"><span data-stu-id="b789e-678">subject :String|[Subject](/javascript/api/outlook/office.subject)</span></span>

<span data-ttu-id="b789e-679">Получает или задает описание, которое отображается в поле темы элемента.</span><span class="sxs-lookup"><span data-stu-id="b789e-679">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="b789e-680">Свойство `subject` получает или задает всю тему элемента для отправки с почтового сервера.</span><span class="sxs-lookup"><span data-stu-id="b789e-680">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="b789e-681">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="b789e-681">Read mode</span></span>

<span data-ttu-id="b789e-p131">Свойство `subject` возвращает строку. С помощью свойства [`normalizedSubject`](#normalizedsubject-string) можно получить тему без начальных префиксов, таких как `RE:` и `FW:`.</span><span class="sxs-lookup"><span data-stu-id="b789e-p131">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```javascript
var subject = Office.context.mailbox.item.subject;
```

##### <a name="compose-mode"></a><span data-ttu-id="b789e-684">Режим создания</span><span class="sxs-lookup"><span data-stu-id="b789e-684">Compose mode</span></span>

<span data-ttu-id="b789e-685">Свойство `subject` возвращает объект `Subject`, который предоставляет методы для получения и задания темы.</span><span class="sxs-lookup"><span data-stu-id="b789e-685">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```javascript
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="b789e-686">Тип:</span><span class="sxs-lookup"><span data-stu-id="b789e-686">Type:</span></span>

*   <span data-ttu-id="b789e-687">String | [Subject](/javascript/api/outlook/office.subject)</span><span class="sxs-lookup"><span data-stu-id="b789e-687">String | [Subject](/javascript/api/outlook/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="b789e-688">Требования</span><span class="sxs-lookup"><span data-stu-id="b789e-688">Requirements</span></span>

|<span data-ttu-id="b789e-689">Requirement</span><span class="sxs-lookup"><span data-stu-id="b789e-689">Requirement</span></span>|<span data-ttu-id="b789e-690">Значение</span><span class="sxs-lookup"><span data-stu-id="b789e-690">Value</span></span>|
|---|---|
|[<span data-ttu-id="b789e-691">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="b789e-691">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="b789e-692">1.0</span><span class="sxs-lookup"><span data-stu-id="b789e-692">1.0</span></span>|
|[<span data-ttu-id="b789e-693">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b789e-693">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="b789e-694">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b789e-694">ReadItem</span></span>|
|[<span data-ttu-id="b789e-695">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b789e-695">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="b789e-696">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="b789e-696">Compose or read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="b789e-697">to :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="b789e-697">to :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="b789e-698">Предоставляет доступ к получателям, указанным в строке **Кому** сообщения.</span><span class="sxs-lookup"><span data-stu-id="b789e-698">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="b789e-699">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="b789e-699">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="b789e-700">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="b789e-700">Read mode</span></span>

<span data-ttu-id="b789e-p133">Свойство `to` возвращает массив, содержащий объект `EmailAddressDetails` для каждого получателя в строке **Кому** сообщения. Коллекция может включать не более 100 элементов.</span><span class="sxs-lookup"><span data-stu-id="b789e-p133">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="b789e-703">Режим создания</span><span class="sxs-lookup"><span data-stu-id="b789e-703">Compose mode</span></span>

<span data-ttu-id="b789e-704">Свойство `to` возвращает объект `Recipients`, предоставляющий методы для получения или обновления получателей, которые указаны в строке **Кому** сообщения.</span><span class="sxs-lookup"><span data-stu-id="b789e-704">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="b789e-705">Тип:</span><span class="sxs-lookup"><span data-stu-id="b789e-705">Type:</span></span>

*   <span data-ttu-id="b789e-706">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="b789e-706">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="b789e-707">Требования</span><span class="sxs-lookup"><span data-stu-id="b789e-707">Requirements</span></span>

|<span data-ttu-id="b789e-708">Требование</span><span class="sxs-lookup"><span data-stu-id="b789e-708">Requirement</span></span>|<span data-ttu-id="b789e-709">Значение</span><span class="sxs-lookup"><span data-stu-id="b789e-709">Value</span></span>|
|---|---|
|[<span data-ttu-id="b789e-710">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="b789e-710">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="b789e-711">1.0</span><span class="sxs-lookup"><span data-stu-id="b789e-711">1.0</span></span>|
|[<span data-ttu-id="b789e-712">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b789e-712">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="b789e-713">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b789e-713">ReadItem</span></span>|
|[<span data-ttu-id="b789e-714">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b789e-714">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="b789e-715">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="b789e-715">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="b789e-716">Пример</span><span class="sxs-lookup"><span data-stu-id="b789e-716">Example</span></span>

```javascript
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

### <a name="methods"></a><span data-ttu-id="b789e-717">Методы</span><span class="sxs-lookup"><span data-stu-id="b789e-717">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="b789e-718">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="b789e-718">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="b789e-719">Добавляет файл в сообщение или встречу в качестве вложения.</span><span class="sxs-lookup"><span data-stu-id="b789e-719">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="b789e-720">Метод `addFileAttachmentAsync` передает файл по указанному универсальному коду ресурса (URI) и вкладывает его в элемент в форме создания.</span><span class="sxs-lookup"><span data-stu-id="b789e-720">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="b789e-721">Идентификатор можно последовательно использовать с методом [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback), чтобы удалить вложение, добавленное во время текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="b789e-721">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b789e-722">Параметры</span><span class="sxs-lookup"><span data-stu-id="b789e-722">Parameters:</span></span>
|<span data-ttu-id="b789e-723">Имя</span><span class="sxs-lookup"><span data-stu-id="b789e-723">Name</span></span>|<span data-ttu-id="b789e-724">Тип</span><span class="sxs-lookup"><span data-stu-id="b789e-724">Type</span></span>|<span data-ttu-id="b789e-725">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="b789e-725">Attributes</span></span>|<span data-ttu-id="b789e-726">Описание</span><span class="sxs-lookup"><span data-stu-id="b789e-726">Description</span></span>|
|---|---|---|---|
|`uri`|<span data-ttu-id="b789e-727">String</span><span class="sxs-lookup"><span data-stu-id="b789e-727">String</span></span>||<span data-ttu-id="b789e-p134">Универсальный код ресурса (URI), представляющий расположение файла, который нужно вложить в сообщение или встречу. Максимальная длина — 2048 символов.</span><span class="sxs-lookup"><span data-stu-id="b789e-p134">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="b789e-730">String</span><span class="sxs-lookup"><span data-stu-id="b789e-730">String</span></span>||<span data-ttu-id="b789e-p135">Имя вложения, которое отображается при передаче вложения. Максимальная длина — 255 символов.</span><span class="sxs-lookup"><span data-stu-id="b789e-p135">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="b789e-733">Object</span><span class="sxs-lookup"><span data-stu-id="b789e-733">Object</span></span>|<span data-ttu-id="b789e-734">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="b789e-734">&lt;optional&gt;</span></span>|<span data-ttu-id="b789e-735">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="b789e-735">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="b789e-736">Object</span><span class="sxs-lookup"><span data-stu-id="b789e-736">Object</span></span>|<span data-ttu-id="b789e-737">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="b789e-737">&lt;optional&gt;</span></span>|<span data-ttu-id="b789e-738">В методе обратного вызова разработчики могут указать любой объект, к которому необходимо получить доступ.</span><span class="sxs-lookup"><span data-stu-id="b789e-738">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="b789e-739">Boolean</span><span class="sxs-lookup"><span data-stu-id="b789e-739">Boolean</span></span>|<span data-ttu-id="b789e-740">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="b789e-740">&lt;optional&gt;</span></span>|<span data-ttu-id="b789e-741">Значение `true` указывает на то, что вложение будет встроено в текст сообщения и не должно отображаться в списке вложений.</span><span class="sxs-lookup"><span data-stu-id="b789e-741">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="b789e-742">function</span><span class="sxs-lookup"><span data-stu-id="b789e-742">function</span></span>|<span data-ttu-id="b789e-743">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="b789e-743">&lt;optional&gt;</span></span>|<span data-ttu-id="b789e-744">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="b789e-744">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="b789e-745">После успешного выполнения идентификатор вложения будет представлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="b789e-745">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="b789e-746">Если добавить вложение не удастся, объект `asyncResult` будет содержать объект `Error` с описанием ошибки.</span><span class="sxs-lookup"><span data-stu-id="b789e-746">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="b789e-747">Ошибки</span><span class="sxs-lookup"><span data-stu-id="b789e-747">Errors</span></span>

|<span data-ttu-id="b789e-748">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="b789e-748">Error code</span></span>|<span data-ttu-id="b789e-749">Описание</span><span class="sxs-lookup"><span data-stu-id="b789e-749">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="b789e-750">Вложение превышает максимальный размер.</span><span class="sxs-lookup"><span data-stu-id="b789e-750">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="b789e-751">Расширение вложения не поддерживается.</span><span class="sxs-lookup"><span data-stu-id="b789e-751">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="b789e-752">Сообщение или встреча содержат слишком много вложений.</span><span class="sxs-lookup"><span data-stu-id="b789e-752">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b789e-753">Требования</span><span class="sxs-lookup"><span data-stu-id="b789e-753">Requirements</span></span>

|<span data-ttu-id="b789e-754">Требование</span><span class="sxs-lookup"><span data-stu-id="b789e-754">Requirement</span></span>|<span data-ttu-id="b789e-755">Значение</span><span class="sxs-lookup"><span data-stu-id="b789e-755">Value</span></span>|
|---|---|
|[<span data-ttu-id="b789e-756">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="b789e-756">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="b789e-757">1.1</span><span class="sxs-lookup"><span data-stu-id="b789e-757">1.1</span></span>|
|[<span data-ttu-id="b789e-758">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b789e-758">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="b789e-759">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="b789e-759">ReadWriteItem</span></span>|
|[<span data-ttu-id="b789e-760">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b789e-760">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="b789e-761">Создание</span><span class="sxs-lookup"><span data-stu-id="b789e-761">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="b789e-762">Примеры</span><span class="sxs-lookup"><span data-stu-id="b789e-762">Examples</span></span>

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

<span data-ttu-id="b789e-763">В приведенном ниже примере файл изображения добавляется в качестве встроенного вложения, а ссылка на вложение добавляется в текст сообщения.</span><span class="sxs-lookup"><span data-stu-id="b789e-763">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

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

#### <a name="addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback"></a><span data-ttu-id="b789e-764">addFileAttachmentFromBase64Async(base64File, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="b789e-764">addFileAttachmentFromBase64Async(base64File, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="b789e-765">Добавляет файл из кодирования base64 в сообщение или встречу в качестве вложения.</span><span class="sxs-lookup"><span data-stu-id="b789e-765">Adds a file from the base64 encoding to a message or appointment as an attachment.</span></span>

<span data-ttu-id="b789e-766">Метод `addFileAttachmentFromBase64Async` передает файл из кодировки base64 и вкладывает его в элемент в форме создания.</span><span class="sxs-lookup"><span data-stu-id="b789e-766">The `addFileAttachmentFromBase64Async` method uploads the file from the base64 encoding and attaches it to the item in the compose form.</span></span> <span data-ttu-id="b789e-767">Этот способ возвращает идентификатор вложения в объекте AsyncResult.value.</span><span class="sxs-lookup"><span data-stu-id="b789e-767">This method returns the attachment identifier in the AsyncResult.value object.</span></span>

<span data-ttu-id="b789e-768">Идентификатор можно использовать с методом [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback), чтобы удалить вложение, добавленное во время текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="b789e-768">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b789e-769">Параметры</span><span class="sxs-lookup"><span data-stu-id="b789e-769">Parameters:</span></span>
|<span data-ttu-id="b789e-770">Имя</span><span class="sxs-lookup"><span data-stu-id="b789e-770">Name</span></span>|<span data-ttu-id="b789e-771">Тип</span><span class="sxs-lookup"><span data-stu-id="b789e-771">Type</span></span>|<span data-ttu-id="b789e-772">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="b789e-772">Attributes</span></span>|<span data-ttu-id="b789e-773">Описание</span><span class="sxs-lookup"><span data-stu-id="b789e-773">Description</span></span>|
|---|---|---|---|
|`base64File`|<span data-ttu-id="b789e-774">String</span><span class="sxs-lookup"><span data-stu-id="b789e-774">String</span></span>||<span data-ttu-id="b789e-775">Закодированное содержимое base64 изображения или файла, которое следует добавить в сообщение электронной почты или событие.</span><span class="sxs-lookup"><span data-stu-id="b789e-775">The base64 encoded content of an image or file to be added to an email or event.</span></span>|
|`attachmentName`|<span data-ttu-id="b789e-776">Строка</span><span class="sxs-lookup"><span data-stu-id="b789e-776">String</span></span>||<span data-ttu-id="b789e-p137">Имя вложения, которое отображается при передаче вложения. Максимальная длина — 255 символов.</span><span class="sxs-lookup"><span data-stu-id="b789e-p137">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="b789e-779">Object</span><span class="sxs-lookup"><span data-stu-id="b789e-779">Object</span></span>|<span data-ttu-id="b789e-780">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="b789e-780">&lt;optional&gt;</span></span>|<span data-ttu-id="b789e-781">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="b789e-781">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="b789e-782">Object</span><span class="sxs-lookup"><span data-stu-id="b789e-782">Object</span></span>|<span data-ttu-id="b789e-783">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="b789e-783">&lt;optional&gt;</span></span>|<span data-ttu-id="b789e-784">В методе обратного вызова разработчики могут указать любой объект, к которому необходимо получить доступ.</span><span class="sxs-lookup"><span data-stu-id="b789e-784">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="b789e-785">Boolean</span><span class="sxs-lookup"><span data-stu-id="b789e-785">Boolean</span></span>|<span data-ttu-id="b789e-786">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="b789e-786">&lt;optional&gt;</span></span>|<span data-ttu-id="b789e-787">Значение `true` указывает на то, что вложение будет встроено в текст сообщения и не должно отображаться в списке вложений.</span><span class="sxs-lookup"><span data-stu-id="b789e-787">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="b789e-788">function</span><span class="sxs-lookup"><span data-stu-id="b789e-788">function</span></span>|<span data-ttu-id="b789e-789">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="b789e-789">&lt;optional&gt;</span></span>|<span data-ttu-id="b789e-790">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="b789e-790">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="b789e-791">После успешного выполнения идентификатор вложения будет представлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="b789e-791">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="b789e-792">Если добавить вложение не удастся, объект `asyncResult` будет содержать объект `Error` с описанием ошибки.</span><span class="sxs-lookup"><span data-stu-id="b789e-792">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="b789e-793">Ошибки</span><span class="sxs-lookup"><span data-stu-id="b789e-793">Errors</span></span>

|<span data-ttu-id="b789e-794">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="b789e-794">Error code</span></span>|<span data-ttu-id="b789e-795">Описание</span><span class="sxs-lookup"><span data-stu-id="b789e-795">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="b789e-796">Вложение превышает максимальный размер.</span><span class="sxs-lookup"><span data-stu-id="b789e-796">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="b789e-797">Расширение вложения не поддерживается.</span><span class="sxs-lookup"><span data-stu-id="b789e-797">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="b789e-798">Сообщение или встреча содержат слишком много вложений.</span><span class="sxs-lookup"><span data-stu-id="b789e-798">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b789e-799">Требования</span><span class="sxs-lookup"><span data-stu-id="b789e-799">Requirements</span></span>

|<span data-ttu-id="b789e-800">Требование</span><span class="sxs-lookup"><span data-stu-id="b789e-800">Requirement</span></span>|<span data-ttu-id="b789e-801">Значение</span><span class="sxs-lookup"><span data-stu-id="b789e-801">Value</span></span>|
|---|---|
|[<span data-ttu-id="b789e-802">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="b789e-802">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="b789e-803">Предварительная версия</span><span class="sxs-lookup"><span data-stu-id="b789e-803">Preview</span></span>|
|[<span data-ttu-id="b789e-804">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b789e-804">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="b789e-805">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="b789e-805">ReadWriteItem</span></span>|
|[<span data-ttu-id="b789e-806">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b789e-806">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="b789e-807">Создание</span><span class="sxs-lookup"><span data-stu-id="b789e-807">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="b789e-808">Примеры</span><span class="sxs-lookup"><span data-stu-id="b789e-808">Examples</span></span>

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

####  <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="b789e-809">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="b789e-809">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="b789e-810">Добавляет обработчик для поддерживаемого события.</span><span class="sxs-lookup"><span data-stu-id="b789e-810">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="b789e-811">Сейчас поддерживаются следующие типы событий: `Office.EventType.AttachmentsChanged`, `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged` и `Office.EventType.RecurrenceChanged`.</span><span class="sxs-lookup"><span data-stu-id="b789e-811">Currently the supported event types are `Office.EventType.AttachmentsChanged`, `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b789e-812">Параметры:</span><span class="sxs-lookup"><span data-stu-id="b789e-812">Parameters:</span></span>

| <span data-ttu-id="b789e-813">Имя</span><span class="sxs-lookup"><span data-stu-id="b789e-813">Name</span></span> | <span data-ttu-id="b789e-814">Тип</span><span class="sxs-lookup"><span data-stu-id="b789e-814">Type</span></span> | <span data-ttu-id="b789e-815">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="b789e-815">Attributes</span></span> | <span data-ttu-id="b789e-816">Описание</span><span class="sxs-lookup"><span data-stu-id="b789e-816">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="b789e-817">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="b789e-817">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="b789e-818">Событие, которое должно вызвать обработчик.</span><span class="sxs-lookup"><span data-stu-id="b789e-818">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="b789e-819">Function</span><span class="sxs-lookup"><span data-stu-id="b789e-819">Function</span></span> || <span data-ttu-id="b789e-p138">Функция для обработки события. Функция должна принимать один параметр, представляющий собой объектный литерал. Значение свойства `type` параметра совпадет со значением параметра `eventType`, переданного методу `addHandlerAsync`.</span><span class="sxs-lookup"><span data-stu-id="b789e-p138">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="b789e-823">Объект</span><span class="sxs-lookup"><span data-stu-id="b789e-823">Object</span></span> | <span data-ttu-id="b789e-824">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="b789e-824">&lt;optional&gt;</span></span> | <span data-ttu-id="b789e-825">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="b789e-825">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="b789e-826">Object</span><span class="sxs-lookup"><span data-stu-id="b789e-826">Object</span></span> | <span data-ttu-id="b789e-827">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="b789e-827">&lt;optional&gt;</span></span> | <span data-ttu-id="b789e-828">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="b789e-828">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="b789e-829">функция</span><span class="sxs-lookup"><span data-stu-id="b789e-829">function</span></span>| <span data-ttu-id="b789e-830">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="b789e-830">&lt;optional&gt;</span></span>|<span data-ttu-id="b789e-831">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="b789e-831">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b789e-832">Требования</span><span class="sxs-lookup"><span data-stu-id="b789e-832">Requirements</span></span>

|<span data-ttu-id="b789e-833">Требование</span><span class="sxs-lookup"><span data-stu-id="b789e-833">Requirement</span></span>| <span data-ttu-id="b789e-834">Значение</span><span class="sxs-lookup"><span data-stu-id="b789e-834">Value</span></span>|
|---|---|
|[<span data-ttu-id="b789e-835">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="b789e-835">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b789e-836">1.7</span><span class="sxs-lookup"><span data-stu-id="b789e-836">1.7</span></span> |
|[<span data-ttu-id="b789e-837">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b789e-837">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b789e-838">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b789e-838">ReadItem</span></span> |
|[<span data-ttu-id="b789e-839">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b789e-839">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b789e-840">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="b789e-840">Compose or read</span></span> |

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="b789e-841">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="b789e-841">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="b789e-842">Добавляет к сообщению элемент Exchange, например сообщение, в виде вложения.</span><span class="sxs-lookup"><span data-stu-id="b789e-842">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="b789e-p139">С помощью метода `addItemAttachmentAsync` можно в элемент формы создания вложить элемент с указанным идентификатором Exchange. Если указать метод обратного вызова, то этот метод вызывается с помощью параметра `asyncResult`, который содержит идентификатор вложения или код, указывающий на ошибки, которые произошли при вложении элемента. При необходимости можно использовать параметр `options` для передачи сведений о состоянии методу обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="b789e-p139">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="b789e-846">Идентификатор можно использовать с методом [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback), чтобы удалить вложение, добавленное во время текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="b789e-846">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="b789e-847">Если ваша надстройка Office выполняется в Outlook Web App, метод `addItemAttachmentAsync` обеспечивает вложение элементов в элементы, отличные от редактируемого. Однако это действие не рекомендуем выполнять, так как оно не поддерживается.</span><span class="sxs-lookup"><span data-stu-id="b789e-847">If your Office Add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b789e-848">Параметры</span><span class="sxs-lookup"><span data-stu-id="b789e-848">Parameters:</span></span>

|<span data-ttu-id="b789e-849">Имя</span><span class="sxs-lookup"><span data-stu-id="b789e-849">Name</span></span>|<span data-ttu-id="b789e-850">Тип</span><span class="sxs-lookup"><span data-stu-id="b789e-850">Type</span></span>|<span data-ttu-id="b789e-851">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="b789e-851">Attributes</span></span>|<span data-ttu-id="b789e-852">Описание</span><span class="sxs-lookup"><span data-stu-id="b789e-852">Description</span></span>|
|---|---|---|---|
|`itemId`|<span data-ttu-id="b789e-853">Строка</span><span class="sxs-lookup"><span data-stu-id="b789e-853">String</span></span>||<span data-ttu-id="b789e-p140">Идентификатор Exchange для вкладываемого элемента. Максимальная длина — 100 символов.</span><span class="sxs-lookup"><span data-stu-id="b789e-p140">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="b789e-856">Строка</span><span class="sxs-lookup"><span data-stu-id="b789e-856">String</span></span>||<span data-ttu-id="b789e-p141">Тема вкладываемого элемента. Максимальная длина — 255 символов.</span><span class="sxs-lookup"><span data-stu-id="b789e-p141">The sujbect of the item to be attached. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="b789e-859">Object</span><span class="sxs-lookup"><span data-stu-id="b789e-859">Object</span></span>|<span data-ttu-id="b789e-860">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="b789e-860">&lt;optional&gt;</span></span>|<span data-ttu-id="b789e-861">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="b789e-861">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="b789e-862">Object</span><span class="sxs-lookup"><span data-stu-id="b789e-862">Object</span></span>|<span data-ttu-id="b789e-863">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="b789e-863">&lt;optional&gt;</span></span>|<span data-ttu-id="b789e-864">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="b789e-864">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="b789e-865">функция</span><span class="sxs-lookup"><span data-stu-id="b789e-865">function</span></span>|<span data-ttu-id="b789e-866">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="b789e-866">&lt;optional&gt;</span></span>|<span data-ttu-id="b789e-867">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="b789e-867">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="b789e-868">После успешного выполнения идентификатор вложения будет представлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="b789e-868">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="b789e-869">Если добавить вложение не удастся, объект `asyncResult` будет содержать объект `Error` с описанием ошибки.</span><span class="sxs-lookup"><span data-stu-id="b789e-869">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="b789e-870">Ошибки</span><span class="sxs-lookup"><span data-stu-id="b789e-870">Errors</span></span>

|<span data-ttu-id="b789e-871">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="b789e-871">Error code</span></span>|<span data-ttu-id="b789e-872">Описание</span><span class="sxs-lookup"><span data-stu-id="b789e-872">Description</span></span>|
|------------|-------------|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="b789e-873">Сообщение или встреча содержат слишком много вложений.</span><span class="sxs-lookup"><span data-stu-id="b789e-873">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b789e-874">Требования</span><span class="sxs-lookup"><span data-stu-id="b789e-874">Requirements</span></span>

|<span data-ttu-id="b789e-875">Требование</span><span class="sxs-lookup"><span data-stu-id="b789e-875">Requirement</span></span>|<span data-ttu-id="b789e-876">Значение</span><span class="sxs-lookup"><span data-stu-id="b789e-876">Value</span></span>|
|---|---|
|[<span data-ttu-id="b789e-877">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="b789e-877">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="b789e-878">1.1</span><span class="sxs-lookup"><span data-stu-id="b789e-878">1.1</span></span>|
|[<span data-ttu-id="b789e-879">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b789e-879">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="b789e-880">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="b789e-880">ReadWriteItem</span></span>|
|[<span data-ttu-id="b789e-881">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b789e-881">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="b789e-882">Создание</span><span class="sxs-lookup"><span data-stu-id="b789e-882">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="b789e-883">Пример</span><span class="sxs-lookup"><span data-stu-id="b789e-883">Example</span></span>

<span data-ttu-id="b789e-884">В следующем примере существующий элемент Outlook добавляется в виде вложения с именем `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="b789e-884">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

####  <a name="close"></a><span data-ttu-id="b789e-885">close()</span><span class="sxs-lookup"><span data-stu-id="b789e-885">close()</span></span>

<span data-ttu-id="b789e-886">Закрывает текущий создаваемый элемент.</span><span class="sxs-lookup"><span data-stu-id="b789e-886">Closes the current item that is being composed.</span></span>

<span data-ttu-id="b789e-p142">Работа метода `close` зависит от текущего состояния создаваемого элемента. Если элемент содержит несохраненные изменения, клиент предложит пользователю сохранить или отклонить их либо отменить действие закрытия.</span><span class="sxs-lookup"><span data-stu-id="b789e-p142">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="b789e-889">Если в Outlook в Интернете элемент представляет собой встречу, ранее сохраненную с помощью метода `saveAsync`, пользователю предлагается сохранить, отклонить или отменить действие, даже если с момента последнего сохранения элемента не вносились какие-либо изменения.</span><span class="sxs-lookup"><span data-stu-id="b789e-889">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="b789e-890">Если в клиенте Outlook для настольных ПК сообщение представляет собой ответ в тексте, метод `close` не работает.</span><span class="sxs-lookup"><span data-stu-id="b789e-890">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="b789e-891">Требования</span><span class="sxs-lookup"><span data-stu-id="b789e-891">Requirements</span></span>

|<span data-ttu-id="b789e-892">Требование</span><span class="sxs-lookup"><span data-stu-id="b789e-892">Requirement</span></span>|<span data-ttu-id="b789e-893">Значение</span><span class="sxs-lookup"><span data-stu-id="b789e-893">Value</span></span>|
|---|---|
|[<span data-ttu-id="b789e-894">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="b789e-894">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="b789e-895">1.3</span><span class="sxs-lookup"><span data-stu-id="b789e-895">1.3</span></span>|
|[<span data-ttu-id="b789e-896">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b789e-896">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="b789e-897">Restricted</span><span class="sxs-lookup"><span data-stu-id="b789e-897">Restricted</span></span>|
|[<span data-ttu-id="b789e-898">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b789e-898">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="b789e-899">Создание</span><span class="sxs-lookup"><span data-stu-id="b789e-899">Compose</span></span>|

#### <a name="displayreplyallformformdata"></a><span data-ttu-id="b789e-900">displayReplyAllForm(formData)</span><span class="sxs-lookup"><span data-stu-id="b789e-900">displayReplyAllForm(formData)</span></span>

<span data-ttu-id="b789e-901">Отображает форму ответа, включающую отправителя и всех получателей выбранного сообщения или организатора и всех участников выбранной встречи.</span><span class="sxs-lookup"><span data-stu-id="b789e-901">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="b789e-902">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="b789e-902">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="b789e-903">В Outlook Web App форма ответа отображается в виде всплывающей формы в представлении с 3 либо 1 или 2 колонками.</span><span class="sxs-lookup"><span data-stu-id="b789e-903">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="b789e-904">Если любой строковый параметр превышает указанные для него ограничения, `displayReplyAllForm` возвращает исключение.</span><span class="sxs-lookup"><span data-stu-id="b789e-904">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="b789e-p143">Если в параметре `formData.attachments` указаны вложения, Outlook и Outlook Web App пытаются скачать их и вложить в форму ответа. Если какие-либо вложения добавить не удается, в форме отображается сообщение об ошибке. Если сообщения об ошибках не предусмотрены, то они не отображаются.</span><span class="sxs-lookup"><span data-stu-id="b789e-p143">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b789e-908">Параметры:</span><span class="sxs-lookup"><span data-stu-id="b789e-908">Parameters:</span></span>

|<span data-ttu-id="b789e-909">Имя</span><span class="sxs-lookup"><span data-stu-id="b789e-909">Name</span></span>|<span data-ttu-id="b789e-910">Тип</span><span class="sxs-lookup"><span data-stu-id="b789e-910">Type</span></span>|<span data-ttu-id="b789e-911">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="b789e-911">Attributes</span></span>|<span data-ttu-id="b789e-912">Описание</span><span class="sxs-lookup"><span data-stu-id="b789e-912">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="b789e-913">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="b789e-913">String &#124; Object</span></span>||<span data-ttu-id="b789e-p144">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="b789e-p144">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="b789e-916">**ИЛИ**</span><span class="sxs-lookup"><span data-stu-id="b789e-916">**OR**</span></span><br/><span data-ttu-id="b789e-p145">Объект, который содержит текст или данные вложения и функцию обратного вызова. Ниже представлено определение этого объекта.</span><span class="sxs-lookup"><span data-stu-id="b789e-p145">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="b789e-919">String</span><span class="sxs-lookup"><span data-stu-id="b789e-919">String</span></span>|<span data-ttu-id="b789e-920">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="b789e-920">&lt;optional&gt;</span></span>|<span data-ttu-id="b789e-p146">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="b789e-p146">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="b789e-923">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="b789e-923">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="b789e-924">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="b789e-924">&lt;optional&gt;</span></span>|<span data-ttu-id="b789e-925">Массив объектов JSON, представляющих собой вложенные файлы или элементы.</span><span class="sxs-lookup"><span data-stu-id="b789e-925">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="b789e-926">Строка</span><span class="sxs-lookup"><span data-stu-id="b789e-926">String</span></span>||<span data-ttu-id="b789e-p147">Указывает тип вложения. Допустимые значения: `file` для вложенного файла и `item` для вложенного элемента.</span><span class="sxs-lookup"><span data-stu-id="b789e-p147">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="b789e-929">Строка</span><span class="sxs-lookup"><span data-stu-id="b789e-929">String</span></span>||<span data-ttu-id="b789e-930">Строка, содержащая имя вложения, длиной до 255 символов.</span><span class="sxs-lookup"><span data-stu-id="b789e-930">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="b789e-931">Строка</span><span class="sxs-lookup"><span data-stu-id="b789e-931">String</span></span>||<span data-ttu-id="b789e-p148">Используется, только если свойству `type` задано значение `file`. URI расположения файла.</span><span class="sxs-lookup"><span data-stu-id="b789e-p148">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="b789e-934">Boolean</span><span class="sxs-lookup"><span data-stu-id="b789e-934">Boolean</span></span>||<span data-ttu-id="b789e-p149">Используется, только если свойству `type` задано значение `file`. Значение `true` указывает на то, что вложение будет встроено в текст сообщения и не должно отображаться в списке вложений.</span><span class="sxs-lookup"><span data-stu-id="b789e-p149">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="b789e-937">String</span><span class="sxs-lookup"><span data-stu-id="b789e-937">String</span></span>||<span data-ttu-id="b789e-p150">Используется, только если свойству `type` задано значение `item`. Идентификатор вложения EWS. Это строка длиной до 100 символов.</span><span class="sxs-lookup"><span data-stu-id="b789e-p150">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="b789e-941">функция</span><span class="sxs-lookup"><span data-stu-id="b789e-941">function</span></span>|<span data-ttu-id="b789e-942">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="b789e-942">&lt;optional&gt;</span></span>|<span data-ttu-id="b789e-943">По завершении работы метода функция, переданная параметру `callback`, вызывается с помощью одного параметра `asyncResult`, представляющего собой объект [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="b789e-943">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b789e-944">Требования</span><span class="sxs-lookup"><span data-stu-id="b789e-944">Requirements</span></span>

|<span data-ttu-id="b789e-945">Требование</span><span class="sxs-lookup"><span data-stu-id="b789e-945">Requirement</span></span>|<span data-ttu-id="b789e-946">Значение</span><span class="sxs-lookup"><span data-stu-id="b789e-946">Value</span></span>|
|---|---|
|[<span data-ttu-id="b789e-947">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="b789e-947">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="b789e-948">1.0</span><span class="sxs-lookup"><span data-stu-id="b789e-948">1.0</span></span>|
|[<span data-ttu-id="b789e-949">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b789e-949">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="b789e-950">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b789e-950">ReadItem</span></span>|
|[<span data-ttu-id="b789e-951">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b789e-951">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="b789e-952">Чтение</span><span class="sxs-lookup"><span data-stu-id="b789e-952">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="b789e-953">Примеры</span><span class="sxs-lookup"><span data-stu-id="b789e-953">Examples</span></span>

<span data-ttu-id="b789e-954">Приведенный ниже код передает строку в функцию `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="b789e-954">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="b789e-955">Ответ с пустым текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="b789e-955">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="b789e-956">Ответ только с текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="b789e-956">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="b789e-957">Ответ с текстом сообщения и вложенным файлом.</span><span class="sxs-lookup"><span data-stu-id="b789e-957">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="b789e-958">Ответ с текстом сообщения и вложенным элементом.</span><span class="sxs-lookup"><span data-stu-id="b789e-958">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="b789e-959">Ответ с текстом сообщения, вложенным файлом, вложенным элементом и обратным вызовом.</span><span class="sxs-lookup"><span data-stu-id="b789e-959">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata"></a><span data-ttu-id="b789e-960">displayReplyForm(formData)</span><span class="sxs-lookup"><span data-stu-id="b789e-960">displayReplyForm(formData)</span></span>

<span data-ttu-id="b789e-961">Отображает форму ответа, включающую только отправителя выбранного сообщения или организатора выбранной встречи.</span><span class="sxs-lookup"><span data-stu-id="b789e-961">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="b789e-962">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="b789e-962">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="b789e-963">В Outlook Web App форма ответа отображается в виде всплывающей формы в представлении с 3 либо 1 или 2 колонками.</span><span class="sxs-lookup"><span data-stu-id="b789e-963">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="b789e-964">Если любой строковый параметр превышает указанные для него ограничения, `displayReplyForm` возвращает исключение.</span><span class="sxs-lookup"><span data-stu-id="b789e-964">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="b789e-p151">Если в параметре `formData.attachments` указаны вложения, Outlook и Outlook Web App пытаются скачать их и вложить в форму ответа. Если какие-либо вложения добавить не удается, в форме отображается сообщение об ошибке. Если сообщения об ошибках не предусмотрены, то они не отображаются.</span><span class="sxs-lookup"><span data-stu-id="b789e-p151">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b789e-968">Параметры</span><span class="sxs-lookup"><span data-stu-id="b789e-968">Parameters:</span></span>

|<span data-ttu-id="b789e-969">Имя</span><span class="sxs-lookup"><span data-stu-id="b789e-969">Name</span></span>|<span data-ttu-id="b789e-970">Тип</span><span class="sxs-lookup"><span data-stu-id="b789e-970">Type</span></span>|<span data-ttu-id="b789e-971">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="b789e-971">Attributes</span></span>|<span data-ttu-id="b789e-972">Описание</span><span class="sxs-lookup"><span data-stu-id="b789e-972">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="b789e-973">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="b789e-973">String &#124; Object</span></span>||<span data-ttu-id="b789e-p152">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="b789e-p152">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="b789e-976">**ИЛИ**</span><span class="sxs-lookup"><span data-stu-id="b789e-976">**OR**</span></span><br/><span data-ttu-id="b789e-p153">Объект, который содержит текст или данные вложения и функцию обратного вызова. Ниже представлено определение этого объекта.</span><span class="sxs-lookup"><span data-stu-id="b789e-p153">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="b789e-979">Строка</span><span class="sxs-lookup"><span data-stu-id="b789e-979">String</span></span>|<span data-ttu-id="b789e-980">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="b789e-980">&lt;optional&gt;</span></span>|<span data-ttu-id="b789e-p154">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="b789e-p154">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="b789e-983">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="b789e-983">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="b789e-984">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="b789e-984">&lt;optional&gt;</span></span>|<span data-ttu-id="b789e-985">Массив объектов JSON, представляющих собой вложенные файлы или элементы.</span><span class="sxs-lookup"><span data-stu-id="b789e-985">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="b789e-986">Строка</span><span class="sxs-lookup"><span data-stu-id="b789e-986">String</span></span>||<span data-ttu-id="b789e-p155">Указывает тип вложения. Допустимые значения: `file` для вложенного файла и `item` для вложенного элемента.</span><span class="sxs-lookup"><span data-stu-id="b789e-p155">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="b789e-989">Строка</span><span class="sxs-lookup"><span data-stu-id="b789e-989">String</span></span>||<span data-ttu-id="b789e-990">Строка, содержащая имя вложения, длиной до 255 символов.</span><span class="sxs-lookup"><span data-stu-id="b789e-990">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="b789e-991">Строка</span><span class="sxs-lookup"><span data-stu-id="b789e-991">String</span></span>||<span data-ttu-id="b789e-p156">Используется, только если свойству `type` задано значение `file`. URI расположения файла.</span><span class="sxs-lookup"><span data-stu-id="b789e-p156">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="b789e-994">Boolean</span><span class="sxs-lookup"><span data-stu-id="b789e-994">Boolean</span></span>||<span data-ttu-id="b789e-p157">Используется, только если свойству `type` задано значение `file`. Значение `true` указывает на то, что вложение будет встроено в текст сообщения и не должно отображаться в списке вложений.</span><span class="sxs-lookup"><span data-stu-id="b789e-p157">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="b789e-997">String</span><span class="sxs-lookup"><span data-stu-id="b789e-997">String</span></span>||<span data-ttu-id="b789e-p158">Используется, только если свойству `type` задано значение `item`. Идентификатор вложения EWS. Это строка длиной до 100 символов.</span><span class="sxs-lookup"><span data-stu-id="b789e-p158">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="b789e-1001">функция</span><span class="sxs-lookup"><span data-stu-id="b789e-1001">function</span></span>|<span data-ttu-id="b789e-1002">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="b789e-1002">&lt;optional&gt;</span></span>|<span data-ttu-id="b789e-1003">По завершении работы метода функция, переданная параметру `callback`, вызывается с помощью одного параметра `asyncResult`, представляющего собой объект [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="b789e-1003">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b789e-1004">Требования</span><span class="sxs-lookup"><span data-stu-id="b789e-1004">Requirements</span></span>

|<span data-ttu-id="b789e-1005">Требование</span><span class="sxs-lookup"><span data-stu-id="b789e-1005">Requirement</span></span>|<span data-ttu-id="b789e-1006">Значение</span><span class="sxs-lookup"><span data-stu-id="b789e-1006">Value</span></span>|
|---|---|
|[<span data-ttu-id="b789e-1007">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="b789e-1007">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="b789e-1008">1.0</span><span class="sxs-lookup"><span data-stu-id="b789e-1008">1.0</span></span>|
|[<span data-ttu-id="b789e-1009">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b789e-1009">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="b789e-1010">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b789e-1010">ReadItem</span></span>|
|[<span data-ttu-id="b789e-1011">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b789e-1011">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="b789e-1012">Чтение</span><span class="sxs-lookup"><span data-stu-id="b789e-1012">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="b789e-1013">Примеры</span><span class="sxs-lookup"><span data-stu-id="b789e-1013">Examples</span></span>

<span data-ttu-id="b789e-1014">Приведенный ниже код передает строку в функцию `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="b789e-1014">The following code passes a string to the `displayReplyForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="b789e-1015">Ответ с пустым текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="b789e-1015">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="b789e-1016">Ответ только с текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="b789e-1016">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="b789e-1017">Ответ с текстом сообщения и вложенным файлом.</span><span class="sxs-lookup"><span data-stu-id="b789e-1017">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="b789e-1018">Ответ с текстом сообщения и вложенным элементом.</span><span class="sxs-lookup"><span data-stu-id="b789e-1018">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="b789e-1019">Ответ с текстом сообщения, вложенным файлом, вложенным элементом и обратным вызовом.</span><span class="sxs-lookup"><span data-stu-id="b789e-1019">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getattachmentcontentasyncattachmentid-options-callback--attachmentcontentjavascriptapioutlookofficeattachmentcontent"></a><span data-ttu-id="b789e-1020">getAttachmentContentAsync(attachmentId, [options], callback) → [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span><span class="sxs-lookup"><span data-stu-id="b789e-1020">getAttachmentContentAsync(attachmentId, [options], callback) → [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span></span>

<span data-ttu-id="b789e-1021">Получает указанное вложение из сообщения или встречи и возвращает в качестве объекта `AttachmentContent`.</span><span class="sxs-lookup"><span data-stu-id="b789e-1021">Gets the specified attachment from a message or appointment and returns it as an `AttachmentContent` object.</span></span>

<span data-ttu-id="b789e-1022">Метод `getAttachmentContentAsync` получает вложение с указанным идентификатором из элемента.</span><span class="sxs-lookup"><span data-stu-id="b789e-1022">The `getAttachmentContentAsync` method gets the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="b789e-1023">Рекомендуется использовать идентификатор для получения вложения в том же сеансе, в котором были получены идентификаторы вложений attachmentIds посредством вызова `getAttachmentsAsync` или `item.attachments`.</span><span class="sxs-lookup"><span data-stu-id="b789e-1023">As a best practice, you should use the identifier to retrieve an attachment in the same session that the attachmentIds were retrieved with the `getAttachmentsAsync` or `item.attachments` call.</span></span> <span data-ttu-id="b789e-1024">В Outlook Web App и Outlook Web App для устройств идентификатор вложения действителен только в рамках одного сеанса.</span><span class="sxs-lookup"><span data-stu-id="b789e-1024">In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="b789e-1025">Сеанс завершается, когда пользователь закрывает приложение или начинает создавать элемент во встроенной форме, а затем переходит из формы в отдельное окно.</span><span class="sxs-lookup"><span data-stu-id="b789e-1025">A session is over when the user closes the app, or if the user starts composing an inline form then subsequently pops out the form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b789e-1026">Параметры:</span><span class="sxs-lookup"><span data-stu-id="b789e-1026">Parameters:</span></span>

|<span data-ttu-id="b789e-1027">Имя</span><span class="sxs-lookup"><span data-stu-id="b789e-1027">Name</span></span>|<span data-ttu-id="b789e-1028">Тип</span><span class="sxs-lookup"><span data-stu-id="b789e-1028">Type</span></span>|<span data-ttu-id="b789e-1029">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="b789e-1029">Attributes</span></span>|<span data-ttu-id="b789e-1030">Описание</span><span class="sxs-lookup"><span data-stu-id="b789e-1030">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="b789e-1031">Строка</span><span class="sxs-lookup"><span data-stu-id="b789e-1031">String</span></span>||<span data-ttu-id="b789e-1032">Идентификатор вложения, который необходимо получить.</span><span class="sxs-lookup"><span data-stu-id="b789e-1032">The identifier of the attachment you want to get.</span></span>|
|`options`|<span data-ttu-id="b789e-1033">Object</span><span class="sxs-lookup"><span data-stu-id="b789e-1033">Object</span></span>|<span data-ttu-id="b789e-1034">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="b789e-1034">&lt;optional&gt;</span></span>|<span data-ttu-id="b789e-1035">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="b789e-1035">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="b789e-1036">Object</span><span class="sxs-lookup"><span data-stu-id="b789e-1036">Object</span></span>|<span data-ttu-id="b789e-1037">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="b789e-1037">&lt;optional&gt;</span></span>|<span data-ttu-id="b789e-1038">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="b789e-1038">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="b789e-1039">функция</span><span class="sxs-lookup"><span data-stu-id="b789e-1039">function</span></span>|<span data-ttu-id="b789e-1040">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="b789e-1040">&lt;optional&gt;</span></span>|<span data-ttu-id="b789e-1041">По завершении работы метода функция, переданная параметру `callback`, вызывается с помощью одного параметра `asyncResult`, представляющего собой объект [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="b789e-1041">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b789e-1042">Требования</span><span class="sxs-lookup"><span data-stu-id="b789e-1042">Requirements</span></span>

|<span data-ttu-id="b789e-1043">Требование</span><span class="sxs-lookup"><span data-stu-id="b789e-1043">Requirement</span></span>|<span data-ttu-id="b789e-1044">Значение</span><span class="sxs-lookup"><span data-stu-id="b789e-1044">Value</span></span>|
|---|---|
|[<span data-ttu-id="b789e-1045">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="b789e-1045">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="b789e-1046">Предварительная версия</span><span class="sxs-lookup"><span data-stu-id="b789e-1046">Preview</span></span>|
|[<span data-ttu-id="b789e-1047">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b789e-1047">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="b789e-1048">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b789e-1048">ReadItem</span></span>|
|[<span data-ttu-id="b789e-1049">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b789e-1049">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="b789e-1050">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="b789e-1050">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="b789e-1051">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="b789e-1051">Returns:</span></span>

<span data-ttu-id="b789e-1052">Тип: [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span><span class="sxs-lookup"><span data-stu-id="b789e-1052">Type: [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span></span>

##### <a name="example"></a><span data-ttu-id="b789e-1053">Пример</span><span class="sxs-lookup"><span data-stu-id="b789e-1053">Example</span></span>

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

#### <a name="getattachmentsasyncoptions-callback--arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails"></a><span data-ttu-id="b789e-1054">getAttachmentsAsync([options], callback) → Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="b789e-1054">getAttachmentsAsync([options], callback) → Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

<span data-ttu-id="b789e-1055">Получает вложения элемента в качестве массива.</span><span class="sxs-lookup"><span data-stu-id="b789e-1055">Gets the item's attachments as an array.</span></span> <span data-ttu-id="b789e-1056">Только в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="b789e-1056">Compose mode only.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b789e-1057">Параметры:</span><span class="sxs-lookup"><span data-stu-id="b789e-1057">Parameters:</span></span>

|<span data-ttu-id="b789e-1058">Имя</span><span class="sxs-lookup"><span data-stu-id="b789e-1058">Name</span></span>|<span data-ttu-id="b789e-1059">Тип</span><span class="sxs-lookup"><span data-stu-id="b789e-1059">Type</span></span>|<span data-ttu-id="b789e-1060">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="b789e-1060">Attributes</span></span>|<span data-ttu-id="b789e-1061">Описание</span><span class="sxs-lookup"><span data-stu-id="b789e-1061">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="b789e-1062">Объект</span><span class="sxs-lookup"><span data-stu-id="b789e-1062">Object</span></span>|<span data-ttu-id="b789e-1063">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="b789e-1063">&lt;optional&gt;</span></span>|<span data-ttu-id="b789e-1064">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="b789e-1064">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="b789e-1065">Object</span><span class="sxs-lookup"><span data-stu-id="b789e-1065">Object</span></span>|<span data-ttu-id="b789e-1066">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="b789e-1066">&lt;optional&gt;</span></span>|<span data-ttu-id="b789e-1067">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="b789e-1067">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="b789e-1068">функция</span><span class="sxs-lookup"><span data-stu-id="b789e-1068">function</span></span>|<span data-ttu-id="b789e-1069">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="b789e-1069">&lt;optional&gt;</span></span>|<span data-ttu-id="b789e-1070">По завершении работы метода функция, переданная параметру `callback`, вызывается с помощью одного параметра `asyncResult`, представляющего собой объект [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="b789e-1070">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b789e-1071">Требования</span><span class="sxs-lookup"><span data-stu-id="b789e-1071">Requirements</span></span>

|<span data-ttu-id="b789e-1072">Требование</span><span class="sxs-lookup"><span data-stu-id="b789e-1072">Requirement</span></span>|<span data-ttu-id="b789e-1073">Значение</span><span class="sxs-lookup"><span data-stu-id="b789e-1073">Value</span></span>|
|---|---|
|[<span data-ttu-id="b789e-1074">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="b789e-1074">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="b789e-1075">Предварительная версия</span><span class="sxs-lookup"><span data-stu-id="b789e-1075">Preview</span></span>|
|[<span data-ttu-id="b789e-1076">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b789e-1076">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="b789e-1077">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b789e-1077">ReadItem</span></span>|
|[<span data-ttu-id="b789e-1078">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b789e-1078">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="b789e-1079">Создание</span><span class="sxs-lookup"><span data-stu-id="b789e-1079">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="b789e-1080">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="b789e-1080">Returns:</span></span>

<span data-ttu-id="b789e-1081">Тип: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="b789e-1081">Type: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

##### <a name="example"></a><span data-ttu-id="b789e-1082">Пример</span><span class="sxs-lookup"><span data-stu-id="b789e-1082">Example</span></span>

<span data-ttu-id="b789e-1083">В приведенном ниже примере создается HTML-строка с подробными сведениями обо всех вложениях для текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="b789e-1083">The following example builds an HTML string with details of all attachments on the current item.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlookofficeentities"></a><span data-ttu-id="b789e-1084">getEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="b789e-1084">getEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span></span>

<span data-ttu-id="b789e-1085">Получает сущности, обнаруженные в теле выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="b789e-1085">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="b789e-1086">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="b789e-1086">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="b789e-1087">Требования</span><span class="sxs-lookup"><span data-stu-id="b789e-1087">Requirements</span></span>

|<span data-ttu-id="b789e-1088">Требование</span><span class="sxs-lookup"><span data-stu-id="b789e-1088">Requirement</span></span>|<span data-ttu-id="b789e-1089">Значение</span><span class="sxs-lookup"><span data-stu-id="b789e-1089">Value</span></span>|
|---|---|
|[<span data-ttu-id="b789e-1090">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="b789e-1090">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="b789e-1091">1.0</span><span class="sxs-lookup"><span data-stu-id="b789e-1091">1.0</span></span>|
|[<span data-ttu-id="b789e-1092">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b789e-1092">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="b789e-1093">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b789e-1093">ReadItem</span></span>|
|[<span data-ttu-id="b789e-1094">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b789e-1094">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="b789e-1095">Чтение</span><span class="sxs-lookup"><span data-stu-id="b789e-1095">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="b789e-1096">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="b789e-1096">Returns:</span></span>

<span data-ttu-id="b789e-1097">Тип: [Entities](/javascript/api/outlook/office.entities)</span><span class="sxs-lookup"><span data-stu-id="b789e-1097">Type: [Entities](/javascript/api/outlook/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="b789e-1098">Пример</span><span class="sxs-lookup"><span data-stu-id="b789e-1098">Example</span></span>

<span data-ttu-id="b789e-1099">Ниже приведен пример получения доступа к сущностям контактов в теле текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="b789e-1099">The following example accesses the contacts entities in the current item's body.</span></span>

```javascript
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion"></a><span data-ttu-id="b789e-1100">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="b789e-1100">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span></span>

<span data-ttu-id="b789e-1101">Получает массив всех сущностей указанного типа, обнаруженных в теле выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="b789e-1101">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="b789e-1102">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="b789e-1102">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b789e-1103">Параметры</span><span class="sxs-lookup"><span data-stu-id="b789e-1103">Parameters:</span></span>

|<span data-ttu-id="b789e-1104">Имя</span><span class="sxs-lookup"><span data-stu-id="b789e-1104">Name</span></span>|<span data-ttu-id="b789e-1105">Тип</span><span class="sxs-lookup"><span data-stu-id="b789e-1105">Type</span></span>|<span data-ttu-id="b789e-1106">Описание</span><span class="sxs-lookup"><span data-stu-id="b789e-1106">Description</span></span>|
|---|---|---|
|`entityType`|[<span data-ttu-id="b789e-1107">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="b789e-1107">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook/office.mailboxenums.entitytype)|<span data-ttu-id="b789e-1108">Одно из значений перечисления EntityType.</span><span class="sxs-lookup"><span data-stu-id="b789e-1108">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b789e-1109">Требования</span><span class="sxs-lookup"><span data-stu-id="b789e-1109">Requirements</span></span>

|<span data-ttu-id="b789e-1110">Требование</span><span class="sxs-lookup"><span data-stu-id="b789e-1110">Requirement</span></span>|<span data-ttu-id="b789e-1111">Значение</span><span class="sxs-lookup"><span data-stu-id="b789e-1111">Value</span></span>|
|---|---|
|[<span data-ttu-id="b789e-1112">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="b789e-1112">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="b789e-1113">1.0</span><span class="sxs-lookup"><span data-stu-id="b789e-1113">1.0</span></span>|
|[<span data-ttu-id="b789e-1114">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b789e-1114">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="b789e-1115">Restricted</span><span class="sxs-lookup"><span data-stu-id="b789e-1115">Restricted</span></span>|
|[<span data-ttu-id="b789e-1116">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b789e-1116">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="b789e-1117">Чтение</span><span class="sxs-lookup"><span data-stu-id="b789e-1117">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="b789e-1118">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="b789e-1118">Returns:</span></span>

<span data-ttu-id="b789e-1119">Если значение, переданное в `entityType`, не является допустимым членом перечисления `EntityType`, метод возвращает значение NULL.</span><span class="sxs-lookup"><span data-stu-id="b789e-1119">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="b789e-1120">Если в теле элемента отсутствуют сущности указанного типа, метод возвращает пустой массив.</span><span class="sxs-lookup"><span data-stu-id="b789e-1120">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="b789e-1121">В противном случае тип объектов в возвращаемом массиве зависит от типа сущности, запрошенной в параметре `entityType`.</span><span class="sxs-lookup"><span data-stu-id="b789e-1121">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="b789e-1122">Хотя минимальный уровень разрешений для использования этого метода — **Restricted**, для некоторых типов сущностей требуется доступ на уровне **ReadItem**, как указано в приведенной ниже таблице.</span><span class="sxs-lookup"><span data-stu-id="b789e-1122">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

|<span data-ttu-id="b789e-1123">Значение параметра `entityType`</span><span class="sxs-lookup"><span data-stu-id="b789e-1123">Value of `entityType`</span></span>|<span data-ttu-id="b789e-1124">Тип объектов в возвращаемом массиве</span><span class="sxs-lookup"><span data-stu-id="b789e-1124">Type of objects in returned array</span></span>|<span data-ttu-id="b789e-1125">Необходимый уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b789e-1125">Required Permission Level</span></span>|
|---|---|---|
|`Address`|<span data-ttu-id="b789e-1126">Строка</span><span class="sxs-lookup"><span data-stu-id="b789e-1126">String</span></span>|<span data-ttu-id="b789e-1127">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="b789e-1127">**Restricted**</span></span>|
|`Contact`|<span data-ttu-id="b789e-1128">Contact</span><span class="sxs-lookup"><span data-stu-id="b789e-1128">Contact</span></span>|<span data-ttu-id="b789e-1129">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="b789e-1129">**ReadItem**</span></span>|
|`EmailAddress`|<span data-ttu-id="b789e-1130">String</span><span class="sxs-lookup"><span data-stu-id="b789e-1130">String</span></span>|<span data-ttu-id="b789e-1131">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="b789e-1131">**ReadItem**</span></span>|
|`MeetingSuggestion`|<span data-ttu-id="b789e-1132">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="b789e-1132">MeetingSuggestion</span></span>|<span data-ttu-id="b789e-1133">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="b789e-1133">**ReadItem**</span></span>|
|`PhoneNumber`|<span data-ttu-id="b789e-1134">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="b789e-1134">PhoneNumber</span></span>|<span data-ttu-id="b789e-1135">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="b789e-1135">**Restricted**</span></span>|
|`TaskSuggestion`|<span data-ttu-id="b789e-1136">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="b789e-1136">TaskSuggestion</span></span>|<span data-ttu-id="b789e-1137">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="b789e-1137">**ReadItem**</span></span>|
|`URL`|<span data-ttu-id="b789e-1138">String</span><span class="sxs-lookup"><span data-stu-id="b789e-1138">String</span></span>|<span data-ttu-id="b789e-1139">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="b789e-1139">**Restricted**</span></span>|

<span data-ttu-id="b789e-1140">Тип: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="b789e-1140">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="b789e-1141">Пример</span><span class="sxs-lookup"><span data-stu-id="b789e-1141">Example</span></span>

<span data-ttu-id="b789e-1142">В примере ниже показано, как получить доступ к массиву строк, которые представляют собой почтовые адреса в теле текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="b789e-1142">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion"></a><span data-ttu-id="b789e-1143">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="b789e-1143">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span></span>

<span data-ttu-id="b789e-1144">Возвращает известные сущности в выбранном элементе, которые проходят через именованный фильтр, определяемый в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="b789e-1144">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="b789e-1145">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="b789e-1145">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="b789e-1146">Метод `getFilteredEntitiesByName` возвращает сущности, соответствующие регулярному выражению, которое определяется в элементе правила [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) в XML-файле манифеста, с использованием указанного значения элемента `FilterName`.</span><span class="sxs-lookup"><span data-stu-id="b789e-1146">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b789e-1147">Параметры:</span><span class="sxs-lookup"><span data-stu-id="b789e-1147">Parameters:</span></span>

|<span data-ttu-id="b789e-1148">Имя</span><span class="sxs-lookup"><span data-stu-id="b789e-1148">Name</span></span>|<span data-ttu-id="b789e-1149">Тип</span><span class="sxs-lookup"><span data-stu-id="b789e-1149">Type</span></span>|<span data-ttu-id="b789e-1150">Описание</span><span class="sxs-lookup"><span data-stu-id="b789e-1150">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="b789e-1151">Строка</span><span class="sxs-lookup"><span data-stu-id="b789e-1151">String</span></span>|<span data-ttu-id="b789e-1152">Имя элемента правила `ItemHasKnownEntity`, определяющее соответствующий фильтр.</span><span class="sxs-lookup"><span data-stu-id="b789e-1152">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b789e-1153">Требования</span><span class="sxs-lookup"><span data-stu-id="b789e-1153">Requirements</span></span>

|<span data-ttu-id="b789e-1154">Требование</span><span class="sxs-lookup"><span data-stu-id="b789e-1154">Requirement</span></span>|<span data-ttu-id="b789e-1155">Значение</span><span class="sxs-lookup"><span data-stu-id="b789e-1155">Value</span></span>|
|---|---|
|[<span data-ttu-id="b789e-1156">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="b789e-1156">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="b789e-1157">1.0</span><span class="sxs-lookup"><span data-stu-id="b789e-1157">1.0</span></span>|
|[<span data-ttu-id="b789e-1158">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b789e-1158">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="b789e-1159">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b789e-1159">ReadItem</span></span>|
|[<span data-ttu-id="b789e-1160">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b789e-1160">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="b789e-1161">Чтение</span><span class="sxs-lookup"><span data-stu-id="b789e-1161">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="b789e-1162">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="b789e-1162">Returns:</span></span>

<span data-ttu-id="b789e-p162">Если в манифесте нет элемента `ItemHasKnownEntity` со значением `FilterName`, соответствующим параметру `name`, метод возвращает `null`. Если параметр `name` соответствует элементу `ItemHasKnownEntity` в манифесте, но при этом в текущем элементе нет соответствующих сущностей, метод возвращает пустой массив.</span><span class="sxs-lookup"><span data-stu-id="b789e-p162">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="b789e-1165">Тип: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="b789e-1165">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span></span>

#### <a name="getinitializationcontextasyncoptions-callback"></a><span data-ttu-id="b789e-1166">getInitializationContextAsync([options], [callback])</span><span class="sxs-lookup"><span data-stu-id="b789e-1166">getInitializationContextAsync([options], [callback])</span></span>

<span data-ttu-id="b789e-1167">Получает данные инициализации, передаваемые при [активации надстройки интерактивным сообщением](https://docs.microsoft.com/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span><span class="sxs-lookup"><span data-stu-id="b789e-1167">Gets initialization data passed when the add-in is [activated by an actionable message](https://docs.microsoft.com/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span></span>

> [!NOTE]
> <span data-ttu-id="b789e-1168">Этот метод поддерживается только версией Outlook 2016 для Windows или более поздней (версии "нажми и работай" с номером больше 16.0.8413.1000) и Outlook в Интернете для Office 365.</span><span class="sxs-lookup"><span data-stu-id="b789e-1168">This method is only supported by Outlook 2016 or later for Windows (Click-to-Run versions later than 16.0.8413.1000) and Outlook on the web for Office 365.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b789e-1169">Параметры:</span><span class="sxs-lookup"><span data-stu-id="b789e-1169">Parameters:</span></span>
|<span data-ttu-id="b789e-1170">Имя</span><span class="sxs-lookup"><span data-stu-id="b789e-1170">Name</span></span>|<span data-ttu-id="b789e-1171">Тип</span><span class="sxs-lookup"><span data-stu-id="b789e-1171">Type</span></span>|<span data-ttu-id="b789e-1172">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="b789e-1172">Attributes</span></span>|<span data-ttu-id="b789e-1173">Описание</span><span class="sxs-lookup"><span data-stu-id="b789e-1173">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="b789e-1174">Объект</span><span class="sxs-lookup"><span data-stu-id="b789e-1174">Object</span></span>|<span data-ttu-id="b789e-1175">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="b789e-1175">&lt;optional&gt;</span></span>|<span data-ttu-id="b789e-1176">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="b789e-1176">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="b789e-1177">Object</span><span class="sxs-lookup"><span data-stu-id="b789e-1177">Object</span></span>|<span data-ttu-id="b789e-1178">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="b789e-1178">&lt;optional&gt;</span></span>|<span data-ttu-id="b789e-1179">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="b789e-1179">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="b789e-1180">функция</span><span class="sxs-lookup"><span data-stu-id="b789e-1180">function</span></span>|<span data-ttu-id="b789e-1181">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="b789e-1181">&lt;optional&gt;</span></span>|<span data-ttu-id="b789e-1182">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="b789e-1182">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="b789e-1183">В случае успешного выполнения данные инициализации предоставляются в свойстве `asyncResult.value` как строка.</span><span class="sxs-lookup"><span data-stu-id="b789e-1183">On success, the initialization data is provided in the `asyncResult.value` property as a string.</span></span><br/><span data-ttu-id="b789e-1184">Если контекст инициализации отсутствует, объект `asyncResult` будет содержать объект `Error`, одному свойству которого (`code`) будет присвоено значение `9020`, а другому (`name`) — значение `GenericResponseError`.</span><span class="sxs-lookup"><span data-stu-id="b789e-1184">If there is no initialization context, the `asyncResult` object will contain an `Error` object with its `code` property set to `9020` and its `name` property set to `GenericResponseError`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b789e-1185">Требования</span><span class="sxs-lookup"><span data-stu-id="b789e-1185">Requirements</span></span>

|<span data-ttu-id="b789e-1186">Требование</span><span class="sxs-lookup"><span data-stu-id="b789e-1186">Requirement</span></span>|<span data-ttu-id="b789e-1187">Значение</span><span class="sxs-lookup"><span data-stu-id="b789e-1187">Value</span></span>|
|---|---|
|[<span data-ttu-id="b789e-1188">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="b789e-1188">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="b789e-1189">Предварительная версия</span><span class="sxs-lookup"><span data-stu-id="b789e-1189">Preview</span></span>|
|[<span data-ttu-id="b789e-1190">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b789e-1190">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="b789e-1191">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b789e-1191">ReadItem</span></span>|
|[<span data-ttu-id="b789e-1192">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b789e-1192">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="b789e-1193">Чтение</span><span class="sxs-lookup"><span data-stu-id="b789e-1193">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b789e-1194">Пример</span><span class="sxs-lookup"><span data-stu-id="b789e-1194">Example</span></span>

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

#### <a name="getregexmatches--object"></a><span data-ttu-id="b789e-1195">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="b789e-1195">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="b789e-1196">Возвращает строковые значения в выбранном элементе, которые соответствуют регулярным выражениям, определенным в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="b789e-1196">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="b789e-1197">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="b789e-1197">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="b789e-p163">Метод `getRegExMatches` возвращает строки, соответствующие регулярному выражению, которое определяется в каждом элементе правила `ItemHasRegularExpressionMatch` или `ItemHasKnownEntity` в XML-файле манифеста. Для правила `ItemHasRegularExpressionMatch` соответствующую строку должно содержать свойство элемента, указанного этим правилом. Простой тип `PropertyName` определяет поддерживаемые свойства.</span><span class="sxs-lookup"><span data-stu-id="b789e-p163">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="b789e-1201">Например, рассмотрим манифест надстройки, который содержит указанный ниже элемент `Rule`.</span><span class="sxs-lookup"><span data-stu-id="b789e-1201">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="b789e-1202">Объект, возвращаемый методом `getRegExMatches`, будет содержать два свойства: `fruits` и `veggies`.</span><span class="sxs-lookup"><span data-stu-id="b789e-1202">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="b789e-p164">Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты. Лучше используйте метод [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) для этого.</span><span class="sxs-lookup"><span data-stu-id="b789e-p164">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="b789e-1206">Requirements</span><span class="sxs-lookup"><span data-stu-id="b789e-1206">Requirements</span></span>

|<span data-ttu-id="b789e-1207">Требование</span><span class="sxs-lookup"><span data-stu-id="b789e-1207">Requirement</span></span>|<span data-ttu-id="b789e-1208">Значение</span><span class="sxs-lookup"><span data-stu-id="b789e-1208">Value</span></span>|
|---|---|
|[<span data-ttu-id="b789e-1209">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="b789e-1209">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="b789e-1210">1.0</span><span class="sxs-lookup"><span data-stu-id="b789e-1210">1.0</span></span>|
|[<span data-ttu-id="b789e-1211">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b789e-1211">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="b789e-1212">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b789e-1212">ReadItem</span></span>|
|[<span data-ttu-id="b789e-1213">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b789e-1213">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="b789e-1214">Чтение</span><span class="sxs-lookup"><span data-stu-id="b789e-1214">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="b789e-1215">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="b789e-1215">Returns:</span></span>

<span data-ttu-id="b789e-p165">Объект, содержащий массив строк, которые соответствуют регулярным выражениям, определяемым в XML-файле манифеста. Имя каждого массива равно соответствующему значению атрибута `RegExName` подходящего правила `ItemHasRegularExpressionMatch` или атрибута `FilterName` соответствующего правила `ItemHasKnownEntity`.</span><span class="sxs-lookup"><span data-stu-id="b789e-p165">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="b789e-1218">

<dt>Тип</dt>

</span><span class="sxs-lookup"><span data-stu-id="b789e-1218">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="b789e-1219">Object</span><span class="sxs-lookup"><span data-stu-id="b789e-1219">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="b789e-1220">Пример</span><span class="sxs-lookup"><span data-stu-id="b789e-1220">Example</span></span>

<span data-ttu-id="b789e-1221">В приведенном ниже примере показано, как получить доступ к массиву совпадений с элементами `fruits` и `veggies` правил активации регулярных выражений, указанными в манифесте.</span><span class="sxs-lookup"><span data-stu-id="b789e-1221">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```javascript
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veges = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="b789e-1222">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="b789e-1222">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="b789e-1223">Возвращает строковые значения в выбранном элементе, которые соответствуют именованному регулярному выражению, определенному в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="b789e-1223">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="b789e-1224">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="b789e-1224">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="b789e-1225">Метод `getRegExMatchesByName` возвращает строки, соответствующие регулярному выражению, которое определяется в элементе правила `ItemHasRegularExpressionMatch` в XML-файле манифеста, с использованием указанного значения элемента `RegExName`.</span><span class="sxs-lookup"><span data-stu-id="b789e-1225">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="b789e-p166">Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты.</span><span class="sxs-lookup"><span data-stu-id="b789e-p166">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b789e-1228">Параметры:</span><span class="sxs-lookup"><span data-stu-id="b789e-1228">Parameters:</span></span>

|<span data-ttu-id="b789e-1229">Имя</span><span class="sxs-lookup"><span data-stu-id="b789e-1229">Name</span></span>|<span data-ttu-id="b789e-1230">Тип</span><span class="sxs-lookup"><span data-stu-id="b789e-1230">Type</span></span>|<span data-ttu-id="b789e-1231">Описание</span><span class="sxs-lookup"><span data-stu-id="b789e-1231">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="b789e-1232">String</span><span class="sxs-lookup"><span data-stu-id="b789e-1232">String</span></span>|<span data-ttu-id="b789e-1233">Имя элемента правила `ItemHasRegularExpressionMatch`, определяющее соответствующий фильтр.</span><span class="sxs-lookup"><span data-stu-id="b789e-1233">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b789e-1234">Требования</span><span class="sxs-lookup"><span data-stu-id="b789e-1234">Requirements</span></span>

|<span data-ttu-id="b789e-1235">Требование</span><span class="sxs-lookup"><span data-stu-id="b789e-1235">Requirement</span></span>|<span data-ttu-id="b789e-1236">Значение</span><span class="sxs-lookup"><span data-stu-id="b789e-1236">Value</span></span>|
|---|---|
|[<span data-ttu-id="b789e-1237">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="b789e-1237">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="b789e-1238">1.0</span><span class="sxs-lookup"><span data-stu-id="b789e-1238">1.0</span></span>|
|[<span data-ttu-id="b789e-1239">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b789e-1239">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="b789e-1240">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b789e-1240">ReadItem</span></span>|
|[<span data-ttu-id="b789e-1241">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b789e-1241">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="b789e-1242">Чтение</span><span class="sxs-lookup"><span data-stu-id="b789e-1242">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="b789e-1243">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="b789e-1243">Returns:</span></span>

<span data-ttu-id="b789e-1244">Массив строк, соответствующих регулярному выражению, определяемому в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="b789e-1244">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="b789e-1245">

<dt>Тип</dt>

</span><span class="sxs-lookup"><span data-stu-id="b789e-1245">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="b789e-1246">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="b789e-1246">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="b789e-1247">Пример</span><span class="sxs-lookup"><span data-stu-id="b789e-1247">Example</span></span>

```javascript
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="b789e-1248">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="b789e-1248">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="b789e-1249">Асинхронно возвращает данные, выбранные в теме или тексте сообщения.</span><span class="sxs-lookup"><span data-stu-id="b789e-1249">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="b789e-p167">Если выделенный фрагмент отсутствует, но курсор находится в тексте или теме, метод возвращает значение NULL для выбранных данных. Если выбраны не текст и не тема, метод возвращает ошибку `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="b789e-p167">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b789e-1252">Параметры</span><span class="sxs-lookup"><span data-stu-id="b789e-1252">Parameters:</span></span>

|<span data-ttu-id="b789e-1253">Имя</span><span class="sxs-lookup"><span data-stu-id="b789e-1253">Name</span></span>|<span data-ttu-id="b789e-1254">Тип</span><span class="sxs-lookup"><span data-stu-id="b789e-1254">Type</span></span>|<span data-ttu-id="b789e-1255">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="b789e-1255">Attributes</span></span>|<span data-ttu-id="b789e-1256">Описание</span><span class="sxs-lookup"><span data-stu-id="b789e-1256">Description</span></span>|
|---|---|---|---|
|`coercionType`|[<span data-ttu-id="b789e-1257">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="b789e-1257">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="b789e-p168">Запрашивает формат данных. Если задано значение Text, метод возвращает обычный текст как строку, удаляя все имеющиеся HTML-теги. Если задано значение HTML, метод возвращает выделенный текст (обычный текст или HTML).</span><span class="sxs-lookup"><span data-stu-id="b789e-p168">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`|<span data-ttu-id="b789e-1261">Object</span><span class="sxs-lookup"><span data-stu-id="b789e-1261">Object</span></span>|<span data-ttu-id="b789e-1262">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="b789e-1262">&lt;optional&gt;</span></span>|<span data-ttu-id="b789e-1263">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="b789e-1263">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="b789e-1264">Object</span><span class="sxs-lookup"><span data-stu-id="b789e-1264">Object</span></span>|<span data-ttu-id="b789e-1265">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="b789e-1265">&lt;optional&gt;</span></span>|<span data-ttu-id="b789e-1266">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="b789e-1266">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="b789e-1267">функция</span><span class="sxs-lookup"><span data-stu-id="b789e-1267">function</span></span>||<span data-ttu-id="b789e-1268">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="b789e-1268">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="b789e-1269">Чтобы получить доступ к выбранным данным из метода обратного вызова, вызовите `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="b789e-1269">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="b789e-1270">Для доступа к исходному свойству, представляющему собой источник выбранных данных, вызовите параметр `asyncResult.value.sourceProperty`, который может иметь значение `body` или `subject`.</span><span class="sxs-lookup"><span data-stu-id="b789e-1270">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b789e-1271">Требования</span><span class="sxs-lookup"><span data-stu-id="b789e-1271">Requirements</span></span>

|<span data-ttu-id="b789e-1272">Требование</span><span class="sxs-lookup"><span data-stu-id="b789e-1272">Requirement</span></span>|<span data-ttu-id="b789e-1273">Значение</span><span class="sxs-lookup"><span data-stu-id="b789e-1273">Value</span></span>|
|---|---|
|[<span data-ttu-id="b789e-1274">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="b789e-1274">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="b789e-1275">1.2</span><span class="sxs-lookup"><span data-stu-id="b789e-1275">1.2</span></span>|
|[<span data-ttu-id="b789e-1276">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b789e-1276">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="b789e-1277">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="b789e-1277">ReadWriteItem</span></span>|
|[<span data-ttu-id="b789e-1278">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b789e-1278">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="b789e-1279">Создание</span><span class="sxs-lookup"><span data-stu-id="b789e-1279">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="b789e-1280">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="b789e-1280">Returns:</span></span>

<span data-ttu-id="b789e-1281">Выбранные данные в виде строки с форматом, определенным в параметре `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="b789e-1281">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="b789e-1282">

<dt>Тип</dt>

</span><span class="sxs-lookup"><span data-stu-id="b789e-1282">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="b789e-1283">String</span><span class="sxs-lookup"><span data-stu-id="b789e-1283">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="b789e-1284">Пример</span><span class="sxs-lookup"><span data-stu-id="b789e-1284">Example</span></span>

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

#### <a name="getselectedentities--entitiesjavascriptapioutlookofficeentities"></a><span data-ttu-id="b789e-1285">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="b789e-1285">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span></span>

<span data-ttu-id="b789e-p170">Возвращает сущности, найденные в выделенном совпадении, выбранном пользователем. Выделенные совпадения применяются к [контекстным надстройкам](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="b789e-p170">Gets the entities found in a highlighted match a user has selected. Highlighted matches apply to [contextual add-ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="b789e-1288">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="b789e-1288">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="b789e-1289">Требования</span><span class="sxs-lookup"><span data-stu-id="b789e-1289">Requirements</span></span>

|<span data-ttu-id="b789e-1290">Требование</span><span class="sxs-lookup"><span data-stu-id="b789e-1290">Requirement</span></span>|<span data-ttu-id="b789e-1291">Значение</span><span class="sxs-lookup"><span data-stu-id="b789e-1291">Value</span></span>|
|---|---|
|[<span data-ttu-id="b789e-1292">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="b789e-1292">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="b789e-1293">1.6</span><span class="sxs-lookup"><span data-stu-id="b789e-1293">1.6</span></span>|
|[<span data-ttu-id="b789e-1294">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b789e-1294">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="b789e-1295">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b789e-1295">ReadItem</span></span>|
|[<span data-ttu-id="b789e-1296">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b789e-1296">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="b789e-1297">Чтение</span><span class="sxs-lookup"><span data-stu-id="b789e-1297">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="b789e-1298">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="b789e-1298">Returns:</span></span>

<span data-ttu-id="b789e-1299">Тип: [Entities](/javascript/api/outlook/office.entities)</span><span class="sxs-lookup"><span data-stu-id="b789e-1299">Type: [Entities](/javascript/api/outlook/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="b789e-1300">Пример</span><span class="sxs-lookup"><span data-stu-id="b789e-1300">Example</span></span>

<span data-ttu-id="b789e-1301">В приведенном ниже примере показано, как получить доступ к сущностям адресов в выделенном совпадении, выбранном пользователем.</span><span class="sxs-lookup"><span data-stu-id="b789e-1301">The following example accesses the addresses entities in the highlighted match selected by the user.</span></span>

```javascript
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

#### <a name="getselectedregexmatches--object"></a><span data-ttu-id="b789e-1302">getSelectedRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="b789e-1302">getSelectedRegExMatches() → {Object}</span></span>

<span data-ttu-id="b789e-p171">Возвращает строковые значения в выделенном совпадении, которые соответствуют регулярным выражениям, определенным в XML-файле манифеста. Выделенные совпадения применяются к [контекстным надстройкам](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="b789e-p171">Returns string values in a highlighted match that match the regular expressions defined in the manifest XML file. Highlighted matches apply to [contextual add-ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="b789e-1305">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="b789e-1305">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="b789e-p172">Метод `getSelectedRegExMatches` возвращает строки, соответствующие регулярному выражению, которое определяется в каждом элементе правила `ItemHasRegularExpressionMatch` или `ItemHasKnownEntity` в XML-файле манифеста. Для правила `ItemHasRegularExpressionMatch` соответствующую строку должно содержать свойство элемента, указанного этим правилом. Простой тип `PropertyName` определяет поддерживаемые свойства.</span><span class="sxs-lookup"><span data-stu-id="b789e-p172">The `getSelectedRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="b789e-1309">Например, рассмотрим манифест надстройки, который содержит указанный ниже элемент `Rule`.</span><span class="sxs-lookup"><span data-stu-id="b789e-1309">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="b789e-1310">Объект, возвращаемый методом `getRegExMatches`, будет содержать два свойства: `fruits` и `veggies`.</span><span class="sxs-lookup"><span data-stu-id="b789e-1310">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="b789e-p173">Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты. Лучше используйте метод [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) для этого.</span><span class="sxs-lookup"><span data-stu-id="b789e-p173">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="b789e-1314">Requirements</span><span class="sxs-lookup"><span data-stu-id="b789e-1314">Requirements</span></span>

|<span data-ttu-id="b789e-1315">Требование</span><span class="sxs-lookup"><span data-stu-id="b789e-1315">Requirement</span></span>|<span data-ttu-id="b789e-1316">Значение</span><span class="sxs-lookup"><span data-stu-id="b789e-1316">Value</span></span>|
|---|---|
|[<span data-ttu-id="b789e-1317">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="b789e-1317">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="b789e-1318">1.6</span><span class="sxs-lookup"><span data-stu-id="b789e-1318">1.6</span></span>|
|[<span data-ttu-id="b789e-1319">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b789e-1319">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="b789e-1320">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b789e-1320">ReadItem</span></span>|
|[<span data-ttu-id="b789e-1321">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b789e-1321">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="b789e-1322">Чтение</span><span class="sxs-lookup"><span data-stu-id="b789e-1322">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="b789e-1323">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="b789e-1323">Returns:</span></span>

<span data-ttu-id="b789e-p174">Объект, содержащий массив строк, которые соответствуют регулярным выражениям, определяемым в XML-файле манифеста. Имя каждого массива равно соответствующему значению атрибута `RegExName` подходящего правила `ItemHasRegularExpressionMatch` или атрибута `FilterName` соответствующего правила `ItemHasKnownEntity`.</span><span class="sxs-lookup"><span data-stu-id="b789e-p174">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

##### <a name="example"></a><span data-ttu-id="b789e-1326">Пример</span><span class="sxs-lookup"><span data-stu-id="b789e-1326">Example</span></span>

<span data-ttu-id="b789e-1327">В приведенном ниже примере показано, как получить доступ к массиву совпадений с элементами `fruits` и `veggies` правил активации регулярных выражений, указанными в манифесте.</span><span class="sxs-lookup"><span data-stu-id="b789e-1327">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```javascript
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

#### <a name="getsharedpropertiesasyncoptions-callback"></a><span data-ttu-id="b789e-1328">getSharedPropertiesAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="b789e-1328">getSharedPropertiesAsync([options], callback)</span></span>

<span data-ttu-id="b789e-1329">Получает свойства выбранного встречи или сообщения в общей папке, календаре или почтовом ящике.</span><span class="sxs-lookup"><span data-stu-id="b789e-1329">Gets the properties of the selected appointment or message in a shared folder, calendar, or mailbox.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b789e-1330">Параметры:</span><span class="sxs-lookup"><span data-stu-id="b789e-1330">Parameters:</span></span>

|<span data-ttu-id="b789e-1331">Имя</span><span class="sxs-lookup"><span data-stu-id="b789e-1331">Name</span></span>|<span data-ttu-id="b789e-1332">Тип</span><span class="sxs-lookup"><span data-stu-id="b789e-1332">Type</span></span>|<span data-ttu-id="b789e-1333">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="b789e-1333">Attributes</span></span>|<span data-ttu-id="b789e-1334">Описание</span><span class="sxs-lookup"><span data-stu-id="b789e-1334">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="b789e-1335">Object</span><span class="sxs-lookup"><span data-stu-id="b789e-1335">Object</span></span>|<span data-ttu-id="b789e-1336">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="b789e-1336">&lt;optional&gt;</span></span>|<span data-ttu-id="b789e-1337">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="b789e-1337">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="b789e-1338">Object</span><span class="sxs-lookup"><span data-stu-id="b789e-1338">Object</span></span>|<span data-ttu-id="b789e-1339">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="b789e-1339">&lt;optional&gt;</span></span>|<span data-ttu-id="b789e-1340">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="b789e-1340">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="b789e-1341">функция</span><span class="sxs-lookup"><span data-stu-id="b789e-1341">function</span></span>||<span data-ttu-id="b789e-1342">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="b789e-1342">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="b789e-1343">Общие свойства предоставляются в виде объекта [`SharedProperties`](/javascript/api/outlook/office.sharedproperties) в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="b789e-1343">The shared properties are provided as a [`SharedProperties`](/javascript/api/outlook/office.sharedproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="b789e-1344">Этот объект можно использовать для получения общих свойств элемента.</span><span class="sxs-lookup"><span data-stu-id="b789e-1344">This object can be used to get the item's shared properties.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b789e-1345">Требования</span><span class="sxs-lookup"><span data-stu-id="b789e-1345">Requirements</span></span>

|<span data-ttu-id="b789e-1346">Требование</span><span class="sxs-lookup"><span data-stu-id="b789e-1346">Requirement</span></span>|<span data-ttu-id="b789e-1347">Значение</span><span class="sxs-lookup"><span data-stu-id="b789e-1347">Value</span></span>|
|---|---|
|[<span data-ttu-id="b789e-1348">Минимальная версия набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="b789e-1348">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="b789e-1349">Предварительная версия</span><span class="sxs-lookup"><span data-stu-id="b789e-1349">Preview</span></span>|
|[<span data-ttu-id="b789e-1350">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b789e-1350">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="b789e-1351">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b789e-1351">ReadItem</span></span>|
|[<span data-ttu-id="b789e-1352">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b789e-1352">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="b789e-1353">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="b789e-1353">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="b789e-1354">Пример</span><span class="sxs-lookup"><span data-stu-id="b789e-1354">Example</span></span>

```js
Office.context.mailbox.item.getSharedPropertiesAsync(callback);
function callback (asyncResult) {
  var context=asyncResult.context;
  var sharedProperties = asyncResult.value;
}
```

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="b789e-1355">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="b789e-1355">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="b789e-1356">Асинхронно загружает настраиваемые свойства для надстройки для выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="b789e-1356">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="b789e-p176">Настраиваемые свойства сохраняются в виде пар "ключ-значение" для каждого приложения и каждого элемента. Этот метод возвращает объект `CustomProperties` при обратном вызове, который предоставляет методы для доступа к настраиваемым свойствам, характерным для текущего элемента и текущей надстройки. Настраиваемые свойства не шифруются для элемента, поэтому этот способ хранения не является безопасным.</span><span class="sxs-lookup"><span data-stu-id="b789e-p176">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b789e-1360">Параметры</span><span class="sxs-lookup"><span data-stu-id="b789e-1360">Parameters:</span></span>

|<span data-ttu-id="b789e-1361">Имя</span><span class="sxs-lookup"><span data-stu-id="b789e-1361">Name</span></span>|<span data-ttu-id="b789e-1362">Тип</span><span class="sxs-lookup"><span data-stu-id="b789e-1362">Type</span></span>|<span data-ttu-id="b789e-1363">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="b789e-1363">Attributes</span></span>|<span data-ttu-id="b789e-1364">Описание</span><span class="sxs-lookup"><span data-stu-id="b789e-1364">Description</span></span>|
|---|---|---|---|
|`callback`|<span data-ttu-id="b789e-1365">функция</span><span class="sxs-lookup"><span data-stu-id="b789e-1365">function</span></span>||<span data-ttu-id="b789e-1366">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="b789e-1366">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="b789e-1367">Настраиваемые свойства предоставляются в виде объекта [`CustomProperties`](/javascript/api/outlook/office.customproperties) в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="b789e-1367">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="b789e-1368">Этот объект позволяет получить, задать и удалить настраиваемые свойства для элемента, а также сохранить изменения, внесенные в набор настраиваемых свойств, на сервере.</span><span class="sxs-lookup"><span data-stu-id="b789e-1368">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`|<span data-ttu-id="b789e-1369">Object</span><span class="sxs-lookup"><span data-stu-id="b789e-1369">Object</span></span>|<span data-ttu-id="b789e-1370">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="b789e-1370">&lt;optional&gt;</span></span>|<span data-ttu-id="b789e-1371">Разработчики могут указать любой объект, к которому необходимо получить доступ, в функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="b789e-1371">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="b789e-1372">Доступ к этому объекту можно получить с помощью свойства `asyncResult.asyncContext` в функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="b789e-1372">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b789e-1373">Требования</span><span class="sxs-lookup"><span data-stu-id="b789e-1373">Requirements</span></span>

|<span data-ttu-id="b789e-1374">Требование</span><span class="sxs-lookup"><span data-stu-id="b789e-1374">Requirement</span></span>|<span data-ttu-id="b789e-1375">Значение</span><span class="sxs-lookup"><span data-stu-id="b789e-1375">Value</span></span>|
|---|---|
|[<span data-ttu-id="b789e-1376">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="b789e-1376">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="b789e-1377">1.0</span><span class="sxs-lookup"><span data-stu-id="b789e-1377">1.0</span></span>|
|[<span data-ttu-id="b789e-1378">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b789e-1378">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="b789e-1379">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b789e-1379">ReadItem</span></span>|
|[<span data-ttu-id="b789e-1380">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b789e-1380">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="b789e-1381">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="b789e-1381">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="b789e-1382">Пример</span><span class="sxs-lookup"><span data-stu-id="b789e-1382">Example</span></span>

<span data-ttu-id="b789e-p179">Приведенный ниже пример кода показывает, как асинхронно загружать настраиваемые свойства, характерные для текущего элемента, с помощью метода `loadCustomPropertiesAsync`. Этот пример также показывает, как сохранять эти свойства на сервере с помощью метода `CustomProperties.saveAsync`. После загрузки настраиваемых свойств в этом примере кода метод `CustomProperties.get` используется для считывания настраиваемого свойства `myProp`, метод `CustomProperties.set` — для записи настраиваемого свойства `otherProp`, а метод `saveAsync` — для сохранения настраиваемых свойств.</span><span class="sxs-lookup"><span data-stu-id="b789e-p179">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="b789e-1386">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="b789e-1386">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="b789e-1387">Удаляет вложение из сообщения или встречи.</span><span class="sxs-lookup"><span data-stu-id="b789e-1387">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="b789e-1388">Метод `removeAttachmentAsync` удаляет из элемента вложение с указанным идентификатором.</span><span class="sxs-lookup"><span data-stu-id="b789e-1388">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="b789e-1389">Идентификатор вложения рекомендуется использовать для удаления вложения, только если оно добавлено тем же почтовым приложением в ходе текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="b789e-1389">As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session.</span></span> <span data-ttu-id="b789e-1390">В Outlook Web App и Outlook Web App для устройств идентификатор вложения действителен только в рамках одного сеанса.</span><span class="sxs-lookup"><span data-stu-id="b789e-1390">In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="b789e-1391">Сеанс завершается, когда пользователь закрывает приложение или начинает создавать элемент во встроенной форме, а затем переходит из формы в отдельное окно.</span><span class="sxs-lookup"><span data-stu-id="b789e-1391">A session is over when the user closes the app, or if the user starts composing an inline form then subsequently pops out the form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b789e-1392">Параметры:</span><span class="sxs-lookup"><span data-stu-id="b789e-1392">Parameters:</span></span>

|<span data-ttu-id="b789e-1393">Имя</span><span class="sxs-lookup"><span data-stu-id="b789e-1393">Name</span></span>|<span data-ttu-id="b789e-1394">Тип</span><span class="sxs-lookup"><span data-stu-id="b789e-1394">Type</span></span>|<span data-ttu-id="b789e-1395">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="b789e-1395">Attributes</span></span>|<span data-ttu-id="b789e-1396">Описание</span><span class="sxs-lookup"><span data-stu-id="b789e-1396">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="b789e-1397">String</span><span class="sxs-lookup"><span data-stu-id="b789e-1397">String</span></span>||<span data-ttu-id="b789e-1398">Идентификатор удаляемого вложения.</span><span class="sxs-lookup"><span data-stu-id="b789e-1398">The identifier of the attachment to remove.</span></span>|
|`options`|<span data-ttu-id="b789e-1399">Object</span><span class="sxs-lookup"><span data-stu-id="b789e-1399">Object</span></span>|<span data-ttu-id="b789e-1400">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="b789e-1400">&lt;optional&gt;</span></span>|<span data-ttu-id="b789e-1401">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="b789e-1401">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="b789e-1402">Object</span><span class="sxs-lookup"><span data-stu-id="b789e-1402">Object</span></span>|<span data-ttu-id="b789e-1403">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="b789e-1403">&lt;optional&gt;</span></span>|<span data-ttu-id="b789e-1404">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="b789e-1404">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="b789e-1405">функция</span><span class="sxs-lookup"><span data-stu-id="b789e-1405">function</span></span>|<span data-ttu-id="b789e-1406">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="b789e-1406">&lt;optional&gt;</span></span>|<span data-ttu-id="b789e-1407">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="b789e-1407">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="b789e-1408">Если удалить вложение не удается, свойство `asyncResult.error` содержит код ошибки с указанием ее причины.</span><span class="sxs-lookup"><span data-stu-id="b789e-1408">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="b789e-1409">Ошибки</span><span class="sxs-lookup"><span data-stu-id="b789e-1409">Errors</span></span>

|<span data-ttu-id="b789e-1410">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="b789e-1410">Error code</span></span>|<span data-ttu-id="b789e-1411">Описание</span><span class="sxs-lookup"><span data-stu-id="b789e-1411">Description</span></span>|
|------------|-------------|
|`InvalidAttachmentId`|<span data-ttu-id="b789e-1412">Идентификатор вложения не существует.</span><span class="sxs-lookup"><span data-stu-id="b789e-1412">The attachment identifier does not exist.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b789e-1413">Требования</span><span class="sxs-lookup"><span data-stu-id="b789e-1413">Requirements</span></span>

|<span data-ttu-id="b789e-1414">Требование</span><span class="sxs-lookup"><span data-stu-id="b789e-1414">Requirement</span></span>|<span data-ttu-id="b789e-1415">Значение</span><span class="sxs-lookup"><span data-stu-id="b789e-1415">Value</span></span>|
|---|---|
|[<span data-ttu-id="b789e-1416">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="b789e-1416">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="b789e-1417">1.1</span><span class="sxs-lookup"><span data-stu-id="b789e-1417">1.1</span></span>|
|[<span data-ttu-id="b789e-1418">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b789e-1418">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="b789e-1419">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="b789e-1419">ReadWriteItem</span></span>|
|[<span data-ttu-id="b789e-1420">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b789e-1420">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="b789e-1421">Создание</span><span class="sxs-lookup"><span data-stu-id="b789e-1421">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="b789e-1422">Пример</span><span class="sxs-lookup"><span data-stu-id="b789e-1422">Example</span></span>

<span data-ttu-id="b789e-1423">Указанный ниже код удаляет вложение с идентификатором "0".</span><span class="sxs-lookup"><span data-stu-id="b789e-1423">The following code removes an attachment with an identifier of '0'.</span></span>

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

####  <a name="removehandlerasynceventtype-options-callback"></a><span data-ttu-id="b789e-1424">removeHandlerAsync(eventType, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="b789e-1424">removeHandlerAsync(eventType, [options], [callback])</span></span>

<span data-ttu-id="b789e-1425">Удаляет обработчиков для поддерживаемого типа события.</span><span class="sxs-lookup"><span data-stu-id="b789e-1425">Removes the event handlers for a supported event type.</span></span>

<span data-ttu-id="b789e-1426">Сейчас поддерживаются следующие типы событий: `Office.EventType.AttachmentsChanged`, `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged` и `Office.EventType.RecurrenceChanged`.</span><span class="sxs-lookup"><span data-stu-id="b789e-1426">Currently the supported event types are `Office.EventType.AttachmentsChanged`, `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b789e-1427">Параметры:</span><span class="sxs-lookup"><span data-stu-id="b789e-1427">Parameters:</span></span>

| <span data-ttu-id="b789e-1428">Имя</span><span class="sxs-lookup"><span data-stu-id="b789e-1428">Name</span></span> | <span data-ttu-id="b789e-1429">Тип</span><span class="sxs-lookup"><span data-stu-id="b789e-1429">Type</span></span> | <span data-ttu-id="b789e-1430">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="b789e-1430">Attributes</span></span> | <span data-ttu-id="b789e-1431">Описание</span><span class="sxs-lookup"><span data-stu-id="b789e-1431">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="b789e-1432">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="b789e-1432">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="b789e-1433">Событие, которое должно отменить обработчик.</span><span class="sxs-lookup"><span data-stu-id="b789e-1433">The event that should revoke the handler.</span></span> |
| `options` | <span data-ttu-id="b789e-1434">Объект</span><span class="sxs-lookup"><span data-stu-id="b789e-1434">Object</span></span> | <span data-ttu-id="b789e-1435">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="b789e-1435">&lt;optional&gt;</span></span> | <span data-ttu-id="b789e-1436">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="b789e-1436">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="b789e-1437">Object</span><span class="sxs-lookup"><span data-stu-id="b789e-1437">Object</span></span> | <span data-ttu-id="b789e-1438">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="b789e-1438">&lt;optional&gt;</span></span> | <span data-ttu-id="b789e-1439">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="b789e-1439">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="b789e-1440">функция</span><span class="sxs-lookup"><span data-stu-id="b789e-1440">function</span></span>| <span data-ttu-id="b789e-1441">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="b789e-1441">&lt;optional&gt;</span></span>|<span data-ttu-id="b789e-1442">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="b789e-1442">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b789e-1443">Требования</span><span class="sxs-lookup"><span data-stu-id="b789e-1443">Requirements</span></span>

|<span data-ttu-id="b789e-1444">Требование</span><span class="sxs-lookup"><span data-stu-id="b789e-1444">Requirement</span></span>| <span data-ttu-id="b789e-1445">Значение</span><span class="sxs-lookup"><span data-stu-id="b789e-1445">Value</span></span>|
|---|---|
|[<span data-ttu-id="b789e-1446">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="b789e-1446">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b789e-1447">1.7</span><span class="sxs-lookup"><span data-stu-id="b789e-1447">1.7</span></span> |
|[<span data-ttu-id="b789e-1448">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b789e-1448">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b789e-1449">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b789e-1449">ReadItem</span></span> |
|[<span data-ttu-id="b789e-1450">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b789e-1450">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b789e-1451">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="b789e-1451">Compose or read</span></span> |

####  <a name="saveasyncoptions-callback"></a><span data-ttu-id="b789e-1452">saveAsync([options], обратный вызов)</span><span class="sxs-lookup"><span data-stu-id="b789e-1452">saveAsync([options], callback)</span></span>

<span data-ttu-id="b789e-1453">Асинхронно сохраняет элемент.</span><span class="sxs-lookup"><span data-stu-id="b789e-1453">Asynchronously saves an item.</span></span>

<span data-ttu-id="b789e-p181">При вызове этот метод сохраняет текущее сообщение в виде черновика и возвращает идентификатор элемента с помощью метода обратного вызова. В Outlook Web App или интерактивном режиме Outlook этот элемент сохраняется на сервере. В Outlook в режиме кэширования этот элемент сохраняется в локальном кэше.</span><span class="sxs-lookup"><span data-stu-id="b789e-p181">When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="b789e-1457">Если в режиме создания надстройка вызывает для элемента метод `saveAsync`, чтобы получить параметр `itemId` для использования с EWS или REST API, необходимо помнить, что синхронизация элемента с сервером может занять много времени, если Outlook работает в режиме кэширования данных.</span><span class="sxs-lookup"><span data-stu-id="b789e-1457">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="b789e-1458">До окончания синхронизации элемента применение параметра `itemId` будет приводить к ошибке.</span><span class="sxs-lookup"><span data-stu-id="b789e-1458">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="b789e-p183">Если метод `saveAsync` вызывается для встречи в режиме создания, она сохраняется как обычная встреча в календаре пользователя, а не как черновик. При сохранении новой встречи приглашения не отправляются. При сохранении существующей встречи уведомления отправляются добавленным или удаленным участникам.</span><span class="sxs-lookup"><span data-stu-id="b789e-p183">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="b789e-1462">Следующие клиенты отличаются другим поведением `saveAsync` в отношении встреч в режиме создания:</span><span class="sxs-lookup"><span data-stu-id="b789e-1462">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="b789e-1463">Outlook для Mac не поддерживает `saveAsync` для собраний в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="b789e-1463">Mac Outlook does not support `saveAsync` on a meeting in compose mode.</span></span> <span data-ttu-id="b789e-1464">При вызове `saveAsync` для собрания в Outlook для Mac возвращается ошибка.</span><span class="sxs-lookup"><span data-stu-id="b789e-1464">Calling `saveAsync` on a meeting in Mac Outlook will return an error.</span></span>
> - <span data-ttu-id="b789e-1465">Outlook в Интернете всегда отправляет приглашение или обновление при вызове `saveAsync` для встречи в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="b789e-1465">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b789e-1466">Параметры</span><span class="sxs-lookup"><span data-stu-id="b789e-1466">Parameters:</span></span>

|<span data-ttu-id="b789e-1467">Имя</span><span class="sxs-lookup"><span data-stu-id="b789e-1467">Name</span></span>|<span data-ttu-id="b789e-1468">Тип</span><span class="sxs-lookup"><span data-stu-id="b789e-1468">Type</span></span>|<span data-ttu-id="b789e-1469">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="b789e-1469">Attributes</span></span>|<span data-ttu-id="b789e-1470">Описание</span><span class="sxs-lookup"><span data-stu-id="b789e-1470">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="b789e-1471">Object</span><span class="sxs-lookup"><span data-stu-id="b789e-1471">Object</span></span>|<span data-ttu-id="b789e-1472">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="b789e-1472">&lt;optional&gt;</span></span>|<span data-ttu-id="b789e-1473">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="b789e-1473">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="b789e-1474">Object</span><span class="sxs-lookup"><span data-stu-id="b789e-1474">Object</span></span>|<span data-ttu-id="b789e-1475">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="b789e-1475">&lt;optional&gt;</span></span>|<span data-ttu-id="b789e-1476">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="b789e-1476">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="b789e-1477">функция</span><span class="sxs-lookup"><span data-stu-id="b789e-1477">function</span></span>||<span data-ttu-id="b789e-1478">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="b789e-1478">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="b789e-1479">После успешного выполнения идентификатор элемента будет предоставлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="b789e-1479">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b789e-1480">Требования</span><span class="sxs-lookup"><span data-stu-id="b789e-1480">Requirements</span></span>

|<span data-ttu-id="b789e-1481">Требование</span><span class="sxs-lookup"><span data-stu-id="b789e-1481">Requirement</span></span>|<span data-ttu-id="b789e-1482">Значение</span><span class="sxs-lookup"><span data-stu-id="b789e-1482">Value</span></span>|
|---|---|
|[<span data-ttu-id="b789e-1483">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="b789e-1483">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="b789e-1484">1.3</span><span class="sxs-lookup"><span data-stu-id="b789e-1484">1.3</span></span>|
|[<span data-ttu-id="b789e-1485">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b789e-1485">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="b789e-1486">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="b789e-1486">ReadWriteItem</span></span>|
|[<span data-ttu-id="b789e-1487">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b789e-1487">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="b789e-1488">Создание</span><span class="sxs-lookup"><span data-stu-id="b789e-1488">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="b789e-1489">Примеры</span><span class="sxs-lookup"><span data-stu-id="b789e-1489">Examples</span></span>

```javascript
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result
  });
```

<span data-ttu-id="b789e-p185">Ниже приведен пример параметра `result`, переданного функции обратного вызова. Свойство `value` содержит идентификатор элемента.</span><span class="sxs-lookup"><span data-stu-id="b789e-p185">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

####  <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="b789e-1492">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="b789e-1492">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="b789e-1493">Асинхронно вставляет данные в текст или тему сообщения.</span><span class="sxs-lookup"><span data-stu-id="b789e-1493">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="b789e-p186">Метод `setSelectedDataAsync` вставляет указанную строку в местоположение курсора в теме или тексте элемента либо, если текст выделен в редакторе, он заменяет выделенный текст. Если курсор находится вне текста или темы элемента, возвращается ошибка. После вставки курсор помещается в конец вставленного содержимого.</span><span class="sxs-lookup"><span data-stu-id="b789e-p186">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b789e-1497">Параметры:</span><span class="sxs-lookup"><span data-stu-id="b789e-1497">Parameters:</span></span>

|<span data-ttu-id="b789e-1498">Имя</span><span class="sxs-lookup"><span data-stu-id="b789e-1498">Name</span></span>|<span data-ttu-id="b789e-1499">Тип</span><span class="sxs-lookup"><span data-stu-id="b789e-1499">Type</span></span>|<span data-ttu-id="b789e-1500">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="b789e-1500">Attributes</span></span>|<span data-ttu-id="b789e-1501">Описание</span><span class="sxs-lookup"><span data-stu-id="b789e-1501">Description</span></span>|
|---|---|---|---|
|`data`|<span data-ttu-id="b789e-1502">String</span><span class="sxs-lookup"><span data-stu-id="b789e-1502">String</span></span>||<span data-ttu-id="b789e-p187">Вставляемые данные. Объем данных не должен превышать 1 000 000 символов. Если передано больше 1 000 000 символов, возвращается исключение `ArgumentOutOfRange`.</span><span class="sxs-lookup"><span data-stu-id="b789e-p187">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`|<span data-ttu-id="b789e-1506">Object</span><span class="sxs-lookup"><span data-stu-id="b789e-1506">Object</span></span>|<span data-ttu-id="b789e-1507">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="b789e-1507">&lt;optional&gt;</span></span>|<span data-ttu-id="b789e-1508">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="b789e-1508">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="b789e-1509">Object</span><span class="sxs-lookup"><span data-stu-id="b789e-1509">Object</span></span>|<span data-ttu-id="b789e-1510">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="b789e-1510">&lt;optional&gt;</span></span>|<span data-ttu-id="b789e-1511">В методе обратного вызова разработчики могут указать любой объект, к которому необходимо получить доступ.</span><span class="sxs-lookup"><span data-stu-id="b789e-1511">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="b789e-1512">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="b789e-1512">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="b789e-1513">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="b789e-1513">&lt;optional&gt;</span></span>|<span data-ttu-id="b789e-p188">Если задано значение `text`, текущий стиль применяется в Outlook Web App и Outlook. Если поле представляет собой редактор HTML, вставляются только текстовые данные, даже если они имеют формат HTML.</span><span class="sxs-lookup"><span data-stu-id="b789e-p188">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="b789e-p189">Если задано значение `html` и поле (не тема) поддерживает HTML, в Outlook Web App применяется текущий стиль, а в Outlook — стиль по умолчанию. Если поле является текстовым, возвращается ошибка `InvalidDataFormat`.</span><span class="sxs-lookup"><span data-stu-id="b789e-p189">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="b789e-1518">Если свойство `coercionType` не задано, результат зависит от поля: если поле имеет формат HTML, используется текст в формате HTML, а если поле текстовое, применяется обычный текст.</span><span class="sxs-lookup"><span data-stu-id="b789e-1518">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`|<span data-ttu-id="b789e-1519">функция</span><span class="sxs-lookup"><span data-stu-id="b789e-1519">function</span></span>||<span data-ttu-id="b789e-1520">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="b789e-1520">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b789e-1521">Требования</span><span class="sxs-lookup"><span data-stu-id="b789e-1521">Requirements</span></span>

|<span data-ttu-id="b789e-1522">Требование</span><span class="sxs-lookup"><span data-stu-id="b789e-1522">Requirement</span></span>|<span data-ttu-id="b789e-1523">Значение</span><span class="sxs-lookup"><span data-stu-id="b789e-1523">Value</span></span>|
|---|---|
|[<span data-ttu-id="b789e-1524">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="b789e-1524">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="b789e-1525">1.2</span><span class="sxs-lookup"><span data-stu-id="b789e-1525">1.2</span></span>|
|[<span data-ttu-id="b789e-1526">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b789e-1526">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="b789e-1527">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="b789e-1527">ReadWriteItem</span></span>|
|[<span data-ttu-id="b789e-1528">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b789e-1528">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="b789e-1529">Создание</span><span class="sxs-lookup"><span data-stu-id="b789e-1529">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="b789e-1530">Пример</span><span class="sxs-lookup"><span data-stu-id="b789e-1530">Example</span></span>

```javascript
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
