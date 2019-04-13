---
title: Office. Context. Mailbox. Item — Предварительная версия набора требований
description: ''
ms.date: 04/12/2019
localization_priority: Normal
ms.openlocfilehash: b74b3aa3c455d33d17767163c960adef7cf783fa
ms.sourcegitcommit: 95ed6dfbfa680dbb40ff9757020fa7e5be4760b6
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/13/2019
ms.locfileid: "31838587"
---
# <a name="item"></a><span data-ttu-id="45fae-102">item</span><span class="sxs-lookup"><span data-stu-id="45fae-102">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="45fae-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="45fae-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="45fae-p101">Пространство имен `item` используется для доступа к выбранному в данный момент сообщению, приглашению на собрание или описанию встречи. Вы можете определить тип пространства имен `item` с помощью свойства [itemType](#itemtype-officemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="45fae-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="45fae-106">Requirements</span><span class="sxs-lookup"><span data-stu-id="45fae-106">Requirements</span></span>

|<span data-ttu-id="45fae-107">Требование</span><span class="sxs-lookup"><span data-stu-id="45fae-107">Requirement</span></span>|<span data-ttu-id="45fae-108">Значение</span><span class="sxs-lookup"><span data-stu-id="45fae-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="45fae-109">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="45fae-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="45fae-110">1.0</span><span class="sxs-lookup"><span data-stu-id="45fae-110">1.0</span></span>|
|[<span data-ttu-id="45fae-111">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="45fae-111">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="45fae-112">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="45fae-112">Restricted</span></span>|
|[<span data-ttu-id="45fae-113">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="45fae-113">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="45fae-114">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="45fae-114">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="45fae-115">Элементы и методы</span><span class="sxs-lookup"><span data-stu-id="45fae-115">Members and methods</span></span>

| <span data-ttu-id="45fae-116">Элемент</span><span class="sxs-lookup"><span data-stu-id="45fae-116">Member</span></span> | <span data-ttu-id="45fae-117">Тип</span><span class="sxs-lookup"><span data-stu-id="45fae-117">Type</span></span> |
|--------|------|
| [<span data-ttu-id="45fae-118">attachments</span><span class="sxs-lookup"><span data-stu-id="45fae-118">attachments</span></span>](#attachments-arrayattachmentdetails) | <span data-ttu-id="45fae-119">Элемент</span><span class="sxs-lookup"><span data-stu-id="45fae-119">Member</span></span> |
| [<span data-ttu-id="45fae-120">bcc</span><span class="sxs-lookup"><span data-stu-id="45fae-120">bcc</span></span>](#bcc-recipients) | <span data-ttu-id="45fae-121">Элемент</span><span class="sxs-lookup"><span data-stu-id="45fae-121">Member</span></span> |
| [<span data-ttu-id="45fae-122">body</span><span class="sxs-lookup"><span data-stu-id="45fae-122">body</span></span>](#body-body) | <span data-ttu-id="45fae-123">Элемент</span><span class="sxs-lookup"><span data-stu-id="45fae-123">Member</span></span> |
| [<span data-ttu-id="45fae-124">cc</span><span class="sxs-lookup"><span data-stu-id="45fae-124">cc</span></span>](#cc-arrayemailaddressdetailsrecipients) | <span data-ttu-id="45fae-125">Элемент</span><span class="sxs-lookup"><span data-stu-id="45fae-125">Member</span></span> |
| [<span data-ttu-id="45fae-126">conversationId</span><span class="sxs-lookup"><span data-stu-id="45fae-126">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="45fae-127">Элемент</span><span class="sxs-lookup"><span data-stu-id="45fae-127">Member</span></span> |
| [<span data-ttu-id="45fae-128">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="45fae-128">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="45fae-129">Элемент</span><span class="sxs-lookup"><span data-stu-id="45fae-129">Member</span></span> |
| [<span data-ttu-id="45fae-130">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="45fae-130">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="45fae-131">Элемент</span><span class="sxs-lookup"><span data-stu-id="45fae-131">Member</span></span> |
| [<span data-ttu-id="45fae-132">end</span><span class="sxs-lookup"><span data-stu-id="45fae-132">end</span></span>](#end-datetime) | <span data-ttu-id="45fae-133">Элемент</span><span class="sxs-lookup"><span data-stu-id="45fae-133">Member</span></span> |
| [<span data-ttu-id="45fae-134">Енханцедлокатион</span><span class="sxs-lookup"><span data-stu-id="45fae-134">enhancedLocation</span></span>](#enhancedlocation-enhancedlocation) | <span data-ttu-id="45fae-135">Элемент</span><span class="sxs-lookup"><span data-stu-id="45fae-135">Member</span></span> |
| [<span data-ttu-id="45fae-136">from</span><span class="sxs-lookup"><span data-stu-id="45fae-136">from</span></span>](#from-emailaddressdetailsfrom) | <span data-ttu-id="45fae-137">Элемент</span><span class="sxs-lookup"><span data-stu-id="45fae-137">Member</span></span> |
| [<span data-ttu-id="45fae-138">Internetheaders:</span><span class="sxs-lookup"><span data-stu-id="45fae-138">internetHeaders</span></span>](#internetheaders-internetheaders) | <span data-ttu-id="45fae-139">Элемент</span><span class="sxs-lookup"><span data-stu-id="45fae-139">Member</span></span> |
| [<span data-ttu-id="45fae-140">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="45fae-140">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="45fae-141">Элемент</span><span class="sxs-lookup"><span data-stu-id="45fae-141">Member</span></span> |
| [<span data-ttu-id="45fae-142">itemClass</span><span class="sxs-lookup"><span data-stu-id="45fae-142">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="45fae-143">Элемент</span><span class="sxs-lookup"><span data-stu-id="45fae-143">Member</span></span> |
| [<span data-ttu-id="45fae-144">itemId</span><span class="sxs-lookup"><span data-stu-id="45fae-144">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="45fae-145">Элемент</span><span class="sxs-lookup"><span data-stu-id="45fae-145">Member</span></span> |
| [<span data-ttu-id="45fae-146">itemType</span><span class="sxs-lookup"><span data-stu-id="45fae-146">itemType</span></span>](#itemtype-officemailboxenumsitemtype) | <span data-ttu-id="45fae-147">Элемент</span><span class="sxs-lookup"><span data-stu-id="45fae-147">Member</span></span> |
| [<span data-ttu-id="45fae-148">location</span><span class="sxs-lookup"><span data-stu-id="45fae-148">location</span></span>](#location-stringlocation) | <span data-ttu-id="45fae-149">Элемент</span><span class="sxs-lookup"><span data-stu-id="45fae-149">Member</span></span> |
| [<span data-ttu-id="45fae-150">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="45fae-150">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="45fae-151">Элемент</span><span class="sxs-lookup"><span data-stu-id="45fae-151">Member</span></span> |
| [<span data-ttu-id="45fae-152">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="45fae-152">notificationMessages</span></span>](#notificationmessages-notificationmessages) | <span data-ttu-id="45fae-153">Элемент</span><span class="sxs-lookup"><span data-stu-id="45fae-153">Member</span></span> |
| [<span data-ttu-id="45fae-154">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="45fae-154">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="45fae-155">Member</span><span class="sxs-lookup"><span data-stu-id="45fae-155">Member</span></span> |
| [<span data-ttu-id="45fae-156">organizer</span><span class="sxs-lookup"><span data-stu-id="45fae-156">organizer</span></span>](#organizer-emailaddressdetailsorganizer) | <span data-ttu-id="45fae-157">Элемент</span><span class="sxs-lookup"><span data-stu-id="45fae-157">Member</span></span> |
| [<span data-ttu-id="45fae-158">recurrence</span><span class="sxs-lookup"><span data-stu-id="45fae-158">recurrence</span></span>](#nullable-recurrence-recurrence) | <span data-ttu-id="45fae-159">Элемент</span><span class="sxs-lookup"><span data-stu-id="45fae-159">Member</span></span> |
| [<span data-ttu-id="45fae-160">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="45fae-160">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="45fae-161">Элемент</span><span class="sxs-lookup"><span data-stu-id="45fae-161">Member</span></span> |
| [<span data-ttu-id="45fae-162">sender</span><span class="sxs-lookup"><span data-stu-id="45fae-162">sender</span></span>](#sender-emailaddressdetails) | <span data-ttu-id="45fae-163">Элемент</span><span class="sxs-lookup"><span data-stu-id="45fae-163">Member</span></span> |
| [<span data-ttu-id="45fae-164">seriesId</span><span class="sxs-lookup"><span data-stu-id="45fae-164">seriesId</span></span>](#nullable-seriesid-string) | <span data-ttu-id="45fae-165">Member</span><span class="sxs-lookup"><span data-stu-id="45fae-165">Member</span></span> |
| [<span data-ttu-id="45fae-166">start</span><span class="sxs-lookup"><span data-stu-id="45fae-166">start</span></span>](#start-datetime) | <span data-ttu-id="45fae-167">Member</span><span class="sxs-lookup"><span data-stu-id="45fae-167">Member</span></span> |
| [<span data-ttu-id="45fae-168">subject</span><span class="sxs-lookup"><span data-stu-id="45fae-168">subject</span></span>](#subject-stringsubject) | <span data-ttu-id="45fae-169">Элемент</span><span class="sxs-lookup"><span data-stu-id="45fae-169">Member</span></span> |
| [<span data-ttu-id="45fae-170">to</span><span class="sxs-lookup"><span data-stu-id="45fae-170">to</span></span>](#to-arrayemailaddressdetailsrecipients) | <span data-ttu-id="45fae-171">Элемент</span><span class="sxs-lookup"><span data-stu-id="45fae-171">Member</span></span> |
| [<span data-ttu-id="45fae-172">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="45fae-172">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="45fae-173">Метод</span><span class="sxs-lookup"><span data-stu-id="45fae-173">Method</span></span> |
| [<span data-ttu-id="45fae-174">addFileAttachmentFromBase64Async</span><span class="sxs-lookup"><span data-stu-id="45fae-174">addFileAttachmentFromBase64Async</span></span>](#addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback) | <span data-ttu-id="45fae-175">Метод</span><span class="sxs-lookup"><span data-stu-id="45fae-175">Method</span></span> |
| [<span data-ttu-id="45fae-176">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="45fae-176">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="45fae-177">Метод</span><span class="sxs-lookup"><span data-stu-id="45fae-177">Method</span></span> |
| [<span data-ttu-id="45fae-178">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="45fae-178">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="45fae-179">Метод</span><span class="sxs-lookup"><span data-stu-id="45fae-179">Method</span></span> |
| [<span data-ttu-id="45fae-180">close</span><span class="sxs-lookup"><span data-stu-id="45fae-180">close</span></span>](#close) | <span data-ttu-id="45fae-181">Метод</span><span class="sxs-lookup"><span data-stu-id="45fae-181">Method</span></span> |
| [<span data-ttu-id="45fae-182">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="45fae-182">displayReplyAllForm</span></span>](#displayreplyallformformdata-callback) | <span data-ttu-id="45fae-183">Метод</span><span class="sxs-lookup"><span data-stu-id="45fae-183">Method</span></span> |
| [<span data-ttu-id="45fae-184">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="45fae-184">displayReplyForm</span></span>](#displayreplyformformdata-callback) | <span data-ttu-id="45fae-185">Метод</span><span class="sxs-lookup"><span data-stu-id="45fae-185">Method</span></span> |
| [<span data-ttu-id="45fae-186">Жетаттачментконтентасинк</span><span class="sxs-lookup"><span data-stu-id="45fae-186">getAttachmentContentAsync</span></span>](#getattachmentcontentasyncattachmentid-options-callback--attachmentcontent) | <span data-ttu-id="45fae-187">Метод</span><span class="sxs-lookup"><span data-stu-id="45fae-187">Method</span></span> |
| [<span data-ttu-id="45fae-188">Жетаттачментсасинк</span><span class="sxs-lookup"><span data-stu-id="45fae-188">getAttachmentsAsync</span></span>](#getattachmentsasyncoptions-callback--arrayattachmentdetails) | <span data-ttu-id="45fae-189">Метод</span><span class="sxs-lookup"><span data-stu-id="45fae-189">Method</span></span> |
| [<span data-ttu-id="45fae-190">getEntities</span><span class="sxs-lookup"><span data-stu-id="45fae-190">getEntities</span></span>](#getentities--entities) | <span data-ttu-id="45fae-191">Метод</span><span class="sxs-lookup"><span data-stu-id="45fae-191">Method</span></span> |
| [<span data-ttu-id="45fae-192">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="45fae-192">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="45fae-193">Метод</span><span class="sxs-lookup"><span data-stu-id="45fae-193">Method</span></span> |
| [<span data-ttu-id="45fae-194">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="45fae-194">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="45fae-195">Метод</span><span class="sxs-lookup"><span data-stu-id="45fae-195">Method</span></span> |
| [<span data-ttu-id="45fae-196">getInitializationContextAsync</span><span class="sxs-lookup"><span data-stu-id="45fae-196">getInitializationContextAsync</span></span>](#getinitializationcontextasyncoptions-callback) | <span data-ttu-id="45fae-197">Метод</span><span class="sxs-lookup"><span data-stu-id="45fae-197">Method</span></span> |
| [<span data-ttu-id="45fae-198">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="45fae-198">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="45fae-199">Метод</span><span class="sxs-lookup"><span data-stu-id="45fae-199">Method</span></span> |
| [<span data-ttu-id="45fae-200">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="45fae-200">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="45fae-201">Метод</span><span class="sxs-lookup"><span data-stu-id="45fae-201">Method</span></span> |
| [<span data-ttu-id="45fae-202">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="45fae-202">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="45fae-203">Метод</span><span class="sxs-lookup"><span data-stu-id="45fae-203">Method</span></span> |
| [<span data-ttu-id="45fae-204">Жетселектедентитиес</span><span class="sxs-lookup"><span data-stu-id="45fae-204">getSelectedEntities</span></span>](#getselectedentities--entities) | <span data-ttu-id="45fae-205">Метод</span><span class="sxs-lookup"><span data-stu-id="45fae-205">Method</span></span> |
| [<span data-ttu-id="45fae-206">Жетселектедрежексматчес</span><span class="sxs-lookup"><span data-stu-id="45fae-206">getSelectedRegExMatches</span></span>](#getselectedregexmatches--object) | <span data-ttu-id="45fae-207">Метод</span><span class="sxs-lookup"><span data-stu-id="45fae-207">Method</span></span> |
| [<span data-ttu-id="45fae-208">Жетшаредпропертиесасинк</span><span class="sxs-lookup"><span data-stu-id="45fae-208">getSharedPropertiesAsync</span></span>](#getsharedpropertiesasyncoptions-callback) | <span data-ttu-id="45fae-209">Метод</span><span class="sxs-lookup"><span data-stu-id="45fae-209">Method</span></span> |
| [<span data-ttu-id="45fae-210">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="45fae-210">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="45fae-211">Метод</span><span class="sxs-lookup"><span data-stu-id="45fae-211">Method</span></span> |
| [<span data-ttu-id="45fae-212">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="45fae-212">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="45fae-213">Метод</span><span class="sxs-lookup"><span data-stu-id="45fae-213">Method</span></span> |
| [<span data-ttu-id="45fae-214">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="45fae-214">removeHandlerAsync</span></span>](#removehandlerasynceventtype-options-callback) | <span data-ttu-id="45fae-215">Метод</span><span class="sxs-lookup"><span data-stu-id="45fae-215">Method</span></span> |
| [<span data-ttu-id="45fae-216">saveAsync</span><span class="sxs-lookup"><span data-stu-id="45fae-216">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="45fae-217">Метод</span><span class="sxs-lookup"><span data-stu-id="45fae-217">Method</span></span> |
| [<span data-ttu-id="45fae-218">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="45fae-218">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="45fae-219">Метод</span><span class="sxs-lookup"><span data-stu-id="45fae-219">Method</span></span> |

### <a name="example"></a><span data-ttu-id="45fae-220">Пример</span><span class="sxs-lookup"><span data-stu-id="45fae-220">Example</span></span>

<span data-ttu-id="45fae-221">В примере кода JavaScript, приведенном ниже, показано, как получить доступ к свойству `subject` текущего элемента в Outlook.</span><span class="sxs-lookup"><span data-stu-id="45fae-221">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="45fae-222">Элементы</span><span class="sxs-lookup"><span data-stu-id="45fae-222">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails"></a><span data-ttu-id="45fae-223">attachments :Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="45fae-223">attachments :Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

<span data-ttu-id="45fae-224">Получает вложения элемента в виде массива.</span><span class="sxs-lookup"><span data-stu-id="45fae-224">Gets the item's attachments as an array.</span></span> <span data-ttu-id="45fae-225">Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="45fae-225">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="45fae-226">Outlook блокирует определенные типы файлов из-за потенциальных проблем с безопасностью, поэтому они не возвращаются.</span><span class="sxs-lookup"><span data-stu-id="45fae-226">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="45fae-227">Дополнительные сведения см. в статье [Блокировка вложений в Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="45fae-227">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="45fae-228">Тип</span><span class="sxs-lookup"><span data-stu-id="45fae-228">Type</span></span>

*   <span data-ttu-id="45fae-229">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="45fae-229">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="45fae-230">Требования</span><span class="sxs-lookup"><span data-stu-id="45fae-230">Requirements</span></span>

|<span data-ttu-id="45fae-231">Требование</span><span class="sxs-lookup"><span data-stu-id="45fae-231">Requirement</span></span>|<span data-ttu-id="45fae-232">Значение</span><span class="sxs-lookup"><span data-stu-id="45fae-232">Value</span></span>|
|---|---|
|[<span data-ttu-id="45fae-233">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="45fae-233">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="45fae-234">1.0</span><span class="sxs-lookup"><span data-stu-id="45fae-234">1.0</span></span>|
|[<span data-ttu-id="45fae-235">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="45fae-235">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="45fae-236">ReadItem</span><span class="sxs-lookup"><span data-stu-id="45fae-236">ReadItem</span></span>|
|[<span data-ttu-id="45fae-237">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="45fae-237">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="45fae-238">Чтение</span><span class="sxs-lookup"><span data-stu-id="45fae-238">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="45fae-239">Пример</span><span class="sxs-lookup"><span data-stu-id="45fae-239">Example</span></span>

<span data-ttu-id="45fae-240">С помощью приведенного ниже кода можно создать HTML-строку с подробными сведениями обо всех вложениях для текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="45fae-240">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

####  <a name="bcc-recipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="45fae-241">bcc :[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="45fae-241">bcc :[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="45fae-242">Получает объект, который предоставляет методы для получения или обновления получателей скрытой копии сообщения.</span><span class="sxs-lookup"><span data-stu-id="45fae-242">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="45fae-243">Только в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="45fae-243">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="45fae-244">Тип</span><span class="sxs-lookup"><span data-stu-id="45fae-244">Type</span></span>

*   [<span data-ttu-id="45fae-245">Получатели</span><span class="sxs-lookup"><span data-stu-id="45fae-245">Recipients</span></span>](/javascript/api/outlook/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="45fae-246">Требования</span><span class="sxs-lookup"><span data-stu-id="45fae-246">Requirements</span></span>

|<span data-ttu-id="45fae-247">Требование</span><span class="sxs-lookup"><span data-stu-id="45fae-247">Requirement</span></span>|<span data-ttu-id="45fae-248">Значение</span><span class="sxs-lookup"><span data-stu-id="45fae-248">Value</span></span>|
|---|---|
|[<span data-ttu-id="45fae-249">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="45fae-249">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="45fae-250">1.1</span><span class="sxs-lookup"><span data-stu-id="45fae-250">1.1</span></span>|
|[<span data-ttu-id="45fae-251">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="45fae-251">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="45fae-252">ReadItem</span><span class="sxs-lookup"><span data-stu-id="45fae-252">ReadItem</span></span>|
|[<span data-ttu-id="45fae-253">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="45fae-253">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="45fae-254">Создание</span><span class="sxs-lookup"><span data-stu-id="45fae-254">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="45fae-255">Пример</span><span class="sxs-lookup"><span data-stu-id="45fae-255">Example</span></span>

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

####  <a name="body-bodyjavascriptapioutlookofficebody"></a><span data-ttu-id="45fae-256">body :[Body](/javascript/api/outlook/office.body)</span><span class="sxs-lookup"><span data-stu-id="45fae-256">body :[Body](/javascript/api/outlook/office.body)</span></span>

<span data-ttu-id="45fae-257">Получает объект, предоставляющий методы для работы с основным текстом элемента.</span><span class="sxs-lookup"><span data-stu-id="45fae-257">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="45fae-258">Тип</span><span class="sxs-lookup"><span data-stu-id="45fae-258">Type</span></span>

*   [<span data-ttu-id="45fae-259">Body</span><span class="sxs-lookup"><span data-stu-id="45fae-259">Body</span></span>](/javascript/api/outlook/office.body)

##### <a name="requirements"></a><span data-ttu-id="45fae-260">Требования</span><span class="sxs-lookup"><span data-stu-id="45fae-260">Requirements</span></span>

|<span data-ttu-id="45fae-261">Требование</span><span class="sxs-lookup"><span data-stu-id="45fae-261">Requirement</span></span>|<span data-ttu-id="45fae-262">Значение</span><span class="sxs-lookup"><span data-stu-id="45fae-262">Value</span></span>|
|---|---|
|[<span data-ttu-id="45fae-263">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="45fae-263">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="45fae-264">1.1</span><span class="sxs-lookup"><span data-stu-id="45fae-264">1.1</span></span>|
|[<span data-ttu-id="45fae-265">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="45fae-265">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="45fae-266">ReadItem</span><span class="sxs-lookup"><span data-stu-id="45fae-266">ReadItem</span></span>|
|[<span data-ttu-id="45fae-267">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="45fae-267">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="45fae-268">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="45fae-268">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="45fae-269">Пример</span><span class="sxs-lookup"><span data-stu-id="45fae-269">Example</span></span>

<span data-ttu-id="45fae-270">В этом примере возвращается текст сообщения в формате обычного текста.</span><span class="sxs-lookup"><span data-stu-id="45fae-270">This example gets the body of the message in plain text.</span></span>

```javascript
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="45fae-271">Ниже приведен пример итогового параметра, переданного функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="45fae-271">The following is an example of the result parameter passed to the callback function.</span></span>

```json
{
  "value": "TEXT of whole body (including threads below)",
  "status": "succeeded",
  "asyncContext": "This is passed to the callback"
}
```

---
---

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="45fae-272">cc :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="45fae-272">cc :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="45fae-273">Предоставляет доступ к получателям копии сообщения.</span><span class="sxs-lookup"><span data-stu-id="45fae-273">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="45fae-274">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="45fae-274">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="45fae-275">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="45fae-275">Read mode</span></span>

<span data-ttu-id="45fae-p106">Свойство `cc` возвращает массив, который содержит объект `EmailAddressDetails` для каждого получателя, указанного в строке **Копия** сообщения. Коллекция может включать не более 100 элементов.</span><span class="sxs-lookup"><span data-stu-id="45fae-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="45fae-278">Режим создания</span><span class="sxs-lookup"><span data-stu-id="45fae-278">Compose mode</span></span>

<span data-ttu-id="45fae-279">Свойство `cc` возвращает объект `Recipients`, предоставляющий методы для получения или обновления получателей, которые указаны в строке **Копия** сообщения.</span><span class="sxs-lookup"><span data-stu-id="45fae-279">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

```javascript
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="45fae-280">Тип</span><span class="sxs-lookup"><span data-stu-id="45fae-280">Type</span></span>

*   <span data-ttu-id="45fae-281">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="45fae-281">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="45fae-282">Требования</span><span class="sxs-lookup"><span data-stu-id="45fae-282">Requirements</span></span>

|<span data-ttu-id="45fae-283">Требование</span><span class="sxs-lookup"><span data-stu-id="45fae-283">Requirement</span></span>|<span data-ttu-id="45fae-284">Значение</span><span class="sxs-lookup"><span data-stu-id="45fae-284">Value</span></span>|
|---|---|
|[<span data-ttu-id="45fae-285">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="45fae-285">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="45fae-286">1.0</span><span class="sxs-lookup"><span data-stu-id="45fae-286">1.0</span></span>|
|[<span data-ttu-id="45fae-287">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="45fae-287">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="45fae-288">ReadItem</span><span class="sxs-lookup"><span data-stu-id="45fae-288">ReadItem</span></span>|
|[<span data-ttu-id="45fae-289">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="45fae-289">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="45fae-290">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="45fae-290">Compose or Read</span></span>|

---
---

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="45fae-291">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="45fae-291">(nullable) conversationId :String</span></span>

<span data-ttu-id="45fae-292">Получает идентификатор разговора по электронной почте, содержащего конкретное сообщение.</span><span class="sxs-lookup"><span data-stu-id="45fae-292">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="45fae-p107">Вы можете получить целочисленное значение этого свойства, если ваше почтовое приложение активируется в формах просмотра или формах создания ответов. Если пользователь изменит тему ответа, после его отправки идентификатор беседы будет изменен, и полученное ранее значение будет недействительным.</span><span class="sxs-lookup"><span data-stu-id="45fae-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="45fae-p108">Это свойство имеет значение NULL для нового элемента в форме создания. Свойство `conversationId` вернет значение, если пользователь задаст тему и сохранит элемент.</span><span class="sxs-lookup"><span data-stu-id="45fae-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="45fae-297">Тип</span><span class="sxs-lookup"><span data-stu-id="45fae-297">Type</span></span>

*   <span data-ttu-id="45fae-298">String</span><span class="sxs-lookup"><span data-stu-id="45fae-298">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="45fae-299">Требования</span><span class="sxs-lookup"><span data-stu-id="45fae-299">Requirements</span></span>

|<span data-ttu-id="45fae-300">Требование</span><span class="sxs-lookup"><span data-stu-id="45fae-300">Requirement</span></span>|<span data-ttu-id="45fae-301">Значение</span><span class="sxs-lookup"><span data-stu-id="45fae-301">Value</span></span>|
|---|---|
|[<span data-ttu-id="45fae-302">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="45fae-302">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="45fae-303">1.0</span><span class="sxs-lookup"><span data-stu-id="45fae-303">1.0</span></span>|
|[<span data-ttu-id="45fae-304">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="45fae-304">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="45fae-305">ReadItem</span><span class="sxs-lookup"><span data-stu-id="45fae-305">ReadItem</span></span>|
|[<span data-ttu-id="45fae-306">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="45fae-306">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="45fae-307">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="45fae-307">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="45fae-308">Пример</span><span class="sxs-lookup"><span data-stu-id="45fae-308">Example</span></span>

```javascript
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

---
---

#### <a name="datetimecreated-date"></a><span data-ttu-id="45fae-309">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="45fae-309">dateTimeCreated :Date</span></span>

<span data-ttu-id="45fae-p109">Получает дату и время создания элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="45fae-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="45fae-312">Тип</span><span class="sxs-lookup"><span data-stu-id="45fae-312">Type</span></span>

*   <span data-ttu-id="45fae-313">Дата</span><span class="sxs-lookup"><span data-stu-id="45fae-313">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="45fae-314">Требования</span><span class="sxs-lookup"><span data-stu-id="45fae-314">Requirements</span></span>

|<span data-ttu-id="45fae-315">Требование</span><span class="sxs-lookup"><span data-stu-id="45fae-315">Requirement</span></span>|<span data-ttu-id="45fae-316">Значение</span><span class="sxs-lookup"><span data-stu-id="45fae-316">Value</span></span>|
|---|---|
|[<span data-ttu-id="45fae-317">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="45fae-317">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="45fae-318">1.0</span><span class="sxs-lookup"><span data-stu-id="45fae-318">1.0</span></span>|
|[<span data-ttu-id="45fae-319">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="45fae-319">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="45fae-320">ReadItem</span><span class="sxs-lookup"><span data-stu-id="45fae-320">ReadItem</span></span>|
|[<span data-ttu-id="45fae-321">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="45fae-321">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="45fae-322">Чтение</span><span class="sxs-lookup"><span data-stu-id="45fae-322">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="45fae-323">Пример</span><span class="sxs-lookup"><span data-stu-id="45fae-323">Example</span></span>

```javascript
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

---
---

#### <a name="datetimemodified-date"></a><span data-ttu-id="45fae-324">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="45fae-324">dateTimeModified :Date</span></span>

<span data-ttu-id="45fae-p110">Получает дату и время последнего изменения элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="45fae-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="45fae-327">Этот элемент не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="45fae-327">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="45fae-328">Тип</span><span class="sxs-lookup"><span data-stu-id="45fae-328">Type</span></span>

*   <span data-ttu-id="45fae-329">Дата</span><span class="sxs-lookup"><span data-stu-id="45fae-329">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="45fae-330">Требования</span><span class="sxs-lookup"><span data-stu-id="45fae-330">Requirements</span></span>

|<span data-ttu-id="45fae-331">Требование</span><span class="sxs-lookup"><span data-stu-id="45fae-331">Requirement</span></span>|<span data-ttu-id="45fae-332">Значение</span><span class="sxs-lookup"><span data-stu-id="45fae-332">Value</span></span>|
|---|---|
|[<span data-ttu-id="45fae-333">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="45fae-333">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="45fae-334">1.0</span><span class="sxs-lookup"><span data-stu-id="45fae-334">1.0</span></span>|
|[<span data-ttu-id="45fae-335">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="45fae-335">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="45fae-336">ReadItem</span><span class="sxs-lookup"><span data-stu-id="45fae-336">ReadItem</span></span>|
|[<span data-ttu-id="45fae-337">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="45fae-337">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="45fae-338">Чтение</span><span class="sxs-lookup"><span data-stu-id="45fae-338">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="45fae-339">Пример</span><span class="sxs-lookup"><span data-stu-id="45fae-339">Example</span></span>

```javascript
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

---
---

####  <a name="end-datetimejavascriptapioutlookofficetime"></a><span data-ttu-id="45fae-340">end :Date|[Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="45fae-340">end :Date|[Time](/javascript/api/outlook/office.time)</span></span>

<span data-ttu-id="45fae-341">Получает или задает дату и время окончания встречи.</span><span class="sxs-lookup"><span data-stu-id="45fae-341">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="45fae-p111">Свойство `end` представлено в виде значения даты и времени в формате UTC. Преобразовать значение свойства end в местные значения даты и времени клиента можно с помощью метода [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime).</span><span class="sxs-lookup"><span data-stu-id="45fae-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="45fae-344">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="45fae-344">Read mode</span></span>

<span data-ttu-id="45fae-345">Свойство `end` возвращает объект `Date`.</span><span class="sxs-lookup"><span data-stu-id="45fae-345">The `end` property returns a `Date` object.</span></span>

```javascript
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="45fae-346">Режим создания</span><span class="sxs-lookup"><span data-stu-id="45fae-346">Compose mode</span></span>

<span data-ttu-id="45fae-347">Свойство `end` возвращает объект `Time`.</span><span class="sxs-lookup"><span data-stu-id="45fae-347">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="45fae-348">Если вы задаете время окончания с помощью метода [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-), необходимо использовать метод [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) для преобразования местного времени на клиенте в формат UTC для сервера.</span><span class="sxs-lookup"><span data-stu-id="45fae-348">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="45fae-349">В примере ниже показано, как с помощью метода [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) объекта `Time` задать время окончания встречи.</span><span class="sxs-lookup"><span data-stu-id="45fae-349">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="45fae-350">Тип</span><span class="sxs-lookup"><span data-stu-id="45fae-350">Type</span></span>

*   <span data-ttu-id="45fae-351">Date | [Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="45fae-351">Date | [Time](/javascript/api/outlook/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="45fae-352">Требования</span><span class="sxs-lookup"><span data-stu-id="45fae-352">Requirements</span></span>

|<span data-ttu-id="45fae-353">Требование</span><span class="sxs-lookup"><span data-stu-id="45fae-353">Requirement</span></span>|<span data-ttu-id="45fae-354">Значение</span><span class="sxs-lookup"><span data-stu-id="45fae-354">Value</span></span>|
|---|---|
|[<span data-ttu-id="45fae-355">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="45fae-355">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="45fae-356">1.0</span><span class="sxs-lookup"><span data-stu-id="45fae-356">1.0</span></span>|
|[<span data-ttu-id="45fae-357">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="45fae-357">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="45fae-358">ReadItem</span><span class="sxs-lookup"><span data-stu-id="45fae-358">ReadItem</span></span>|
|[<span data-ttu-id="45fae-359">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="45fae-359">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="45fae-360">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="45fae-360">Compose or Read</span></span>|

---
---

#### <a name="enhancedlocation-enhancedlocationjavascriptapioutlookofficeenhancedlocation"></a><span data-ttu-id="45fae-361">Енханцедлокатион:[енханцедлокатион](/javascript/api/outlook/office.enhancedlocation)</span><span class="sxs-lookup"><span data-stu-id="45fae-361">enhancedLocation :[EnhancedLocation](/javascript/api/outlook/office.enhancedlocation)</span></span>

<span data-ttu-id="45fae-362">Получает или задает расположение встречи.</span><span class="sxs-lookup"><span data-stu-id="45fae-362">Gets or sets the locations of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="45fae-363">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="45fae-363">Read mode</span></span>

<span data-ttu-id="45fae-364">Свойство возвращает объект [енханцедлокатион](/javascript/api/outlook/office.enhancedlocation) , который позволяет получить набор расположений (каждый, представленный объектом локатиондетаилс), связанный с встречей. [](/javascript/api/outlook/office.locationdetails) `enhancedLocation`</span><span class="sxs-lookup"><span data-stu-id="45fae-364">The `enhancedLocation` property returns an [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation) object that allows you to get the set of locations (each represented by a [LocationDetails](/javascript/api/outlook/office.locationdetails) object) associated with the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="45fae-365">Режим создания</span><span class="sxs-lookup"><span data-stu-id="45fae-365">Compose mode</span></span>

<span data-ttu-id="45fae-366">`enhancedLocation` Свойство возвращает объект [енханцедлокатион](/javascript/api/outlook/office.enhancedlocation) , который предоставляет методы для получения, удаления или добавления расположений для встречи.</span><span class="sxs-lookup"><span data-stu-id="45fae-366">The `enhancedLocation` property returns an [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation) object that provides methods to get, remove, or add locations on an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="45fae-367">Тип</span><span class="sxs-lookup"><span data-stu-id="45fae-367">Type</span></span>

*   [<span data-ttu-id="45fae-368">EnhancedLocation</span><span class="sxs-lookup"><span data-stu-id="45fae-368">EnhancedLocation</span></span>](/javascript/api/outlook/office.enhancedlocation)

##### <a name="requirements"></a><span data-ttu-id="45fae-369">Требования</span><span class="sxs-lookup"><span data-stu-id="45fae-369">Requirements</span></span>

|<span data-ttu-id="45fae-370">Требование</span><span class="sxs-lookup"><span data-stu-id="45fae-370">Requirement</span></span>|<span data-ttu-id="45fae-371">Значение</span><span class="sxs-lookup"><span data-stu-id="45fae-371">Value</span></span>|
|---|---|
|[<span data-ttu-id="45fae-372">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="45fae-372">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="45fae-373">Предварительная версия</span><span class="sxs-lookup"><span data-stu-id="45fae-373">Preview</span></span>|
|[<span data-ttu-id="45fae-374">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="45fae-374">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="45fae-375">ReadItem</span><span class="sxs-lookup"><span data-stu-id="45fae-375">ReadItem</span></span>|
|[<span data-ttu-id="45fae-376">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="45fae-376">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="45fae-377">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="45fae-377">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="45fae-378">Пример</span><span class="sxs-lookup"><span data-stu-id="45fae-378">Example</span></span>

<span data-ttu-id="45fae-379">В следующем примере показано получение текущих расположений, связанных с встречей.</span><span class="sxs-lookup"><span data-stu-id="45fae-379">The following example gets the current locations associated with the appointment.</span></span>

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

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom"></a><span data-ttu-id="45fae-380">от:[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[](/javascript/api/outlook/office.from)</span><span class="sxs-lookup"><span data-stu-id="45fae-380">from :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[From](/javascript/api/outlook/office.from)</span></span>

<span data-ttu-id="45fae-381">Получает электронный адрес отправителя сообщения.</span><span class="sxs-lookup"><span data-stu-id="45fae-381">Gets the email address of the sender of a message.</span></span>

<span data-ttu-id="45fae-p112">Свойства `from` и [`sender`](#sender-emailaddressdetails) представляют одно лицо, если сообщение не отправлено представителем. В противном случае свойство `from` представляет лицо, делегировавшее полномочия, а свойство sender — представителя.</span><span class="sxs-lookup"><span data-stu-id="45fae-p112">The `from` and [`sender`](#sender-emailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="45fae-384">Свойству `recipientType`, принадлежащему объекту `EmailAddressDetails` в свойстве `from`, задано значение `undefined`.</span><span class="sxs-lookup"><span data-stu-id="45fae-384">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="45fae-385">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="45fae-385">Read mode</span></span>

<span data-ttu-id="45fae-386">`from` Свойство возвращает `EmailAddressDetails` объект.</span><span class="sxs-lookup"><span data-stu-id="45fae-386">The `from` property returns an `EmailAddressDetails` object.</span></span>

```javascript
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

##### <a name="compose-mode"></a><span data-ttu-id="45fae-387">Режим создания</span><span class="sxs-lookup"><span data-stu-id="45fae-387">Compose mode</span></span>

<span data-ttu-id="45fae-388">`from` Свойство возвращает `From` объект, который предоставляет метод для получения значения From.</span><span class="sxs-lookup"><span data-stu-id="45fae-388">The `from` property returns a `From` object that provides a method to get the from value.</span></span>

```javascript
Office.context.mailbox.item.from.getAsync(callback);

function callback(asyncResult) {
  var from = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="45fae-389">Тип</span><span class="sxs-lookup"><span data-stu-id="45fae-389">Type</span></span>

*   <span data-ttu-id="45fae-390">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [из](/javascript/api/outlook/office.from)</span><span class="sxs-lookup"><span data-stu-id="45fae-390">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [From](/javascript/api/outlook/office.from)</span></span>

##### <a name="requirements"></a><span data-ttu-id="45fae-391">Требования</span><span class="sxs-lookup"><span data-stu-id="45fae-391">Requirements</span></span>

|<span data-ttu-id="45fae-392">Требование</span><span class="sxs-lookup"><span data-stu-id="45fae-392">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="45fae-393">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="45fae-393">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="45fae-394">1.0</span><span class="sxs-lookup"><span data-stu-id="45fae-394">1.0</span></span>|<span data-ttu-id="45fae-395">1.7</span><span class="sxs-lookup"><span data-stu-id="45fae-395">1.7</span></span>|
|[<span data-ttu-id="45fae-396">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="45fae-396">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="45fae-397">ReadItem</span><span class="sxs-lookup"><span data-stu-id="45fae-397">ReadItem</span></span>|<span data-ttu-id="45fae-398">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="45fae-398">ReadWriteItem</span></span>|
|[<span data-ttu-id="45fae-399">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="45fae-399">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="45fae-400">Чтение</span><span class="sxs-lookup"><span data-stu-id="45fae-400">Read</span></span>|<span data-ttu-id="45fae-401">Создание</span><span class="sxs-lookup"><span data-stu-id="45fae-401">Compose</span></span>|

---
---

#### <a name="internetheaders-internetheadersjavascriptapioutlookofficeinternetheaders"></a><span data-ttu-id="45fae-402">Internetheaders::[internetheaders:](/javascript/api/outlook/office.internetheaders)</span><span class="sxs-lookup"><span data-stu-id="45fae-402">internetHeaders :[InternetHeaders](/javascript/api/outlook/office.internetheaders)</span></span>

<span data-ttu-id="45fae-403">Возвращает или задает заголовки Интернета сообщения.</span><span class="sxs-lookup"><span data-stu-id="45fae-403">Gets or sets the internet headers of a message.</span></span>

##### <a name="type"></a><span data-ttu-id="45fae-404">Тип</span><span class="sxs-lookup"><span data-stu-id="45fae-404">Type</span></span>

*   [<span data-ttu-id="45fae-405">InternetHeaders</span><span class="sxs-lookup"><span data-stu-id="45fae-405">InternetHeaders</span></span>](/javascript/api/outlook/office.internetheaders)

##### <a name="requirements"></a><span data-ttu-id="45fae-406">Требования</span><span class="sxs-lookup"><span data-stu-id="45fae-406">Requirements</span></span>

|<span data-ttu-id="45fae-407">Требование</span><span class="sxs-lookup"><span data-stu-id="45fae-407">Requirement</span></span>|<span data-ttu-id="45fae-408">Значение</span><span class="sxs-lookup"><span data-stu-id="45fae-408">Value</span></span>|
|---|---|
|[<span data-ttu-id="45fae-409">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="45fae-409">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="45fae-410">Предварительная версия</span><span class="sxs-lookup"><span data-stu-id="45fae-410">Preview</span></span>|
|[<span data-ttu-id="45fae-411">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="45fae-411">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="45fae-412">ReadItem</span><span class="sxs-lookup"><span data-stu-id="45fae-412">ReadItem</span></span>|
|[<span data-ttu-id="45fae-413">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="45fae-413">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="45fae-414">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="45fae-414">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="45fae-415">Пример</span><span class="sxs-lookup"><span data-stu-id="45fae-415">Example</span></span>

```javascript
Office.context.mailbox.item.internetHeaders.getAsync(["header1", "header2"], callback);

function callback(asyncResult) {
  var dictionary = asyncResult.value;
  var header1_value = dictionary["header1"];
}
```

---
---

#### <a name="internetmessageid-string"></a><span data-ttu-id="45fae-416">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="45fae-416">internetMessageId :String</span></span>

<span data-ttu-id="45fae-p113">Получает идентификатор интернет-сообщения для электронного сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="45fae-p113">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="45fae-419">Тип</span><span class="sxs-lookup"><span data-stu-id="45fae-419">Type</span></span>

*   <span data-ttu-id="45fae-420">String</span><span class="sxs-lookup"><span data-stu-id="45fae-420">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="45fae-421">Требования</span><span class="sxs-lookup"><span data-stu-id="45fae-421">Requirements</span></span>

|<span data-ttu-id="45fae-422">Требование</span><span class="sxs-lookup"><span data-stu-id="45fae-422">Requirement</span></span>|<span data-ttu-id="45fae-423">Значение</span><span class="sxs-lookup"><span data-stu-id="45fae-423">Value</span></span>|
|---|---|
|[<span data-ttu-id="45fae-424">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="45fae-424">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="45fae-425">1.0</span><span class="sxs-lookup"><span data-stu-id="45fae-425">1.0</span></span>|
|[<span data-ttu-id="45fae-426">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="45fae-426">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="45fae-427">ReadItem</span><span class="sxs-lookup"><span data-stu-id="45fae-427">ReadItem</span></span>|
|[<span data-ttu-id="45fae-428">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="45fae-428">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="45fae-429">Чтение</span><span class="sxs-lookup"><span data-stu-id="45fae-429">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="45fae-430">Пример</span><span class="sxs-lookup"><span data-stu-id="45fae-430">Example</span></span>

```javascript
var internetMessageId = Office.context.mailbox.item.internetMessageId;
console.log("internetMessageId: " + internetMessageId);
```

---
---

#### <a name="itemclass-string"></a><span data-ttu-id="45fae-431">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="45fae-431">itemClass :String</span></span>

<span data-ttu-id="45fae-p114">Получает класс элемента веб-служб Exchange для выбранного элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="45fae-p114">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="45fae-p115">Свойство `itemClass` указывает класс сообщения выбранного элемента. Ниже приводятся классы сообщения по умолчанию для элемента сообщения или встречи.</span><span class="sxs-lookup"><span data-stu-id="45fae-p115">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

|<span data-ttu-id="45fae-436">Тип</span><span class="sxs-lookup"><span data-stu-id="45fae-436">Type</span></span>|<span data-ttu-id="45fae-437">Описание</span><span class="sxs-lookup"><span data-stu-id="45fae-437">Description</span></span>|<span data-ttu-id="45fae-438">Класс элемента</span><span class="sxs-lookup"><span data-stu-id="45fae-438">item class</span></span>|
|---|---|---|
|<span data-ttu-id="45fae-439">Элементы встречи</span><span class="sxs-lookup"><span data-stu-id="45fae-439">Appointment items</span></span>|<span data-ttu-id="45fae-440">Это элементы календаря для класса элемента `IPM.Appointment` или `IPM.Appointment.Occurrence`.</span><span class="sxs-lookup"><span data-stu-id="45fae-440">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span>|`IPM.Appointment`<br />`IPM.Appointment.Occurrence`|
|<span data-ttu-id="45fae-441">Элементы сообщения</span><span class="sxs-lookup"><span data-stu-id="45fae-441">Message items</span></span>|<span data-ttu-id="45fae-442">Сюда входят электронные сообщения, для которых по умолчанию задан класс сообщения `IPM.Note`, а также приглашения на собрания, ответы на них и уведомления об их отмене, использующие `IPM.Schedule.Meeting` в качестве базового класса сообщения.</span><span class="sxs-lookup"><span data-stu-id="45fae-442">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span>|`IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled`|

<span data-ttu-id="45fae-443">Можно создавать настраиваемые классы сообщения, расширяющие классы сообщения по умолчанию, например настраиваемый класс сообщения о встрече `IPM.Appointment.Contoso`.</span><span class="sxs-lookup"><span data-stu-id="45fae-443">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="45fae-444">Тип</span><span class="sxs-lookup"><span data-stu-id="45fae-444">Type</span></span>

*   <span data-ttu-id="45fae-445">String</span><span class="sxs-lookup"><span data-stu-id="45fae-445">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="45fae-446">Требования</span><span class="sxs-lookup"><span data-stu-id="45fae-446">Requirements</span></span>

|<span data-ttu-id="45fae-447">Требование</span><span class="sxs-lookup"><span data-stu-id="45fae-447">Requirement</span></span>|<span data-ttu-id="45fae-448">Значение</span><span class="sxs-lookup"><span data-stu-id="45fae-448">Value</span></span>|
|---|---|
|[<span data-ttu-id="45fae-449">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="45fae-449">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="45fae-450">1.0</span><span class="sxs-lookup"><span data-stu-id="45fae-450">1.0</span></span>|
|[<span data-ttu-id="45fae-451">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="45fae-451">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="45fae-452">ReadItem</span><span class="sxs-lookup"><span data-stu-id="45fae-452">ReadItem</span></span>|
|[<span data-ttu-id="45fae-453">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="45fae-453">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="45fae-454">Чтение</span><span class="sxs-lookup"><span data-stu-id="45fae-454">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="45fae-455">Пример</span><span class="sxs-lookup"><span data-stu-id="45fae-455">Example</span></span>

```javascript
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

---
---

#### <a name="nullable-itemid-string"></a><span data-ttu-id="45fae-456">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="45fae-456">(nullable) itemId :String</span></span>

<span data-ttu-id="45fae-p116">Получает идентификатор элемента веб-служб Exchange для текущего элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="45fae-p116">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="45fae-459">Идентификатор, возвращаемый свойством `itemId`, совпадает с идентификатором элемента веб-служб Exchange.</span><span class="sxs-lookup"><span data-stu-id="45fae-459">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="45fae-460">Свойство `itemId` не совпадает с идентификатором записи Outlook, а также идентификатором, который используется REST API Outlook.</span><span class="sxs-lookup"><span data-stu-id="45fae-460">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="45fae-461">Прежде чем совершать вызовы REST API, используя это значение, его необходимо преобразовать с помощью [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="45fae-461">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="45fae-462">Дополнительные сведения см. в статье [Использование REST API Outlook из надстройки Outlook](/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="45fae-462">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="45fae-p118">Свойство `itemId` недоступно в режиме создания. Если требуется идентификатор элемента, с помощью метода [`saveAsync`](#saveasyncoptions-callback) можно сохранить элемент в хранилище. При этом в параметре [`AsyncResult.value`](/javascript/api/office/office.asyncresult) функции обратного вызова возвращается идентификатор элемента.</span><span class="sxs-lookup"><span data-stu-id="45fae-p118">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="45fae-465">Тип</span><span class="sxs-lookup"><span data-stu-id="45fae-465">Type</span></span>

*   <span data-ttu-id="45fae-466">String</span><span class="sxs-lookup"><span data-stu-id="45fae-466">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="45fae-467">Требования</span><span class="sxs-lookup"><span data-stu-id="45fae-467">Requirements</span></span>

|<span data-ttu-id="45fae-468">Требование</span><span class="sxs-lookup"><span data-stu-id="45fae-468">Requirement</span></span>|<span data-ttu-id="45fae-469">Значение</span><span class="sxs-lookup"><span data-stu-id="45fae-469">Value</span></span>|
|---|---|
|[<span data-ttu-id="45fae-470">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="45fae-470">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="45fae-471">1.0</span><span class="sxs-lookup"><span data-stu-id="45fae-471">1.0</span></span>|
|[<span data-ttu-id="45fae-472">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="45fae-472">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="45fae-473">ReadItem</span><span class="sxs-lookup"><span data-stu-id="45fae-473">ReadItem</span></span>|
|[<span data-ttu-id="45fae-474">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="45fae-474">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="45fae-475">Чтение</span><span class="sxs-lookup"><span data-stu-id="45fae-475">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="45fae-476">Пример</span><span class="sxs-lookup"><span data-stu-id="45fae-476">Example</span></span>

<span data-ttu-id="45fae-p119">Указанный ниже код проверяет наличие идентификатора элемента. Если свойство `itemId` возвращает значение `null` или `undefined`, элемент будет сохранен в хранилище, а из асинхронного результата будет получен идентификатор элемента.</span><span class="sxs-lookup"><span data-stu-id="45fae-p119">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

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

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype"></a><span data-ttu-id="45fae-479">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="45fae-479">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="45fae-480">Получает тип элемента, который представляет экземпляр.</span><span class="sxs-lookup"><span data-stu-id="45fae-480">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="45fae-481">Свойство `itemType` возвращает одно из значений перечисления `ItemType`, которое указывает, является ли экземпляр объекта `item` сообщением или встречей.</span><span class="sxs-lookup"><span data-stu-id="45fae-481">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="45fae-482">Тип</span><span class="sxs-lookup"><span data-stu-id="45fae-482">Type</span></span>

*   [<span data-ttu-id="45fae-483">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="45fae-483">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="45fae-484">Требования</span><span class="sxs-lookup"><span data-stu-id="45fae-484">Requirements</span></span>

|<span data-ttu-id="45fae-485">Требование</span><span class="sxs-lookup"><span data-stu-id="45fae-485">Requirement</span></span>|<span data-ttu-id="45fae-486">Значение</span><span class="sxs-lookup"><span data-stu-id="45fae-486">Value</span></span>|
|---|---|
|[<span data-ttu-id="45fae-487">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="45fae-487">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="45fae-488">1.0</span><span class="sxs-lookup"><span data-stu-id="45fae-488">1.0</span></span>|
|[<span data-ttu-id="45fae-489">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="45fae-489">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="45fae-490">ReadItem</span><span class="sxs-lookup"><span data-stu-id="45fae-490">ReadItem</span></span>|
|[<span data-ttu-id="45fae-491">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="45fae-491">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="45fae-492">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="45fae-492">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="45fae-493">Пример</span><span class="sxs-lookup"><span data-stu-id="45fae-493">Example</span></span>

```javascript
if (Office.context.mailbox.item.itemType === Office.MailboxEnums.ItemType.Message) {
  // Do something.
} else {
  // Do something else.
}
```

---
---

####  <a name="location-stringlocationjavascriptapioutlookofficelocation"></a><span data-ttu-id="45fae-494">location :String|[Location](/javascript/api/outlook/office.location)</span><span class="sxs-lookup"><span data-stu-id="45fae-494">location :String|[Location](/javascript/api/outlook/office.location)</span></span>

<span data-ttu-id="45fae-495">Получает или задает место встречи.</span><span class="sxs-lookup"><span data-stu-id="45fae-495">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="45fae-496">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="45fae-496">Read mode</span></span>

<span data-ttu-id="45fae-497">Свойство `location` возвращает строку, содержащую сведения о месте встречи.</span><span class="sxs-lookup"><span data-stu-id="45fae-497">The `location` property returns a string that contains the location of the appointment.</span></span>

```javascript
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="45fae-498">Режим создания</span><span class="sxs-lookup"><span data-stu-id="45fae-498">Compose mode</span></span>

<span data-ttu-id="45fae-499">Свойство `location` возвращает объект `Location`, предоставляющий методы, которые используются для получения и задания места встречи.</span><span class="sxs-lookup"><span data-stu-id="45fae-499">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```javascript
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="45fae-500">Тип</span><span class="sxs-lookup"><span data-stu-id="45fae-500">Type</span></span>

*   <span data-ttu-id="45fae-501">String | [Location](/javascript/api/outlook/office.location)</span><span class="sxs-lookup"><span data-stu-id="45fae-501">String | [Location](/javascript/api/outlook/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="45fae-502">Требования</span><span class="sxs-lookup"><span data-stu-id="45fae-502">Requirements</span></span>

|<span data-ttu-id="45fae-503">Требование</span><span class="sxs-lookup"><span data-stu-id="45fae-503">Requirement</span></span>|<span data-ttu-id="45fae-504">Значение</span><span class="sxs-lookup"><span data-stu-id="45fae-504">Value</span></span>|
|---|---|
|[<span data-ttu-id="45fae-505">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="45fae-505">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="45fae-506">1.0</span><span class="sxs-lookup"><span data-stu-id="45fae-506">1.0</span></span>|
|[<span data-ttu-id="45fae-507">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="45fae-507">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="45fae-508">ReadItem</span><span class="sxs-lookup"><span data-stu-id="45fae-508">ReadItem</span></span>|
|[<span data-ttu-id="45fae-509">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="45fae-509">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="45fae-510">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="45fae-510">Compose or Read</span></span>|

---
---

#### <a name="normalizedsubject-string"></a><span data-ttu-id="45fae-511">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="45fae-511">normalizedSubject :String</span></span>

<span data-ttu-id="45fae-p120">Получает тему элемента со всеми удаленными префиксами (включая `RE:` и `FWD:`). Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="45fae-p120">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="45fae-p121">Свойство normalizedSubject получает тему элемента со стандартными префиксами (такими как `RE:` и `FW:`), добавляемыми почтовыми программами. Для получения темы элемента с неизмененными префиксами используйте свойство [`subject`](#subject-stringsubject).</span><span class="sxs-lookup"><span data-stu-id="45fae-p121">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="45fae-516">Тип</span><span class="sxs-lookup"><span data-stu-id="45fae-516">Type</span></span>

*   <span data-ttu-id="45fae-517">String</span><span class="sxs-lookup"><span data-stu-id="45fae-517">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="45fae-518">Требования</span><span class="sxs-lookup"><span data-stu-id="45fae-518">Requirements</span></span>

|<span data-ttu-id="45fae-519">Требование</span><span class="sxs-lookup"><span data-stu-id="45fae-519">Requirement</span></span>|<span data-ttu-id="45fae-520">Значение</span><span class="sxs-lookup"><span data-stu-id="45fae-520">Value</span></span>|
|---|---|
|[<span data-ttu-id="45fae-521">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="45fae-521">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="45fae-522">1.0</span><span class="sxs-lookup"><span data-stu-id="45fae-522">1.0</span></span>|
|[<span data-ttu-id="45fae-523">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="45fae-523">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="45fae-524">ReadItem</span><span class="sxs-lookup"><span data-stu-id="45fae-524">ReadItem</span></span>|
|[<span data-ttu-id="45fae-525">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="45fae-525">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="45fae-526">Чтение</span><span class="sxs-lookup"><span data-stu-id="45fae-526">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="45fae-527">Пример</span><span class="sxs-lookup"><span data-stu-id="45fae-527">Example</span></span>

```javascript
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

---
---

####  <a name="notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessages"></a><span data-ttu-id="45fae-528">notificationMessages :[NotificationMessages](/javascript/api/outlook/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="45fae-528">notificationMessages :[NotificationMessages](/javascript/api/outlook/office.notificationmessages)</span></span>

<span data-ttu-id="45fae-529">Получает сообщения уведомления для элемента.</span><span class="sxs-lookup"><span data-stu-id="45fae-529">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="45fae-530">Тип</span><span class="sxs-lookup"><span data-stu-id="45fae-530">Type</span></span>

*   [<span data-ttu-id="45fae-531">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="45fae-531">NotificationMessages</span></span>](/javascript/api/outlook/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="45fae-532">Требования</span><span class="sxs-lookup"><span data-stu-id="45fae-532">Requirements</span></span>

|<span data-ttu-id="45fae-533">Требование</span><span class="sxs-lookup"><span data-stu-id="45fae-533">Requirement</span></span>|<span data-ttu-id="45fae-534">Значение</span><span class="sxs-lookup"><span data-stu-id="45fae-534">Value</span></span>|
|---|---|
|[<span data-ttu-id="45fae-535">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="45fae-535">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="45fae-536">1.3</span><span class="sxs-lookup"><span data-stu-id="45fae-536">1.3</span></span>|
|[<span data-ttu-id="45fae-537">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="45fae-537">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="45fae-538">ReadItem</span><span class="sxs-lookup"><span data-stu-id="45fae-538">ReadItem</span></span>|
|[<span data-ttu-id="45fae-539">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="45fae-539">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="45fae-540">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="45fae-540">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="45fae-541">Пример</span><span class="sxs-lookup"><span data-stu-id="45fae-541">Example</span></span>

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

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="45fae-542">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="45fae-542">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="45fae-543">Предоставляет доступ к необязательным участникам события.</span><span class="sxs-lookup"><span data-stu-id="45fae-543">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="45fae-544">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="45fae-544">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="45fae-545">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="45fae-545">Read mode</span></span>

<span data-ttu-id="45fae-546">Свойство `optionalAttendees` возвращает массив, содержащий объект `EmailAddressDetails` для каждого необязательного участника собрания.</span><span class="sxs-lookup"><span data-stu-id="45fae-546">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

```javascript
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="45fae-547">Режим создания</span><span class="sxs-lookup"><span data-stu-id="45fae-547">Compose mode</span></span>

<span data-ttu-id="45fae-548">Свойство `optionalAttendees` возвращает объект `Recipients`, который предоставляет методы для получения или обновления необязательных участников собрания.</span><span class="sxs-lookup"><span data-stu-id="45fae-548">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="45fae-549">Тип</span><span class="sxs-lookup"><span data-stu-id="45fae-549">Type</span></span>

*   <span data-ttu-id="45fae-550">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="45fae-550">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="45fae-551">Требования</span><span class="sxs-lookup"><span data-stu-id="45fae-551">Requirements</span></span>

|<span data-ttu-id="45fae-552">Требование</span><span class="sxs-lookup"><span data-stu-id="45fae-552">Requirement</span></span>|<span data-ttu-id="45fae-553">Значение</span><span class="sxs-lookup"><span data-stu-id="45fae-553">Value</span></span>|
|---|---|
|[<span data-ttu-id="45fae-554">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="45fae-554">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="45fae-555">1.0</span><span class="sxs-lookup"><span data-stu-id="45fae-555">1.0</span></span>|
|[<span data-ttu-id="45fae-556">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="45fae-556">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="45fae-557">ReadItem</span><span class="sxs-lookup"><span data-stu-id="45fae-557">ReadItem</span></span>|
|[<span data-ttu-id="45fae-558">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="45fae-558">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="45fae-559">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="45fae-559">Compose or Read</span></span>|

---
---

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsorganizerjavascriptapioutlookofficeorganizer"></a><span data-ttu-id="45fae-560">Организатор:[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[Организатор](/javascript/api/outlook/office.organizer)</span><span class="sxs-lookup"><span data-stu-id="45fae-560">organizer :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[Organizer](/javascript/api/outlook/office.organizer)</span></span>

<span data-ttu-id="45fae-561">Получает адрес электронной почты организатора для указанного собрания.</span><span class="sxs-lookup"><span data-stu-id="45fae-561">Gets the email address of the organizer for a specified meeting.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="45fae-562">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="45fae-562">Read mode</span></span>

<span data-ttu-id="45fae-563">`organizer` Свойство возвращает объект [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) , представляющий организатора собрания.</span><span class="sxs-lookup"><span data-stu-id="45fae-563">The `organizer` property returns an [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) object that represents the meeting organizer.</span></span>

```javascript
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

##### <a name="compose-mode"></a><span data-ttu-id="45fae-564">Режим создания</span><span class="sxs-lookup"><span data-stu-id="45fae-564">Compose mode</span></span>

<span data-ttu-id="45fae-565">Свойство возвращает объект организатора, который предоставляет метод для получения значения организатора. [](/javascript/api/outlook/office.organizer) `organizer`</span><span class="sxs-lookup"><span data-stu-id="45fae-565">The `organizer` property returns an [Organizer](/javascript/api/outlook/office.organizer) object that provides a method to get the organizer value.</span></span>

```javascript
Office.context.mailbox.item.organizer.getAsync(
  function(asyncResult) {
    console.log(JSON.stringify(asyncResult));
  }
);
```

##### <a name="type"></a><span data-ttu-id="45fae-566">Тип</span><span class="sxs-lookup"><span data-stu-id="45fae-566">Type</span></span>

*   <span data-ttu-id="45fae-567">[](/javascript/api/outlook/office.emailaddressdetails) | [Организатор](/javascript/api/outlook/office.organizer) EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="45fae-567">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [Organizer](/javascript/api/outlook/office.organizer)</span></span>

##### <a name="requirements"></a><span data-ttu-id="45fae-568">Требования</span><span class="sxs-lookup"><span data-stu-id="45fae-568">Requirements</span></span>

|<span data-ttu-id="45fae-569">Требование</span><span class="sxs-lookup"><span data-stu-id="45fae-569">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="45fae-570">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="45fae-570">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="45fae-571">1.0</span><span class="sxs-lookup"><span data-stu-id="45fae-571">1.0</span></span>|<span data-ttu-id="45fae-572">1.7</span><span class="sxs-lookup"><span data-stu-id="45fae-572">1.7</span></span>|
|[<span data-ttu-id="45fae-573">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="45fae-573">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="45fae-574">ReadItem</span><span class="sxs-lookup"><span data-stu-id="45fae-574">ReadItem</span></span>|<span data-ttu-id="45fae-575">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="45fae-575">ReadWriteItem</span></span>|
|[<span data-ttu-id="45fae-576">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="45fae-576">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="45fae-577">Чтение</span><span class="sxs-lookup"><span data-stu-id="45fae-577">Read</span></span>|<span data-ttu-id="45fae-578">Создание</span><span class="sxs-lookup"><span data-stu-id="45fae-578">Compose</span></span>|

---
---

#### <a name="nullable-recurrence-recurrencejavascriptapioutlookofficerecurrence"></a><span data-ttu-id="45fae-579">(Nullable) повторение[](/javascript/api/outlook/office.recurrence) : повторение</span><span class="sxs-lookup"><span data-stu-id="45fae-579">(nullable) recurrence :[Recurrence](/javascript/api/outlook/office.recurrence)</span></span>

<span data-ttu-id="45fae-580">Получает или задает шаблон повторения встречи.</span><span class="sxs-lookup"><span data-stu-id="45fae-580">Gets or sets the recurrence pattern of an appointment.</span></span> <span data-ttu-id="45fae-581">Получает шаблон повторения приглашения на собрание.</span><span class="sxs-lookup"><span data-stu-id="45fae-581">Gets the recurrence pattern of a meeting request.</span></span> <span data-ttu-id="45fae-582">Режимы чтения и создания для элементов встречи.</span><span class="sxs-lookup"><span data-stu-id="45fae-582">Read and compose modes for appointment items.</span></span> <span data-ttu-id="45fae-583">Режим чтения для элементов приглашения на собрания.</span><span class="sxs-lookup"><span data-stu-id="45fae-583">Read mode for meeting request items.</span></span>

<span data-ttu-id="45fae-584">`recurrence` Свойство возвращает объект [повторения](/javascript/api/outlook/office.recurrence) для повторяющихся встреч или приглашений на собрания, если элемент представляет собой серию или экземпляр в ряду.</span><span class="sxs-lookup"><span data-stu-id="45fae-584">The `recurrence` property returns a [recurrence](/javascript/api/outlook/office.recurrence) object for recurring appointments or meetings requests if an item is a series or an instance in a series.</span></span> <span data-ttu-id="45fae-585">`null`возвращается для отдельных встреч и приглашений на собрание для отдельных встреч.</span><span class="sxs-lookup"><span data-stu-id="45fae-585">`null` is returned for single appointments and meeting requests of single appointments.</span></span> <span data-ttu-id="45fae-586">`undefined`возвращается для сообщений, которые не являются приглашениями на собрания.</span><span class="sxs-lookup"><span data-stu-id="45fae-586">`undefined` is returned for messages that are not meeting requests.</span></span>

> <span data-ttu-id="45fae-587">Note: приглашения на `itemClass` собрания имеют значение IPM. Schedule. Meeting. Request.</span><span class="sxs-lookup"><span data-stu-id="45fae-587">Note: Meeting requests have an `itemClass` value of IPM.Schedule.Meeting.Request.</span></span>

> <span data-ttu-id="45fae-588">Note: при наличии объекта `null`повторения это указывает на то, что объект является одной встречей или приглашением на собрание одной встречи, а не частью ряда.</span><span class="sxs-lookup"><span data-stu-id="45fae-588">Note: If the recurrence object is `null`, this indicates that the object is a single appointment or a meeting request of a single appointment and NOT a part of a series.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="45fae-589">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="45fae-589">Read mode</span></span>

<span data-ttu-id="45fae-590">`recurrence` Свойство возвращает объект [повторения](/javascript/api/outlook/office.recurrence) , представляющий повторение встречи.</span><span class="sxs-lookup"><span data-stu-id="45fae-590">The `recurrence` property returns a [Recurrence](/javascript/api/outlook/office.recurrence) object that represents the appointment recurrence.</span></span> <span data-ttu-id="45fae-591">Оно доступно для встреч и приглашений на собрания.</span><span class="sxs-lookup"><span data-stu-id="45fae-591">This is available for appointments and meeting requests.</span></span>

```javascript
var recurrence = Office.context.mailbox.item.recurrence;
console.log("Recurrence: " + JSON.stringify(recurrence));
```

##### <a name="compose-mode"></a><span data-ttu-id="45fae-592">Режим создания</span><span class="sxs-lookup"><span data-stu-id="45fae-592">Compose mode</span></span>

<span data-ttu-id="45fae-593">`recurrence` Свойство возвращает объект [повторения](/javascript/api/outlook/office.recurrence) , который предоставляет методы для управления повторением встречи.</span><span class="sxs-lookup"><span data-stu-id="45fae-593">The `recurrence` property returns a [Recurrence](/javascript/api/outlook/office.recurrence) object that provides methods to manage the appointment recurrence.</span></span> <span data-ttu-id="45fae-594">Оно доступно для встреч.</span><span class="sxs-lookup"><span data-stu-id="45fae-594">This is available for appointments.</span></span>

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

##### <a name="type"></a><span data-ttu-id="45fae-595">Тип</span><span class="sxs-lookup"><span data-stu-id="45fae-595">Type</span></span>

* [<span data-ttu-id="45fae-596">Повторения</span><span class="sxs-lookup"><span data-stu-id="45fae-596">Recurrence</span></span>](/javascript/api/outlook/office.recurrence)

|<span data-ttu-id="45fae-597">Требование</span><span class="sxs-lookup"><span data-stu-id="45fae-597">Requirement</span></span>|<span data-ttu-id="45fae-598">Значение</span><span class="sxs-lookup"><span data-stu-id="45fae-598">Value</span></span>|
|---|---|
|[<span data-ttu-id="45fae-599">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="45fae-599">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="45fae-600">1.7</span><span class="sxs-lookup"><span data-stu-id="45fae-600">1.7</span></span>|
|[<span data-ttu-id="45fae-601">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="45fae-601">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="45fae-602">ReadItem</span><span class="sxs-lookup"><span data-stu-id="45fae-602">ReadItem</span></span>|
|[<span data-ttu-id="45fae-603">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="45fae-603">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="45fae-604">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="45fae-604">Compose or Read</span></span>|

---
---

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="45fae-605">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="45fae-605">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="45fae-606">Предоставляет доступ к обязательным участникам события.</span><span class="sxs-lookup"><span data-stu-id="45fae-606">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="45fae-607">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="45fae-607">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="45fae-608">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="45fae-608">Read mode</span></span>

<span data-ttu-id="45fae-609">Свойство `requiredAttendees` возвращает массив, содержащий объект `EmailAddressDetails` для каждого обязательного участника собрания.</span><span class="sxs-lookup"><span data-stu-id="45fae-609">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

```javascript
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="45fae-610">Режим создания</span><span class="sxs-lookup"><span data-stu-id="45fae-610">Compose mode</span></span>

<span data-ttu-id="45fae-611">Свойство `requiredAttendees` возвращает объект `Recipients`, предоставляющий методы, с помощью которых можно получить или обновить сведения об обязательных участниках собрания.</span><span class="sxs-lookup"><span data-stu-id="45fae-611">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="45fae-612">Тип</span><span class="sxs-lookup"><span data-stu-id="45fae-612">Type</span></span>

*   <span data-ttu-id="45fae-613">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="45fae-613">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="45fae-614">Требования</span><span class="sxs-lookup"><span data-stu-id="45fae-614">Requirements</span></span>

|<span data-ttu-id="45fae-615">Требование</span><span class="sxs-lookup"><span data-stu-id="45fae-615">Requirement</span></span>|<span data-ttu-id="45fae-616">Значение</span><span class="sxs-lookup"><span data-stu-id="45fae-616">Value</span></span>|
|---|---|
|[<span data-ttu-id="45fae-617">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="45fae-617">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="45fae-618">1.0</span><span class="sxs-lookup"><span data-stu-id="45fae-618">1.0</span></span>|
|[<span data-ttu-id="45fae-619">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="45fae-619">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="45fae-620">ReadItem</span><span class="sxs-lookup"><span data-stu-id="45fae-620">ReadItem</span></span>|
|[<span data-ttu-id="45fae-621">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="45fae-621">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="45fae-622">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="45fae-622">Compose or Read</span></span>|

---
---

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails"></a><span data-ttu-id="45fae-623">sender :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="45fae-623">sender :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)</span></span>

<span data-ttu-id="45fae-p128">Получает электронный адрес отправителя электронного сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="45fae-p128">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="45fae-p129">Свойства [`from`](#from-emailaddressdetailsfrom) и `sender` представляют одно лицо, если сообщение не отправлено представителем. В противном случае свойство `from` представляет лицо, делегировавшее полномочия, а свойство sender — представителя.</span><span class="sxs-lookup"><span data-stu-id="45fae-p129">The [`from`](#from-emailaddressdetailsfrom) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="45fae-628">Свойству `recipientType`, принадлежащему объекту `EmailAddressDetails` в свойстве `sender`, задано значение `undefined`.</span><span class="sxs-lookup"><span data-stu-id="45fae-628">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="45fae-629">Тип</span><span class="sxs-lookup"><span data-stu-id="45fae-629">Type</span></span>

*   [<span data-ttu-id="45fae-630">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="45fae-630">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="45fae-631">Требования</span><span class="sxs-lookup"><span data-stu-id="45fae-631">Requirements</span></span>

|<span data-ttu-id="45fae-632">Требование</span><span class="sxs-lookup"><span data-stu-id="45fae-632">Requirement</span></span>|<span data-ttu-id="45fae-633">Значение</span><span class="sxs-lookup"><span data-stu-id="45fae-633">Value</span></span>|
|---|---|
|[<span data-ttu-id="45fae-634">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="45fae-634">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="45fae-635">1.0</span><span class="sxs-lookup"><span data-stu-id="45fae-635">1.0</span></span>|
|[<span data-ttu-id="45fae-636">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="45fae-636">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="45fae-637">ReadItem</span><span class="sxs-lookup"><span data-stu-id="45fae-637">ReadItem</span></span>|
|[<span data-ttu-id="45fae-638">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="45fae-638">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="45fae-639">Чтение</span><span class="sxs-lookup"><span data-stu-id="45fae-639">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="45fae-640">Пример</span><span class="sxs-lookup"><span data-stu-id="45fae-640">Example</span></span>

```javascript
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

---
---

#### <a name="nullable-seriesid-string"></a><span data-ttu-id="45fae-641">(Nullable) seriesId: строка</span><span class="sxs-lookup"><span data-stu-id="45fae-641">(nullable) seriesId :String</span></span>

<span data-ttu-id="45fae-642">Получает идентификатор ряда, к которому принадлежит экземпляр.</span><span class="sxs-lookup"><span data-stu-id="45fae-642">Gets the id of the series that an instance belongs to.</span></span>

<span data-ttu-id="45fae-643">В OWA и Outlook `seriesId` возвращается идентификатор веб-служб Exchange (EWS) родительского элемента (ряда), к которому принадлежит этот элемент.</span><span class="sxs-lookup"><span data-stu-id="45fae-643">In OWA and Outlook, the `seriesId` returns the Exchange Web Services (EWS) ID of the parent (series) item that this item belongs to.</span></span> <span data-ttu-id="45fae-644">Однако в iOS и Android `seriesId` возвращается идентификатор REST родительского элемента.</span><span class="sxs-lookup"><span data-stu-id="45fae-644">However, in iOS and Android, the `seriesId` returns the REST ID of the parent item.</span></span>

> [!NOTE]
> <span data-ttu-id="45fae-645">Идентификатор, возвращаемый свойством `seriesId`, совпадает с идентификатором элемента веб-служб Exchange.</span><span class="sxs-lookup"><span data-stu-id="45fae-645">The identifier returned by the `seriesId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="45fae-646">`seriesId` Свойство не совпадает с идентификаторами Outlook, используемыми в REST API Outlook.</span><span class="sxs-lookup"><span data-stu-id="45fae-646">The `seriesId` property is not identical to the Outlook IDs used by the Outlook REST API.</span></span> <span data-ttu-id="45fae-647">Прежде чем совершать вызовы REST API, используя это значение, его необходимо преобразовать с помощью [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="45fae-647">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="45fae-648">Дополнительные сведения см. в статье [Использование REST API Outlook из надстройки Outlook](/outlook/add-ins/use-rest-api).</span><span class="sxs-lookup"><span data-stu-id="45fae-648">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api).</span></span>

<span data-ttu-id="45fae-649">`seriesId` Свойство возвращает `null` элементы, у которых нет родительских элементов, таких как одиночные встречи, элементы ряда или приглашения на собрание, `undefined` и возвращаемые для других элементов, не являющиеся приглашениями на собрания.</span><span class="sxs-lookup"><span data-stu-id="45fae-649">The `seriesId` property returns `null` for items that do not have parent items such as single appointments, series items, or meeting requests and returns `undefined` for any other items that are not meeting requests.</span></span>

##### <a name="type"></a><span data-ttu-id="45fae-650">Тип</span><span class="sxs-lookup"><span data-stu-id="45fae-650">Type</span></span>

* <span data-ttu-id="45fae-651">String</span><span class="sxs-lookup"><span data-stu-id="45fae-651">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="45fae-652">Требования</span><span class="sxs-lookup"><span data-stu-id="45fae-652">Requirements</span></span>

|<span data-ttu-id="45fae-653">Требование</span><span class="sxs-lookup"><span data-stu-id="45fae-653">Requirement</span></span>|<span data-ttu-id="45fae-654">Значение</span><span class="sxs-lookup"><span data-stu-id="45fae-654">Value</span></span>|
|---|---|
|[<span data-ttu-id="45fae-655">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="45fae-655">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="45fae-656">1.7</span><span class="sxs-lookup"><span data-stu-id="45fae-656">1.7</span></span>|
|[<span data-ttu-id="45fae-657">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="45fae-657">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="45fae-658">ReadItem</span><span class="sxs-lookup"><span data-stu-id="45fae-658">ReadItem</span></span>|
|[<span data-ttu-id="45fae-659">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="45fae-659">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="45fae-660">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="45fae-660">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="45fae-661">Пример</span><span class="sxs-lookup"><span data-stu-id="45fae-661">Example</span></span>

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

####  <a name="start-datetimejavascriptapioutlookofficetime"></a><span data-ttu-id="45fae-662">start :Date|[Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="45fae-662">start :Date|[Time](/javascript/api/outlook/office.time)</span></span>

<span data-ttu-id="45fae-663">Получает или задает дату и время начала встречи.</span><span class="sxs-lookup"><span data-stu-id="45fae-663">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="45fae-p132">Свойство `start` представлено в виде значения даты и времени в формате UTC. Это значение можно преобразовать в местные значения даты и времени клиента с помощью метода [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime).</span><span class="sxs-lookup"><span data-stu-id="45fae-p132">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="45fae-666">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="45fae-666">Read mode</span></span>

<span data-ttu-id="45fae-667">Свойство `start` возвращает объект `Date`.</span><span class="sxs-lookup"><span data-stu-id="45fae-667">The `start` property returns a `Date` object.</span></span>

```javascript
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="45fae-668">Режим создания</span><span class="sxs-lookup"><span data-stu-id="45fae-668">Compose mode</span></span>

<span data-ttu-id="45fae-669">Свойство `start` возвращает объект `Time`.</span><span class="sxs-lookup"><span data-stu-id="45fae-669">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="45fae-670">Если вы задаете время начала с помощью метода [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-), необходимо использовать метод [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) для преобразования местного времени на клиенте в формат UTC для сервера.</span><span class="sxs-lookup"><span data-stu-id="45fae-670">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="45fae-671">В примере ниже с помощью метода [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) объекта `Time` задается время начала встречи в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="45fae-671">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="45fae-672">Тип</span><span class="sxs-lookup"><span data-stu-id="45fae-672">Type</span></span>

*   <span data-ttu-id="45fae-673">Date | [Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="45fae-673">Date | [Time](/javascript/api/outlook/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="45fae-674">Требования</span><span class="sxs-lookup"><span data-stu-id="45fae-674">Requirements</span></span>

|<span data-ttu-id="45fae-675">Требование</span><span class="sxs-lookup"><span data-stu-id="45fae-675">Requirement</span></span>|<span data-ttu-id="45fae-676">Значение</span><span class="sxs-lookup"><span data-stu-id="45fae-676">Value</span></span>|
|---|---|
|[<span data-ttu-id="45fae-677">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="45fae-677">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="45fae-678">1.0</span><span class="sxs-lookup"><span data-stu-id="45fae-678">1.0</span></span>|
|[<span data-ttu-id="45fae-679">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="45fae-679">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="45fae-680">ReadItem</span><span class="sxs-lookup"><span data-stu-id="45fae-680">ReadItem</span></span>|
|[<span data-ttu-id="45fae-681">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="45fae-681">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="45fae-682">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="45fae-682">Compose or Read</span></span>|

---
---

####  <a name="subject-stringsubjectjavascriptapioutlookofficesubject"></a><span data-ttu-id="45fae-683">subject :String|[Subject](/javascript/api/outlook/office.subject)</span><span class="sxs-lookup"><span data-stu-id="45fae-683">subject :String|[Subject](/javascript/api/outlook/office.subject)</span></span>

<span data-ttu-id="45fae-684">Получает или задает описание, которое отображается в поле темы элемента.</span><span class="sxs-lookup"><span data-stu-id="45fae-684">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="45fae-685">Свойство `subject` получает или задает всю тему элемента для отправки с почтового сервера.</span><span class="sxs-lookup"><span data-stu-id="45fae-685">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="45fae-686">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="45fae-686">Read mode</span></span>

<span data-ttu-id="45fae-p133">Свойство `subject` возвращает строку. С помощью свойства [`normalizedSubject`](#normalizedsubject-string) можно получить тему без начальных префиксов, таких как `RE:` и `FW:`.</span><span class="sxs-lookup"><span data-stu-id="45fae-p133">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

<span data-ttu-id="45fae-689">В примере кода JavaScript, приведенном ниже, показано, как получить доступ к свойству `subject` текущего элемента в Outlook.</span><span class="sxs-lookup"><span data-stu-id="45fae-689">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

```javascript
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="45fae-690">Режим создания</span><span class="sxs-lookup"><span data-stu-id="45fae-690">Compose mode</span></span>
<span data-ttu-id="45fae-691">Свойство `subject` возвращает объект `Subject`, который предоставляет методы для получения и задания темы.</span><span class="sxs-lookup"><span data-stu-id="45fae-691">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```javascript
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="45fae-692">Тип</span><span class="sxs-lookup"><span data-stu-id="45fae-692">Type</span></span>

*   <span data-ttu-id="45fae-693">String | [Subject](/javascript/api/outlook/office.subject)</span><span class="sxs-lookup"><span data-stu-id="45fae-693">String | [Subject](/javascript/api/outlook/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="45fae-694">Требования</span><span class="sxs-lookup"><span data-stu-id="45fae-694">Requirements</span></span>

|<span data-ttu-id="45fae-695">Требование</span><span class="sxs-lookup"><span data-stu-id="45fae-695">Requirement</span></span>|<span data-ttu-id="45fae-696">Значение</span><span class="sxs-lookup"><span data-stu-id="45fae-696">Value</span></span>|
|---|---|
|[<span data-ttu-id="45fae-697">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="45fae-697">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="45fae-698">1.0</span><span class="sxs-lookup"><span data-stu-id="45fae-698">1.0</span></span>|
|[<span data-ttu-id="45fae-699">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="45fae-699">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="45fae-700">ReadItem</span><span class="sxs-lookup"><span data-stu-id="45fae-700">ReadItem</span></span>|
|[<span data-ttu-id="45fae-701">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="45fae-701">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="45fae-702">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="45fae-702">Compose or Read</span></span>|

---
---

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="45fae-703">to :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="45fae-703">to :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="45fae-704">Предоставляет доступ к получателям, указанным в строке **Кому** сообщения.</span><span class="sxs-lookup"><span data-stu-id="45fae-704">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="45fae-705">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="45fae-705">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="45fae-706">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="45fae-706">Read mode</span></span>

<span data-ttu-id="45fae-p135">Свойство `to` возвращает массив, содержащий объект `EmailAddressDetails` для каждого получателя в строке **Кому** сообщения. Коллекция может включать не более 100 элементов.</span><span class="sxs-lookup"><span data-stu-id="45fae-p135">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="45fae-709">Режим создания</span><span class="sxs-lookup"><span data-stu-id="45fae-709">Compose mode</span></span>

<span data-ttu-id="45fae-710">Свойство `to` возвращает объект `Recipients`, предоставляющий методы для получения или обновления получателей, которые указаны в строке **Кому** сообщения.</span><span class="sxs-lookup"><span data-stu-id="45fae-710">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

```javascript
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="45fae-711">Тип</span><span class="sxs-lookup"><span data-stu-id="45fae-711">Type</span></span>

*   <span data-ttu-id="45fae-712">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="45fae-712">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="45fae-713">Требования</span><span class="sxs-lookup"><span data-stu-id="45fae-713">Requirements</span></span>

|<span data-ttu-id="45fae-714">Требование</span><span class="sxs-lookup"><span data-stu-id="45fae-714">Requirement</span></span>|<span data-ttu-id="45fae-715">Значение</span><span class="sxs-lookup"><span data-stu-id="45fae-715">Value</span></span>|
|---|---|
|[<span data-ttu-id="45fae-716">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="45fae-716">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="45fae-717">1.0</span><span class="sxs-lookup"><span data-stu-id="45fae-717">1.0</span></span>|
|[<span data-ttu-id="45fae-718">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="45fae-718">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="45fae-719">ReadItem</span><span class="sxs-lookup"><span data-stu-id="45fae-719">ReadItem</span></span>|
|[<span data-ttu-id="45fae-720">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="45fae-720">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="45fae-721">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="45fae-721">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="45fae-722">Методы</span><span class="sxs-lookup"><span data-stu-id="45fae-722">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="45fae-723">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="45fae-723">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="45fae-724">Добавляет файл в сообщение или встречу в качестве вложения.</span><span class="sxs-lookup"><span data-stu-id="45fae-724">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="45fae-725">Метод `addFileAttachmentAsync` передает файл по указанному универсальному коду ресурса (URI) и вкладывает его в элемент в форме создания.</span><span class="sxs-lookup"><span data-stu-id="45fae-725">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="45fae-726">Идентификатор можно использовать с методом [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback), чтобы удалить вложение, добавленное во время текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="45fae-726">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="45fae-727">Параметры</span><span class="sxs-lookup"><span data-stu-id="45fae-727">Parameters</span></span>
|<span data-ttu-id="45fae-728">Имя</span><span class="sxs-lookup"><span data-stu-id="45fae-728">Name</span></span>|<span data-ttu-id="45fae-729">Тип</span><span class="sxs-lookup"><span data-stu-id="45fae-729">Type</span></span>|<span data-ttu-id="45fae-730">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="45fae-730">Attributes</span></span>|<span data-ttu-id="45fae-731">Описание</span><span class="sxs-lookup"><span data-stu-id="45fae-731">Description</span></span>|
|---|---|---|---|
|`uri`|<span data-ttu-id="45fae-732">String</span><span class="sxs-lookup"><span data-stu-id="45fae-732">String</span></span>||<span data-ttu-id="45fae-p136">Универсальный код ресурса (URI), представляющий расположение файла, который нужно вложить в сообщение или встречу. Максимальная длина — 2048 символов.</span><span class="sxs-lookup"><span data-stu-id="45fae-p136">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="45fae-735">String</span><span class="sxs-lookup"><span data-stu-id="45fae-735">String</span></span>||<span data-ttu-id="45fae-p137">Имя вложения, которое отображается при передаче вложения. Максимальная длина — 255 символов.</span><span class="sxs-lookup"><span data-stu-id="45fae-p137">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="45fae-738">Объект</span><span class="sxs-lookup"><span data-stu-id="45fae-738">Object</span></span>|<span data-ttu-id="45fae-739">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="45fae-739">&lt;optional&gt;</span></span>|<span data-ttu-id="45fae-740">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="45fae-740">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="45fae-741">Object</span><span class="sxs-lookup"><span data-stu-id="45fae-741">Object</span></span>|<span data-ttu-id="45fae-742">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="45fae-742">&lt;optional&gt;</span></span>|<span data-ttu-id="45fae-743">В методе обратного вызова разработчики могут указать любой объект, к которому необходимо получить доступ.</span><span class="sxs-lookup"><span data-stu-id="45fae-743">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="45fae-744">Boolean</span><span class="sxs-lookup"><span data-stu-id="45fae-744">Boolean</span></span>|<span data-ttu-id="45fae-745">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="45fae-745">&lt;optional&gt;</span></span>|<span data-ttu-id="45fae-746">Значение `true` указывает на то, что вложение будет встроено в текст сообщения и не должно отображаться в списке вложений.</span><span class="sxs-lookup"><span data-stu-id="45fae-746">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="45fae-747">function</span><span class="sxs-lookup"><span data-stu-id="45fae-747">function</span></span>|<span data-ttu-id="45fae-748">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="45fae-748">&lt;optional&gt;</span></span>|<span data-ttu-id="45fae-749">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="45fae-749">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="45fae-750">После успешного выполнения идентификатор вложения будет представлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="45fae-750">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="45fae-751">Если добавить вложение не удастся, объект `asyncResult` будет содержать объект `Error` с описанием ошибки.</span><span class="sxs-lookup"><span data-stu-id="45fae-751">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="45fae-752">Ошибки</span><span class="sxs-lookup"><span data-stu-id="45fae-752">Errors</span></span>

|<span data-ttu-id="45fae-753">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="45fae-753">Error code</span></span>|<span data-ttu-id="45fae-754">Описание</span><span class="sxs-lookup"><span data-stu-id="45fae-754">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="45fae-755">Вложение превышает максимальный размер.</span><span class="sxs-lookup"><span data-stu-id="45fae-755">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="45fae-756">Расширение вложения не поддерживается.</span><span class="sxs-lookup"><span data-stu-id="45fae-756">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="45fae-757">Сообщение или встреча содержат слишком много вложений.</span><span class="sxs-lookup"><span data-stu-id="45fae-757">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="45fae-758">Требования</span><span class="sxs-lookup"><span data-stu-id="45fae-758">Requirements</span></span>

|<span data-ttu-id="45fae-759">Требование</span><span class="sxs-lookup"><span data-stu-id="45fae-759">Requirement</span></span>|<span data-ttu-id="45fae-760">Значение</span><span class="sxs-lookup"><span data-stu-id="45fae-760">Value</span></span>|
|---|---|
|[<span data-ttu-id="45fae-761">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="45fae-761">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="45fae-762">1.1</span><span class="sxs-lookup"><span data-stu-id="45fae-762">1.1</span></span>|
|[<span data-ttu-id="45fae-763">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="45fae-763">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="45fae-764">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="45fae-764">ReadWriteItem</span></span>|
|[<span data-ttu-id="45fae-765">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="45fae-765">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="45fae-766">Создание</span><span class="sxs-lookup"><span data-stu-id="45fae-766">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="45fae-767">Примеры</span><span class="sxs-lookup"><span data-stu-id="45fae-767">Examples</span></span>

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

<span data-ttu-id="45fae-768">В приведенном ниже примере файл изображения добавляется в качестве встроенного вложения, а ссылка на вложение добавляется в текст сообщения.</span><span class="sxs-lookup"><span data-stu-id="45fae-768">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

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

#### <a name="addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback"></a><span data-ttu-id="45fae-769">addFileAttachmentFromBase64Async (base64File, Аттачментнаме, [параметры], [обратный вызов])</span><span class="sxs-lookup"><span data-stu-id="45fae-769">addFileAttachmentFromBase64Async(base64File, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="45fae-770">Добавляет файл из кодировки Base64 в сообщение или встречу в качестве вложения.</span><span class="sxs-lookup"><span data-stu-id="45fae-770">Adds a file from the base64 encoding to a message or appointment as an attachment.</span></span>

<span data-ttu-id="45fae-771">`addFileAttachmentFromBase64Async` Метод передает файл из кодировки Base64 и вкладывает его в элемент в форме создания.</span><span class="sxs-lookup"><span data-stu-id="45fae-771">The `addFileAttachmentFromBase64Async` method uploads the file from the base64 encoding and attaches it to the item in the compose form.</span></span> <span data-ttu-id="45fae-772">Этот метод возвращает идентификатор вложения в объекте AsyncResult. Value.</span><span class="sxs-lookup"><span data-stu-id="45fae-772">This method returns the attachment identifier in the AsyncResult.value object.</span></span>

<span data-ttu-id="45fae-773">Идентификатор можно использовать с методом [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback), чтобы удалить вложение, добавленное во время текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="45fae-773">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="45fae-774">Параметры</span><span class="sxs-lookup"><span data-stu-id="45fae-774">Parameters</span></span>

|<span data-ttu-id="45fae-775">Имя</span><span class="sxs-lookup"><span data-stu-id="45fae-775">Name</span></span>|<span data-ttu-id="45fae-776">Тип</span><span class="sxs-lookup"><span data-stu-id="45fae-776">Type</span></span>|<span data-ttu-id="45fae-777">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="45fae-777">Attributes</span></span>|<span data-ttu-id="45fae-778">Описание</span><span class="sxs-lookup"><span data-stu-id="45fae-778">Description</span></span>|
|---|---|---|---|
|`base64File`|<span data-ttu-id="45fae-779">String</span><span class="sxs-lookup"><span data-stu-id="45fae-779">String</span></span>||<span data-ttu-id="45fae-780">Содержимое изображения или файла в кодировке Base64, которое добавляется в сообщение электронной почты или событие.</span><span class="sxs-lookup"><span data-stu-id="45fae-780">The base64 encoded content of an image or file to be added to an email or event.</span></span>|
|`attachmentName`|<span data-ttu-id="45fae-781">String</span><span class="sxs-lookup"><span data-stu-id="45fae-781">String</span></span>||<span data-ttu-id="45fae-p139">Имя вложения, которое отображается при передаче вложения. Максимальная длина — 255 символов.</span><span class="sxs-lookup"><span data-stu-id="45fae-p139">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="45fae-784">Объект</span><span class="sxs-lookup"><span data-stu-id="45fae-784">Object</span></span>|<span data-ttu-id="45fae-785">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="45fae-785">&lt;optional&gt;</span></span>|<span data-ttu-id="45fae-786">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="45fae-786">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="45fae-787">Object</span><span class="sxs-lookup"><span data-stu-id="45fae-787">Object</span></span>|<span data-ttu-id="45fae-788">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="45fae-788">&lt;optional&gt;</span></span>|<span data-ttu-id="45fae-789">В методе обратного вызова разработчики могут указать любой объект, к которому необходимо получить доступ.</span><span class="sxs-lookup"><span data-stu-id="45fae-789">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="45fae-790">Boolean</span><span class="sxs-lookup"><span data-stu-id="45fae-790">Boolean</span></span>|<span data-ttu-id="45fae-791">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="45fae-791">&lt;optional&gt;</span></span>|<span data-ttu-id="45fae-792">Значение `true` указывает на то, что вложение будет встроено в текст сообщения и не должно отображаться в списке вложений.</span><span class="sxs-lookup"><span data-stu-id="45fae-792">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="45fae-793">function</span><span class="sxs-lookup"><span data-stu-id="45fae-793">function</span></span>|<span data-ttu-id="45fae-794">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="45fae-794">&lt;optional&gt;</span></span>|<span data-ttu-id="45fae-795">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="45fae-795">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="45fae-796">После успешного выполнения идентификатор вложения будет представлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="45fae-796">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="45fae-797">Если добавить вложение не удастся, объект `asyncResult` будет содержать объект `Error` с описанием ошибки.</span><span class="sxs-lookup"><span data-stu-id="45fae-797">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="45fae-798">Ошибки</span><span class="sxs-lookup"><span data-stu-id="45fae-798">Errors</span></span>

|<span data-ttu-id="45fae-799">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="45fae-799">Error code</span></span>|<span data-ttu-id="45fae-800">Описание</span><span class="sxs-lookup"><span data-stu-id="45fae-800">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="45fae-801">Вложение превышает максимальный размер.</span><span class="sxs-lookup"><span data-stu-id="45fae-801">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="45fae-802">Расширение вложения не поддерживается.</span><span class="sxs-lookup"><span data-stu-id="45fae-802">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="45fae-803">Сообщение или встреча содержат слишком много вложений.</span><span class="sxs-lookup"><span data-stu-id="45fae-803">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="45fae-804">Требования</span><span class="sxs-lookup"><span data-stu-id="45fae-804">Requirements</span></span>

|<span data-ttu-id="45fae-805">Требование</span><span class="sxs-lookup"><span data-stu-id="45fae-805">Requirement</span></span>|<span data-ttu-id="45fae-806">Значение</span><span class="sxs-lookup"><span data-stu-id="45fae-806">Value</span></span>|
|---|---|
|[<span data-ttu-id="45fae-807">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="45fae-807">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="45fae-808">Предварительная версия</span><span class="sxs-lookup"><span data-stu-id="45fae-808">Preview</span></span>|
|[<span data-ttu-id="45fae-809">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="45fae-809">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="45fae-810">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="45fae-810">ReadWriteItem</span></span>|
|[<span data-ttu-id="45fae-811">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="45fae-811">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="45fae-812">Создание</span><span class="sxs-lookup"><span data-stu-id="45fae-812">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="45fae-813">Примеры</span><span class="sxs-lookup"><span data-stu-id="45fae-813">Examples</span></span>

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

####  <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="45fae-814">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="45fae-814">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="45fae-815">Добавляет обработчик для поддерживаемого события.</span><span class="sxs-lookup"><span data-stu-id="45fae-815">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="45fae-816">В настоящее время поддерживаются типы `Office.EventType.AttachmentsChanged`событий `Office.EventType.AppointmentTimeChanged`, `Office.EventType.EnhancedLocationsChanged` `Office.EventType.RecipientsChanged`,, и `Office.EventType.RecurrenceChanged`.</span><span class="sxs-lookup"><span data-stu-id="45fae-816">Currently the supported event types are `Office.EventType.AttachmentsChanged`, `Office.EventType.AppointmentTimeChanged`, `Office.EventType.EnhancedLocationsChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="45fae-817">Параметры</span><span class="sxs-lookup"><span data-stu-id="45fae-817">Parameters</span></span>

| <span data-ttu-id="45fae-818">Имя</span><span class="sxs-lookup"><span data-stu-id="45fae-818">Name</span></span> | <span data-ttu-id="45fae-819">Тип</span><span class="sxs-lookup"><span data-stu-id="45fae-819">Type</span></span> | <span data-ttu-id="45fae-820">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="45fae-820">Attributes</span></span> | <span data-ttu-id="45fae-821">Описание</span><span class="sxs-lookup"><span data-stu-id="45fae-821">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="45fae-822">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="45fae-822">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="45fae-823">Событие, которое должно вызвать обработчик.</span><span class="sxs-lookup"><span data-stu-id="45fae-823">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="45fae-824">Function</span><span class="sxs-lookup"><span data-stu-id="45fae-824">Function</span></span> || <span data-ttu-id="45fae-p140">Функция для обработки события. Функция должна принимать один параметр, представляющий собой объектный литерал. Значение свойства `type` параметра совпадет со значением параметра `eventType`, переданного методу `addHandlerAsync`.</span><span class="sxs-lookup"><span data-stu-id="45fae-p140">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="45fae-828">Объект</span><span class="sxs-lookup"><span data-stu-id="45fae-828">Object</span></span> | <span data-ttu-id="45fae-829">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="45fae-829">&lt;optional&gt;</span></span> | <span data-ttu-id="45fae-830">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="45fae-830">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="45fae-831">Объект</span><span class="sxs-lookup"><span data-stu-id="45fae-831">Object</span></span> | <span data-ttu-id="45fae-832">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="45fae-832">&lt;optional&gt;</span></span> | <span data-ttu-id="45fae-833">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="45fae-833">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="45fae-834">функция</span><span class="sxs-lookup"><span data-stu-id="45fae-834">function</span></span>| <span data-ttu-id="45fae-835">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="45fae-835">&lt;optional&gt;</span></span>|<span data-ttu-id="45fae-836">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="45fae-836">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="45fae-837">Требования</span><span class="sxs-lookup"><span data-stu-id="45fae-837">Requirements</span></span>

|<span data-ttu-id="45fae-838">Требование</span><span class="sxs-lookup"><span data-stu-id="45fae-838">Requirement</span></span>| <span data-ttu-id="45fae-839">Значение</span><span class="sxs-lookup"><span data-stu-id="45fae-839">Value</span></span>|
|---|---|
|[<span data-ttu-id="45fae-840">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="45fae-840">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="45fae-841">1.7</span><span class="sxs-lookup"><span data-stu-id="45fae-841">1.7</span></span> |
|[<span data-ttu-id="45fae-842">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="45fae-842">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="45fae-843">ReadItem</span><span class="sxs-lookup"><span data-stu-id="45fae-843">ReadItem</span></span> |
|[<span data-ttu-id="45fae-844">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="45fae-844">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="45fae-845">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="45fae-845">Compose or Read</span></span> |

##### <a name="example"></a><span data-ttu-id="45fae-846">Пример</span><span class="sxs-lookup"><span data-stu-id="45fae-846">Example</span></span>

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

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="45fae-847">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="45fae-847">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="45fae-848">Добавляет к сообщению элемент Exchange, например сообщение, в виде вложения.</span><span class="sxs-lookup"><span data-stu-id="45fae-848">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="45fae-p141">С помощью метода `addItemAttachmentAsync` можно в элемент формы создания вложить элемент с указанным идентификатором Exchange. Если указать метод обратного вызова, то этот метод вызывается с помощью параметра `asyncResult`, который содержит идентификатор вложения или код, указывающий на ошибки, которые произошли при вложении элемента. При необходимости можно использовать параметр `options` для передачи сведений о состоянии методу обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="45fae-p141">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="45fae-852">Идентификатор можно использовать с методом [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback), чтобы удалить вложение, добавленное во время текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="45fae-852">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="45fae-853">Если ваша надстройка Office выполняется в Outlook Web App, метод `addItemAttachmentAsync` обеспечивает вложение элементов в элементы, отличные от редактируемого. Однако это действие не рекомендуем выполнять, так как оно не поддерживается.</span><span class="sxs-lookup"><span data-stu-id="45fae-853">If your Office Add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="45fae-854">Параметры</span><span class="sxs-lookup"><span data-stu-id="45fae-854">Parameters</span></span>

|<span data-ttu-id="45fae-855">Имя</span><span class="sxs-lookup"><span data-stu-id="45fae-855">Name</span></span>|<span data-ttu-id="45fae-856">Тип</span><span class="sxs-lookup"><span data-stu-id="45fae-856">Type</span></span>|<span data-ttu-id="45fae-857">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="45fae-857">Attributes</span></span>|<span data-ttu-id="45fae-858">Описание</span><span class="sxs-lookup"><span data-stu-id="45fae-858">Description</span></span>|
|---|---|---|---|
|`itemId`|<span data-ttu-id="45fae-859">String</span><span class="sxs-lookup"><span data-stu-id="45fae-859">String</span></span>||<span data-ttu-id="45fae-p142">Идентификатор Exchange для вкладываемого элемента. Максимальная длина — 100 символов.</span><span class="sxs-lookup"><span data-stu-id="45fae-p142">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="45fae-862">String</span><span class="sxs-lookup"><span data-stu-id="45fae-862">String</span></span>||<span data-ttu-id="45fae-863">Тема вкладываемого элемента.</span><span class="sxs-lookup"><span data-stu-id="45fae-863">The subject of the item to be attached.</span></span> <span data-ttu-id="45fae-864">Максимальная длина: 255 символов.</span><span class="sxs-lookup"><span data-stu-id="45fae-864">The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="45fae-865">Object</span><span class="sxs-lookup"><span data-stu-id="45fae-865">Object</span></span>|<span data-ttu-id="45fae-866">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="45fae-866">&lt;optional&gt;</span></span>|<span data-ttu-id="45fae-867">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="45fae-867">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="45fae-868">Объект</span><span class="sxs-lookup"><span data-stu-id="45fae-868">Object</span></span>|<span data-ttu-id="45fae-869">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="45fae-869">&lt;optional&gt;</span></span>|<span data-ttu-id="45fae-870">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="45fae-870">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="45fae-871">функция</span><span class="sxs-lookup"><span data-stu-id="45fae-871">function</span></span>|<span data-ttu-id="45fae-872">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="45fae-872">&lt;optional&gt;</span></span>|<span data-ttu-id="45fae-873">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="45fae-873">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="45fae-874">После успешного выполнения идентификатор вложения будет представлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="45fae-874">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="45fae-875">Если добавить вложение не удастся, объект `asyncResult` будет содержать объект `Error` с описанием ошибки.</span><span class="sxs-lookup"><span data-stu-id="45fae-875">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="45fae-876">Ошибки</span><span class="sxs-lookup"><span data-stu-id="45fae-876">Errors</span></span>

|<span data-ttu-id="45fae-877">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="45fae-877">Error code</span></span>|<span data-ttu-id="45fae-878">Описание</span><span class="sxs-lookup"><span data-stu-id="45fae-878">Description</span></span>|
|------------|-------------|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="45fae-879">Сообщение или встреча содержат слишком много вложений.</span><span class="sxs-lookup"><span data-stu-id="45fae-879">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="45fae-880">Требования</span><span class="sxs-lookup"><span data-stu-id="45fae-880">Requirements</span></span>

|<span data-ttu-id="45fae-881">Требование</span><span class="sxs-lookup"><span data-stu-id="45fae-881">Requirement</span></span>|<span data-ttu-id="45fae-882">Значение</span><span class="sxs-lookup"><span data-stu-id="45fae-882">Value</span></span>|
|---|---|
|[<span data-ttu-id="45fae-883">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="45fae-883">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="45fae-884">1.1</span><span class="sxs-lookup"><span data-stu-id="45fae-884">1.1</span></span>|
|[<span data-ttu-id="45fae-885">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="45fae-885">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="45fae-886">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="45fae-886">ReadWriteItem</span></span>|
|[<span data-ttu-id="45fae-887">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="45fae-887">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="45fae-888">Создание</span><span class="sxs-lookup"><span data-stu-id="45fae-888">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="45fae-889">Пример</span><span class="sxs-lookup"><span data-stu-id="45fae-889">Example</span></span>

<span data-ttu-id="45fae-890">В следующем примере существующий элемент Outlook добавляется в виде вложения с именем `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="45fae-890">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

####  <a name="close"></a><span data-ttu-id="45fae-891">close()</span><span class="sxs-lookup"><span data-stu-id="45fae-891">close()</span></span>

<span data-ttu-id="45fae-892">Закрывает текущий создаваемый элемент.</span><span class="sxs-lookup"><span data-stu-id="45fae-892">Closes the current item that is being composed.</span></span>

<span data-ttu-id="45fae-p144">Работа метода `close` зависит от текущего состояния создаваемого элемента. Если элемент содержит несохраненные изменения, клиент предложит пользователю сохранить или отклонить их либо отменить действие закрытия.</span><span class="sxs-lookup"><span data-stu-id="45fae-p144">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="45fae-895">Если в Outlook в Интернете элемент представляет собой встречу, ранее сохраненную с помощью метода `saveAsync`, пользователю предлагается сохранить, отклонить или отменить действие, даже если с момента последнего сохранения элемента не вносились какие-либо изменения.</span><span class="sxs-lookup"><span data-stu-id="45fae-895">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="45fae-896">Если в клиенте Outlook для настольных ПК сообщение представляет собой ответ в тексте, метод `close` не работает.</span><span class="sxs-lookup"><span data-stu-id="45fae-896">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="45fae-897">Требования</span><span class="sxs-lookup"><span data-stu-id="45fae-897">Requirements</span></span>

|<span data-ttu-id="45fae-898">Требование</span><span class="sxs-lookup"><span data-stu-id="45fae-898">Requirement</span></span>|<span data-ttu-id="45fae-899">Значение</span><span class="sxs-lookup"><span data-stu-id="45fae-899">Value</span></span>|
|---|---|
|[<span data-ttu-id="45fae-900">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="45fae-900">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="45fae-901">1.3</span><span class="sxs-lookup"><span data-stu-id="45fae-901">1.3</span></span>|
|[<span data-ttu-id="45fae-902">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="45fae-902">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="45fae-903">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="45fae-903">Restricted</span></span>|
|[<span data-ttu-id="45fae-904">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="45fae-904">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="45fae-905">Создание</span><span class="sxs-lookup"><span data-stu-id="45fae-905">Compose</span></span>|

---
---

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="45fae-906">displayReplyAllForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="45fae-906">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="45fae-907">Отображает форму ответа, включающую отправителя и всех получателей выбранного сообщения или организатора и всех участников выбранной встречи.</span><span class="sxs-lookup"><span data-stu-id="45fae-907">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="45fae-908">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="45fae-908">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="45fae-909">В Outlook Web App форма ответа отображается в виде всплывающей формы в представлении с 3 либо 1 или 2 колонками.</span><span class="sxs-lookup"><span data-stu-id="45fae-909">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="45fae-910">Если любой строковый параметр превышает указанные для него ограничения, `displayReplyAllForm` возвращает исключение.</span><span class="sxs-lookup"><span data-stu-id="45fae-910">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="45fae-p145">Если в параметре `formData.attachments` указаны вложения, Outlook и Outlook Web App пытаются скачать их и вложить в форму ответа. Если какие-либо вложения добавить не удается, в форме отображается сообщение об ошибке. Если сообщения об ошибках не предусмотрены, то они не отображаются.</span><span class="sxs-lookup"><span data-stu-id="45fae-p145">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="45fae-914">Параметры</span><span class="sxs-lookup"><span data-stu-id="45fae-914">Parameters</span></span>

|<span data-ttu-id="45fae-915">Имя</span><span class="sxs-lookup"><span data-stu-id="45fae-915">Name</span></span>|<span data-ttu-id="45fae-916">Тип</span><span class="sxs-lookup"><span data-stu-id="45fae-916">Type</span></span>|<span data-ttu-id="45fae-917">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="45fae-917">Attributes</span></span>|<span data-ttu-id="45fae-918">Описание</span><span class="sxs-lookup"><span data-stu-id="45fae-918">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="45fae-919">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="45fae-919">String &#124; Object</span></span>||<span data-ttu-id="45fae-p146">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="45fae-p146">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="45fae-922">**ИЛИ**</span><span class="sxs-lookup"><span data-stu-id="45fae-922">**OR**</span></span><br/><span data-ttu-id="45fae-p147">Объект, который содержит текст или данные вложения и функцию обратного вызова. Ниже представлено определение этого объекта.</span><span class="sxs-lookup"><span data-stu-id="45fae-p147">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="45fae-925">String</span><span class="sxs-lookup"><span data-stu-id="45fae-925">String</span></span>|<span data-ttu-id="45fae-926">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="45fae-926">&lt;optional&gt;</span></span>|<span data-ttu-id="45fae-p148">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="45fae-p148">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="45fae-929">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="45fae-929">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="45fae-930">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="45fae-930">&lt;optional&gt;</span></span>|<span data-ttu-id="45fae-931">Массив объектов JSON, представляющих собой вложенные файлы или элементы.</span><span class="sxs-lookup"><span data-stu-id="45fae-931">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="45fae-932">String</span><span class="sxs-lookup"><span data-stu-id="45fae-932">String</span></span>||<span data-ttu-id="45fae-p149">Указывает тип вложения. Допустимые значения: `file` для вложенного файла и `item` для вложенного элемента.</span><span class="sxs-lookup"><span data-stu-id="45fae-p149">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="45fae-935">Строка</span><span class="sxs-lookup"><span data-stu-id="45fae-935">String</span></span>||<span data-ttu-id="45fae-936">Строка, содержащая имя вложения, длиной до 255 символов.</span><span class="sxs-lookup"><span data-stu-id="45fae-936">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="45fae-937">Строка</span><span class="sxs-lookup"><span data-stu-id="45fae-937">String</span></span>||<span data-ttu-id="45fae-p150">Используется, только если свойству `type` задано значение `file`. URI расположения файла.</span><span class="sxs-lookup"><span data-stu-id="45fae-p150">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="45fae-940">Логический</span><span class="sxs-lookup"><span data-stu-id="45fae-940">Boolean</span></span>||<span data-ttu-id="45fae-p151">Используется, только если свойству `type` задано значение `file`. Значение `true` указывает на то, что вложение будет встроено в текст сообщения и не должно отображаться в списке вложений.</span><span class="sxs-lookup"><span data-stu-id="45fae-p151">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="45fae-943">String</span><span class="sxs-lookup"><span data-stu-id="45fae-943">String</span></span>||<span data-ttu-id="45fae-p152">Используется, только если свойству `type` задано значение `item`. Идентификатор вложения EWS. Это строка длиной до 100 символов.</span><span class="sxs-lookup"><span data-stu-id="45fae-p152">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="45fae-947">function</span><span class="sxs-lookup"><span data-stu-id="45fae-947">function</span></span>|<span data-ttu-id="45fae-948">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="45fae-948">&lt;optional&gt;</span></span>|<span data-ttu-id="45fae-949">По завершении работы метода функция, переданная параметру `callback`, вызывается с помощью одного параметра `asyncResult`, представляющего собой объект [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="45fae-949">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="45fae-950">Требования</span><span class="sxs-lookup"><span data-stu-id="45fae-950">Requirements</span></span>

|<span data-ttu-id="45fae-951">Требование</span><span class="sxs-lookup"><span data-stu-id="45fae-951">Requirement</span></span>|<span data-ttu-id="45fae-952">Значение</span><span class="sxs-lookup"><span data-stu-id="45fae-952">Value</span></span>|
|---|---|
|[<span data-ttu-id="45fae-953">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="45fae-953">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="45fae-954">1.0</span><span class="sxs-lookup"><span data-stu-id="45fae-954">1.0</span></span>|
|[<span data-ttu-id="45fae-955">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="45fae-955">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="45fae-956">ReadItem</span><span class="sxs-lookup"><span data-stu-id="45fae-956">ReadItem</span></span>|
|[<span data-ttu-id="45fae-957">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="45fae-957">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="45fae-958">Чтение</span><span class="sxs-lookup"><span data-stu-id="45fae-958">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="45fae-959">Примеры</span><span class="sxs-lookup"><span data-stu-id="45fae-959">Examples</span></span>

<span data-ttu-id="45fae-960">Приведенный ниже код передает строку в функцию `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="45fae-960">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="45fae-961">Ответ с пустым текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="45fae-961">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="45fae-962">Ответ только с текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="45fae-962">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="45fae-963">Ответ с текстом сообщения и вложенным файлом.</span><span class="sxs-lookup"><span data-stu-id="45fae-963">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="45fae-964">Ответ с текстом сообщения и вложенным элементом.</span><span class="sxs-lookup"><span data-stu-id="45fae-964">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="45fae-965">Ответ с текстом сообщения, вложенным файлом, вложенным элементом и обратным вызовом.</span><span class="sxs-lookup"><span data-stu-id="45fae-965">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="45fae-966">displayReplyForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="45fae-966">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="45fae-967">Отображает форму ответа, включающую только отправителя выбранного сообщения или организатора выбранной встречи.</span><span class="sxs-lookup"><span data-stu-id="45fae-967">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="45fae-968">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="45fae-968">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="45fae-969">В Outlook Web App форма ответа отображается в виде всплывающей формы в представлении с 3 либо 1 или 2 колонками.</span><span class="sxs-lookup"><span data-stu-id="45fae-969">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="45fae-970">Если любой строковый параметр превышает указанные для него ограничения, `displayReplyForm` возвращает исключение.</span><span class="sxs-lookup"><span data-stu-id="45fae-970">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="45fae-p153">Если в параметре `formData.attachments` указаны вложения, Outlook и Outlook Web App пытаются скачать их и вложить в форму ответа. Если какие-либо вложения добавить не удается, в форме отображается сообщение об ошибке. Если сообщения об ошибках не предусмотрены, то они не отображаются.</span><span class="sxs-lookup"><span data-stu-id="45fae-p153">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="45fae-974">Параметры</span><span class="sxs-lookup"><span data-stu-id="45fae-974">Parameters</span></span>

|<span data-ttu-id="45fae-975">Имя</span><span class="sxs-lookup"><span data-stu-id="45fae-975">Name</span></span>|<span data-ttu-id="45fae-976">Тип</span><span class="sxs-lookup"><span data-stu-id="45fae-976">Type</span></span>|<span data-ttu-id="45fae-977">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="45fae-977">Attributes</span></span>|<span data-ttu-id="45fae-978">Описание</span><span class="sxs-lookup"><span data-stu-id="45fae-978">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="45fae-979">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="45fae-979">String &#124; Object</span></span>||<span data-ttu-id="45fae-p154">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="45fae-p154">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="45fae-982">**ИЛИ**</span><span class="sxs-lookup"><span data-stu-id="45fae-982">**OR**</span></span><br/><span data-ttu-id="45fae-p155">Объект, который содержит текст или данные вложения и функцию обратного вызова. Ниже представлено определение этого объекта.</span><span class="sxs-lookup"><span data-stu-id="45fae-p155">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="45fae-985">String</span><span class="sxs-lookup"><span data-stu-id="45fae-985">String</span></span>|<span data-ttu-id="45fae-986">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="45fae-986">&lt;optional&gt;</span></span>|<span data-ttu-id="45fae-p156">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="45fae-p156">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="45fae-989">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="45fae-989">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="45fae-990">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="45fae-990">&lt;optional&gt;</span></span>|<span data-ttu-id="45fae-991">Массив объектов JSON, представляющих собой вложенные файлы или элементы.</span><span class="sxs-lookup"><span data-stu-id="45fae-991">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="45fae-992">String</span><span class="sxs-lookup"><span data-stu-id="45fae-992">String</span></span>||<span data-ttu-id="45fae-p157">Указывает тип вложения. Допустимые значения: `file` для вложенного файла и `item` для вложенного элемента.</span><span class="sxs-lookup"><span data-stu-id="45fae-p157">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="45fae-995">Строка</span><span class="sxs-lookup"><span data-stu-id="45fae-995">String</span></span>||<span data-ttu-id="45fae-996">Строка, содержащая имя вложения, длиной до 255 символов.</span><span class="sxs-lookup"><span data-stu-id="45fae-996">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="45fae-997">Строка</span><span class="sxs-lookup"><span data-stu-id="45fae-997">String</span></span>||<span data-ttu-id="45fae-p158">Используется, только если свойству `type` задано значение `file`. URI расположения файла.</span><span class="sxs-lookup"><span data-stu-id="45fae-p158">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="45fae-1000">Логический</span><span class="sxs-lookup"><span data-stu-id="45fae-1000">Boolean</span></span>||<span data-ttu-id="45fae-p159">Используется, только если свойству `type` задано значение `file`. Значение `true` указывает на то, что вложение будет встроено в текст сообщения и не должно отображаться в списке вложений.</span><span class="sxs-lookup"><span data-stu-id="45fae-p159">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="45fae-1003">String</span><span class="sxs-lookup"><span data-stu-id="45fae-1003">String</span></span>||<span data-ttu-id="45fae-p160">Используется, только если свойству `type` задано значение `item`. Идентификатор вложения EWS. Это строка длиной до 100 символов.</span><span class="sxs-lookup"><span data-stu-id="45fae-p160">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="45fae-1007">function</span><span class="sxs-lookup"><span data-stu-id="45fae-1007">function</span></span>|<span data-ttu-id="45fae-1008">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="45fae-1008">&lt;optional&gt;</span></span>|<span data-ttu-id="45fae-1009">По завершении работы метода функция, переданная параметру `callback`, вызывается с помощью одного параметра `asyncResult`, представляющего собой объект [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="45fae-1009">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="45fae-1010">Требования</span><span class="sxs-lookup"><span data-stu-id="45fae-1010">Requirements</span></span>

|<span data-ttu-id="45fae-1011">Требование</span><span class="sxs-lookup"><span data-stu-id="45fae-1011">Requirement</span></span>|<span data-ttu-id="45fae-1012">Значение</span><span class="sxs-lookup"><span data-stu-id="45fae-1012">Value</span></span>|
|---|---|
|[<span data-ttu-id="45fae-1013">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="45fae-1013">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="45fae-1014">1.0</span><span class="sxs-lookup"><span data-stu-id="45fae-1014">1.0</span></span>|
|[<span data-ttu-id="45fae-1015">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="45fae-1015">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="45fae-1016">ReadItem</span><span class="sxs-lookup"><span data-stu-id="45fae-1016">ReadItem</span></span>|
|[<span data-ttu-id="45fae-1017">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="45fae-1017">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="45fae-1018">Чтение</span><span class="sxs-lookup"><span data-stu-id="45fae-1018">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="45fae-1019">Примеры</span><span class="sxs-lookup"><span data-stu-id="45fae-1019">Examples</span></span>

<span data-ttu-id="45fae-1020">Приведенный ниже код передает строку в функцию `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="45fae-1020">The following code passes a string to the `displayReplyForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="45fae-1021">Ответ с пустым текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="45fae-1021">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="45fae-1022">Ответ только с текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="45fae-1022">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="45fae-1023">Ответ с текстом сообщения и вложенным файлом.</span><span class="sxs-lookup"><span data-stu-id="45fae-1023">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="45fae-1024">Ответ с текстом сообщения и вложенным элементом.</span><span class="sxs-lookup"><span data-stu-id="45fae-1024">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="45fae-1025">Ответ с текстом сообщения, вложенным файлом, вложенным элементом и обратным вызовом.</span><span class="sxs-lookup"><span data-stu-id="45fae-1025">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getattachmentcontentasyncattachmentid-options-callback--attachmentcontentjavascriptapioutlookofficeattachmentcontent"></a><span data-ttu-id="45fae-1026">Жетаттачментконтентасинк (attachmentId, [параметры], [callback]) → [вложениеимеет содержимое](/javascript/api/outlook/office.attachmentcontent)</span><span class="sxs-lookup"><span data-stu-id="45fae-1026">getAttachmentContentAsync(attachmentId, [options], [callback]) → [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span></span>

<span data-ttu-id="45fae-1027">Получает указанное вложение из сообщения или встречи и возвращает его в виде `AttachmentContent` объекта.</span><span class="sxs-lookup"><span data-stu-id="45fae-1027">Gets the specified attachment from a message or appointment and returns it as an `AttachmentContent` object.</span></span>

<span data-ttu-id="45fae-1028">`getAttachmentContentAsync` Метод получает вложение с указанным идентификатором из элемента.</span><span class="sxs-lookup"><span data-stu-id="45fae-1028">The `getAttachmentContentAsync` method gets the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="45fae-1029">Рекомендуется использовать идентификатор для получения вложения в том же сеансе, когда Аттачментидс был получен с помощью вызова `getAttachmentsAsync` или. `item.attachments`</span><span class="sxs-lookup"><span data-stu-id="45fae-1029">As a best practice, you should use the identifier to retrieve an attachment in the same session that the attachmentIds were retrieved with the `getAttachmentsAsync` or `item.attachments` call.</span></span> <span data-ttu-id="45fae-1030">В Outlook Web App и Outlook Web App для устройств идентификатор вложения действителен только в рамках одного сеанса.</span><span class="sxs-lookup"><span data-stu-id="45fae-1030">In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="45fae-1031">Сеанс переходит к моменту, когда пользователь закрывает приложение, или если пользователь начинает создание встроенной формы, затем извлекает форму, чтобы продолжить работу в отдельном окне.</span><span class="sxs-lookup"><span data-stu-id="45fae-1031">A session is over when the user closes the app, or if the user starts composing an inline form then subsequently pops out the form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="45fae-1032">Параметры</span><span class="sxs-lookup"><span data-stu-id="45fae-1032">Parameters</span></span>

|<span data-ttu-id="45fae-1033">Имя</span><span class="sxs-lookup"><span data-stu-id="45fae-1033">Name</span></span>|<span data-ttu-id="45fae-1034">Тип</span><span class="sxs-lookup"><span data-stu-id="45fae-1034">Type</span></span>|<span data-ttu-id="45fae-1035">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="45fae-1035">Attributes</span></span>|<span data-ttu-id="45fae-1036">Описание</span><span class="sxs-lookup"><span data-stu-id="45fae-1036">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="45fae-1037">String</span><span class="sxs-lookup"><span data-stu-id="45fae-1037">String</span></span>||<span data-ttu-id="45fae-1038">Идентификатор вложения, которое требуется получить.</span><span class="sxs-lookup"><span data-stu-id="45fae-1038">The identifier of the attachment you want to get.</span></span>|
|`options`|<span data-ttu-id="45fae-1039">Объект</span><span class="sxs-lookup"><span data-stu-id="45fae-1039">Object</span></span>|<span data-ttu-id="45fae-1040">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="45fae-1040">&lt;optional&gt;</span></span>|<span data-ttu-id="45fae-1041">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="45fae-1041">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="45fae-1042">Object</span><span class="sxs-lookup"><span data-stu-id="45fae-1042">Object</span></span>|<span data-ttu-id="45fae-1043">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="45fae-1043">&lt;optional&gt;</span></span>|<span data-ttu-id="45fae-1044">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="45fae-1044">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="45fae-1045">функция</span><span class="sxs-lookup"><span data-stu-id="45fae-1045">function</span></span>|<span data-ttu-id="45fae-1046">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="45fae-1046">&lt;optional&gt;</span></span>|<span data-ttu-id="45fae-1047">По завершении работы метода функция, переданная параметру `callback`, вызывается с помощью одного параметра `asyncResult`, представляющего собой объект [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="45fae-1047">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="45fae-1048">Требования</span><span class="sxs-lookup"><span data-stu-id="45fae-1048">Requirements</span></span>

|<span data-ttu-id="45fae-1049">Требование</span><span class="sxs-lookup"><span data-stu-id="45fae-1049">Requirement</span></span>|<span data-ttu-id="45fae-1050">Значение</span><span class="sxs-lookup"><span data-stu-id="45fae-1050">Value</span></span>|
|---|---|
|[<span data-ttu-id="45fae-1051">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="45fae-1051">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="45fae-1052">Предварительная версия</span><span class="sxs-lookup"><span data-stu-id="45fae-1052">Preview</span></span>|
|[<span data-ttu-id="45fae-1053">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="45fae-1053">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="45fae-1054">ReadItem</span><span class="sxs-lookup"><span data-stu-id="45fae-1054">ReadItem</span></span>|
|[<span data-ttu-id="45fae-1055">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="45fae-1055">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="45fae-1056">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="45fae-1056">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="45fae-1057">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="45fae-1057">Returns:</span></span>

<span data-ttu-id="45fae-1058">Тип: [вложениеимеет содержимое](/javascript/api/outlook/office.attachmentcontent)</span><span class="sxs-lookup"><span data-stu-id="45fae-1058">Type: [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span></span>

##### <a name="example"></a><span data-ttu-id="45fae-1059">Пример</span><span class="sxs-lookup"><span data-stu-id="45fae-1059">Example</span></span>

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

#### <a name="getattachmentsasyncoptions-callback--arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails"></a><span data-ttu-id="45fae-1060">Жетаттачментсасинк ([параметры], [обратный вызов]) → Array. _Лт_[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="45fae-1060">getAttachmentsAsync([options], [callback]) → Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

<span data-ttu-id="45fae-1061">Получает вложения элемента в виде массива.</span><span class="sxs-lookup"><span data-stu-id="45fae-1061">Gets the item's attachments as an array.</span></span> <span data-ttu-id="45fae-1062">Только в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="45fae-1062">Compose mode only.</span></span>

##### <a name="parameters"></a><span data-ttu-id="45fae-1063">Параметры</span><span class="sxs-lookup"><span data-stu-id="45fae-1063">Parameters</span></span>

|<span data-ttu-id="45fae-1064">Имя</span><span class="sxs-lookup"><span data-stu-id="45fae-1064">Name</span></span>|<span data-ttu-id="45fae-1065">Тип</span><span class="sxs-lookup"><span data-stu-id="45fae-1065">Type</span></span>|<span data-ttu-id="45fae-1066">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="45fae-1066">Attributes</span></span>|<span data-ttu-id="45fae-1067">Описание</span><span class="sxs-lookup"><span data-stu-id="45fae-1067">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="45fae-1068">Object</span><span class="sxs-lookup"><span data-stu-id="45fae-1068">Object</span></span>|<span data-ttu-id="45fae-1069">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="45fae-1069">&lt;optional&gt;</span></span>|<span data-ttu-id="45fae-1070">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="45fae-1070">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="45fae-1071">Объект</span><span class="sxs-lookup"><span data-stu-id="45fae-1071">Object</span></span>|<span data-ttu-id="45fae-1072">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="45fae-1072">&lt;optional&gt;</span></span>|<span data-ttu-id="45fae-1073">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="45fae-1073">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="45fae-1074">функция</span><span class="sxs-lookup"><span data-stu-id="45fae-1074">function</span></span>|<span data-ttu-id="45fae-1075">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="45fae-1075">&lt;optional&gt;</span></span>|<span data-ttu-id="45fae-1076">По завершении работы метода функция, переданная параметру `callback`, вызывается с помощью одного параметра `asyncResult`, представляющего собой объект [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="45fae-1076">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="45fae-1077">Требования</span><span class="sxs-lookup"><span data-stu-id="45fae-1077">Requirements</span></span>

|<span data-ttu-id="45fae-1078">Требование</span><span class="sxs-lookup"><span data-stu-id="45fae-1078">Requirement</span></span>|<span data-ttu-id="45fae-1079">Значение</span><span class="sxs-lookup"><span data-stu-id="45fae-1079">Value</span></span>|
|---|---|
|[<span data-ttu-id="45fae-1080">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="45fae-1080">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="45fae-1081">Предварительная версия</span><span class="sxs-lookup"><span data-stu-id="45fae-1081">Preview</span></span>|
|[<span data-ttu-id="45fae-1082">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="45fae-1082">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="45fae-1083">ReadItem</span><span class="sxs-lookup"><span data-stu-id="45fae-1083">ReadItem</span></span>|
|[<span data-ttu-id="45fae-1084">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="45fae-1084">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="45fae-1085">Создание</span><span class="sxs-lookup"><span data-stu-id="45fae-1085">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="45fae-1086">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="45fae-1086">Returns:</span></span>

<span data-ttu-id="45fae-1087">Тип: Array. _Лт_[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="45fae-1087">Type: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

##### <a name="example"></a><span data-ttu-id="45fae-1088">Пример</span><span class="sxs-lookup"><span data-stu-id="45fae-1088">Example</span></span>

<span data-ttu-id="45fae-1089">В приведенном ниже примере создается строка HTML со сведениями обо всех вложениях в текущем элементе.</span><span class="sxs-lookup"><span data-stu-id="45fae-1089">The following example builds an HTML string with details of all attachments on the current item.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlookofficeentities"></a><span data-ttu-id="45fae-1090">getEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="45fae-1090">getEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span></span>

<span data-ttu-id="45fae-1091">Получает сущности, обнаруженные в теле выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="45fae-1091">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="45fae-1092">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="45fae-1092">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="45fae-1093">Требования</span><span class="sxs-lookup"><span data-stu-id="45fae-1093">Requirements</span></span>

|<span data-ttu-id="45fae-1094">Требование</span><span class="sxs-lookup"><span data-stu-id="45fae-1094">Requirement</span></span>|<span data-ttu-id="45fae-1095">Значение</span><span class="sxs-lookup"><span data-stu-id="45fae-1095">Value</span></span>|
|---|---|
|[<span data-ttu-id="45fae-1096">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="45fae-1096">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="45fae-1097">1.0</span><span class="sxs-lookup"><span data-stu-id="45fae-1097">1.0</span></span>|
|[<span data-ttu-id="45fae-1098">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="45fae-1098">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="45fae-1099">ReadItem</span><span class="sxs-lookup"><span data-stu-id="45fae-1099">ReadItem</span></span>|
|[<span data-ttu-id="45fae-1100">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="45fae-1100">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="45fae-1101">Чтение</span><span class="sxs-lookup"><span data-stu-id="45fae-1101">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="45fae-1102">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="45fae-1102">Returns:</span></span>

<span data-ttu-id="45fae-1103">Тип: [Entities](/javascript/api/outlook/office.entities)</span><span class="sxs-lookup"><span data-stu-id="45fae-1103">Type: [Entities](/javascript/api/outlook/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="45fae-1104">Пример</span><span class="sxs-lookup"><span data-stu-id="45fae-1104">Example</span></span>

<span data-ttu-id="45fae-1105">Ниже приведен пример получения доступа к сущностям контактов в теле текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="45fae-1105">The following example accesses the contacts entities in the current item's body.</span></span>

```javascript
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

---
---

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion"></a><span data-ttu-id="45fae-1106">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="45fae-1106">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span></span>

<span data-ttu-id="45fae-1107">Получает массив всех сущностей указанного типа, обнаруженных в теле выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="45fae-1107">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="45fae-1108">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="45fae-1108">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="45fae-1109">Параметры</span><span class="sxs-lookup"><span data-stu-id="45fae-1109">Parameters</span></span>

|<span data-ttu-id="45fae-1110">Имя</span><span class="sxs-lookup"><span data-stu-id="45fae-1110">Name</span></span>|<span data-ttu-id="45fae-1111">Тип</span><span class="sxs-lookup"><span data-stu-id="45fae-1111">Type</span></span>|<span data-ttu-id="45fae-1112">Описание</span><span class="sxs-lookup"><span data-stu-id="45fae-1112">Description</span></span>|
|---|---|---|
|`entityType`|[<span data-ttu-id="45fae-1113">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="45fae-1113">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook/office.mailboxenums.entitytype)|<span data-ttu-id="45fae-1114">Одно из значений перечисления EntityType.</span><span class="sxs-lookup"><span data-stu-id="45fae-1114">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="45fae-1115">Требования</span><span class="sxs-lookup"><span data-stu-id="45fae-1115">Requirements</span></span>

|<span data-ttu-id="45fae-1116">Требование</span><span class="sxs-lookup"><span data-stu-id="45fae-1116">Requirement</span></span>|<span data-ttu-id="45fae-1117">Значение</span><span class="sxs-lookup"><span data-stu-id="45fae-1117">Value</span></span>|
|---|---|
|[<span data-ttu-id="45fae-1118">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="45fae-1118">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="45fae-1119">1.0</span><span class="sxs-lookup"><span data-stu-id="45fae-1119">1.0</span></span>|
|[<span data-ttu-id="45fae-1120">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="45fae-1120">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="45fae-1121">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="45fae-1121">Restricted</span></span>|
|[<span data-ttu-id="45fae-1122">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="45fae-1122">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="45fae-1123">Чтение</span><span class="sxs-lookup"><span data-stu-id="45fae-1123">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="45fae-1124">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="45fae-1124">Returns:</span></span>

<span data-ttu-id="45fae-1125">Если значение, переданное в `entityType`, не является допустимым членом перечисления `EntityType`, метод возвращает значение NULL.</span><span class="sxs-lookup"><span data-stu-id="45fae-1125">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="45fae-1126">Если в теле элемента отсутствуют сущности указанного типа, метод возвращает пустой массив.</span><span class="sxs-lookup"><span data-stu-id="45fae-1126">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="45fae-1127">В противном случае тип объектов в возвращаемом массиве зависит от типа сущности, запрошенной в параметре `entityType`.</span><span class="sxs-lookup"><span data-stu-id="45fae-1127">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="45fae-1128">Хотя минимальный уровень разрешений для использования этого метода — **Restricted**, для некоторых типов сущностей требуется доступ на уровне **ReadItem**, как указано в приведенной ниже таблице.</span><span class="sxs-lookup"><span data-stu-id="45fae-1128">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

|<span data-ttu-id="45fae-1129">Значение параметра `entityType`</span><span class="sxs-lookup"><span data-stu-id="45fae-1129">Value of `entityType`</span></span>|<span data-ttu-id="45fae-1130">Тип объектов в возвращаемом массиве</span><span class="sxs-lookup"><span data-stu-id="45fae-1130">Type of objects in returned array</span></span>|<span data-ttu-id="45fae-1131">Необходимый уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="45fae-1131">Required Permission Level</span></span>|
|---|---|---|
|`Address`|<span data-ttu-id="45fae-1132">String</span><span class="sxs-lookup"><span data-stu-id="45fae-1132">String</span></span>|<span data-ttu-id="45fae-1133">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="45fae-1133">**Restricted**</span></span>|
|`Contact`|<span data-ttu-id="45fae-1134">Contact</span><span class="sxs-lookup"><span data-stu-id="45fae-1134">Contact</span></span>|<span data-ttu-id="45fae-1135">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="45fae-1135">**ReadItem**</span></span>|
|`EmailAddress`|<span data-ttu-id="45fae-1136">String</span><span class="sxs-lookup"><span data-stu-id="45fae-1136">String</span></span>|<span data-ttu-id="45fae-1137">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="45fae-1137">**ReadItem**</span></span>|
|`MeetingSuggestion`|<span data-ttu-id="45fae-1138">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="45fae-1138">MeetingSuggestion</span></span>|<span data-ttu-id="45fae-1139">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="45fae-1139">**ReadItem**</span></span>|
|`PhoneNumber`|<span data-ttu-id="45fae-1140">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="45fae-1140">PhoneNumber</span></span>|<span data-ttu-id="45fae-1141">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="45fae-1141">**Restricted**</span></span>|
|`TaskSuggestion`|<span data-ttu-id="45fae-1142">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="45fae-1142">TaskSuggestion</span></span>|<span data-ttu-id="45fae-1143">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="45fae-1143">**ReadItem**</span></span>|
|`URL`|<span data-ttu-id="45fae-1144">String</span><span class="sxs-lookup"><span data-stu-id="45fae-1144">String</span></span>|<span data-ttu-id="45fae-1145">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="45fae-1145">**Restricted**</span></span>|

<span data-ttu-id="45fae-1146">Тип: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="45fae-1146">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="45fae-1147">Пример</span><span class="sxs-lookup"><span data-stu-id="45fae-1147">Example</span></span>

<span data-ttu-id="45fae-1148">В примере ниже показано, как получить доступ к массиву строк, которые представляют собой почтовые адреса в теле текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="45fae-1148">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion"></a><span data-ttu-id="45fae-1149">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="45fae-1149">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span></span>

<span data-ttu-id="45fae-1150">Возвращает известные сущности в выбранном элементе, которые проходят через именованный фильтр, определяемый в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="45fae-1150">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="45fae-1151">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="45fae-1151">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="45fae-1152">Метод `getFilteredEntitiesByName` возвращает сущности, соответствующие регулярному выражению, которое определяется в элементе правила [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) в XML-файле манифеста, с использованием указанного значения элемента `FilterName`.</span><span class="sxs-lookup"><span data-stu-id="45fae-1152">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="45fae-1153">Параметры</span><span class="sxs-lookup"><span data-stu-id="45fae-1153">Parameters</span></span>

|<span data-ttu-id="45fae-1154">Имя</span><span class="sxs-lookup"><span data-stu-id="45fae-1154">Name</span></span>|<span data-ttu-id="45fae-1155">Тип</span><span class="sxs-lookup"><span data-stu-id="45fae-1155">Type</span></span>|<span data-ttu-id="45fae-1156">Описание</span><span class="sxs-lookup"><span data-stu-id="45fae-1156">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="45fae-1157">String</span><span class="sxs-lookup"><span data-stu-id="45fae-1157">String</span></span>|<span data-ttu-id="45fae-1158">Имя элемента правила `ItemHasKnownEntity`, определяющее соответствующий фильтр.</span><span class="sxs-lookup"><span data-stu-id="45fae-1158">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="45fae-1159">Требования</span><span class="sxs-lookup"><span data-stu-id="45fae-1159">Requirements</span></span>

|<span data-ttu-id="45fae-1160">Требование</span><span class="sxs-lookup"><span data-stu-id="45fae-1160">Requirement</span></span>|<span data-ttu-id="45fae-1161">Значение</span><span class="sxs-lookup"><span data-stu-id="45fae-1161">Value</span></span>|
|---|---|
|[<span data-ttu-id="45fae-1162">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="45fae-1162">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="45fae-1163">1.0</span><span class="sxs-lookup"><span data-stu-id="45fae-1163">1.0</span></span>|
|[<span data-ttu-id="45fae-1164">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="45fae-1164">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="45fae-1165">ReadItem</span><span class="sxs-lookup"><span data-stu-id="45fae-1165">ReadItem</span></span>|
|[<span data-ttu-id="45fae-1166">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="45fae-1166">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="45fae-1167">Чтение</span><span class="sxs-lookup"><span data-stu-id="45fae-1167">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="45fae-1168">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="45fae-1168">Returns:</span></span>

<span data-ttu-id="45fae-p164">Если в манифесте нет элемента `ItemHasKnownEntity` со значением `FilterName`, соответствующим параметру `name`, метод возвращает `null`. Если параметр `name` соответствует элементу `ItemHasKnownEntity` в манифесте, но при этом в текущем элементе нет соответствующих сущностей, метод возвращает пустой массив.</span><span class="sxs-lookup"><span data-stu-id="45fae-p164">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="45fae-1171">Тип: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="45fae-1171">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span></span>

---
---

#### <a name="getinitializationcontextasyncoptions-callback"></a><span data-ttu-id="45fae-1172">getInitializationContextAsync ([параметры], [обратный вызов])</span><span class="sxs-lookup"><span data-stu-id="45fae-1172">getInitializationContextAsync([options], [callback])</span></span>

<span data-ttu-id="45fae-1173">Получает данные инициализации, передаваемые при активации надстройки [сообщением с действиями](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span><span class="sxs-lookup"><span data-stu-id="45fae-1173">Gets initialization data passed when the add-in is [activated by an actionable message](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span></span>

> [!NOTE]
> <span data-ttu-id="45fae-1174">Этот метод поддерживается только Outlook 2016 или более поздней версии для Windows (версии "нажми и работай" более поздней версии, чем 16.0.8413.1000) и Outlook в Интернете для Office 365.</span><span class="sxs-lookup"><span data-stu-id="45fae-1174">This method is only supported by Outlook 2016 or later for Windows (Click-to-Run versions later than 16.0.8413.1000) and Outlook on the web for Office 365.</span></span>

##### <a name="parameters"></a><span data-ttu-id="45fae-1175">Параметры</span><span class="sxs-lookup"><span data-stu-id="45fae-1175">Parameters</span></span>

|<span data-ttu-id="45fae-1176">Имя</span><span class="sxs-lookup"><span data-stu-id="45fae-1176">Name</span></span>|<span data-ttu-id="45fae-1177">Тип</span><span class="sxs-lookup"><span data-stu-id="45fae-1177">Type</span></span>|<span data-ttu-id="45fae-1178">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="45fae-1178">Attributes</span></span>|<span data-ttu-id="45fae-1179">Описание</span><span class="sxs-lookup"><span data-stu-id="45fae-1179">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="45fae-1180">Объект</span><span class="sxs-lookup"><span data-stu-id="45fae-1180">Object</span></span>|<span data-ttu-id="45fae-1181">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="45fae-1181">&lt;optional&gt;</span></span>|<span data-ttu-id="45fae-1182">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="45fae-1182">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="45fae-1183">Объект</span><span class="sxs-lookup"><span data-stu-id="45fae-1183">Object</span></span>|<span data-ttu-id="45fae-1184">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="45fae-1184">&lt;optional&gt;</span></span>|<span data-ttu-id="45fae-1185">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="45fae-1185">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="45fae-1186">функция</span><span class="sxs-lookup"><span data-stu-id="45fae-1186">function</span></span>|<span data-ttu-id="45fae-1187">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="45fae-1187">&lt;optional&gt;</span></span>|<span data-ttu-id="45fae-1188">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="45fae-1188">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="45fae-1189">При успешном выполнении данные инициализации предоставляются в `asyncResult.value` свойстве в виде строки.</span><span class="sxs-lookup"><span data-stu-id="45fae-1189">On success, the initialization data is provided in the `asyncResult.value` property as a string.</span></span><br/><span data-ttu-id="45fae-1190">Если `asyncResult` контекст инициализации отсутствует, объект будет содержать `Error` объект со `code` свойством, `9020` `name` для свойства которого задано значение. `GenericResponseError`</span><span class="sxs-lookup"><span data-stu-id="45fae-1190">If there is no initialization context, the `asyncResult` object will contain an `Error` object with its `code` property set to `9020` and its `name` property set to `GenericResponseError`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="45fae-1191">Требования</span><span class="sxs-lookup"><span data-stu-id="45fae-1191">Requirements</span></span>

|<span data-ttu-id="45fae-1192">Требование</span><span class="sxs-lookup"><span data-stu-id="45fae-1192">Requirement</span></span>|<span data-ttu-id="45fae-1193">Значение</span><span class="sxs-lookup"><span data-stu-id="45fae-1193">Value</span></span>|
|---|---|
|[<span data-ttu-id="45fae-1194">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="45fae-1194">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="45fae-1195">Предварительная версия</span><span class="sxs-lookup"><span data-stu-id="45fae-1195">Preview</span></span>|
|[<span data-ttu-id="45fae-1196">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="45fae-1196">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="45fae-1197">ReadItem</span><span class="sxs-lookup"><span data-stu-id="45fae-1197">ReadItem</span></span>|
|[<span data-ttu-id="45fae-1198">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="45fae-1198">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="45fae-1199">Чтение</span><span class="sxs-lookup"><span data-stu-id="45fae-1199">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="45fae-1200">Пример</span><span class="sxs-lookup"><span data-stu-id="45fae-1200">Example</span></span>

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

#### <a name="getregexmatches--object"></a><span data-ttu-id="45fae-1201">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="45fae-1201">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="45fae-1202">Возвращает строковые значения в выбранном элементе, которые соответствуют регулярным выражениям, определенным в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="45fae-1202">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="45fae-1203">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="45fae-1203">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="45fae-p165">Метод `getRegExMatches` возвращает строки, соответствующие регулярному выражению, которое определяется в каждом элементе правила `ItemHasRegularExpressionMatch` или `ItemHasKnownEntity` в XML-файле манифеста. Для правила `ItemHasRegularExpressionMatch` соответствующую строку должно содержать свойство элемента, указанного этим правилом. Простой тип `PropertyName` определяет поддерживаемые свойства.</span><span class="sxs-lookup"><span data-stu-id="45fae-p165">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="45fae-1207">Например, рассмотрим манифест надстройки, который содержит указанный ниже элемент `Rule`.</span><span class="sxs-lookup"><span data-stu-id="45fae-1207">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="45fae-1208">Объект, возвращаемый методом `getRegExMatches`, будет содержать два свойства: `fruits` и `veggies`.</span><span class="sxs-lookup"><span data-stu-id="45fae-1208">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="45fae-p166">Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты. Лучше используйте метод [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) для этого.</span><span class="sxs-lookup"><span data-stu-id="45fae-p166">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="45fae-1212">Requirements</span><span class="sxs-lookup"><span data-stu-id="45fae-1212">Requirements</span></span>

|<span data-ttu-id="45fae-1213">Требование</span><span class="sxs-lookup"><span data-stu-id="45fae-1213">Requirement</span></span>|<span data-ttu-id="45fae-1214">Значение</span><span class="sxs-lookup"><span data-stu-id="45fae-1214">Value</span></span>|
|---|---|
|[<span data-ttu-id="45fae-1215">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="45fae-1215">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="45fae-1216">1.0</span><span class="sxs-lookup"><span data-stu-id="45fae-1216">1.0</span></span>|
|[<span data-ttu-id="45fae-1217">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="45fae-1217">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="45fae-1218">ReadItem</span><span class="sxs-lookup"><span data-stu-id="45fae-1218">ReadItem</span></span>|
|[<span data-ttu-id="45fae-1219">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="45fae-1219">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="45fae-1220">Чтение</span><span class="sxs-lookup"><span data-stu-id="45fae-1220">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="45fae-1221">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="45fae-1221">Returns:</span></span>

<span data-ttu-id="45fae-p167">Объект, содержащий массив строк, которые соответствуют регулярным выражениям, определяемым в XML-файле манифеста. Имя каждого массива равно соответствующему значению атрибута `RegExName` подходящего правила `ItemHasRegularExpressionMatch` или атрибута `FilterName` соответствующего правила `ItemHasKnownEntity`.</span><span class="sxs-lookup"><span data-stu-id="45fae-p167">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="45fae-1224">Тип:</span><span class="sxs-lookup"><span data-stu-id="45fae-1224">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="45fae-1225">Object</span><span class="sxs-lookup"><span data-stu-id="45fae-1225">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="45fae-1226">Пример</span><span class="sxs-lookup"><span data-stu-id="45fae-1226">Example</span></span>

<span data-ttu-id="45fae-1227">В приведенном ниже примере показано, как получить доступ к массиву совпадений с элементами `fruits` и `veggies` правил активации регулярных выражений, указанными в манифесте.</span><span class="sxs-lookup"><span data-stu-id="45fae-1227">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```javascript
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

---
---

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="45fae-1228">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="45fae-1228">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="45fae-1229">Возвращает строковые значения в выбранном элементе, которые соответствуют именованному регулярному выражению, определенному в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="45fae-1229">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="45fae-1230">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="45fae-1230">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="45fae-1231">Метод `getRegExMatchesByName` возвращает строки, соответствующие регулярному выражению, которое определяется в элементе правила `ItemHasRegularExpressionMatch` в XML-файле манифеста, с использованием указанного значения элемента `RegExName`.</span><span class="sxs-lookup"><span data-stu-id="45fae-1231">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="45fae-p168">Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты.</span><span class="sxs-lookup"><span data-stu-id="45fae-p168">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="45fae-1234">Параметры</span><span class="sxs-lookup"><span data-stu-id="45fae-1234">Parameters</span></span>

|<span data-ttu-id="45fae-1235">Имя</span><span class="sxs-lookup"><span data-stu-id="45fae-1235">Name</span></span>|<span data-ttu-id="45fae-1236">Тип</span><span class="sxs-lookup"><span data-stu-id="45fae-1236">Type</span></span>|<span data-ttu-id="45fae-1237">Описание</span><span class="sxs-lookup"><span data-stu-id="45fae-1237">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="45fae-1238">String</span><span class="sxs-lookup"><span data-stu-id="45fae-1238">String</span></span>|<span data-ttu-id="45fae-1239">Имя элемента правила `ItemHasRegularExpressionMatch`, определяющее соответствующий фильтр.</span><span class="sxs-lookup"><span data-stu-id="45fae-1239">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="45fae-1240">Требования</span><span class="sxs-lookup"><span data-stu-id="45fae-1240">Requirements</span></span>

|<span data-ttu-id="45fae-1241">Требование</span><span class="sxs-lookup"><span data-stu-id="45fae-1241">Requirement</span></span>|<span data-ttu-id="45fae-1242">Значение</span><span class="sxs-lookup"><span data-stu-id="45fae-1242">Value</span></span>|
|---|---|
|[<span data-ttu-id="45fae-1243">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="45fae-1243">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="45fae-1244">1.0</span><span class="sxs-lookup"><span data-stu-id="45fae-1244">1.0</span></span>|
|[<span data-ttu-id="45fae-1245">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="45fae-1245">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="45fae-1246">ReadItem</span><span class="sxs-lookup"><span data-stu-id="45fae-1246">ReadItem</span></span>|
|[<span data-ttu-id="45fae-1247">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="45fae-1247">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="45fae-1248">Чтение</span><span class="sxs-lookup"><span data-stu-id="45fae-1248">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="45fae-1249">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="45fae-1249">Returns:</span></span>

<span data-ttu-id="45fae-1250">Массив строк, соответствующих регулярному выражению, определяемому в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="45fae-1250">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="45fae-1251">

<dt>Тип</dt>

</span><span class="sxs-lookup"><span data-stu-id="45fae-1251">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="45fae-1252">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="45fae-1252">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="45fae-1253">Пример</span><span class="sxs-lookup"><span data-stu-id="45fae-1253">Example</span></span>

```javascript
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

---
---

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="45fae-1254">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="45fae-1254">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="45fae-1255">Асинхронно возвращает данные, выбранные в теме или тексте сообщения.</span><span class="sxs-lookup"><span data-stu-id="45fae-1255">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="45fae-p169">Если выделенный фрагмент отсутствует, но курсор находится в тексте или теме, метод возвращает значение NULL для выбранных данных. Если выбраны не текст и не тема, метод возвращает ошибку `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="45fae-p169">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="45fae-1258">Параметры</span><span class="sxs-lookup"><span data-stu-id="45fae-1258">Parameters</span></span>

|<span data-ttu-id="45fae-1259">Имя</span><span class="sxs-lookup"><span data-stu-id="45fae-1259">Name</span></span>|<span data-ttu-id="45fae-1260">Тип</span><span class="sxs-lookup"><span data-stu-id="45fae-1260">Type</span></span>|<span data-ttu-id="45fae-1261">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="45fae-1261">Attributes</span></span>|<span data-ttu-id="45fae-1262">Описание</span><span class="sxs-lookup"><span data-stu-id="45fae-1262">Description</span></span>|
|---|---|---|---|
|`coercionType`|[<span data-ttu-id="45fae-1263">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="45fae-1263">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="45fae-p170">Запрашивает формат данных. Если задано значение Text, метод возвращает обычный текст как строку, удаляя все имеющиеся HTML-теги. Если задано значение HTML, метод возвращает выделенный текст (обычный текст или HTML).</span><span class="sxs-lookup"><span data-stu-id="45fae-p170">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`|<span data-ttu-id="45fae-1267">Объект</span><span class="sxs-lookup"><span data-stu-id="45fae-1267">Object</span></span>|<span data-ttu-id="45fae-1268">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="45fae-1268">&lt;optional&gt;</span></span>|<span data-ttu-id="45fae-1269">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="45fae-1269">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="45fae-1270">Объект</span><span class="sxs-lookup"><span data-stu-id="45fae-1270">Object</span></span>|<span data-ttu-id="45fae-1271">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="45fae-1271">&lt;optional&gt;</span></span>|<span data-ttu-id="45fae-1272">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="45fae-1272">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="45fae-1273">функция</span><span class="sxs-lookup"><span data-stu-id="45fae-1273">function</span></span>||<span data-ttu-id="45fae-1274">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="45fae-1274">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="45fae-1275">Чтобы получить доступ к выбранным данным из метода обратного вызова, вызовите `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="45fae-1275">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="45fae-1276">Для доступа к исходному свойству, представляющему собой источник выбранных данных, вызовите параметр `asyncResult.value.sourceProperty`, который может иметь значение `body` или `subject`.</span><span class="sxs-lookup"><span data-stu-id="45fae-1276">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="45fae-1277">Требования</span><span class="sxs-lookup"><span data-stu-id="45fae-1277">Requirements</span></span>

|<span data-ttu-id="45fae-1278">Требование</span><span class="sxs-lookup"><span data-stu-id="45fae-1278">Requirement</span></span>|<span data-ttu-id="45fae-1279">Значение</span><span class="sxs-lookup"><span data-stu-id="45fae-1279">Value</span></span>|
|---|---|
|[<span data-ttu-id="45fae-1280">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="45fae-1280">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="45fae-1281">1.2</span><span class="sxs-lookup"><span data-stu-id="45fae-1281">1.2</span></span>|
|[<span data-ttu-id="45fae-1282">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="45fae-1282">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="45fae-1283">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="45fae-1283">ReadWriteItem</span></span>|
|[<span data-ttu-id="45fae-1284">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="45fae-1284">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="45fae-1285">Создание</span><span class="sxs-lookup"><span data-stu-id="45fae-1285">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="45fae-1286">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="45fae-1286">Returns:</span></span>

<span data-ttu-id="45fae-1287">Выбранные данные в виде строки с форматом, определенным в параметре `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="45fae-1287">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="45fae-1288">

<dt>Тип</dt>

</span><span class="sxs-lookup"><span data-stu-id="45fae-1288">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="45fae-1289">String</span><span class="sxs-lookup"><span data-stu-id="45fae-1289">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="45fae-1290">Пример</span><span class="sxs-lookup"><span data-stu-id="45fae-1290">Example</span></span>

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

#### <a name="getselectedentities--entitiesjavascriptapioutlookofficeentities"></a><span data-ttu-id="45fae-1291">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="45fae-1291">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span></span>

<span data-ttu-id="45fae-1292">Возвращает сущности, найденные в выделенном совпадении, выбранном пользователем.</span><span class="sxs-lookup"><span data-stu-id="45fae-1292">Gets the entities found in a highlighted match a user has selected.</span></span> <span data-ttu-id="45fae-1293">Выделенные совпадения применяются к [контекстным надстройкам](/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="45fae-1293">Highlighted matches apply to [contextual add-ins](/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="45fae-1294">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="45fae-1294">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="45fae-1295">Требования</span><span class="sxs-lookup"><span data-stu-id="45fae-1295">Requirements</span></span>

|<span data-ttu-id="45fae-1296">Требование</span><span class="sxs-lookup"><span data-stu-id="45fae-1296">Requirement</span></span>|<span data-ttu-id="45fae-1297">Значение</span><span class="sxs-lookup"><span data-stu-id="45fae-1297">Value</span></span>|
|---|---|
|[<span data-ttu-id="45fae-1298">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="45fae-1298">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="45fae-1299">1.6</span><span class="sxs-lookup"><span data-stu-id="45fae-1299">1.6</span></span>|
|[<span data-ttu-id="45fae-1300">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="45fae-1300">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="45fae-1301">ReadItem</span><span class="sxs-lookup"><span data-stu-id="45fae-1301">ReadItem</span></span>|
|[<span data-ttu-id="45fae-1302">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="45fae-1302">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="45fae-1303">Чтение</span><span class="sxs-lookup"><span data-stu-id="45fae-1303">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="45fae-1304">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="45fae-1304">Returns:</span></span>

<span data-ttu-id="45fae-1305">Тип: [Entities](/javascript/api/outlook/office.entities)</span><span class="sxs-lookup"><span data-stu-id="45fae-1305">Type: [Entities](/javascript/api/outlook/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="45fae-1306">Пример</span><span class="sxs-lookup"><span data-stu-id="45fae-1306">Example</span></span>

<span data-ttu-id="45fae-1307">В приведенном ниже примере показано, как получить доступ к сущностям адресов в выделенном совпадении, выбранном пользователем.</span><span class="sxs-lookup"><span data-stu-id="45fae-1307">The following example accesses the addresses entities in the highlighted match selected by the user.</span></span>

```javascript
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

---
---

#### <a name="getselectedregexmatches--object"></a><span data-ttu-id="45fae-1308">getSelectedRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="45fae-1308">getSelectedRegExMatches() → {Object}</span></span>

<span data-ttu-id="45fae-p173">Возвращает строковые значения в выделенном совпадении, которые соответствуют регулярным выражениям, определенным в XML-файле манифеста. Выделенные совпадения применяются к [контекстным надстройкам](/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="45fae-p173">Returns string values in a highlighted match that match the regular expressions defined in the manifest XML file. Highlighted matches apply to [contextual add-ins](/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="45fae-1311">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="45fae-1311">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="45fae-p174">Метод `getSelectedRegExMatches` возвращает строки, соответствующие регулярному выражению, которое определяется в каждом элементе правила `ItemHasRegularExpressionMatch` или `ItemHasKnownEntity` в XML-файле манифеста. Для правила `ItemHasRegularExpressionMatch` соответствующую строку должно содержать свойство элемента, указанного этим правилом. Простой тип `PropertyName` определяет поддерживаемые свойства.</span><span class="sxs-lookup"><span data-stu-id="45fae-p174">The `getSelectedRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="45fae-1315">Например, рассмотрим манифест надстройки, который содержит указанный ниже элемент `Rule`.</span><span class="sxs-lookup"><span data-stu-id="45fae-1315">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="45fae-1316">Объект, возвращаемый методом `getRegExMatches`, будет содержать два свойства: `fruits` и `veggies`.</span><span class="sxs-lookup"><span data-stu-id="45fae-1316">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="45fae-p175">Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты. Лучше используйте метод [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) для этого.</span><span class="sxs-lookup"><span data-stu-id="45fae-p175">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="45fae-1320">Requirements</span><span class="sxs-lookup"><span data-stu-id="45fae-1320">Requirements</span></span>

|<span data-ttu-id="45fae-1321">Требование</span><span class="sxs-lookup"><span data-stu-id="45fae-1321">Requirement</span></span>|<span data-ttu-id="45fae-1322">Значение</span><span class="sxs-lookup"><span data-stu-id="45fae-1322">Value</span></span>|
|---|---|
|[<span data-ttu-id="45fae-1323">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="45fae-1323">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="45fae-1324">1.6</span><span class="sxs-lookup"><span data-stu-id="45fae-1324">1.6</span></span>|
|[<span data-ttu-id="45fae-1325">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="45fae-1325">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="45fae-1326">ReadItem</span><span class="sxs-lookup"><span data-stu-id="45fae-1326">ReadItem</span></span>|
|[<span data-ttu-id="45fae-1327">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="45fae-1327">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="45fae-1328">Чтение</span><span class="sxs-lookup"><span data-stu-id="45fae-1328">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="45fae-1329">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="45fae-1329">Returns:</span></span>

<span data-ttu-id="45fae-p176">Объект, содержащий массив строк, которые соответствуют регулярным выражениям, определяемым в XML-файле манифеста. Имя каждого массива равно соответствующему значению атрибута `RegExName` подходящего правила `ItemHasRegularExpressionMatch` или атрибута `FilterName` соответствующего правила `ItemHasKnownEntity`.</span><span class="sxs-lookup"><span data-stu-id="45fae-p176">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

##### <a name="example"></a><span data-ttu-id="45fae-1332">Пример</span><span class="sxs-lookup"><span data-stu-id="45fae-1332">Example</span></span>

<span data-ttu-id="45fae-1333">В приведенном ниже примере показано, как получить доступ к массиву совпадений с элементами `fruits` и `veggies` правил активации регулярных выражений, указанными в манифесте.</span><span class="sxs-lookup"><span data-stu-id="45fae-1333">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```javascript
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

---
---

#### <a name="getsharedpropertiesasyncoptions-callback"></a><span data-ttu-id="45fae-1334">Жетшаредпропертиесасинк ([параметры], обратный вызов)</span><span class="sxs-lookup"><span data-stu-id="45fae-1334">getSharedPropertiesAsync([options], callback)</span></span>

<span data-ttu-id="45fae-1335">Получает свойства выбранной встречи или сообщения в общей папке, календаре или почтовом ящике.</span><span class="sxs-lookup"><span data-stu-id="45fae-1335">Gets the properties of the selected appointment or message in a shared folder, calendar, or mailbox.</span></span>

##### <a name="parameters"></a><span data-ttu-id="45fae-1336">Параметры</span><span class="sxs-lookup"><span data-stu-id="45fae-1336">Parameters</span></span>

|<span data-ttu-id="45fae-1337">Имя</span><span class="sxs-lookup"><span data-stu-id="45fae-1337">Name</span></span>|<span data-ttu-id="45fae-1338">Тип</span><span class="sxs-lookup"><span data-stu-id="45fae-1338">Type</span></span>|<span data-ttu-id="45fae-1339">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="45fae-1339">Attributes</span></span>|<span data-ttu-id="45fae-1340">Описание</span><span class="sxs-lookup"><span data-stu-id="45fae-1340">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="45fae-1341">Object</span><span class="sxs-lookup"><span data-stu-id="45fae-1341">Object</span></span>|<span data-ttu-id="45fae-1342">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="45fae-1342">&lt;optional&gt;</span></span>|<span data-ttu-id="45fae-1343">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="45fae-1343">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="45fae-1344">Объект</span><span class="sxs-lookup"><span data-stu-id="45fae-1344">Object</span></span>|<span data-ttu-id="45fae-1345">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="45fae-1345">&lt;optional&gt;</span></span>|<span data-ttu-id="45fae-1346">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="45fae-1346">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="45fae-1347">функция</span><span class="sxs-lookup"><span data-stu-id="45fae-1347">function</span></span>||<span data-ttu-id="45fae-1348">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="45fae-1348">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="45fae-1349">Общие свойства предоставляются в виде [`SharedProperties`](/javascript/api/outlook/office.sharedproperties) объекта в `asyncResult.value` свойстве.</span><span class="sxs-lookup"><span data-stu-id="45fae-1349">The shared properties are provided as a [`SharedProperties`](/javascript/api/outlook/office.sharedproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="45fae-1350">Этот объект можно использовать для получения общих свойств элемента.</span><span class="sxs-lookup"><span data-stu-id="45fae-1350">This object can be used to get the item's shared properties.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="45fae-1351">Требования</span><span class="sxs-lookup"><span data-stu-id="45fae-1351">Requirements</span></span>

|<span data-ttu-id="45fae-1352">Требование</span><span class="sxs-lookup"><span data-stu-id="45fae-1352">Requirement</span></span>|<span data-ttu-id="45fae-1353">Значение</span><span class="sxs-lookup"><span data-stu-id="45fae-1353">Value</span></span>|
|---|---|
|[<span data-ttu-id="45fae-1354">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="45fae-1354">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="45fae-1355">Предварительная версия</span><span class="sxs-lookup"><span data-stu-id="45fae-1355">Preview</span></span>|
|[<span data-ttu-id="45fae-1356">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="45fae-1356">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="45fae-1357">ReadItem</span><span class="sxs-lookup"><span data-stu-id="45fae-1357">ReadItem</span></span>|
|[<span data-ttu-id="45fae-1358">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="45fae-1358">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="45fae-1359">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="45fae-1359">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="45fae-1360">Пример</span><span class="sxs-lookup"><span data-stu-id="45fae-1360">Example</span></span>

```javascript
Office.context.mailbox.item.getSharedPropertiesAsync(callback);

function callback (asyncResult) {
  var context = asyncResult.context;
  var sharedProperties = asyncResult.value;
}
```

---
---

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="45fae-1361">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="45fae-1361">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="45fae-1362">Асинхронно загружает настраиваемые свойства для надстройки для выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="45fae-1362">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="45fae-p178">Настраиваемые свойства сохраняются в виде пар "ключ-значение" для каждого приложения и каждого элемента. Этот метод возвращает объект `CustomProperties` при обратном вызове, который предоставляет методы для доступа к настраиваемым свойствам, характерным для текущего элемента и текущей надстройки. Настраиваемые свойства не шифруются для элемента, поэтому этот способ хранения не является безопасным.</span><span class="sxs-lookup"><span data-stu-id="45fae-p178">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="45fae-1366">Параметры</span><span class="sxs-lookup"><span data-stu-id="45fae-1366">Parameters</span></span>

|<span data-ttu-id="45fae-1367">Имя</span><span class="sxs-lookup"><span data-stu-id="45fae-1367">Name</span></span>|<span data-ttu-id="45fae-1368">Тип</span><span class="sxs-lookup"><span data-stu-id="45fae-1368">Type</span></span>|<span data-ttu-id="45fae-1369">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="45fae-1369">Attributes</span></span>|<span data-ttu-id="45fae-1370">Описание</span><span class="sxs-lookup"><span data-stu-id="45fae-1370">Description</span></span>|
|---|---|---|---|
|`callback`|<span data-ttu-id="45fae-1371">function</span><span class="sxs-lookup"><span data-stu-id="45fae-1371">function</span></span>||<span data-ttu-id="45fae-1372">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="45fae-1372">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="45fae-1373">Настраиваемые свойства предоставляются в виде объекта [`CustomProperties`](/javascript/api/outlook/office.customproperties) в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="45fae-1373">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="45fae-1374">Этот объект позволяет получить, задать и удалить настраиваемые свойства для элемента, а также сохранить изменения, внесенные в набор настраиваемых свойств, на сервере.</span><span class="sxs-lookup"><span data-stu-id="45fae-1374">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`|<span data-ttu-id="45fae-1375">Объект</span><span class="sxs-lookup"><span data-stu-id="45fae-1375">Object</span></span>|<span data-ttu-id="45fae-1376">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="45fae-1376">&lt;optional&gt;</span></span>|<span data-ttu-id="45fae-1377">Разработчики могут указать любой объект, к которому необходимо получить доступ, в функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="45fae-1377">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="45fae-1378">Доступ к этому объекту можно получить с помощью свойства `asyncResult.asyncContext` в функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="45fae-1378">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="45fae-1379">Требования</span><span class="sxs-lookup"><span data-stu-id="45fae-1379">Requirements</span></span>

|<span data-ttu-id="45fae-1380">Требование</span><span class="sxs-lookup"><span data-stu-id="45fae-1380">Requirement</span></span>|<span data-ttu-id="45fae-1381">Значение</span><span class="sxs-lookup"><span data-stu-id="45fae-1381">Value</span></span>|
|---|---|
|[<span data-ttu-id="45fae-1382">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="45fae-1382">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="45fae-1383">1.0</span><span class="sxs-lookup"><span data-stu-id="45fae-1383">1.0</span></span>|
|[<span data-ttu-id="45fae-1384">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="45fae-1384">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="45fae-1385">ReadItem</span><span class="sxs-lookup"><span data-stu-id="45fae-1385">ReadItem</span></span>|
|[<span data-ttu-id="45fae-1386">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="45fae-1386">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="45fae-1387">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="45fae-1387">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="45fae-1388">Пример</span><span class="sxs-lookup"><span data-stu-id="45fae-1388">Example</span></span>

<span data-ttu-id="45fae-p181">Приведенный ниже пример кода показывает, как асинхронно загружать настраиваемые свойства, характерные для текущего элемента, с помощью метода `loadCustomPropertiesAsync`. Этот пример также показывает, как сохранять эти свойства на сервере с помощью метода `CustomProperties.saveAsync`. После загрузки настраиваемых свойств в этом примере кода метод `CustomProperties.get` используется для считывания настраиваемого свойства `myProp`, метод `CustomProperties.set` — для записи настраиваемого свойства `otherProp`, а метод `saveAsync` — для сохранения настраиваемых свойств.</span><span class="sxs-lookup"><span data-stu-id="45fae-p181">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="45fae-1392">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="45fae-1392">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="45fae-1393">Удаляет вложение из сообщения или встречи.</span><span class="sxs-lookup"><span data-stu-id="45fae-1393">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="45fae-1394">Метод `removeAttachmentAsync` удаляет из элемента вложение с указанным идентификатором.</span><span class="sxs-lookup"><span data-stu-id="45fae-1394">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="45fae-1395">Идентификатор вложения рекомендуется использовать для удаления вложения, только если оно добавлено тем же почтовым приложением в ходе текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="45fae-1395">As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session.</span></span> <span data-ttu-id="45fae-1396">В Outlook Web App и Outlook Web App для устройств идентификатор вложения действителен только в рамках одного сеанса.</span><span class="sxs-lookup"><span data-stu-id="45fae-1396">In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="45fae-1397">Сеанс переходит к моменту, когда пользователь закрывает приложение, или если пользователь начинает создание встроенной формы, затем извлекает форму, чтобы продолжить работу в отдельном окне.</span><span class="sxs-lookup"><span data-stu-id="45fae-1397">A session is over when the user closes the app, or if the user starts composing an inline form then subsequently pops out the form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="45fae-1398">Параметры</span><span class="sxs-lookup"><span data-stu-id="45fae-1398">Parameters</span></span>

|<span data-ttu-id="45fae-1399">Имя</span><span class="sxs-lookup"><span data-stu-id="45fae-1399">Name</span></span>|<span data-ttu-id="45fae-1400">Тип</span><span class="sxs-lookup"><span data-stu-id="45fae-1400">Type</span></span>|<span data-ttu-id="45fae-1401">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="45fae-1401">Attributes</span></span>|<span data-ttu-id="45fae-1402">Описание</span><span class="sxs-lookup"><span data-stu-id="45fae-1402">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="45fae-1403">String</span><span class="sxs-lookup"><span data-stu-id="45fae-1403">String</span></span>||<span data-ttu-id="45fae-1404">Идентификатор удаляемого вложения.</span><span class="sxs-lookup"><span data-stu-id="45fae-1404">The identifier of the attachment to remove.</span></span>|
|`options`|<span data-ttu-id="45fae-1405">Объект</span><span class="sxs-lookup"><span data-stu-id="45fae-1405">Object</span></span>|<span data-ttu-id="45fae-1406">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="45fae-1406">&lt;optional&gt;</span></span>|<span data-ttu-id="45fae-1407">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="45fae-1407">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="45fae-1408">Объект</span><span class="sxs-lookup"><span data-stu-id="45fae-1408">Object</span></span>|<span data-ttu-id="45fae-1409">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="45fae-1409">&lt;optional&gt;</span></span>|<span data-ttu-id="45fae-1410">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="45fae-1410">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="45fae-1411">функция</span><span class="sxs-lookup"><span data-stu-id="45fae-1411">function</span></span>|<span data-ttu-id="45fae-1412">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="45fae-1412">&lt;optional&gt;</span></span>|<span data-ttu-id="45fae-1413">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="45fae-1413">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="45fae-1414">Если удалить вложение не удается, свойство `asyncResult.error` содержит код ошибки с указанием ее причины.</span><span class="sxs-lookup"><span data-stu-id="45fae-1414">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="45fae-1415">Ошибки</span><span class="sxs-lookup"><span data-stu-id="45fae-1415">Errors</span></span>

|<span data-ttu-id="45fae-1416">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="45fae-1416">Error code</span></span>|<span data-ttu-id="45fae-1417">Описание</span><span class="sxs-lookup"><span data-stu-id="45fae-1417">Description</span></span>|
|------------|-------------|
|`InvalidAttachmentId`|<span data-ttu-id="45fae-1418">Идентификатор вложения не существует.</span><span class="sxs-lookup"><span data-stu-id="45fae-1418">The attachment identifier does not exist.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="45fae-1419">Требования</span><span class="sxs-lookup"><span data-stu-id="45fae-1419">Requirements</span></span>

|<span data-ttu-id="45fae-1420">Требование</span><span class="sxs-lookup"><span data-stu-id="45fae-1420">Requirement</span></span>|<span data-ttu-id="45fae-1421">Значение</span><span class="sxs-lookup"><span data-stu-id="45fae-1421">Value</span></span>|
|---|---|
|[<span data-ttu-id="45fae-1422">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="45fae-1422">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="45fae-1423">1.1</span><span class="sxs-lookup"><span data-stu-id="45fae-1423">1.1</span></span>|
|[<span data-ttu-id="45fae-1424">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="45fae-1424">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="45fae-1425">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="45fae-1425">ReadWriteItem</span></span>|
|[<span data-ttu-id="45fae-1426">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="45fae-1426">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="45fae-1427">Создание</span><span class="sxs-lookup"><span data-stu-id="45fae-1427">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="45fae-1428">Пример</span><span class="sxs-lookup"><span data-stu-id="45fae-1428">Example</span></span>

<span data-ttu-id="45fae-1429">Указанный ниже код удаляет вложение с идентификатором "0".</span><span class="sxs-lookup"><span data-stu-id="45fae-1429">The following code removes an attachment with an identifier of '0'.</span></span>

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

####  <a name="removehandlerasynceventtype-options-callback"></a><span data-ttu-id="45fae-1430">removeHandlerAsync(eventType, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="45fae-1430">removeHandlerAsync(eventType, [options], [callback])</span></span>

<span data-ttu-id="45fae-1431">Удаляет обработчиков для поддерживаемого типа события.</span><span class="sxs-lookup"><span data-stu-id="45fae-1431">Removes the event handlers for a supported event type.</span></span>

<span data-ttu-id="45fae-1432">В настоящее время поддерживаются типы `Office.EventType.AttachmentsChanged`событий `Office.EventType.AppointmentTimeChanged`, `Office.EventType.EnhancedLocationsChanged` `Office.EventType.RecipientsChanged`,, и `Office.EventType.RecurrenceChanged`.</span><span class="sxs-lookup"><span data-stu-id="45fae-1432">Currently the supported event types are `Office.EventType.AttachmentsChanged`, `Office.EventType.AppointmentTimeChanged`, `Office.EventType.EnhancedLocationsChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="45fae-1433">Параметры</span><span class="sxs-lookup"><span data-stu-id="45fae-1433">Parameters</span></span>

| <span data-ttu-id="45fae-1434">Имя</span><span class="sxs-lookup"><span data-stu-id="45fae-1434">Name</span></span> | <span data-ttu-id="45fae-1435">Тип</span><span class="sxs-lookup"><span data-stu-id="45fae-1435">Type</span></span> | <span data-ttu-id="45fae-1436">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="45fae-1436">Attributes</span></span> | <span data-ttu-id="45fae-1437">Описание</span><span class="sxs-lookup"><span data-stu-id="45fae-1437">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="45fae-1438">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="45fae-1438">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="45fae-1439">Событие, которое должно отменить обработчик.</span><span class="sxs-lookup"><span data-stu-id="45fae-1439">The event that should revoke the handler.</span></span> |
| `options` | <span data-ttu-id="45fae-1440">Объект</span><span class="sxs-lookup"><span data-stu-id="45fae-1440">Object</span></span> | <span data-ttu-id="45fae-1441">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="45fae-1441">&lt;optional&gt;</span></span> | <span data-ttu-id="45fae-1442">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="45fae-1442">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="45fae-1443">Объект</span><span class="sxs-lookup"><span data-stu-id="45fae-1443">Object</span></span> | <span data-ttu-id="45fae-1444">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="45fae-1444">&lt;optional&gt;</span></span> | <span data-ttu-id="45fae-1445">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="45fae-1445">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="45fae-1446">функция</span><span class="sxs-lookup"><span data-stu-id="45fae-1446">function</span></span>| <span data-ttu-id="45fae-1447">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="45fae-1447">&lt;optional&gt;</span></span>|<span data-ttu-id="45fae-1448">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="45fae-1448">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="45fae-1449">Требования</span><span class="sxs-lookup"><span data-stu-id="45fae-1449">Requirements</span></span>

|<span data-ttu-id="45fae-1450">Требование</span><span class="sxs-lookup"><span data-stu-id="45fae-1450">Requirement</span></span>| <span data-ttu-id="45fae-1451">Значение</span><span class="sxs-lookup"><span data-stu-id="45fae-1451">Value</span></span>|
|---|---|
|[<span data-ttu-id="45fae-1452">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="45fae-1452">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="45fae-1453">1.7</span><span class="sxs-lookup"><span data-stu-id="45fae-1453">1.7</span></span> |
|[<span data-ttu-id="45fae-1454">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="45fae-1454">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="45fae-1455">ReadItem</span><span class="sxs-lookup"><span data-stu-id="45fae-1455">ReadItem</span></span> |
|[<span data-ttu-id="45fae-1456">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="45fae-1456">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="45fae-1457">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="45fae-1457">Compose or Read</span></span> |

---
---

####  <a name="saveasyncoptions-callback"></a><span data-ttu-id="45fae-1458">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="45fae-1458">saveAsync([options], callback)</span></span>

<span data-ttu-id="45fae-1459">Асинхронно сохраняет элемент.</span><span class="sxs-lookup"><span data-stu-id="45fae-1459">Asynchronously saves an item.</span></span>

<span data-ttu-id="45fae-p183">При вызове этот метод сохраняет текущее сообщение в виде черновика и возвращает идентификатор элемента с помощью метода обратного вызова. В Outlook Web App или интерактивном режиме Outlook этот элемент сохраняется на сервере. В Outlook в режиме кэширования этот элемент сохраняется в локальном кэше.</span><span class="sxs-lookup"><span data-stu-id="45fae-p183">When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="45fae-1463">Если в режиме создания надстройка вызывает для элемента метод `saveAsync`, чтобы получить параметр `itemId` для использования с EWS или REST API, необходимо помнить, что синхронизация элемента с сервером может занять много времени, если Outlook работает в режиме кэширования данных.</span><span class="sxs-lookup"><span data-stu-id="45fae-1463">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="45fae-1464">До окончания синхронизации элемента применение параметра `itemId` будет приводить к ошибке.</span><span class="sxs-lookup"><span data-stu-id="45fae-1464">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="45fae-p185">Если метод `saveAsync` вызывается для встречи в режиме создания, она сохраняется как обычная встреча в календаре пользователя, а не как черновик. При сохранении новой встречи приглашения не отправляются. При сохранении существующей встречи уведомления отправляются добавленным или удаленным участникам.</span><span class="sxs-lookup"><span data-stu-id="45fae-p185">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="45fae-1468">Следующие клиенты отличаются другим поведением `saveAsync` в отношении встреч в режиме создания:</span><span class="sxs-lookup"><span data-stu-id="45fae-1468">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="45fae-1469">Outlook для Mac не поддерживает `saveAsync` для собраний в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="45fae-1469">Mac Outlook does not support `saveAsync` on a meeting in compose mode.</span></span> <span data-ttu-id="45fae-1470">При вызове `saveAsync` для собрания в Outlook для Mac возвращается ошибка.</span><span class="sxs-lookup"><span data-stu-id="45fae-1470">Calling `saveAsync` on a meeting in Mac Outlook will return an error.</span></span>
> - <span data-ttu-id="45fae-1471">Outlook в Интернете всегда отправляет приглашение или обновление при вызове `saveAsync` для встречи в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="45fae-1471">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="45fae-1472">Параметры</span><span class="sxs-lookup"><span data-stu-id="45fae-1472">Parameters</span></span>

|<span data-ttu-id="45fae-1473">Имя</span><span class="sxs-lookup"><span data-stu-id="45fae-1473">Name</span></span>|<span data-ttu-id="45fae-1474">Тип</span><span class="sxs-lookup"><span data-stu-id="45fae-1474">Type</span></span>|<span data-ttu-id="45fae-1475">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="45fae-1475">Attributes</span></span>|<span data-ttu-id="45fae-1476">Описание</span><span class="sxs-lookup"><span data-stu-id="45fae-1476">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="45fae-1477">Object</span><span class="sxs-lookup"><span data-stu-id="45fae-1477">Object</span></span>|<span data-ttu-id="45fae-1478">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="45fae-1478">&lt;optional&gt;</span></span>|<span data-ttu-id="45fae-1479">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="45fae-1479">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="45fae-1480">Объект</span><span class="sxs-lookup"><span data-stu-id="45fae-1480">Object</span></span>|<span data-ttu-id="45fae-1481">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="45fae-1481">&lt;optional&gt;</span></span>|<span data-ttu-id="45fae-1482">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="45fae-1482">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="45fae-1483">функция</span><span class="sxs-lookup"><span data-stu-id="45fae-1483">function</span></span>||<span data-ttu-id="45fae-1484">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="45fae-1484">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="45fae-1485">После успешного выполнения идентификатор элемента будет предоставлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="45fae-1485">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="45fae-1486">Требования</span><span class="sxs-lookup"><span data-stu-id="45fae-1486">Requirements</span></span>

|<span data-ttu-id="45fae-1487">Требование</span><span class="sxs-lookup"><span data-stu-id="45fae-1487">Requirement</span></span>|<span data-ttu-id="45fae-1488">Значение</span><span class="sxs-lookup"><span data-stu-id="45fae-1488">Value</span></span>|
|---|---|
|[<span data-ttu-id="45fae-1489">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="45fae-1489">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="45fae-1490">1.3</span><span class="sxs-lookup"><span data-stu-id="45fae-1490">1.3</span></span>|
|[<span data-ttu-id="45fae-1491">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="45fae-1491">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="45fae-1492">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="45fae-1492">ReadWriteItem</span></span>|
|[<span data-ttu-id="45fae-1493">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="45fae-1493">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="45fae-1494">Создание</span><span class="sxs-lookup"><span data-stu-id="45fae-1494">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="45fae-1495">Примеры</span><span class="sxs-lookup"><span data-stu-id="45fae-1495">Examples</span></span>

```javascript
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result.
  });
```

<span data-ttu-id="45fae-p187">Ниже приведен пример параметра `result`, переданного функции обратного вызова. Свойство `value` содержит идентификатор элемента.</span><span class="sxs-lookup"><span data-stu-id="45fae-p187">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

---
---

####  <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="45fae-1498">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="45fae-1498">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="45fae-1499">Асинхронно вставляет данные в текст или тему сообщения.</span><span class="sxs-lookup"><span data-stu-id="45fae-1499">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="45fae-p188">Метод `setSelectedDataAsync` вставляет указанную строку в местоположение курсора в теме или тексте элемента либо, если текст выделен в редакторе, он заменяет выделенный текст. Если курсор находится вне текста или темы элемента, возвращается ошибка. После вставки курсор помещается в конец вставленного содержимого.</span><span class="sxs-lookup"><span data-stu-id="45fae-p188">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="45fae-1503">Параметры</span><span class="sxs-lookup"><span data-stu-id="45fae-1503">Parameters</span></span>

|<span data-ttu-id="45fae-1504">Имя</span><span class="sxs-lookup"><span data-stu-id="45fae-1504">Name</span></span>|<span data-ttu-id="45fae-1505">Тип</span><span class="sxs-lookup"><span data-stu-id="45fae-1505">Type</span></span>|<span data-ttu-id="45fae-1506">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="45fae-1506">Attributes</span></span>|<span data-ttu-id="45fae-1507">Описание</span><span class="sxs-lookup"><span data-stu-id="45fae-1507">Description</span></span>|
|---|---|---|---|
|`data`|<span data-ttu-id="45fae-1508">String</span><span class="sxs-lookup"><span data-stu-id="45fae-1508">String</span></span>||<span data-ttu-id="45fae-p189">Вставляемые данные. Объем данных не должен превышать 1 000 000 символов. Если передано больше 1 000 000 символов, возвращается исключение `ArgumentOutOfRange`.</span><span class="sxs-lookup"><span data-stu-id="45fae-p189">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`|<span data-ttu-id="45fae-1512">Object</span><span class="sxs-lookup"><span data-stu-id="45fae-1512">Object</span></span>|<span data-ttu-id="45fae-1513">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="45fae-1513">&lt;optional&gt;</span></span>|<span data-ttu-id="45fae-1514">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="45fae-1514">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="45fae-1515">Объект</span><span class="sxs-lookup"><span data-stu-id="45fae-1515">Object</span></span>|<span data-ttu-id="45fae-1516">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="45fae-1516">&lt;optional&gt;</span></span>|<span data-ttu-id="45fae-1517">В методе обратного вызова разработчики могут указать любой объект, к которому необходимо получить доступ.</span><span class="sxs-lookup"><span data-stu-id="45fae-1517">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="45fae-1518">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="45fae-1518">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="45fae-1519">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="45fae-1519">&lt;optional&gt;</span></span>|<span data-ttu-id="45fae-p190">Если задано значение `text`, текущий стиль применяется в Outlook Web App и Outlook. Если поле представляет собой редактор HTML, вставляются только текстовые данные, даже если они имеют формат HTML.</span><span class="sxs-lookup"><span data-stu-id="45fae-p190">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="45fae-p191">Если задано значение `html` и поле (не тема) поддерживает HTML, в Outlook Web App применяется текущий стиль, а в Outlook — стиль по умолчанию. Если поле является текстовым, возвращается ошибка `InvalidDataFormat`.</span><span class="sxs-lookup"><span data-stu-id="45fae-p191">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="45fae-1524">Если свойство `coercionType` не задано, результат зависит от поля: если поле имеет формат HTML, используется текст в формате HTML, а если поле текстовое, применяется обычный текст.</span><span class="sxs-lookup"><span data-stu-id="45fae-1524">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`|<span data-ttu-id="45fae-1525">функция</span><span class="sxs-lookup"><span data-stu-id="45fae-1525">function</span></span>||<span data-ttu-id="45fae-1526">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="45fae-1526">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="45fae-1527">Требования</span><span class="sxs-lookup"><span data-stu-id="45fae-1527">Requirements</span></span>

|<span data-ttu-id="45fae-1528">Требование</span><span class="sxs-lookup"><span data-stu-id="45fae-1528">Requirement</span></span>|<span data-ttu-id="45fae-1529">Значение</span><span class="sxs-lookup"><span data-stu-id="45fae-1529">Value</span></span>|
|---|---|
|[<span data-ttu-id="45fae-1530">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="45fae-1530">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="45fae-1531">1.2</span><span class="sxs-lookup"><span data-stu-id="45fae-1531">1.2</span></span>|
|[<span data-ttu-id="45fae-1532">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="45fae-1532">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="45fae-1533">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="45fae-1533">ReadWriteItem</span></span>|
|[<span data-ttu-id="45fae-1534">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="45fae-1534">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="45fae-1535">Создание</span><span class="sxs-lookup"><span data-stu-id="45fae-1535">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="45fae-1536">Пример</span><span class="sxs-lookup"><span data-stu-id="45fae-1536">Example</span></span>

```javascript
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
