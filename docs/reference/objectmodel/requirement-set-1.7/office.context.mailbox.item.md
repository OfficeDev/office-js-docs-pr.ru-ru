---
title: Office.Context.Mailbox.Item - требование задать 1.7
description: ''
ms.date: 01/16/2019
localization_priority: Normal
ms.openlocfilehash: dfc86d8a118ab5f5c32968c567a2eec6b9e7d267
ms.sourcegitcommit: d1aa7201820176ed986b9f00bb9c88e055906c77
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 01/23/2019
ms.locfileid: "29389587"
---
# <a name="item"></a><span data-ttu-id="1d898-102">item</span><span class="sxs-lookup"><span data-stu-id="1d898-102">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="1d898-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="1d898-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="1d898-p101">Пространство имен `item` используется для доступа к выбранному в данный момент сообщению, приглашению на собрание или описанию встречи. Вы можете определить тип пространства имен `item` с помощью свойства [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook17officemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="1d898-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook17officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="1d898-106">Requirements</span><span class="sxs-lookup"><span data-stu-id="1d898-106">Requirements</span></span>

|<span data-ttu-id="1d898-107">Требование</span><span class="sxs-lookup"><span data-stu-id="1d898-107">Requirement</span></span>|<span data-ttu-id="1d898-108">Значение</span><span class="sxs-lookup"><span data-stu-id="1d898-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="1d898-109">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="1d898-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1d898-110">1.0</span><span class="sxs-lookup"><span data-stu-id="1d898-110">1.0</span></span>|
|[<span data-ttu-id="1d898-111">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="1d898-111">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1d898-112">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="1d898-112">Restricted</span></span>|
|[<span data-ttu-id="1d898-113">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="1d898-113">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="1d898-114">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="1d898-114">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="1d898-115">Элементы и методы</span><span class="sxs-lookup"><span data-stu-id="1d898-115">Members and methods</span></span>

| <span data-ttu-id="1d898-116">Элемент</span><span class="sxs-lookup"><span data-stu-id="1d898-116">Member</span></span> | <span data-ttu-id="1d898-117">Тип</span><span class="sxs-lookup"><span data-stu-id="1d898-117">Type</span></span> |
|--------|------|
| [<span data-ttu-id="1d898-118">attachments</span><span class="sxs-lookup"><span data-stu-id="1d898-118">attachments</span></span>](#attachments-arrayattachmentdetailsjavascriptapioutlook17officeattachmentdetails) | <span data-ttu-id="1d898-119">Элемент</span><span class="sxs-lookup"><span data-stu-id="1d898-119">Member</span></span> |
| [<span data-ttu-id="1d898-120">bcc</span><span class="sxs-lookup"><span data-stu-id="1d898-120">bcc</span></span>](#bcc-recipientsjavascriptapioutlook17officerecipients) | <span data-ttu-id="1d898-121">Элемент</span><span class="sxs-lookup"><span data-stu-id="1d898-121">Member</span></span> |
| [<span data-ttu-id="1d898-122">body</span><span class="sxs-lookup"><span data-stu-id="1d898-122">body</span></span>](#body-bodyjavascriptapioutlook17officebody) | <span data-ttu-id="1d898-123">Элемент</span><span class="sxs-lookup"><span data-stu-id="1d898-123">Member</span></span> |
| [<span data-ttu-id="1d898-124">cc</span><span class="sxs-lookup"><span data-stu-id="1d898-124">cc</span></span>](#cc-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients) | <span data-ttu-id="1d898-125">Элемент</span><span class="sxs-lookup"><span data-stu-id="1d898-125">Member</span></span> |
| [<span data-ttu-id="1d898-126">conversationId</span><span class="sxs-lookup"><span data-stu-id="1d898-126">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="1d898-127">Элемент</span><span class="sxs-lookup"><span data-stu-id="1d898-127">Member</span></span> |
| [<span data-ttu-id="1d898-128">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="1d898-128">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="1d898-129">Элемент</span><span class="sxs-lookup"><span data-stu-id="1d898-129">Member</span></span> |
| [<span data-ttu-id="1d898-130">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="1d898-130">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="1d898-131">Элемент</span><span class="sxs-lookup"><span data-stu-id="1d898-131">Member</span></span> |
| [<span data-ttu-id="1d898-132">end</span><span class="sxs-lookup"><span data-stu-id="1d898-132">end</span></span>](#end-datetimejavascriptapioutlook17officetime) | <span data-ttu-id="1d898-133">Элемент</span><span class="sxs-lookup"><span data-stu-id="1d898-133">Member</span></span> |
| [<span data-ttu-id="1d898-134">from</span><span class="sxs-lookup"><span data-stu-id="1d898-134">from</span></span>](#from-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsfromjavascriptapioutlook17officefrom) | <span data-ttu-id="1d898-135">Элемент</span><span class="sxs-lookup"><span data-stu-id="1d898-135">Member</span></span> |
| [<span data-ttu-id="1d898-136">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="1d898-136">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="1d898-137">Элемент</span><span class="sxs-lookup"><span data-stu-id="1d898-137">Member</span></span> |
| [<span data-ttu-id="1d898-138">itemClass</span><span class="sxs-lookup"><span data-stu-id="1d898-138">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="1d898-139">Элемент</span><span class="sxs-lookup"><span data-stu-id="1d898-139">Member</span></span> |
| [<span data-ttu-id="1d898-140">itemId</span><span class="sxs-lookup"><span data-stu-id="1d898-140">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="1d898-141">Элемент</span><span class="sxs-lookup"><span data-stu-id="1d898-141">Member</span></span> |
| [<span data-ttu-id="1d898-142">itemType</span><span class="sxs-lookup"><span data-stu-id="1d898-142">itemType</span></span>](#itemtype-officemailboxenumsitemtypejavascriptapioutlook17officemailboxenumsitemtype) | <span data-ttu-id="1d898-143">Элемент</span><span class="sxs-lookup"><span data-stu-id="1d898-143">Member</span></span> |
| [<span data-ttu-id="1d898-144">location</span><span class="sxs-lookup"><span data-stu-id="1d898-144">location</span></span>](#location-stringlocationjavascriptapioutlook17officelocation) | <span data-ttu-id="1d898-145">Элемент</span><span class="sxs-lookup"><span data-stu-id="1d898-145">Member</span></span> |
| [<span data-ttu-id="1d898-146">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="1d898-146">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="1d898-147">Элемент</span><span class="sxs-lookup"><span data-stu-id="1d898-147">Member</span></span> |
| [<span data-ttu-id="1d898-148">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="1d898-148">notificationMessages</span></span>](#notificationmessages-notificationmessagesjavascriptapioutlook17officenotificationmessages) | <span data-ttu-id="1d898-149">Элемент</span><span class="sxs-lookup"><span data-stu-id="1d898-149">Member</span></span> |
| [<span data-ttu-id="1d898-150">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="1d898-150">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients) | <span data-ttu-id="1d898-151">Элемент</span><span class="sxs-lookup"><span data-stu-id="1d898-151">Member</span></span> |
| [<span data-ttu-id="1d898-152">organizer</span><span class="sxs-lookup"><span data-stu-id="1d898-152">organizer</span></span>](#organizer-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsorganizerjavascriptapioutlook17officeorganizer) | <span data-ttu-id="1d898-153">Элемент</span><span class="sxs-lookup"><span data-stu-id="1d898-153">Member</span></span> |
| [<span data-ttu-id="1d898-154">recurrence</span><span class="sxs-lookup"><span data-stu-id="1d898-154">recurrence</span></span>](#nullable-recurrence-recurrencejavascriptapioutlook17officerecurrence) | <span data-ttu-id="1d898-155">Элемент</span><span class="sxs-lookup"><span data-stu-id="1d898-155">Member</span></span> |
| [<span data-ttu-id="1d898-156">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="1d898-156">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients) | <span data-ttu-id="1d898-157">Элемент</span><span class="sxs-lookup"><span data-stu-id="1d898-157">Member</span></span> |
| [<span data-ttu-id="1d898-158">sender</span><span class="sxs-lookup"><span data-stu-id="1d898-158">sender</span></span>](#sender-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetails) | <span data-ttu-id="1d898-159">Элемент</span><span class="sxs-lookup"><span data-stu-id="1d898-159">Member</span></span> |
| [<span data-ttu-id="1d898-160">seriesId</span><span class="sxs-lookup"><span data-stu-id="1d898-160">seriesId</span></span>](#nullable-seriesid-string) | <span data-ttu-id="1d898-161">Элемент</span><span class="sxs-lookup"><span data-stu-id="1d898-161">Member</span></span> |
| [<span data-ttu-id="1d898-162">start</span><span class="sxs-lookup"><span data-stu-id="1d898-162">start</span></span>](#start-datetimejavascriptapioutlook17officetime) | <span data-ttu-id="1d898-163">Элемент</span><span class="sxs-lookup"><span data-stu-id="1d898-163">Member</span></span> |
| [<span data-ttu-id="1d898-164">subject</span><span class="sxs-lookup"><span data-stu-id="1d898-164">subject</span></span>](#subject-stringsubjectjavascriptapioutlook17officesubject) | <span data-ttu-id="1d898-165">Элемент</span><span class="sxs-lookup"><span data-stu-id="1d898-165">Member</span></span> |
| [<span data-ttu-id="1d898-166">to</span><span class="sxs-lookup"><span data-stu-id="1d898-166">to</span></span>](#to-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients) | <span data-ttu-id="1d898-167">Элемент</span><span class="sxs-lookup"><span data-stu-id="1d898-167">Member</span></span> |
| [<span data-ttu-id="1d898-168">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="1d898-168">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="1d898-169">Метод</span><span class="sxs-lookup"><span data-stu-id="1d898-169">Method</span></span> |
| [<span data-ttu-id="1d898-170">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="1d898-170">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="1d898-171">Метод</span><span class="sxs-lookup"><span data-stu-id="1d898-171">Method</span></span> |
| [<span data-ttu-id="1d898-172">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="1d898-172">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="1d898-173">Метод</span><span class="sxs-lookup"><span data-stu-id="1d898-173">Method</span></span> |
| [<span data-ttu-id="1d898-174">close</span><span class="sxs-lookup"><span data-stu-id="1d898-174">close</span></span>](#close) | <span data-ttu-id="1d898-175">Метод</span><span class="sxs-lookup"><span data-stu-id="1d898-175">Method</span></span> |
| [<span data-ttu-id="1d898-176">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="1d898-176">displayReplyAllForm</span></span>](#displayreplyallformformdata) | <span data-ttu-id="1d898-177">Метод</span><span class="sxs-lookup"><span data-stu-id="1d898-177">Method</span></span> |
| [<span data-ttu-id="1d898-178">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="1d898-178">displayReplyForm</span></span>](#displayreplyformformdata) | <span data-ttu-id="1d898-179">Метод</span><span class="sxs-lookup"><span data-stu-id="1d898-179">Method</span></span> |
| [<span data-ttu-id="1d898-180">getEntities</span><span class="sxs-lookup"><span data-stu-id="1d898-180">getEntities</span></span>](#getentities--entitiesjavascriptapioutlook17officeentities) | <span data-ttu-id="1d898-181">Метод</span><span class="sxs-lookup"><span data-stu-id="1d898-181">Method</span></span> |
| [<span data-ttu-id="1d898-182">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="1d898-182">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook17officecontactmeetingsuggestionjavascriptapioutlook17officemeetingsuggestionphonenumberjavascriptapioutlook17officephonenumbertasksuggestionjavascriptapioutlook17officetasksuggestion) | <span data-ttu-id="1d898-183">Метод</span><span class="sxs-lookup"><span data-stu-id="1d898-183">Method</span></span> |
| [<span data-ttu-id="1d898-184">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="1d898-184">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook17officecontactmeetingsuggestionjavascriptapioutlook17officemeetingsuggestionphonenumberjavascriptapioutlook17officephonenumbertasksuggestionjavascriptapioutlook17officetasksuggestion) | <span data-ttu-id="1d898-185">Метод</span><span class="sxs-lookup"><span data-stu-id="1d898-185">Method</span></span> |
| [<span data-ttu-id="1d898-186">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="1d898-186">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="1d898-187">Метод</span><span class="sxs-lookup"><span data-stu-id="1d898-187">Method</span></span> |
| [<span data-ttu-id="1d898-188">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="1d898-188">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="1d898-189">Метод</span><span class="sxs-lookup"><span data-stu-id="1d898-189">Method</span></span> |
| [<span data-ttu-id="1d898-190">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="1d898-190">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="1d898-191">Метод</span><span class="sxs-lookup"><span data-stu-id="1d898-191">Method</span></span> |
| [<span data-ttu-id="1d898-192">getSelectedEntities</span><span class="sxs-lookup"><span data-stu-id="1d898-192">getSelectedEntities</span></span>](#getselectedentities--entitiesjavascriptapioutlook17officeentities) | <span data-ttu-id="1d898-193">Метод</span><span class="sxs-lookup"><span data-stu-id="1d898-193">Method</span></span> |
| [<span data-ttu-id="1d898-194">getSelectedRegExMatches</span><span class="sxs-lookup"><span data-stu-id="1d898-194">getSelectedRegExMatches</span></span>](#getselectedregexmatches--object) | <span data-ttu-id="1d898-195">Метод</span><span class="sxs-lookup"><span data-stu-id="1d898-195">Method</span></span> |
| [<span data-ttu-id="1d898-196">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="1d898-196">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="1d898-197">Метод</span><span class="sxs-lookup"><span data-stu-id="1d898-197">Method</span></span> |
| [<span data-ttu-id="1d898-198">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="1d898-198">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="1d898-199">Метод</span><span class="sxs-lookup"><span data-stu-id="1d898-199">Method</span></span> |
| [<span data-ttu-id="1d898-200">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="1d898-200">removeHandlerAsync</span></span>](#removehandlerasynceventtype-options-callback) | <span data-ttu-id="1d898-201">Метод</span><span class="sxs-lookup"><span data-stu-id="1d898-201">Method</span></span> |
| [<span data-ttu-id="1d898-202">saveAsync</span><span class="sxs-lookup"><span data-stu-id="1d898-202">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="1d898-203">Метод</span><span class="sxs-lookup"><span data-stu-id="1d898-203">Method</span></span> |
| [<span data-ttu-id="1d898-204">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="1d898-204">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="1d898-205">Метод</span><span class="sxs-lookup"><span data-stu-id="1d898-205">Method</span></span> |

### <a name="example"></a><span data-ttu-id="1d898-206">Пример</span><span class="sxs-lookup"><span data-stu-id="1d898-206">Example</span></span>

<span data-ttu-id="1d898-207">В примере кода JavaScript, приведенном ниже, показано, как получить доступ к свойству `subject` текущего элемента в Outlook.</span><span class="sxs-lookup"><span data-stu-id="1d898-207">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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
}
```

### <a name="members"></a><span data-ttu-id="1d898-208">Элементы</span><span class="sxs-lookup"><span data-stu-id="1d898-208">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook17officeattachmentdetails"></a><span data-ttu-id="1d898-209">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_7/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="1d898-209">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_7/office.attachmentdetails)></span></span>

<span data-ttu-id="1d898-p102">Получает массив вложений для элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="1d898-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="1d898-212">Outlook блокирует определенные типы файлов из-за потенциальных проблем с безопасностью, поэтому они не возвращаются.</span><span class="sxs-lookup"><span data-stu-id="1d898-212">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="1d898-213">Дополнительные сведения см. в статье [Блокировка вложений в Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="1d898-213">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="1d898-214">Тип:</span><span class="sxs-lookup"><span data-stu-id="1d898-214">Type:</span></span>

*   <span data-ttu-id="1d898-215">Array.<[AttachmentDetails](/javascript/api/outlook_1_7/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="1d898-215">Array.<[AttachmentDetails](/javascript/api/outlook_1_7/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="1d898-216">Требования</span><span class="sxs-lookup"><span data-stu-id="1d898-216">Requirements</span></span>

|<span data-ttu-id="1d898-217">Требование</span><span class="sxs-lookup"><span data-stu-id="1d898-217">Requirement</span></span>|<span data-ttu-id="1d898-218">Значение</span><span class="sxs-lookup"><span data-stu-id="1d898-218">Value</span></span>|
|---|---|
|[<span data-ttu-id="1d898-219">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="1d898-219">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1d898-220">1.0</span><span class="sxs-lookup"><span data-stu-id="1d898-220">1.0</span></span>|
|[<span data-ttu-id="1d898-221">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="1d898-221">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1d898-222">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1d898-222">ReadItem</span></span>|
|[<span data-ttu-id="1d898-223">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="1d898-223">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="1d898-224">Чтение</span><span class="sxs-lookup"><span data-stu-id="1d898-224">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1d898-225">Пример</span><span class="sxs-lookup"><span data-stu-id="1d898-225">Example</span></span>

<span data-ttu-id="1d898-226">С помощью приведенного ниже кода можно создать HTML-строку с подробными сведениями обо всех вложениях для текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="1d898-226">The following code builds an HTML string with details of all attachments on the current item.</span></span>

```js
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

####  <a name="bcc-recipientsjavascriptapioutlook17officerecipients"></a><span data-ttu-id="1d898-227">bcc :[Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="1d898-227">bcc :[Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

<span data-ttu-id="1d898-228">Получает объект, который предоставляет методы для получения или обновления получателей скрытой копии сообщения.</span><span class="sxs-lookup"><span data-stu-id="1d898-228">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="1d898-229">Только в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="1d898-229">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="1d898-230">Тип:</span><span class="sxs-lookup"><span data-stu-id="1d898-230">Type:</span></span>

*   [<span data-ttu-id="1d898-231">Recipients</span><span class="sxs-lookup"><span data-stu-id="1d898-231">Recipients</span></span>](/javascript/api/outlook_1_7/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="1d898-232">Требования</span><span class="sxs-lookup"><span data-stu-id="1d898-232">Requirements</span></span>

|<span data-ttu-id="1d898-233">Требование</span><span class="sxs-lookup"><span data-stu-id="1d898-233">Requirement</span></span>|<span data-ttu-id="1d898-234">Значение</span><span class="sxs-lookup"><span data-stu-id="1d898-234">Value</span></span>|
|---|---|
|[<span data-ttu-id="1d898-235">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="1d898-235">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1d898-236">1.1</span><span class="sxs-lookup"><span data-stu-id="1d898-236">1.1</span></span>|
|[<span data-ttu-id="1d898-237">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="1d898-237">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1d898-238">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1d898-238">ReadItem</span></span>|
|[<span data-ttu-id="1d898-239">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="1d898-239">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="1d898-240">Создание</span><span class="sxs-lookup"><span data-stu-id="1d898-240">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="1d898-241">Пример</span><span class="sxs-lookup"><span data-stu-id="1d898-241">Example</span></span>

```js
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlook17officebody"></a><span data-ttu-id="1d898-242">body :[Body](/javascript/api/outlook_1_7/office.body)</span><span class="sxs-lookup"><span data-stu-id="1d898-242">body :[Body](/javascript/api/outlook_1_7/office.body)</span></span>

<span data-ttu-id="1d898-243">Получает объект, предоставляющий методы для работы с основным текстом элемента.</span><span class="sxs-lookup"><span data-stu-id="1d898-243">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="1d898-244">Тип:</span><span class="sxs-lookup"><span data-stu-id="1d898-244">Type:</span></span>

*   [<span data-ttu-id="1d898-245">Body</span><span class="sxs-lookup"><span data-stu-id="1d898-245">Body</span></span>](/javascript/api/outlook_1_7/office.body)

##### <a name="requirements"></a><span data-ttu-id="1d898-246">Требования</span><span class="sxs-lookup"><span data-stu-id="1d898-246">Requirements</span></span>

|<span data-ttu-id="1d898-247">Требование</span><span class="sxs-lookup"><span data-stu-id="1d898-247">Requirement</span></span>|<span data-ttu-id="1d898-248">Значение</span><span class="sxs-lookup"><span data-stu-id="1d898-248">Value</span></span>|
|---|---|
|[<span data-ttu-id="1d898-249">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="1d898-249">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1d898-250">1.1</span><span class="sxs-lookup"><span data-stu-id="1d898-250">1.1</span></span>|
|[<span data-ttu-id="1d898-251">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="1d898-251">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1d898-252">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1d898-252">ReadItem</span></span>|
|[<span data-ttu-id="1d898-253">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="1d898-253">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="1d898-254">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="1d898-254">Compose or read</span></span>|

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients"></a><span data-ttu-id="1d898-255">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="1d898-255">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

<span data-ttu-id="1d898-256">Предоставляет доступ к получателям копии сообщения.</span><span class="sxs-lookup"><span data-stu-id="1d898-256">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="1d898-257">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="1d898-257">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="1d898-258">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="1d898-258">Read mode</span></span>

<span data-ttu-id="1d898-p106">Свойство `cc` возвращает массив, который содержит объект `EmailAddressDetails` для каждого получателя, указанного в строке **Копия** сообщения. Коллекция может включать не более 100 элементов.</span><span class="sxs-lookup"><span data-stu-id="1d898-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="1d898-261">Режим создания</span><span class="sxs-lookup"><span data-stu-id="1d898-261">Compose mode</span></span>

<span data-ttu-id="1d898-262">Свойство `cc` возвращает объект `Recipients`, предоставляющий методы для получения или обновления получателей, которые указаны в строке **Копия** сообщения.</span><span class="sxs-lookup"><span data-stu-id="1d898-262">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="1d898-263">Тип:</span><span class="sxs-lookup"><span data-stu-id="1d898-263">Type:</span></span>

*   <span data-ttu-id="1d898-264">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="1d898-264">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="1d898-265">Требования</span><span class="sxs-lookup"><span data-stu-id="1d898-265">Requirements</span></span>

|<span data-ttu-id="1d898-266">Требование</span><span class="sxs-lookup"><span data-stu-id="1d898-266">Requirement</span></span>|<span data-ttu-id="1d898-267">Значение</span><span class="sxs-lookup"><span data-stu-id="1d898-267">Value</span></span>|
|---|---|
|[<span data-ttu-id="1d898-268">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="1d898-268">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1d898-269">1.0</span><span class="sxs-lookup"><span data-stu-id="1d898-269">1.0</span></span>|
|[<span data-ttu-id="1d898-270">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="1d898-270">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1d898-271">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1d898-271">ReadItem</span></span>|
|[<span data-ttu-id="1d898-272">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="1d898-272">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="1d898-273">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="1d898-273">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="1d898-274">Пример</span><span class="sxs-lookup"><span data-stu-id="1d898-274">Example</span></span>

```js
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="1d898-275">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="1d898-275">(nullable) conversationId :String</span></span>

<span data-ttu-id="1d898-276">Получает идентификатор разговора по электронной почте, содержащего конкретное сообщение.</span><span class="sxs-lookup"><span data-stu-id="1d898-276">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="1d898-p107">Вы можете получить целочисленное значение этого свойства, если ваше почтовое приложение активируется в формах просмотра или формах создания ответов. Если пользователь изменит тему ответа, после его отправки идентификатор беседы будет изменен, и полученное ранее значение будет недействительным.</span><span class="sxs-lookup"><span data-stu-id="1d898-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="1d898-p108">Это свойство имеет значение NULL для нового элемента в форме создания. Свойство `conversationId` вернет значение, если пользователь задаст тему и сохранит элемент.</span><span class="sxs-lookup"><span data-stu-id="1d898-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="1d898-281">Тип:</span><span class="sxs-lookup"><span data-stu-id="1d898-281">Type:</span></span>

*   <span data-ttu-id="1d898-282">String</span><span class="sxs-lookup"><span data-stu-id="1d898-282">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="1d898-283">Требования</span><span class="sxs-lookup"><span data-stu-id="1d898-283">Requirements</span></span>

|<span data-ttu-id="1d898-284">Требование</span><span class="sxs-lookup"><span data-stu-id="1d898-284">Requirement</span></span>|<span data-ttu-id="1d898-285">Значение</span><span class="sxs-lookup"><span data-stu-id="1d898-285">Value</span></span>|
|---|---|
|[<span data-ttu-id="1d898-286">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="1d898-286">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1d898-287">1.0</span><span class="sxs-lookup"><span data-stu-id="1d898-287">1.0</span></span>|
|[<span data-ttu-id="1d898-288">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="1d898-288">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1d898-289">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1d898-289">ReadItem</span></span>|
|[<span data-ttu-id="1d898-290">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="1d898-290">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="1d898-291">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="1d898-291">Compose or read</span></span>|

#### <a name="datetimecreated-date"></a><span data-ttu-id="1d898-292">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="1d898-292">dateTimeCreated :Date</span></span>

<span data-ttu-id="1d898-p109">Получает дату и время создания элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="1d898-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="1d898-295">Тип:</span><span class="sxs-lookup"><span data-stu-id="1d898-295">Type:</span></span>

*   <span data-ttu-id="1d898-296">Date</span><span class="sxs-lookup"><span data-stu-id="1d898-296">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="1d898-297">Требования</span><span class="sxs-lookup"><span data-stu-id="1d898-297">Requirements</span></span>

|<span data-ttu-id="1d898-298">Требование</span><span class="sxs-lookup"><span data-stu-id="1d898-298">Requirement</span></span>|<span data-ttu-id="1d898-299">Значение</span><span class="sxs-lookup"><span data-stu-id="1d898-299">Value</span></span>|
|---|---|
|[<span data-ttu-id="1d898-300">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="1d898-300">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1d898-301">1.0</span><span class="sxs-lookup"><span data-stu-id="1d898-301">1.0</span></span>|
|[<span data-ttu-id="1d898-302">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="1d898-302">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1d898-303">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1d898-303">ReadItem</span></span>|
|[<span data-ttu-id="1d898-304">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="1d898-304">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="1d898-305">Чтение</span><span class="sxs-lookup"><span data-stu-id="1d898-305">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1d898-306">Пример</span><span class="sxs-lookup"><span data-stu-id="1d898-306">Example</span></span>

```js
var created = Office.context.mailbox.item.dateTimeCreated;
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="1d898-307">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="1d898-307">dateTimeModified :Date</span></span>

<span data-ttu-id="1d898-p110">Получает дату и время последнего изменения элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="1d898-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="1d898-310">Этот элемент не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="1d898-310">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="1d898-311">Тип:</span><span class="sxs-lookup"><span data-stu-id="1d898-311">Type:</span></span>

*   <span data-ttu-id="1d898-312">Date</span><span class="sxs-lookup"><span data-stu-id="1d898-312">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="1d898-313">Требования</span><span class="sxs-lookup"><span data-stu-id="1d898-313">Requirements</span></span>

|<span data-ttu-id="1d898-314">Требование</span><span class="sxs-lookup"><span data-stu-id="1d898-314">Requirement</span></span>|<span data-ttu-id="1d898-315">Значение</span><span class="sxs-lookup"><span data-stu-id="1d898-315">Value</span></span>|
|---|---|
|[<span data-ttu-id="1d898-316">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="1d898-316">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1d898-317">1.0</span><span class="sxs-lookup"><span data-stu-id="1d898-317">1.0</span></span>|
|[<span data-ttu-id="1d898-318">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="1d898-318">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1d898-319">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1d898-319">ReadItem</span></span>|
|[<span data-ttu-id="1d898-320">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="1d898-320">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="1d898-321">Чтение</span><span class="sxs-lookup"><span data-stu-id="1d898-321">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1d898-322">Пример</span><span class="sxs-lookup"><span data-stu-id="1d898-322">Example</span></span>

```js
var modified = Office.context.mailbox.item.dateTimeModified;
```

####  <a name="end-datetimejavascriptapioutlook17officetime"></a><span data-ttu-id="1d898-323">end :Date|[Time](/javascript/api/outlook_1_7/office.time)</span><span class="sxs-lookup"><span data-stu-id="1d898-323">end :Date|[Time](/javascript/api/outlook_1_7/office.time)</span></span>

<span data-ttu-id="1d898-324">Получает или задает дату и время окончания встречи.</span><span class="sxs-lookup"><span data-stu-id="1d898-324">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="1d898-p111">Свойство `end` представлено в виде значения даты и времени в формате UTC. Преобразовать значение свойства end в местные значения даты и времени клиента можно с помощью метода [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook17officelocalclienttime).</span><span class="sxs-lookup"><span data-stu-id="1d898-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook17officelocalclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="1d898-327">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="1d898-327">Read mode</span></span>

<span data-ttu-id="1d898-328">Свойство `end` возвращает объект `Date`.</span><span class="sxs-lookup"><span data-stu-id="1d898-328">The `end` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="1d898-329">Режим создания</span><span class="sxs-lookup"><span data-stu-id="1d898-329">Compose mode</span></span>

<span data-ttu-id="1d898-330">Свойство `end` возвращает объект `Time`.</span><span class="sxs-lookup"><span data-stu-id="1d898-330">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="1d898-331">Если вы задаете время окончания с помощью метода [`Time.setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-), необходимо использовать метод [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) для преобразования местного времени на клиенте в формат UTC для сервера.</span><span class="sxs-lookup"><span data-stu-id="1d898-331">When you use the [`Time.setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="1d898-332">Тип:</span><span class="sxs-lookup"><span data-stu-id="1d898-332">Type:</span></span>

*   <span data-ttu-id="1d898-333">Date | [Time](/javascript/api/outlook_1_7/office.time)</span><span class="sxs-lookup"><span data-stu-id="1d898-333">Date | [Time](/javascript/api/outlook_1_7/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="1d898-334">Требования</span><span class="sxs-lookup"><span data-stu-id="1d898-334">Requirements</span></span>

|<span data-ttu-id="1d898-335">Требование</span><span class="sxs-lookup"><span data-stu-id="1d898-335">Requirement</span></span>|<span data-ttu-id="1d898-336">Значение</span><span class="sxs-lookup"><span data-stu-id="1d898-336">Value</span></span>|
|---|---|
|[<span data-ttu-id="1d898-337">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="1d898-337">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1d898-338">1.0</span><span class="sxs-lookup"><span data-stu-id="1d898-338">1.0</span></span>|
|[<span data-ttu-id="1d898-339">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="1d898-339">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1d898-340">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1d898-340">ReadItem</span></span>|
|[<span data-ttu-id="1d898-341">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="1d898-341">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="1d898-342">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="1d898-342">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="1d898-343">Пример</span><span class="sxs-lookup"><span data-stu-id="1d898-343">Example</span></span>

<span data-ttu-id="1d898-344">В примере ниже показано, как с помощью метода [`setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) объекта `Time` задать время окончания встречи в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="1d898-344">The following example sets the end time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

```js
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

#### <a name="from-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsfromjavascriptapioutlook17officefrom"></a><span data-ttu-id="1d898-345">from :[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)|[From](/javascript/api/outlook_1_7/office.from)</span><span class="sxs-lookup"><span data-stu-id="1d898-345">from :[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)|[From](/javascript/api/outlook_1_7/office.from)</span></span>

<span data-ttu-id="1d898-346">Получает адрес электронной почты отправителя сообщения.</span><span class="sxs-lookup"><span data-stu-id="1d898-346">Gets the email address of the sender of a message.</span></span>

<span data-ttu-id="1d898-p112">Свойства `from` и [`sender`](#sender-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetails) представляют одно лицо, если сообщение не отправлено представителем. В противном случае свойство `from` представляет лицо, делегировавшее полномочия, а свойство sender — представителя.</span><span class="sxs-lookup"><span data-stu-id="1d898-p112">The `from` and [`sender`](#sender-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="1d898-349">Свойству `recipientType`, принадлежащему объекту `EmailAddressDetails` в свойстве `from`, задано значение `undefined`.</span><span class="sxs-lookup"><span data-stu-id="1d898-349">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="1d898-350">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="1d898-350">Read mode</span></span>

<span data-ttu-id="1d898-351">Свойство `from` возвращает объект `EmailAddressDetails`.</span><span class="sxs-lookup"><span data-stu-id="1d898-351">The `from` property returns an `EmailAddressDetails` object.</span></span>

```js
var subject = Office.context.mailbox.item.from;
```

##### <a name="compose-mode"></a><span data-ttu-id="1d898-352">Режим создания</span><span class="sxs-lookup"><span data-stu-id="1d898-352">Compose mode</span></span>

<span data-ttu-id="1d898-353">Свойство `from` возвращает объект `From`, который предоставляет метод для получения значения отправителя.</span><span class="sxs-lookup"><span data-stu-id="1d898-353">The `from` property returns a `From` object that provides a method to get the from value.</span></span>

```js
Office.context.mailbox.item.from.getAsync(callback);

function callback(asyncResult) {
  var from = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="1d898-354">Тип:</span><span class="sxs-lookup"><span data-stu-id="1d898-354">Type:</span></span>

*   <span data-ttu-id="1d898-355">[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails) | [From](/javascript/api/outlook_1_7/office.from)</span><span class="sxs-lookup"><span data-stu-id="1d898-355">[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails) | [From](/javascript/api/outlook_1_7/office.from)</span></span>

##### <a name="requirements"></a><span data-ttu-id="1d898-356">Требования</span><span class="sxs-lookup"><span data-stu-id="1d898-356">Requirements</span></span>

|<span data-ttu-id="1d898-357">Требование</span><span class="sxs-lookup"><span data-stu-id="1d898-357">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="1d898-358">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="1d898-358">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1d898-359">1.0</span><span class="sxs-lookup"><span data-stu-id="1d898-359">1.0</span></span>|<span data-ttu-id="1d898-360">1.7</span><span class="sxs-lookup"><span data-stu-id="1d898-360">1.7</span></span>|
|[<span data-ttu-id="1d898-361">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="1d898-361">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1d898-362">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1d898-362">ReadItem</span></span>|<span data-ttu-id="1d898-363">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="1d898-363">ReadWriteItem</span></span>|
|[<span data-ttu-id="1d898-364">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="1d898-364">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="1d898-365">Чтение</span><span class="sxs-lookup"><span data-stu-id="1d898-365">Read</span></span>|<span data-ttu-id="1d898-366">Создание</span><span class="sxs-lookup"><span data-stu-id="1d898-366">Compose</span></span>|

#### <a name="internetmessageid-string"></a><span data-ttu-id="1d898-367">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="1d898-367">internetMessageId :String</span></span>

<span data-ttu-id="1d898-p113">Получает идентификатор интернет-сообщения для электронного сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="1d898-p113">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="1d898-370">Тип:</span><span class="sxs-lookup"><span data-stu-id="1d898-370">Type:</span></span>

*   <span data-ttu-id="1d898-371">String</span><span class="sxs-lookup"><span data-stu-id="1d898-371">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="1d898-372">Требования</span><span class="sxs-lookup"><span data-stu-id="1d898-372">Requirements</span></span>

|<span data-ttu-id="1d898-373">Требование</span><span class="sxs-lookup"><span data-stu-id="1d898-373">Requirement</span></span>|<span data-ttu-id="1d898-374">Значение</span><span class="sxs-lookup"><span data-stu-id="1d898-374">Value</span></span>|
|---|---|
|[<span data-ttu-id="1d898-375">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="1d898-375">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1d898-376">1.0</span><span class="sxs-lookup"><span data-stu-id="1d898-376">1.0</span></span>|
|[<span data-ttu-id="1d898-377">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="1d898-377">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1d898-378">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1d898-378">ReadItem</span></span>|
|[<span data-ttu-id="1d898-379">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="1d898-379">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="1d898-380">Чтение</span><span class="sxs-lookup"><span data-stu-id="1d898-380">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1d898-381">Пример</span><span class="sxs-lookup"><span data-stu-id="1d898-381">Example</span></span>

```js
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="1d898-382">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="1d898-382">itemClass :String</span></span>

<span data-ttu-id="1d898-p114">Получает класс элемента веб-служб Exchange для выбранного элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="1d898-p114">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="1d898-p115">Свойство `itemClass` указывает класс сообщения выбранного элемента. Ниже приводятся классы сообщения по умолчанию для элемента сообщения или встречи.</span><span class="sxs-lookup"><span data-stu-id="1d898-p115">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

|<span data-ttu-id="1d898-387">Тип</span><span class="sxs-lookup"><span data-stu-id="1d898-387">Type</span></span>|<span data-ttu-id="1d898-388">Описание</span><span class="sxs-lookup"><span data-stu-id="1d898-388">Description</span></span>|<span data-ttu-id="1d898-389">Класс элемента</span><span class="sxs-lookup"><span data-stu-id="1d898-389">item class</span></span>|
|---|---|---|
|<span data-ttu-id="1d898-390">Элементы встречи</span><span class="sxs-lookup"><span data-stu-id="1d898-390">Appointment items</span></span>|<span data-ttu-id="1d898-391">Это элементы календаря для класса элемента `IPM.Appointment` или `IPM.Appointment.Occurence`.</span><span class="sxs-lookup"><span data-stu-id="1d898-391">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurence`.</span></span>|`IPM.Appointment`<br />`IPM.Appointment.Occurence`|
|<span data-ttu-id="1d898-392">Элементы сообщения</span><span class="sxs-lookup"><span data-stu-id="1d898-392">Message items</span></span>|<span data-ttu-id="1d898-393">Сюда входят электронные сообщения, для которых по умолчанию задан класс сообщения `IPM.Note`, а также приглашения на собрания, ответы на них и уведомления об их отмене, использующие `IPM.Schedule.Meeting` в качестве базового класса сообщения.</span><span class="sxs-lookup"><span data-stu-id="1d898-393">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span>|`IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled`|

<span data-ttu-id="1d898-394">Можно создавать настраиваемые классы сообщения, расширяющие классы сообщения по умолчанию, например настраиваемый класс сообщения о встрече `IPM.Appointment.Contoso`.</span><span class="sxs-lookup"><span data-stu-id="1d898-394">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="1d898-395">Тип:</span><span class="sxs-lookup"><span data-stu-id="1d898-395">Type:</span></span>

*   <span data-ttu-id="1d898-396">String</span><span class="sxs-lookup"><span data-stu-id="1d898-396">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="1d898-397">Требования</span><span class="sxs-lookup"><span data-stu-id="1d898-397">Requirements</span></span>

|<span data-ttu-id="1d898-398">Требование</span><span class="sxs-lookup"><span data-stu-id="1d898-398">Requirement</span></span>|<span data-ttu-id="1d898-399">Значение</span><span class="sxs-lookup"><span data-stu-id="1d898-399">Value</span></span>|
|---|---|
|[<span data-ttu-id="1d898-400">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="1d898-400">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1d898-401">1.0</span><span class="sxs-lookup"><span data-stu-id="1d898-401">1.0</span></span>|
|[<span data-ttu-id="1d898-402">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="1d898-402">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1d898-403">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1d898-403">ReadItem</span></span>|
|[<span data-ttu-id="1d898-404">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="1d898-404">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="1d898-405">Чтение</span><span class="sxs-lookup"><span data-stu-id="1d898-405">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1d898-406">Пример</span><span class="sxs-lookup"><span data-stu-id="1d898-406">Example</span></span>

```js
var itemClass = Office.context.mailbox.item.itemClass;
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="1d898-407">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="1d898-407">(nullable) itemId :String</span></span>

<span data-ttu-id="1d898-p116">Получает идентификатор элемента веб-служб Exchange для текущего элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="1d898-p116">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="1d898-410">Идентификатор, возвращаемый свойством `itemId`, совпадает с идентификатором элемента веб-служб Exchange.</span><span class="sxs-lookup"><span data-stu-id="1d898-410">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="1d898-411">Свойство `itemId` не совпадает с идентификатором записи Outlook, а также идентификатором, который используется REST API Outlook.</span><span class="sxs-lookup"><span data-stu-id="1d898-411">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="1d898-412">Прежде чем совершать вызовы REST API, используя это значение, его необходимо преобразовать с помощью [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="1d898-412">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="1d898-413">Дополнительные сведения см. в статье [Использование REST API Outlook из надстройки Outlook](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="1d898-413">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="1d898-p118">Свойство `itemId` недоступно в режиме создания. Если требуется идентификатор элемента, с помощью метода [`saveAsync`](#saveasyncoptions-callback) можно сохранить элемент в хранилище. При этом в параметре [`AsyncResult.value`](/javascript/api/office/office.asyncresult) функции обратного вызова возвращается идентификатор элемента.</span><span class="sxs-lookup"><span data-stu-id="1d898-p118">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="1d898-416">Тип:</span><span class="sxs-lookup"><span data-stu-id="1d898-416">Type:</span></span>

*   <span data-ttu-id="1d898-417">String</span><span class="sxs-lookup"><span data-stu-id="1d898-417">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="1d898-418">Требования</span><span class="sxs-lookup"><span data-stu-id="1d898-418">Requirements</span></span>

|<span data-ttu-id="1d898-419">Требование</span><span class="sxs-lookup"><span data-stu-id="1d898-419">Requirement</span></span>|<span data-ttu-id="1d898-420">Значение</span><span class="sxs-lookup"><span data-stu-id="1d898-420">Value</span></span>|
|---|---|
|[<span data-ttu-id="1d898-421">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="1d898-421">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1d898-422">1.0</span><span class="sxs-lookup"><span data-stu-id="1d898-422">1.0</span></span>|
|[<span data-ttu-id="1d898-423">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="1d898-423">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1d898-424">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1d898-424">ReadItem</span></span>|
|[<span data-ttu-id="1d898-425">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="1d898-425">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="1d898-426">Чтение</span><span class="sxs-lookup"><span data-stu-id="1d898-426">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1d898-427">Пример</span><span class="sxs-lookup"><span data-stu-id="1d898-427">Example</span></span>

<span data-ttu-id="1d898-p119">Указанный ниже код проверяет наличие идентификатора элемента. Если свойство `itemId` возвращает значение `null` или `undefined`, элемент будет сохранен в хранилище, а из асинхронного результата будет получен идентификатор элемента.</span><span class="sxs-lookup"><span data-stu-id="1d898-p119">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```js
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result){
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook17officemailboxenumsitemtype"></a><span data-ttu-id="1d898-430">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_7/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="1d898-430">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_7/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="1d898-431">Получает тип элемента, который представляет экземпляр.</span><span class="sxs-lookup"><span data-stu-id="1d898-431">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="1d898-432">Свойство `itemType` возвращает одно из значений перечисления `ItemType`, которое указывает, является ли экземпляр объекта `item` сообщением или встречей.</span><span class="sxs-lookup"><span data-stu-id="1d898-432">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="1d898-433">Тип:</span><span class="sxs-lookup"><span data-stu-id="1d898-433">Type:</span></span>

*   [<span data-ttu-id="1d898-434">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="1d898-434">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook_1_7/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="1d898-435">Требования</span><span class="sxs-lookup"><span data-stu-id="1d898-435">Requirements</span></span>

|<span data-ttu-id="1d898-436">Требование</span><span class="sxs-lookup"><span data-stu-id="1d898-436">Requirement</span></span>|<span data-ttu-id="1d898-437">Значение</span><span class="sxs-lookup"><span data-stu-id="1d898-437">Value</span></span>|
|---|---|
|[<span data-ttu-id="1d898-438">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="1d898-438">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1d898-439">1.0</span><span class="sxs-lookup"><span data-stu-id="1d898-439">1.0</span></span>|
|[<span data-ttu-id="1d898-440">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="1d898-440">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1d898-441">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1d898-441">ReadItem</span></span>|
|[<span data-ttu-id="1d898-442">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="1d898-442">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="1d898-443">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="1d898-443">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="1d898-444">Пример</span><span class="sxs-lookup"><span data-stu-id="1d898-444">Example</span></span>

```js
if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Message)
  // do something
else
  // do something else
```

####  <a name="location-stringlocationjavascriptapioutlook17officelocation"></a><span data-ttu-id="1d898-445">location :String|[Location](/javascript/api/outlook_1_7/office.location)</span><span class="sxs-lookup"><span data-stu-id="1d898-445">location :String|[Location](/javascript/api/outlook_1_7/office.location)</span></span>

<span data-ttu-id="1d898-446">Получает или задает место встречи.</span><span class="sxs-lookup"><span data-stu-id="1d898-446">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="1d898-447">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="1d898-447">Read mode</span></span>

<span data-ttu-id="1d898-448">Свойство `location` возвращает строку, содержащую сведения о месте встречи.</span><span class="sxs-lookup"><span data-stu-id="1d898-448">The `location` property returns a string that contains the location of the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="1d898-449">Режим создания</span><span class="sxs-lookup"><span data-stu-id="1d898-449">Compose mode</span></span>

<span data-ttu-id="1d898-450">Свойство `location` возвращает объект `Location`, предоставляющий методы, которые используются для получения и задания места встречи.</span><span class="sxs-lookup"><span data-stu-id="1d898-450">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="1d898-451">Тип:</span><span class="sxs-lookup"><span data-stu-id="1d898-451">Type:</span></span>

*   <span data-ttu-id="1d898-452">String | [Location](/javascript/api/outlook_1_7/office.location)</span><span class="sxs-lookup"><span data-stu-id="1d898-452">String | [Location](/javascript/api/outlook_1_7/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="1d898-453">Требования</span><span class="sxs-lookup"><span data-stu-id="1d898-453">Requirements</span></span>

|<span data-ttu-id="1d898-454">Требование</span><span class="sxs-lookup"><span data-stu-id="1d898-454">Requirement</span></span>|<span data-ttu-id="1d898-455">Значение</span><span class="sxs-lookup"><span data-stu-id="1d898-455">Value</span></span>|
|---|---|
|[<span data-ttu-id="1d898-456">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="1d898-456">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1d898-457">1.0</span><span class="sxs-lookup"><span data-stu-id="1d898-457">1.0</span></span>|
|[<span data-ttu-id="1d898-458">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="1d898-458">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1d898-459">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1d898-459">ReadItem</span></span>|
|[<span data-ttu-id="1d898-460">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="1d898-460">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="1d898-461">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="1d898-461">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="1d898-462">Пример</span><span class="sxs-lookup"><span data-stu-id="1d898-462">Example</span></span>

```js
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

#### <a name="normalizedsubject-string"></a><span data-ttu-id="1d898-463">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="1d898-463">normalizedSubject :String</span></span>

<span data-ttu-id="1d898-p120">Получает тему элемента со всеми удаленными префиксами (включая `RE:` и `FWD:`). Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="1d898-p120">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="1d898-p121">Свойство normalizedSubject получает тему элемента со стандартными префиксами (такими как `RE:` и `FW:`), добавляемыми почтовыми программами. Для получения темы элемента с неизмененными префиксами используйте свойство [`subject`](#subject-stringsubjectjavascriptapioutlook17officesubject).</span><span class="sxs-lookup"><span data-stu-id="1d898-p121">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubjectjavascriptapioutlook17officesubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="1d898-468">Тип:</span><span class="sxs-lookup"><span data-stu-id="1d898-468">Type:</span></span>

*   <span data-ttu-id="1d898-469">String</span><span class="sxs-lookup"><span data-stu-id="1d898-469">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="1d898-470">Требования</span><span class="sxs-lookup"><span data-stu-id="1d898-470">Requirements</span></span>

|<span data-ttu-id="1d898-471">Требование</span><span class="sxs-lookup"><span data-stu-id="1d898-471">Requirement</span></span>|<span data-ttu-id="1d898-472">Значение</span><span class="sxs-lookup"><span data-stu-id="1d898-472">Value</span></span>|
|---|---|
|[<span data-ttu-id="1d898-473">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="1d898-473">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1d898-474">1.0</span><span class="sxs-lookup"><span data-stu-id="1d898-474">1.0</span></span>|
|[<span data-ttu-id="1d898-475">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="1d898-475">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1d898-476">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1d898-476">ReadItem</span></span>|
|[<span data-ttu-id="1d898-477">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="1d898-477">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="1d898-478">Чтение</span><span class="sxs-lookup"><span data-stu-id="1d898-478">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1d898-479">Пример</span><span class="sxs-lookup"><span data-stu-id="1d898-479">Example</span></span>

```js
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
```

####  <a name="notificationmessages-notificationmessagesjavascriptapioutlook17officenotificationmessages"></a><span data-ttu-id="1d898-480">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_7/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="1d898-480">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_7/office.notificationmessages)</span></span>

<span data-ttu-id="1d898-481">Получает сообщения уведомления для элемента.</span><span class="sxs-lookup"><span data-stu-id="1d898-481">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="1d898-482">Тип:</span><span class="sxs-lookup"><span data-stu-id="1d898-482">Type:</span></span>

*   [<span data-ttu-id="1d898-483">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="1d898-483">NotificationMessages</span></span>](/javascript/api/outlook_1_7/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="1d898-484">Требования</span><span class="sxs-lookup"><span data-stu-id="1d898-484">Requirements</span></span>

|<span data-ttu-id="1d898-485">Требование</span><span class="sxs-lookup"><span data-stu-id="1d898-485">Requirement</span></span>|<span data-ttu-id="1d898-486">Значение</span><span class="sxs-lookup"><span data-stu-id="1d898-486">Value</span></span>|
|---|---|
|[<span data-ttu-id="1d898-487">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="1d898-487">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1d898-488">1.3</span><span class="sxs-lookup"><span data-stu-id="1d898-488">1.3</span></span>|
|[<span data-ttu-id="1d898-489">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="1d898-489">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1d898-490">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1d898-490">ReadItem</span></span>|
|[<span data-ttu-id="1d898-491">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="1d898-491">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="1d898-492">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="1d898-492">Compose or read</span></span>|

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients"></a><span data-ttu-id="1d898-493">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="1d898-493">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

<span data-ttu-id="1d898-494">Предоставляет доступ к необязательным участникам события.</span><span class="sxs-lookup"><span data-stu-id="1d898-494">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="1d898-495">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="1d898-495">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="1d898-496">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="1d898-496">Read mode</span></span>

<span data-ttu-id="1d898-497">Свойство `optionalAttendees` возвращает массив, содержащий объект `EmailAddressDetails` для каждого необязательного участника собрания.</span><span class="sxs-lookup"><span data-stu-id="1d898-497">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="1d898-498">Режим создания</span><span class="sxs-lookup"><span data-stu-id="1d898-498">Compose mode</span></span>

<span data-ttu-id="1d898-499">Свойство `optionalAttendees` возвращает объект `Recipients`, который предоставляет методы для получения или обновления необязательных участников собрания.</span><span class="sxs-lookup"><span data-stu-id="1d898-499">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="1d898-500">Тип:</span><span class="sxs-lookup"><span data-stu-id="1d898-500">Type:</span></span>

*   <span data-ttu-id="1d898-501">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="1d898-501">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="1d898-502">Требования</span><span class="sxs-lookup"><span data-stu-id="1d898-502">Requirements</span></span>

|<span data-ttu-id="1d898-503">Требование</span><span class="sxs-lookup"><span data-stu-id="1d898-503">Requirement</span></span>|<span data-ttu-id="1d898-504">Значение</span><span class="sxs-lookup"><span data-stu-id="1d898-504">Value</span></span>|
|---|---|
|[<span data-ttu-id="1d898-505">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="1d898-505">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1d898-506">1.0</span><span class="sxs-lookup"><span data-stu-id="1d898-506">1.0</span></span>|
|[<span data-ttu-id="1d898-507">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="1d898-507">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1d898-508">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1d898-508">ReadItem</span></span>|
|[<span data-ttu-id="1d898-509">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="1d898-509">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="1d898-510">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="1d898-510">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="1d898-511">Пример</span><span class="sxs-lookup"><span data-stu-id="1d898-511">Example</span></span>

```js
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsorganizerjavascriptapioutlook17officeorganizer"></a><span data-ttu-id="1d898-512">organizer :[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)|[Organizer](/javascript/api/outlook_1_7/office.organizer)</span><span class="sxs-lookup"><span data-stu-id="1d898-512">organizer :[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)|[Organizer](/javascript/api/outlook_1_7/office.organizer)</span></span>

<span data-ttu-id="1d898-513">Получает адрес электронной почты организатора указанного собрания.</span><span class="sxs-lookup"><span data-stu-id="1d898-513">Gets the email address of the organizer for a specified meeting.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="1d898-514">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="1d898-514">Read mode</span></span>

<span data-ttu-id="1d898-515">Свойство `organizer` возвращает объект [EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails), представляющий организатора собрания.</span><span class="sxs-lookup"><span data-stu-id="1d898-515">The `organizer` property returns an [EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails) object that represents the meeting organizer.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="1d898-516">Режим создания</span><span class="sxs-lookup"><span data-stu-id="1d898-516">Compose mode</span></span>

<span data-ttu-id="1d898-517">Свойство `organizer` возвращает объект [Organizer](/javascript/api/outlook_1_7/office.organizer), который предоставляет метод для получения значения организатора.</span><span class="sxs-lookup"><span data-stu-id="1d898-517">The `organizer` property returns an [Organizer](/javascript/api/outlook_1_7/office.organizer) object that provides a method to get the organizer value.</span></span>

##### <a name="type"></a><span data-ttu-id="1d898-518">Тип:</span><span class="sxs-lookup"><span data-stu-id="1d898-518">Type:</span></span>

*   <span data-ttu-id="1d898-519">[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails) | [Organizer](/javascript/api/outlook_1_7/office.organizer)</span><span class="sxs-lookup"><span data-stu-id="1d898-519">[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails) | [Organizer](/javascript/api/outlook_1_7/office.organizer)</span></span>

##### <a name="requirements"></a><span data-ttu-id="1d898-520">Требования</span><span class="sxs-lookup"><span data-stu-id="1d898-520">Requirements</span></span>

|<span data-ttu-id="1d898-521">Требование</span><span class="sxs-lookup"><span data-stu-id="1d898-521">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="1d898-522">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="1d898-522">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1d898-523">1.0</span><span class="sxs-lookup"><span data-stu-id="1d898-523">1.0</span></span>|<span data-ttu-id="1d898-524">1.7</span><span class="sxs-lookup"><span data-stu-id="1d898-524">1.7</span></span>|
|[<span data-ttu-id="1d898-525">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="1d898-525">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1d898-526">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1d898-526">ReadItem</span></span>|<span data-ttu-id="1d898-527">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="1d898-527">ReadWriteItem</span></span>|
|[<span data-ttu-id="1d898-528">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="1d898-528">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="1d898-529">Чтение</span><span class="sxs-lookup"><span data-stu-id="1d898-529">Read</span></span>|<span data-ttu-id="1d898-530">Создание</span><span class="sxs-lookup"><span data-stu-id="1d898-530">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="1d898-531">Пример</span><span class="sxs-lookup"><span data-stu-id="1d898-531">Example</span></span>

```js
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
```

#### <a name="nullable-recurrence-recurrencejavascriptapioutlook17officerecurrence"></a><span data-ttu-id="1d898-532">(nullable) recurrence :[Recurrence](/javascript/api/outlook_1_7/office.recurrence)</span><span class="sxs-lookup"><span data-stu-id="1d898-532">(nullable) recurrence :[Recurrence](/javascript/api/outlook_1_7/office.recurrence)</span></span>

<span data-ttu-id="1d898-533">Получает или задает расписание повторения для встречи.</span><span class="sxs-lookup"><span data-stu-id="1d898-533">Gets or sets the recurrence pattern of an appointment.</span></span> <span data-ttu-id="1d898-534">Получает расписание повторения для приглашения на собрание.</span><span class="sxs-lookup"><span data-stu-id="1d898-534">Gets the recurrence pattern of a meeting request.</span></span> <span data-ttu-id="1d898-535">Доступно в режимах чтения и создания для элементов встречи.</span><span class="sxs-lookup"><span data-stu-id="1d898-535">Read and compose modes for appointment items.</span></span> <span data-ttu-id="1d898-536">Доступно в режиме чтения для элементов приглашения на собрание.</span><span class="sxs-lookup"><span data-stu-id="1d898-536">Read mode for meeting request items.</span></span>

<span data-ttu-id="1d898-537">Свойство `recurrence` возвращает объект [recurrence](/javascript/api/outlook_1_7/office.recurrence) для повторяющихся встреч или приглашений на собрание, если элемент представляет собой серию или экземпляр в пределах серии.</span><span class="sxs-lookup"><span data-stu-id="1d898-537">The `recurrence` property returns a [recurrence](/javascript/api/outlook_1_7/office.recurrence) object for recurring appointments or meetings requests if an item is a series or an instance in a series.</span></span> <span data-ttu-id="1d898-538">Значение `null` возвращается для отдельных встреч и приглашений на собрания, связанных с одной встречей.</span><span class="sxs-lookup"><span data-stu-id="1d898-538">`null` is returned for single appointments and meeting requests of single appointments.</span></span> <span data-ttu-id="1d898-539">Значение `undefined` возвращается для сообщений, которые не являются приглашениями на собрания.</span><span class="sxs-lookup"><span data-stu-id="1d898-539">`undefined` is returned for messages that are not meeting requests.</span></span>

> <span data-ttu-id="1d898-540">Примечание. Приглашения на собрания имеют значение `itemClass` для класса IPM.Schedule.Meeting.Request.</span><span class="sxs-lookup"><span data-stu-id="1d898-540">Note: Meeting requests have an `itemClass` value of IPM.Schedule.Meeting.Request.</span></span>

> <span data-ttu-id="1d898-541">Примечание. Если объект recurrence имеет значение `null`, он представляет собой отдельную встречу или приглашение на собрание, связанное с одной встречей, и НЕ входит в серию.</span><span class="sxs-lookup"><span data-stu-id="1d898-541">Note: If the recurrence object is `null`, this indicates that the object is a single appointment or a meeting request of a single appointment and NOT a part of a series.</span></span>

##### <a name="type"></a><span data-ttu-id="1d898-542">Тип:</span><span class="sxs-lookup"><span data-stu-id="1d898-542">Type:</span></span>

* [<span data-ttu-id="1d898-543">Recurrence</span><span class="sxs-lookup"><span data-stu-id="1d898-543">Recurrence</span></span>](/javascript/api/outlook_1_7/office.recurrence)

|<span data-ttu-id="1d898-544">Требование</span><span class="sxs-lookup"><span data-stu-id="1d898-544">Requirement</span></span>|<span data-ttu-id="1d898-545">Значение</span><span class="sxs-lookup"><span data-stu-id="1d898-545">Value</span></span>|
|---|---|
|[<span data-ttu-id="1d898-546">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="1d898-546">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1d898-547">1.7</span><span class="sxs-lookup"><span data-stu-id="1d898-547">1.7</span></span>|
|[<span data-ttu-id="1d898-548">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="1d898-548">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1d898-549">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1d898-549">ReadItem</span></span>|
|[<span data-ttu-id="1d898-550">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="1d898-550">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="1d898-551">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="1d898-551">Compose or read</span></span>|

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients"></a><span data-ttu-id="1d898-552">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="1d898-552">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

<span data-ttu-id="1d898-553">Предоставляет доступ к обязательным участникам события.</span><span class="sxs-lookup"><span data-stu-id="1d898-553">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="1d898-554">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="1d898-554">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="1d898-555">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="1d898-555">Read mode</span></span>

<span data-ttu-id="1d898-556">Свойство `requiredAttendees` возвращает массив, содержащий объект `EmailAddressDetails` для каждого обязательного участника собрания.</span><span class="sxs-lookup"><span data-stu-id="1d898-556">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="1d898-557">Режим создания</span><span class="sxs-lookup"><span data-stu-id="1d898-557">Compose mode</span></span>

<span data-ttu-id="1d898-558">Свойство `requiredAttendees` возвращает объект `Recipients`, предоставляющий методы, с помощью которых можно получить или обновить сведения об обязательных участниках собрания.</span><span class="sxs-lookup"><span data-stu-id="1d898-558">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="1d898-559">Тип:</span><span class="sxs-lookup"><span data-stu-id="1d898-559">Type:</span></span>

*   <span data-ttu-id="1d898-560">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="1d898-560">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="1d898-561">Требования</span><span class="sxs-lookup"><span data-stu-id="1d898-561">Requirements</span></span>

|<span data-ttu-id="1d898-562">Требование</span><span class="sxs-lookup"><span data-stu-id="1d898-562">Requirement</span></span>|<span data-ttu-id="1d898-563">Значение</span><span class="sxs-lookup"><span data-stu-id="1d898-563">Value</span></span>|
|---|---|
|[<span data-ttu-id="1d898-564">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="1d898-564">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1d898-565">1.0</span><span class="sxs-lookup"><span data-stu-id="1d898-565">1.0</span></span>|
|[<span data-ttu-id="1d898-566">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="1d898-566">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1d898-567">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1d898-567">ReadItem</span></span>|
|[<span data-ttu-id="1d898-568">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="1d898-568">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="1d898-569">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="1d898-569">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="1d898-570">Пример</span><span class="sxs-lookup"><span data-stu-id="1d898-570">Example</span></span>

```js
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
}
```

#### <a name="sender-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetails"></a><span data-ttu-id="1d898-571">sender :[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="1d898-571">sender :[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)</span></span>

<span data-ttu-id="1d898-p126">Получает электронный адрес отправителя электронного сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="1d898-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="1d898-p127">Свойства [`from`](#from-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsfromjavascriptapioutlook17officefrom) и `sender` представляют одно лицо, если сообщение не отправлено представителем. В противном случае свойство `from` представляет лицо, делегировавшее полномочия, а свойство sender — представителя.</span><span class="sxs-lookup"><span data-stu-id="1d898-p127">The [`from`](#from-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsfromjavascriptapioutlook17officefrom) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="1d898-576">Свойству `recipientType`, принадлежащему объекту `EmailAddressDetails` в свойстве `sender`, задано значение `undefined`.</span><span class="sxs-lookup"><span data-stu-id="1d898-576">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="1d898-577">Тип:</span><span class="sxs-lookup"><span data-stu-id="1d898-577">Type:</span></span>

*   [<span data-ttu-id="1d898-578">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="1d898-578">EmailAddressDetails</span></span>](/javascript/api/outlook_1_7/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="1d898-579">Требования</span><span class="sxs-lookup"><span data-stu-id="1d898-579">Requirements</span></span>

|<span data-ttu-id="1d898-580">Требование</span><span class="sxs-lookup"><span data-stu-id="1d898-580">Requirement</span></span>|<span data-ttu-id="1d898-581">Значение</span><span class="sxs-lookup"><span data-stu-id="1d898-581">Value</span></span>|
|---|---|
|[<span data-ttu-id="1d898-582">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="1d898-582">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1d898-583">1.0</span><span class="sxs-lookup"><span data-stu-id="1d898-583">1.0</span></span>|
|[<span data-ttu-id="1d898-584">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="1d898-584">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1d898-585">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1d898-585">ReadItem</span></span>|
|[<span data-ttu-id="1d898-586">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="1d898-586">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="1d898-587">Чтение</span><span class="sxs-lookup"><span data-stu-id="1d898-587">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1d898-588">Пример</span><span class="sxs-lookup"><span data-stu-id="1d898-588">Example</span></span>

```js
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
```

#### <a name="nullable-seriesid-string"></a><span data-ttu-id="1d898-589">(nullable) seriesId :String</span><span class="sxs-lookup"><span data-stu-id="1d898-589">(nullable) seriesId :String</span></span>

<span data-ttu-id="1d898-590">Получает идентификатор серии, к которой относится экземпляр.</span><span class="sxs-lookup"><span data-stu-id="1d898-590">Gets the id of the series that an instance belongs to.</span></span>

<span data-ttu-id="1d898-591">В Outlook Web App и Outlook свойство `seriesId` возвращает идентификатор веб-служб Exchange (EWS) родительского элемента (серии), к которому относится этот элемент.</span><span class="sxs-lookup"><span data-stu-id="1d898-591">In OWA and Outlook, the `seriesId` returns the Exchange Web Services (EWS) ID of the parent (series) item that this item belongs to.</span></span> <span data-ttu-id="1d898-592">Однако в iOS и Android свойство `seriesId` возвращает идентификатор REST родительского элемента.</span><span class="sxs-lookup"><span data-stu-id="1d898-592">However, in iOS and Android, the `seriesId` returns the REST ID of the parent item.</span></span>

> [!NOTE]
> <span data-ttu-id="1d898-593">Идентификатор, возвращаемый свойством `seriesId`, совпадает с идентификатором элемента веб-служб Exchange.</span><span class="sxs-lookup"><span data-stu-id="1d898-593">The identifier returned by the `seriesId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="1d898-594">Свойство `seriesId` не совпадает с идентификаторами Outlook, которые используются в REST API Outlook.</span><span class="sxs-lookup"><span data-stu-id="1d898-594">The `seriesId` property is not identical to the Outlook IDs used by the Outlook REST API.</span></span> <span data-ttu-id="1d898-595">Прежде чем совершать вызовы REST API, используя это значение, его необходимо преобразовать с помощью [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="1d898-595">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="1d898-596">Дополнительные сведения см. в статье [Использование REST API Outlook из надстройки Outlook](https://docs.microsoft.com/outlook/add-ins/use-rest-api).</span><span class="sxs-lookup"><span data-stu-id="1d898-596">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api).</span></span>

<span data-ttu-id="1d898-597">Свойство `seriesId` возвращает значение `null` для элементов, у которых нет родительских элементов, например отдельных встреч, элементов серий или приглашений на собрания, и возвращает значение `undefined` для всех других элементов, которые не представляют собой приглашения на собрания.</span><span class="sxs-lookup"><span data-stu-id="1d898-597">The `seriesId` property returns `null` for items that do not have parent items such as single appointments, series items, or meeting requests and returns `undefined` for any other items that are not meeting requests.</span></span>

##### <a name="type"></a><span data-ttu-id="1d898-598">Тип:</span><span class="sxs-lookup"><span data-stu-id="1d898-598">Type:</span></span>

* <span data-ttu-id="1d898-599">String</span><span class="sxs-lookup"><span data-stu-id="1d898-599">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="1d898-600">Требования</span><span class="sxs-lookup"><span data-stu-id="1d898-600">Requirements</span></span>

|<span data-ttu-id="1d898-601">Требование</span><span class="sxs-lookup"><span data-stu-id="1d898-601">Requirement</span></span>|<span data-ttu-id="1d898-602">Значение</span><span class="sxs-lookup"><span data-stu-id="1d898-602">Value</span></span>|
|---|---|
|[<span data-ttu-id="1d898-603">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="1d898-603">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1d898-604">1.7</span><span class="sxs-lookup"><span data-stu-id="1d898-604">1.7</span></span>|
|[<span data-ttu-id="1d898-605">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="1d898-605">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1d898-606">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1d898-606">ReadItem</span></span>|
|[<span data-ttu-id="1d898-607">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="1d898-607">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="1d898-608">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="1d898-608">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="1d898-609">Пример</span><span class="sxs-lookup"><span data-stu-id="1d898-609">Example</span></span>

```js
var seriesId = Office.context.mailbox.item.seriesId;
var isSeries = (seriesId == null);
```

####  <a name="start-datetimejavascriptapioutlook17officetime"></a><span data-ttu-id="1d898-610">start :Date|[Time](/javascript/api/outlook_1_7/office.time)</span><span class="sxs-lookup"><span data-stu-id="1d898-610">start :Date|[Time](/javascript/api/outlook_1_7/office.time)</span></span>

<span data-ttu-id="1d898-611">Получает или задает дату и время начала встречи.</span><span class="sxs-lookup"><span data-stu-id="1d898-611">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="1d898-p130">Свойство `start` представлено в виде значения даты и времени в формате UTC. Это значение можно преобразовать в местные значения даты и времени клиента с помощью метода [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook17officelocalclienttime).</span><span class="sxs-lookup"><span data-stu-id="1d898-p130">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook17officelocalclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="1d898-614">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="1d898-614">Read mode</span></span>

<span data-ttu-id="1d898-615">Свойство `start` возвращает объект `Date`.</span><span class="sxs-lookup"><span data-stu-id="1d898-615">The `start` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="1d898-616">Режим создания</span><span class="sxs-lookup"><span data-stu-id="1d898-616">Compose mode</span></span>

<span data-ttu-id="1d898-617">Свойство `start` возвращает объект `Time`.</span><span class="sxs-lookup"><span data-stu-id="1d898-617">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="1d898-618">Если вы задаете время начала с помощью метода [`Time.setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-), необходимо использовать метод [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) для преобразования местного времени на клиенте в формат UTC для сервера.</span><span class="sxs-lookup"><span data-stu-id="1d898-618">When you use the [`Time.setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="1d898-619">Тип:</span><span class="sxs-lookup"><span data-stu-id="1d898-619">Type:</span></span>

*   <span data-ttu-id="1d898-620">Date | [Time](/javascript/api/outlook_1_7/office.time)</span><span class="sxs-lookup"><span data-stu-id="1d898-620">Date | [Time](/javascript/api/outlook_1_7/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="1d898-621">Требования</span><span class="sxs-lookup"><span data-stu-id="1d898-621">Requirements</span></span>

|<span data-ttu-id="1d898-622">Требование</span><span class="sxs-lookup"><span data-stu-id="1d898-622">Requirement</span></span>|<span data-ttu-id="1d898-623">Значение</span><span class="sxs-lookup"><span data-stu-id="1d898-623">Value</span></span>|
|---|---|
|[<span data-ttu-id="1d898-624">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="1d898-624">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1d898-625">1.0</span><span class="sxs-lookup"><span data-stu-id="1d898-625">1.0</span></span>|
|[<span data-ttu-id="1d898-626">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="1d898-626">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1d898-627">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1d898-627">ReadItem</span></span>|
|[<span data-ttu-id="1d898-628">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="1d898-628">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="1d898-629">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="1d898-629">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="1d898-630">Пример</span><span class="sxs-lookup"><span data-stu-id="1d898-630">Example</span></span>

<span data-ttu-id="1d898-631">В примере ниже с помощью метода [`setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) объекта `Time` задается время начала встречи в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="1d898-631">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

```js
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

####  <a name="subject-stringsubjectjavascriptapioutlook17officesubject"></a><span data-ttu-id="1d898-632">subject :String|[Subject](/javascript/api/outlook_1_7/office.subject)</span><span class="sxs-lookup"><span data-stu-id="1d898-632">subject :String|[Subject](/javascript/api/outlook_1_7/office.subject)</span></span>

<span data-ttu-id="1d898-633">Получает или задает описание, которое отображается в поле темы элемента.</span><span class="sxs-lookup"><span data-stu-id="1d898-633">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="1d898-634">Свойство `subject` получает или задает всю тему элемента для отправки с почтового сервера.</span><span class="sxs-lookup"><span data-stu-id="1d898-634">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="1d898-635">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="1d898-635">Read mode</span></span>

<span data-ttu-id="1d898-p131">Свойство `subject` возвращает строку. С помощью свойства [`normalizedSubject`](#normalizedsubject-string) можно получить тему без начальных префиксов, таких как `RE:` и `FW:`.</span><span class="sxs-lookup"><span data-stu-id="1d898-p131">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```js
var subject = Office.context.mailbox.item.subject;
```

##### <a name="compose-mode"></a><span data-ttu-id="1d898-638">Режим создания</span><span class="sxs-lookup"><span data-stu-id="1d898-638">Compose mode</span></span>

<span data-ttu-id="1d898-639">Свойство `subject` возвращает объект `Subject`, который предоставляет методы для получения и задания темы.</span><span class="sxs-lookup"><span data-stu-id="1d898-639">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```js
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="1d898-640">Тип:</span><span class="sxs-lookup"><span data-stu-id="1d898-640">Type:</span></span>

*   <span data-ttu-id="1d898-641">String | [Subject](/javascript/api/outlook_1_7/office.subject)</span><span class="sxs-lookup"><span data-stu-id="1d898-641">String | [Subject](/javascript/api/outlook_1_7/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="1d898-642">Требования</span><span class="sxs-lookup"><span data-stu-id="1d898-642">Requirements</span></span>

|<span data-ttu-id="1d898-643">Требование</span><span class="sxs-lookup"><span data-stu-id="1d898-643">Requirement</span></span>|<span data-ttu-id="1d898-644">Значение</span><span class="sxs-lookup"><span data-stu-id="1d898-644">Value</span></span>|
|---|---|
|[<span data-ttu-id="1d898-645">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="1d898-645">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1d898-646">1.0</span><span class="sxs-lookup"><span data-stu-id="1d898-646">1.0</span></span>|
|[<span data-ttu-id="1d898-647">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="1d898-647">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1d898-648">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1d898-648">ReadItem</span></span>|
|[<span data-ttu-id="1d898-649">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="1d898-649">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="1d898-650">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="1d898-650">Compose or read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients"></a><span data-ttu-id="1d898-651">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="1d898-651">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

<span data-ttu-id="1d898-652">Предоставляет доступ к получателям, указанным в строке **Кому** сообщения.</span><span class="sxs-lookup"><span data-stu-id="1d898-652">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="1d898-653">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="1d898-653">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="1d898-654">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="1d898-654">Read mode</span></span>

<span data-ttu-id="1d898-p133">Свойство `to` возвращает массив, содержащий объект `EmailAddressDetails` для каждого получателя в строке **Кому** сообщения. Коллекция может включать не более 100 элементов.</span><span class="sxs-lookup"><span data-stu-id="1d898-p133">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="1d898-657">Режим создания</span><span class="sxs-lookup"><span data-stu-id="1d898-657">Compose mode</span></span>

<span data-ttu-id="1d898-658">Свойство `to` возвращает объект `Recipients`, предоставляющий методы для получения или обновления получателей, которые указаны в строке **Кому** сообщения.</span><span class="sxs-lookup"><span data-stu-id="1d898-658">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="1d898-659">Тип:</span><span class="sxs-lookup"><span data-stu-id="1d898-659">Type:</span></span>

*   <span data-ttu-id="1d898-660">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="1d898-660">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="1d898-661">Требования</span><span class="sxs-lookup"><span data-stu-id="1d898-661">Requirements</span></span>

|<span data-ttu-id="1d898-662">Требование</span><span class="sxs-lookup"><span data-stu-id="1d898-662">Requirement</span></span>|<span data-ttu-id="1d898-663">Значение</span><span class="sxs-lookup"><span data-stu-id="1d898-663">Value</span></span>|
|---|---|
|[<span data-ttu-id="1d898-664">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="1d898-664">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1d898-665">1.0</span><span class="sxs-lookup"><span data-stu-id="1d898-665">1.0</span></span>|
|[<span data-ttu-id="1d898-666">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="1d898-666">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1d898-667">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1d898-667">ReadItem</span></span>|
|[<span data-ttu-id="1d898-668">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="1d898-668">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="1d898-669">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="1d898-669">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="1d898-670">Пример</span><span class="sxs-lookup"><span data-stu-id="1d898-670">Example</span></span>

```js
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

### <a name="methods"></a><span data-ttu-id="1d898-671">Методы</span><span class="sxs-lookup"><span data-stu-id="1d898-671">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="1d898-672">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="1d898-672">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="1d898-673">Добавляет файл в сообщение или встречу в качестве вложения.</span><span class="sxs-lookup"><span data-stu-id="1d898-673">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="1d898-674">Метод `addFileAttachmentAsync` передает файл по указанному универсальному коду ресурса (URI) и вкладывает его в элемент в форме создания.</span><span class="sxs-lookup"><span data-stu-id="1d898-674">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="1d898-675">Идентификатор можно последовательно использовать с методом [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback), чтобы удалить вложение, добавленное во время текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="1d898-675">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1d898-676">Параметры</span><span class="sxs-lookup"><span data-stu-id="1d898-676">Parameters:</span></span>
|<span data-ttu-id="1d898-677">Имя</span><span class="sxs-lookup"><span data-stu-id="1d898-677">Name</span></span>|<span data-ttu-id="1d898-678">Тип</span><span class="sxs-lookup"><span data-stu-id="1d898-678">Type</span></span>|<span data-ttu-id="1d898-679">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="1d898-679">Attributes</span></span>|<span data-ttu-id="1d898-680">Описание</span><span class="sxs-lookup"><span data-stu-id="1d898-680">Description</span></span>|
|---|---|---|---|
|`uri`|<span data-ttu-id="1d898-681">String</span><span class="sxs-lookup"><span data-stu-id="1d898-681">String</span></span>||<span data-ttu-id="1d898-p134">Универсальный код ресурса (URI), представляющий расположение файла, который нужно вложить в сообщение или встречу. Максимальная длина — 2048 символов.</span><span class="sxs-lookup"><span data-stu-id="1d898-p134">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="1d898-684">String</span><span class="sxs-lookup"><span data-stu-id="1d898-684">String</span></span>||<span data-ttu-id="1d898-p135">Имя вложения, которое отображается при передаче вложения. Максимальная длина — 255 символов.</span><span class="sxs-lookup"><span data-stu-id="1d898-p135">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="1d898-687">Object</span><span class="sxs-lookup"><span data-stu-id="1d898-687">Object</span></span>|<span data-ttu-id="1d898-688">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="1d898-688">&lt;optional&gt;</span></span>|<span data-ttu-id="1d898-689">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="1d898-689">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="1d898-690">Object</span><span class="sxs-lookup"><span data-stu-id="1d898-690">Object</span></span>|<span data-ttu-id="1d898-691">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="1d898-691">&lt;optional&gt;</span></span>|<span data-ttu-id="1d898-692">В методе обратного вызова разработчики могут указать любой объект, к которому необходимо получить доступ.</span><span class="sxs-lookup"><span data-stu-id="1d898-692">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="1d898-693">Boolean</span><span class="sxs-lookup"><span data-stu-id="1d898-693">Boolean</span></span>|<span data-ttu-id="1d898-694">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="1d898-694">&lt;optional&gt;</span></span>|<span data-ttu-id="1d898-695">Значение `true` указывает на то, что вложение будет встроено в текст сообщения и не должно отображаться в списке вложений.</span><span class="sxs-lookup"><span data-stu-id="1d898-695">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="1d898-696">function</span><span class="sxs-lookup"><span data-stu-id="1d898-696">function</span></span>|<span data-ttu-id="1d898-697">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="1d898-697">&lt;optional&gt;</span></span>|<span data-ttu-id="1d898-698">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="1d898-698">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="1d898-699">После успешного выполнения идентификатор вложения будет представлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="1d898-699">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="1d898-700">Если добавить вложение не удастся, объект `asyncResult` будет содержать объект `Error` с описанием ошибки.</span><span class="sxs-lookup"><span data-stu-id="1d898-700">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="1d898-701">Ошибки</span><span class="sxs-lookup"><span data-stu-id="1d898-701">Errors</span></span>

|<span data-ttu-id="1d898-702">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="1d898-702">Error code</span></span>|<span data-ttu-id="1d898-703">Описание</span><span class="sxs-lookup"><span data-stu-id="1d898-703">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="1d898-704">Вложение превышает максимальный размер.</span><span class="sxs-lookup"><span data-stu-id="1d898-704">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="1d898-705">Расширение вложения не поддерживается.</span><span class="sxs-lookup"><span data-stu-id="1d898-705">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="1d898-706">Сообщение или встреча содержат слишком много вложений.</span><span class="sxs-lookup"><span data-stu-id="1d898-706">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1d898-707">Требования</span><span class="sxs-lookup"><span data-stu-id="1d898-707">Requirements</span></span>

|<span data-ttu-id="1d898-708">Требование</span><span class="sxs-lookup"><span data-stu-id="1d898-708">Requirement</span></span>|<span data-ttu-id="1d898-709">Значение</span><span class="sxs-lookup"><span data-stu-id="1d898-709">Value</span></span>|
|---|---|
|[<span data-ttu-id="1d898-710">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="1d898-710">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1d898-711">1.1</span><span class="sxs-lookup"><span data-stu-id="1d898-711">1.1</span></span>|
|[<span data-ttu-id="1d898-712">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="1d898-712">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1d898-713">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="1d898-713">ReadWriteItem</span></span>|
|[<span data-ttu-id="1d898-714">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="1d898-714">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="1d898-715">Создание</span><span class="sxs-lookup"><span data-stu-id="1d898-715">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="1d898-716">Примеры</span><span class="sxs-lookup"><span data-stu-id="1d898-716">Examples</span></span>

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

<span data-ttu-id="1d898-717">В примере ниже файл изображения добавляется в качестве встроенного вложения, а ссылка на вложение добавляется в текст сообщения.</span><span class="sxs-lookup"><span data-stu-id="1d898-717">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

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

####  <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="1d898-718">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="1d898-718">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="1d898-719">Добавляет обработчик для поддерживаемого события.</span><span class="sxs-lookup"><span data-stu-id="1d898-719">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="1d898-720">Сейчас поддерживаются следующие типы событий: `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged` и `Office.EventType.RecurrenceChanged`.</span><span class="sxs-lookup"><span data-stu-id="1d898-720">Currently the supported event types are `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`</span></span>

##### <a name="parameters"></a><span data-ttu-id="1d898-721">Параметры</span><span class="sxs-lookup"><span data-stu-id="1d898-721">Parameters:</span></span>

| <span data-ttu-id="1d898-722">Имя</span><span class="sxs-lookup"><span data-stu-id="1d898-722">Name</span></span> | <span data-ttu-id="1d898-723">Тип</span><span class="sxs-lookup"><span data-stu-id="1d898-723">Type</span></span> | <span data-ttu-id="1d898-724">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="1d898-724">Attributes</span></span> | <span data-ttu-id="1d898-725">Описание</span><span class="sxs-lookup"><span data-stu-id="1d898-725">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="1d898-726">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="1d898-726">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="1d898-727">Событие, которое должно вызвать обработчик.</span><span class="sxs-lookup"><span data-stu-id="1d898-727">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="1d898-728">Function</span><span class="sxs-lookup"><span data-stu-id="1d898-728">Function</span></span> || <span data-ttu-id="1d898-p136">Функция для обработки события. Функция должна принимать один параметр, представляющий собой объектный литерал. Значение свойства `type` параметра совпадет со значением параметра `eventType`, переданного методу `addHandlerAsync`.</span><span class="sxs-lookup"><span data-stu-id="1d898-p136">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="1d898-732">Объект</span><span class="sxs-lookup"><span data-stu-id="1d898-732">Object</span></span> | <span data-ttu-id="1d898-733">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="1d898-733">&lt;optional&gt;</span></span> | <span data-ttu-id="1d898-734">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="1d898-734">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="1d898-735">Object</span><span class="sxs-lookup"><span data-stu-id="1d898-735">Object</span></span> | <span data-ttu-id="1d898-736">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="1d898-736">&lt;optional&gt;</span></span> | <span data-ttu-id="1d898-737">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="1d898-737">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="1d898-738">функция</span><span class="sxs-lookup"><span data-stu-id="1d898-738">function</span></span>| <span data-ttu-id="1d898-739">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="1d898-739">&lt;optional&gt;</span></span>|<span data-ttu-id="1d898-740">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="1d898-740">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1d898-741">Требования</span><span class="sxs-lookup"><span data-stu-id="1d898-741">Requirements</span></span>

|<span data-ttu-id="1d898-742">Требование</span><span class="sxs-lookup"><span data-stu-id="1d898-742">Requirement</span></span>| <span data-ttu-id="1d898-743">Значение</span><span class="sxs-lookup"><span data-stu-id="1d898-743">Value</span></span>|
|---|---|
|[<span data-ttu-id="1d898-744">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="1d898-744">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1d898-745">1.7</span><span class="sxs-lookup"><span data-stu-id="1d898-745">1.7</span></span> |
|[<span data-ttu-id="1d898-746">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="1d898-746">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1d898-747">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1d898-747">ReadItem</span></span> |
|[<span data-ttu-id="1d898-748">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="1d898-748">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1d898-749">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="1d898-749">Compose or read</span></span> |

##### <a name="example"></a><span data-ttu-id="1d898-750">Пример</span><span class="sxs-lookup"><span data-stu-id="1d898-750">Example</span></span>

```js
Office.initialize = function (reason) {
  $(document).ready(function () {
    Office.context.mailbox.item.addHandlerAsync(Office.EventType.RecurrenceChanged, loadNewItem, function (result) {
      if (result.status === Office.AsyncResultStatus.Failed) {
        // Handle error
      }
    });
  });
};

function loadNewItem(eventArgs) {
  // Load the properties of the newly selected item
  loadProps(Office.context.mailbox.item);
};
```

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="1d898-751">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="1d898-751">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="1d898-752">Добавляет к сообщению элемент Exchange, например сообщение, в виде вложения.</span><span class="sxs-lookup"><span data-stu-id="1d898-752">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="1d898-p137">С помощью метода `addItemAttachmentAsync` можно в элемент формы создания вложить элемент с указанным идентификатором Exchange. Если указать метод обратного вызова, то этот метод вызывается с помощью параметра `asyncResult`, который содержит идентификатор вложения или код, указывающий на ошибки, которые произошли при вложении элемента. При необходимости можно использовать параметр `options` для передачи сведений о состоянии методу обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="1d898-p137">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="1d898-756">Идентификатор можно использовать с методом [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback), чтобы удалить вложение, добавленное во время текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="1d898-756">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="1d898-757">Если ваша надстройка Office выполняется в Outlook Web App, метод `addItemAttachmentAsync` обеспечивает вложение элементов в элементы, отличные от редактируемого. Однако это действие не рекомендуем выполнять, так как оно не поддерживается.</span><span class="sxs-lookup"><span data-stu-id="1d898-757">If your Office Add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1d898-758">Параметры:</span><span class="sxs-lookup"><span data-stu-id="1d898-758">Parameters:</span></span>

|<span data-ttu-id="1d898-759">Имя</span><span class="sxs-lookup"><span data-stu-id="1d898-759">Name</span></span>|<span data-ttu-id="1d898-760">Тип</span><span class="sxs-lookup"><span data-stu-id="1d898-760">Type</span></span>|<span data-ttu-id="1d898-761">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="1d898-761">Attributes</span></span>|<span data-ttu-id="1d898-762">Описание</span><span class="sxs-lookup"><span data-stu-id="1d898-762">Description</span></span>|
|---|---|---|---|
|`itemId`|<span data-ttu-id="1d898-763">String</span><span class="sxs-lookup"><span data-stu-id="1d898-763">String</span></span>||<span data-ttu-id="1d898-p138">Идентификатор Exchange для вкладываемого элемента. Максимальная длина — 100 символов.</span><span class="sxs-lookup"><span data-stu-id="1d898-p138">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="1d898-766">String</span><span class="sxs-lookup"><span data-stu-id="1d898-766">String</span></span>||<span data-ttu-id="1d898-p139">Тема вкладываемого элемента. Максимальная длина — 255 символов.</span><span class="sxs-lookup"><span data-stu-id="1d898-p139">The sujbect of the item to be attached. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="1d898-769">Object</span><span class="sxs-lookup"><span data-stu-id="1d898-769">Object</span></span>|<span data-ttu-id="1d898-770">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="1d898-770">&lt;optional&gt;</span></span>|<span data-ttu-id="1d898-771">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="1d898-771">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="1d898-772">Object</span><span class="sxs-lookup"><span data-stu-id="1d898-772">Object</span></span>|<span data-ttu-id="1d898-773">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="1d898-773">&lt;optional&gt;</span></span>|<span data-ttu-id="1d898-774">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="1d898-774">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="1d898-775">функция</span><span class="sxs-lookup"><span data-stu-id="1d898-775">function</span></span>|<span data-ttu-id="1d898-776">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="1d898-776">&lt;optional&gt;</span></span>|<span data-ttu-id="1d898-777">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="1d898-777">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="1d898-778">После успешного выполнения идентификатор вложения будет представлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="1d898-778">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="1d898-779">Если добавить вложение не удастся, объект `asyncResult` будет содержать объект `Error` с описанием ошибки.</span><span class="sxs-lookup"><span data-stu-id="1d898-779">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="1d898-780">Ошибки</span><span class="sxs-lookup"><span data-stu-id="1d898-780">Errors</span></span>

|<span data-ttu-id="1d898-781">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="1d898-781">Error code</span></span>|<span data-ttu-id="1d898-782">Описание</span><span class="sxs-lookup"><span data-stu-id="1d898-782">Description</span></span>|
|------------|-------------|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="1d898-783">Сообщение или встреча содержат слишком много вложений.</span><span class="sxs-lookup"><span data-stu-id="1d898-783">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1d898-784">Требования</span><span class="sxs-lookup"><span data-stu-id="1d898-784">Requirements</span></span>

|<span data-ttu-id="1d898-785">Требование</span><span class="sxs-lookup"><span data-stu-id="1d898-785">Requirement</span></span>|<span data-ttu-id="1d898-786">Значение</span><span class="sxs-lookup"><span data-stu-id="1d898-786">Value</span></span>|
|---|---|
|[<span data-ttu-id="1d898-787">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="1d898-787">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1d898-788">1.1</span><span class="sxs-lookup"><span data-stu-id="1d898-788">1.1</span></span>|
|[<span data-ttu-id="1d898-789">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="1d898-789">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1d898-790">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="1d898-790">ReadWriteItem</span></span>|
|[<span data-ttu-id="1d898-791">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="1d898-791">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="1d898-792">Создание</span><span class="sxs-lookup"><span data-stu-id="1d898-792">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="1d898-793">Пример</span><span class="sxs-lookup"><span data-stu-id="1d898-793">Example</span></span>

<span data-ttu-id="1d898-794">В следующем примере существующий элемент Outlook добавляется в виде вложения с именем `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="1d898-794">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

```js
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

####  <a name="close"></a><span data-ttu-id="1d898-795">close()</span><span class="sxs-lookup"><span data-stu-id="1d898-795">close()</span></span>

<span data-ttu-id="1d898-796">Закрывает текущий создаваемый элемент.</span><span class="sxs-lookup"><span data-stu-id="1d898-796">Closes the current item that is being composed.</span></span>

<span data-ttu-id="1d898-p140">Работа метода `close` зависит от текущего состояния создаваемого элемента. Если элемент содержит несохраненные изменения, клиент предложит пользователю сохранить или отклонить их либо отменить действие закрытия.</span><span class="sxs-lookup"><span data-stu-id="1d898-p140">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="1d898-799">Если в Outlook в Интернете элемент представляет собой встречу, ранее сохраненную с помощью метода `saveAsync`, пользователю предлагается сохранить, отклонить или отменить действие, даже если с момента последнего сохранения элемента не вносились какие-либо изменения.</span><span class="sxs-lookup"><span data-stu-id="1d898-799">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="1d898-800">Если в клиенте Outlook для настольных ПК сообщение представляет собой ответ в тексте, метод `close` не работает.</span><span class="sxs-lookup"><span data-stu-id="1d898-800">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="1d898-801">Требования</span><span class="sxs-lookup"><span data-stu-id="1d898-801">Requirements</span></span>

|<span data-ttu-id="1d898-802">Требование</span><span class="sxs-lookup"><span data-stu-id="1d898-802">Requirement</span></span>|<span data-ttu-id="1d898-803">Значение</span><span class="sxs-lookup"><span data-stu-id="1d898-803">Value</span></span>|
|---|---|
|[<span data-ttu-id="1d898-804">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="1d898-804">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1d898-805">1.3</span><span class="sxs-lookup"><span data-stu-id="1d898-805">1.3</span></span>|
|[<span data-ttu-id="1d898-806">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="1d898-806">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1d898-807">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="1d898-807">Restricted</span></span>|
|[<span data-ttu-id="1d898-808">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="1d898-808">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="1d898-809">Создание</span><span class="sxs-lookup"><span data-stu-id="1d898-809">Compose</span></span>|

#### <a name="displayreplyallformformdata"></a><span data-ttu-id="1d898-810">displayReplyAllForm(formData)</span><span class="sxs-lookup"><span data-stu-id="1d898-810">displayReplyAllForm(formData)</span></span>

<span data-ttu-id="1d898-811">Отображает форму ответа, включающую отправителя и всех получателей выбранного сообщения или организатора и всех участников выбранной встречи.</span><span class="sxs-lookup"><span data-stu-id="1d898-811">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="1d898-812">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="1d898-812">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="1d898-813">В Outlook Web App форма ответа отображается в виде всплывающей формы в представлении с 3 либо 1 или 2 колонками.</span><span class="sxs-lookup"><span data-stu-id="1d898-813">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="1d898-814">Если любой строковый параметр превышает указанные для него ограничения, `displayReplyAllForm` возвращает исключение.</span><span class="sxs-lookup"><span data-stu-id="1d898-814">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="1d898-p141">Если в параметре `formData.attachments` указаны вложения, Outlook и Outlook Web App пытаются скачать их и вложить в форму ответа. Если какие-либо вложения добавить не удается, в форме отображается сообщение об ошибке. Если сообщения об ошибках не предусмотрены, то они не отображаются.</span><span class="sxs-lookup"><span data-stu-id="1d898-p141">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1d898-818">Параметры</span><span class="sxs-lookup"><span data-stu-id="1d898-818">Parameters:</span></span>

|<span data-ttu-id="1d898-819">Имя</span><span class="sxs-lookup"><span data-stu-id="1d898-819">Name</span></span>|<span data-ttu-id="1d898-820">Тип</span><span class="sxs-lookup"><span data-stu-id="1d898-820">Type</span></span>|<span data-ttu-id="1d898-821">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="1d898-821">Attributes</span></span>|<span data-ttu-id="1d898-822">Описание</span><span class="sxs-lookup"><span data-stu-id="1d898-822">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="1d898-823">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="1d898-823">String &#124; Object</span></span>||<span data-ttu-id="1d898-p142">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="1d898-p142">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="1d898-826">**ИЛИ**</span><span class="sxs-lookup"><span data-stu-id="1d898-826">**OR**</span></span><br/><span data-ttu-id="1d898-p143">Объект, который содержит текст или данные вложения и функцию обратного вызова. Ниже представлено определение этого объекта.</span><span class="sxs-lookup"><span data-stu-id="1d898-p143">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="1d898-829">String</span><span class="sxs-lookup"><span data-stu-id="1d898-829">String</span></span>|<span data-ttu-id="1d898-830">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="1d898-830">&lt;optional&gt;</span></span>|<span data-ttu-id="1d898-p144">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="1d898-p144">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="1d898-833">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="1d898-833">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="1d898-834">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="1d898-834">&lt;optional&gt;</span></span>|<span data-ttu-id="1d898-835">Массив объектов JSON, представляющих собой вложенные файлы или элементы.</span><span class="sxs-lookup"><span data-stu-id="1d898-835">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="1d898-836">String</span><span class="sxs-lookup"><span data-stu-id="1d898-836">String</span></span>||<span data-ttu-id="1d898-p145">Указывает тип вложения. Допустимые значения: `file` для вложенного файла и `item` для вложенного элемента.</span><span class="sxs-lookup"><span data-stu-id="1d898-p145">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="1d898-839">String</span><span class="sxs-lookup"><span data-stu-id="1d898-839">String</span></span>||<span data-ttu-id="1d898-840">Строка, содержащая имя вложения, длиной до 255 символов.</span><span class="sxs-lookup"><span data-stu-id="1d898-840">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="1d898-841">String</span><span class="sxs-lookup"><span data-stu-id="1d898-841">String</span></span>||<span data-ttu-id="1d898-p146">Используется, только если свойству `type` задано значение `file`. URI расположения файла.</span><span class="sxs-lookup"><span data-stu-id="1d898-p146">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="1d898-844">Boolean</span><span class="sxs-lookup"><span data-stu-id="1d898-844">Boolean</span></span>||<span data-ttu-id="1d898-p147">Используется, только если свойству `type` задано значение `file`. Значение `true` указывает на то, что вложение будет встроено в текст сообщения и не должно отображаться в списке вложений.</span><span class="sxs-lookup"><span data-stu-id="1d898-p147">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="1d898-847">String</span><span class="sxs-lookup"><span data-stu-id="1d898-847">String</span></span>||<span data-ttu-id="1d898-p148">Используется, только если свойству `type` задано значение `item`. Идентификатор вложения EWS. Это строка длиной до 100 символов.</span><span class="sxs-lookup"><span data-stu-id="1d898-p148">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="1d898-851">function</span><span class="sxs-lookup"><span data-stu-id="1d898-851">function</span></span>|<span data-ttu-id="1d898-852">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="1d898-852">&lt;optional&gt;</span></span>|<span data-ttu-id="1d898-853">По завершении работы метода функция, переданная параметру `callback`, вызывается с помощью одного параметра `asyncResult`, представляющего собой объект [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="1d898-853">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1d898-854">Требования</span><span class="sxs-lookup"><span data-stu-id="1d898-854">Requirements</span></span>

|<span data-ttu-id="1d898-855">Требование</span><span class="sxs-lookup"><span data-stu-id="1d898-855">Requirement</span></span>|<span data-ttu-id="1d898-856">Значение</span><span class="sxs-lookup"><span data-stu-id="1d898-856">Value</span></span>|
|---|---|
|[<span data-ttu-id="1d898-857">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="1d898-857">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1d898-858">1.0</span><span class="sxs-lookup"><span data-stu-id="1d898-858">1.0</span></span>|
|[<span data-ttu-id="1d898-859">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="1d898-859">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1d898-860">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1d898-860">ReadItem</span></span>|
|[<span data-ttu-id="1d898-861">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="1d898-861">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="1d898-862">Чтение</span><span class="sxs-lookup"><span data-stu-id="1d898-862">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="1d898-863">Примеры</span><span class="sxs-lookup"><span data-stu-id="1d898-863">Examples</span></span>

<span data-ttu-id="1d898-864">Приведенный ниже код передает строку в функцию `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="1d898-864">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="1d898-865">Ответ с пустым текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="1d898-865">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="1d898-866">Ответ только с текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="1d898-866">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="1d898-867">Ответ с текстом сообщения и вложенным файлом.</span><span class="sxs-lookup"><span data-stu-id="1d898-867">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="1d898-868">Ответ с текстом сообщения и вложенным элементом.</span><span class="sxs-lookup"><span data-stu-id="1d898-868">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="1d898-869">Ответ с текстом сообщения, вложенным файлом, вложенным элементом и обратным вызовом.</span><span class="sxs-lookup"><span data-stu-id="1d898-869">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata"></a><span data-ttu-id="1d898-870">displayReplyForm(formData)</span><span class="sxs-lookup"><span data-stu-id="1d898-870">displayReplyForm(formData)</span></span>

<span data-ttu-id="1d898-871">Отображает форму ответа, включающую только отправителя выбранного сообщения или организатора выбранной встречи.</span><span class="sxs-lookup"><span data-stu-id="1d898-871">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="1d898-872">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="1d898-872">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="1d898-873">В Outlook Web App форма ответа отображается в виде всплывающей формы в представлении с 3 либо 1 или 2 колонками.</span><span class="sxs-lookup"><span data-stu-id="1d898-873">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="1d898-874">Если любой строковый параметр превышает указанные для него ограничения, `displayReplyForm` возвращает исключение.</span><span class="sxs-lookup"><span data-stu-id="1d898-874">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="1d898-p149">Если в параметре `formData.attachments` указаны вложения, Outlook и Outlook Web App пытаются скачать их и вложить в форму ответа. Если какие-либо вложения добавить не удается, в форме отображается сообщение об ошибке. Если сообщения об ошибках не предусмотрены, то они не отображаются.</span><span class="sxs-lookup"><span data-stu-id="1d898-p149">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1d898-878">Параметры</span><span class="sxs-lookup"><span data-stu-id="1d898-878">Parameters:</span></span>

|<span data-ttu-id="1d898-879">Имя</span><span class="sxs-lookup"><span data-stu-id="1d898-879">Name</span></span>|<span data-ttu-id="1d898-880">Тип</span><span class="sxs-lookup"><span data-stu-id="1d898-880">Type</span></span>|<span data-ttu-id="1d898-881">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="1d898-881">Attributes</span></span>|<span data-ttu-id="1d898-882">Описание</span><span class="sxs-lookup"><span data-stu-id="1d898-882">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="1d898-883">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="1d898-883">String &#124; Object</span></span>||<span data-ttu-id="1d898-p150">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="1d898-p150">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="1d898-886">**ИЛИ**</span><span class="sxs-lookup"><span data-stu-id="1d898-886">**OR**</span></span><br/><span data-ttu-id="1d898-p151">Объект, который содержит текст или данные вложения и функцию обратного вызова. Ниже представлено определение этого объекта.</span><span class="sxs-lookup"><span data-stu-id="1d898-p151">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="1d898-889">String</span><span class="sxs-lookup"><span data-stu-id="1d898-889">String</span></span>|<span data-ttu-id="1d898-890">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="1d898-890">&lt;optional&gt;</span></span>|<span data-ttu-id="1d898-p152">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="1d898-p152">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="1d898-893">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="1d898-893">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="1d898-894">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="1d898-894">&lt;optional&gt;</span></span>|<span data-ttu-id="1d898-895">Массив объектов JSON, представляющих собой вложенные файлы или элементы.</span><span class="sxs-lookup"><span data-stu-id="1d898-895">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="1d898-896">String</span><span class="sxs-lookup"><span data-stu-id="1d898-896">String</span></span>||<span data-ttu-id="1d898-p153">Указывает тип вложения. Допустимые значения: `file` для вложенного файла и `item` для вложенного элемента.</span><span class="sxs-lookup"><span data-stu-id="1d898-p153">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="1d898-899">String</span><span class="sxs-lookup"><span data-stu-id="1d898-899">String</span></span>||<span data-ttu-id="1d898-900">Строка, содержащая имя вложения, длиной до 255 символов.</span><span class="sxs-lookup"><span data-stu-id="1d898-900">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="1d898-901">String</span><span class="sxs-lookup"><span data-stu-id="1d898-901">String</span></span>||<span data-ttu-id="1d898-p154">Используется, только если свойству `type` задано значение `file`. URI расположения файла.</span><span class="sxs-lookup"><span data-stu-id="1d898-p154">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="1d898-904">Boolean</span><span class="sxs-lookup"><span data-stu-id="1d898-904">Boolean</span></span>||<span data-ttu-id="1d898-p155">Используется, только если свойству `type` задано значение `file`. Значение `true` указывает на то, что вложение будет встроено в текст сообщения и не должно отображаться в списке вложений.</span><span class="sxs-lookup"><span data-stu-id="1d898-p155">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="1d898-907">String</span><span class="sxs-lookup"><span data-stu-id="1d898-907">String</span></span>||<span data-ttu-id="1d898-p156">Используется, только если свойству `type` задано значение `item`. Идентификатор вложения EWS. Это строка длиной до 100 символов.</span><span class="sxs-lookup"><span data-stu-id="1d898-p156">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="1d898-911">function</span><span class="sxs-lookup"><span data-stu-id="1d898-911">function</span></span>|<span data-ttu-id="1d898-912">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="1d898-912">&lt;optional&gt;</span></span>|<span data-ttu-id="1d898-913">По завершении работы метода функция, переданная параметру `callback`, вызывается с помощью одного параметра `asyncResult`, представляющего собой объект [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="1d898-913">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1d898-914">Требования</span><span class="sxs-lookup"><span data-stu-id="1d898-914">Requirements</span></span>

|<span data-ttu-id="1d898-915">Требование</span><span class="sxs-lookup"><span data-stu-id="1d898-915">Requirement</span></span>|<span data-ttu-id="1d898-916">Значение</span><span class="sxs-lookup"><span data-stu-id="1d898-916">Value</span></span>|
|---|---|
|[<span data-ttu-id="1d898-917">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="1d898-917">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1d898-918">1.0</span><span class="sxs-lookup"><span data-stu-id="1d898-918">1.0</span></span>|
|[<span data-ttu-id="1d898-919">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="1d898-919">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1d898-920">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1d898-920">ReadItem</span></span>|
|[<span data-ttu-id="1d898-921">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="1d898-921">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="1d898-922">Чтение</span><span class="sxs-lookup"><span data-stu-id="1d898-922">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="1d898-923">Примеры</span><span class="sxs-lookup"><span data-stu-id="1d898-923">Examples</span></span>

<span data-ttu-id="1d898-924">Приведенный ниже код передает строку в функцию `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="1d898-924">The following code passes a string to the `displayReplyForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="1d898-925">Ответ с пустым текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="1d898-925">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="1d898-926">Ответ только с текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="1d898-926">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="1d898-927">Ответ с текстом сообщения и вложенным файлом.</span><span class="sxs-lookup"><span data-stu-id="1d898-927">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="1d898-928">Ответ с текстом сообщения и вложенным элементом.</span><span class="sxs-lookup"><span data-stu-id="1d898-928">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="1d898-929">Ответ с текстом сообщения, вложенным файлом, вложенным элементом и обратным вызовом.</span><span class="sxs-lookup"><span data-stu-id="1d898-929">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlook17officeentities"></a><span data-ttu-id="1d898-930">getEntities() → {[Entities](/javascript/api/outlook_1_7/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="1d898-930">getEntities() → {[Entities](/javascript/api/outlook_1_7/office.entities)}</span></span>

<span data-ttu-id="1d898-931">Получает сущности, обнаруженные в теле выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="1d898-931">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="1d898-932">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="1d898-932">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="1d898-933">Требования</span><span class="sxs-lookup"><span data-stu-id="1d898-933">Requirements</span></span>

|<span data-ttu-id="1d898-934">Требование</span><span class="sxs-lookup"><span data-stu-id="1d898-934">Requirement</span></span>|<span data-ttu-id="1d898-935">Значение</span><span class="sxs-lookup"><span data-stu-id="1d898-935">Value</span></span>|
|---|---|
|[<span data-ttu-id="1d898-936">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="1d898-936">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1d898-937">1.0</span><span class="sxs-lookup"><span data-stu-id="1d898-937">1.0</span></span>|
|[<span data-ttu-id="1d898-938">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="1d898-938">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1d898-939">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1d898-939">ReadItem</span></span>|
|[<span data-ttu-id="1d898-940">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="1d898-940">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="1d898-941">Чтение</span><span class="sxs-lookup"><span data-stu-id="1d898-941">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="1d898-942">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="1d898-942">Returns:</span></span>

<span data-ttu-id="1d898-943">Тип: [Entities](/javascript/api/outlook_1_7/office.entities)</span><span class="sxs-lookup"><span data-stu-id="1d898-943">Type: [Entities](/javascript/api/outlook_1_7/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="1d898-944">Пример</span><span class="sxs-lookup"><span data-stu-id="1d898-944">Example</span></span>

<span data-ttu-id="1d898-945">Ниже приведен пример получения доступа к сущностям контактов в теле текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="1d898-945">The following example accesses the contacts entities in the current item's body.</span></span>

```js
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook17officecontactmeetingsuggestionjavascriptapioutlook17officemeetingsuggestionphonenumberjavascriptapioutlook17officephonenumbertasksuggestionjavascriptapioutlook17officetasksuggestion"></a><span data-ttu-id="1d898-946">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="1d898-946">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))>}</span></span>

<span data-ttu-id="1d898-947">Получает массив всех сущностей указанного типа, обнаруженных в теле выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="1d898-947">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="1d898-948">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="1d898-948">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1d898-949">Параметры</span><span class="sxs-lookup"><span data-stu-id="1d898-949">Parameters:</span></span>

|<span data-ttu-id="1d898-950">Имя</span><span class="sxs-lookup"><span data-stu-id="1d898-950">Name</span></span>|<span data-ttu-id="1d898-951">Тип</span><span class="sxs-lookup"><span data-stu-id="1d898-951">Type</span></span>|<span data-ttu-id="1d898-952">Описание</span><span class="sxs-lookup"><span data-stu-id="1d898-952">Description</span></span>|
|---|---|---|
|`entityType`|[<span data-ttu-id="1d898-953">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="1d898-953">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook_1_7/office.mailboxenums.entitytype)|<span data-ttu-id="1d898-954">Одно из значений перечисления EntityType.</span><span class="sxs-lookup"><span data-stu-id="1d898-954">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1d898-955">Требования</span><span class="sxs-lookup"><span data-stu-id="1d898-955">Requirements</span></span>

|<span data-ttu-id="1d898-956">Требование</span><span class="sxs-lookup"><span data-stu-id="1d898-956">Requirement</span></span>|<span data-ttu-id="1d898-957">Значение</span><span class="sxs-lookup"><span data-stu-id="1d898-957">Value</span></span>|
|---|---|
|[<span data-ttu-id="1d898-958">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="1d898-958">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1d898-959">1.0</span><span class="sxs-lookup"><span data-stu-id="1d898-959">1.0</span></span>|
|[<span data-ttu-id="1d898-960">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="1d898-960">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1d898-961">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="1d898-961">Restricted</span></span>|
|[<span data-ttu-id="1d898-962">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="1d898-962">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="1d898-963">Чтение</span><span class="sxs-lookup"><span data-stu-id="1d898-963">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="1d898-964">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="1d898-964">Returns:</span></span>

<span data-ttu-id="1d898-965">Если значение, переданное в `entityType`, не является допустимым членом перечисления `EntityType`, метод возвращает значение NULL.</span><span class="sxs-lookup"><span data-stu-id="1d898-965">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="1d898-966">Если в теле элемента отсутствуют сущности указанного типа, метод возвращает пустой массив.</span><span class="sxs-lookup"><span data-stu-id="1d898-966">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="1d898-967">В противном случае тип объектов в возвращаемом массиве зависит от типа сущности, запрошенной в параметре `entityType`.</span><span class="sxs-lookup"><span data-stu-id="1d898-967">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="1d898-968">Хотя минимальный уровень разрешений для использования этого метода — **Restricted**, для некоторых типов сущностей требуется доступ на уровне **ReadItem**, как указано в приведенной ниже таблице.</span><span class="sxs-lookup"><span data-stu-id="1d898-968">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

|<span data-ttu-id="1d898-969">Значение параметра `entityType`</span><span class="sxs-lookup"><span data-stu-id="1d898-969">Value of `entityType`</span></span>|<span data-ttu-id="1d898-970">Тип объектов в возвращаемом массиве</span><span class="sxs-lookup"><span data-stu-id="1d898-970">Type of objects in returned array</span></span>|<span data-ttu-id="1d898-971">Необходимый уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="1d898-971">Required Permission Level</span></span>|
|---|---|---|
|`Address`|<span data-ttu-id="1d898-972">String</span><span class="sxs-lookup"><span data-stu-id="1d898-972">String</span></span>|<span data-ttu-id="1d898-973">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="1d898-973">**Restricted**</span></span>|
|`Contact`|<span data-ttu-id="1d898-974">Contact</span><span class="sxs-lookup"><span data-stu-id="1d898-974">Contact</span></span>|<span data-ttu-id="1d898-975">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="1d898-975">**ReadItem**</span></span>|
|`EmailAddress`|<span data-ttu-id="1d898-976">String</span><span class="sxs-lookup"><span data-stu-id="1d898-976">String</span></span>|<span data-ttu-id="1d898-977">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="1d898-977">**ReadItem**</span></span>|
|`MeetingSuggestion`|<span data-ttu-id="1d898-978">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="1d898-978">MeetingSuggestion</span></span>|<span data-ttu-id="1d898-979">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="1d898-979">**ReadItem**</span></span>|
|`PhoneNumber`|<span data-ttu-id="1d898-980">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="1d898-980">PhoneNumber</span></span>|<span data-ttu-id="1d898-981">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="1d898-981">**Restricted**</span></span>|
|`TaskSuggestion`|<span data-ttu-id="1d898-982">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="1d898-982">TaskSuggestion</span></span>|<span data-ttu-id="1d898-983">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="1d898-983">**ReadItem**</span></span>|
|`URL`|<span data-ttu-id="1d898-984">String</span><span class="sxs-lookup"><span data-stu-id="1d898-984">String</span></span>|<span data-ttu-id="1d898-985">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="1d898-985">**Restricted**</span></span>|

<span data-ttu-id="1d898-986">Тип: Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="1d898-986">Type: Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="1d898-987">Пример</span><span class="sxs-lookup"><span data-stu-id="1d898-987">Example</span></span>

<span data-ttu-id="1d898-988">В примере ниже показано, как получить доступ к массиву строк, которые представляют собой почтовые адреса в теле текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="1d898-988">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook17officecontactmeetingsuggestionjavascriptapioutlook17officemeetingsuggestionphonenumberjavascriptapioutlook17officephonenumbertasksuggestionjavascriptapioutlook17officetasksuggestion"></a><span data-ttu-id="1d898-989">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="1d898-989">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))>}</span></span>

<span data-ttu-id="1d898-990">Возвращает известные сущности в выбранном элементе, которые проходят через именованный фильтр, определяемый в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="1d898-990">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="1d898-991">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="1d898-991">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="1d898-992">Метод `getFilteredEntitiesByName` возвращает сущности, соответствующие регулярному выражению, которое определяется в элементе правила [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) в XML-файле манифеста, с использованием указанного значения элемента `FilterName`.</span><span class="sxs-lookup"><span data-stu-id="1d898-992">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1d898-993">Параметры</span><span class="sxs-lookup"><span data-stu-id="1d898-993">Parameters:</span></span>

|<span data-ttu-id="1d898-994">Имя</span><span class="sxs-lookup"><span data-stu-id="1d898-994">Name</span></span>|<span data-ttu-id="1d898-995">Тип</span><span class="sxs-lookup"><span data-stu-id="1d898-995">Type</span></span>|<span data-ttu-id="1d898-996">Описание</span><span class="sxs-lookup"><span data-stu-id="1d898-996">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="1d898-997">String</span><span class="sxs-lookup"><span data-stu-id="1d898-997">String</span></span>|<span data-ttu-id="1d898-998">Имя элемента правила `ItemHasKnownEntity`, определяющее соответствующий фильтр.</span><span class="sxs-lookup"><span data-stu-id="1d898-998">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1d898-999">Требования</span><span class="sxs-lookup"><span data-stu-id="1d898-999">Requirements</span></span>

|<span data-ttu-id="1d898-1000">Требование</span><span class="sxs-lookup"><span data-stu-id="1d898-1000">Requirement</span></span>|<span data-ttu-id="1d898-1001">Значение</span><span class="sxs-lookup"><span data-stu-id="1d898-1001">Value</span></span>|
|---|---|
|[<span data-ttu-id="1d898-1002">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="1d898-1002">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1d898-1003">1.0</span><span class="sxs-lookup"><span data-stu-id="1d898-1003">1.0</span></span>|
|[<span data-ttu-id="1d898-1004">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="1d898-1004">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1d898-1005">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1d898-1005">ReadItem</span></span>|
|[<span data-ttu-id="1d898-1006">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="1d898-1006">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="1d898-1007">Чтение</span><span class="sxs-lookup"><span data-stu-id="1d898-1007">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="1d898-1008">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="1d898-1008">Returns:</span></span>

<span data-ttu-id="1d898-p158">Если в манифесте нет элемента `ItemHasKnownEntity` со значением `FilterName`, соответствующим параметру `name`, метод возвращает `null`. Если параметр `name` соответствует элементу `ItemHasKnownEntity` в манифесте, но при этом в текущем элементе нет соответствующих сущностей, метод возвращает пустой массив.</span><span class="sxs-lookup"><span data-stu-id="1d898-p158">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="1d898-1011">Тип: Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="1d898-1011">Type: Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))></span></span>

#### <a name="getregexmatches--object"></a><span data-ttu-id="1d898-1012">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="1d898-1012">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="1d898-1013">Возвращает строковые значения в выбранном элементе, которые соответствуют регулярным выражениям, определенным в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="1d898-1013">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="1d898-1014">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="1d898-1014">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="1d898-p159">Метод `getRegExMatches` возвращает строки, соответствующие регулярному выражению, которое определяется в каждом элементе правила `ItemHasRegularExpressionMatch` или `ItemHasKnownEntity` в XML-файле манифеста. Для правила `ItemHasRegularExpressionMatch` соответствующую строку должно содержать свойство элемента, указанного этим правилом. Простой тип `PropertyName` определяет поддерживаемые свойства.</span><span class="sxs-lookup"><span data-stu-id="1d898-p159">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="1d898-1018">Например, рассмотрим манифест надстройки, который содержит указанный ниже элемент `Rule`.</span><span class="sxs-lookup"><span data-stu-id="1d898-1018">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="1d898-1019">Объект, возвращаемый методом `getRegExMatches`, будет содержать два свойства: `fruits` и `veggies`.</span><span class="sxs-lookup"><span data-stu-id="1d898-1019">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="1d898-p160">Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты. Лучше используйте метод [`Body.getAsync`](/javascript/api/outlook_1_7/office.body#getasync-coerciontype--options--callback-) для этого.</span><span class="sxs-lookup"><span data-stu-id="1d898-p160">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook_1_7/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="1d898-1023">Requirements</span><span class="sxs-lookup"><span data-stu-id="1d898-1023">Requirements</span></span>

|<span data-ttu-id="1d898-1024">Требование</span><span class="sxs-lookup"><span data-stu-id="1d898-1024">Requirement</span></span>|<span data-ttu-id="1d898-1025">Значение</span><span class="sxs-lookup"><span data-stu-id="1d898-1025">Value</span></span>|
|---|---|
|[<span data-ttu-id="1d898-1026">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="1d898-1026">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1d898-1027">1.0</span><span class="sxs-lookup"><span data-stu-id="1d898-1027">1.0</span></span>|
|[<span data-ttu-id="1d898-1028">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="1d898-1028">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1d898-1029">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1d898-1029">ReadItem</span></span>|
|[<span data-ttu-id="1d898-1030">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="1d898-1030">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="1d898-1031">Чтение</span><span class="sxs-lookup"><span data-stu-id="1d898-1031">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="1d898-1032">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="1d898-1032">Returns:</span></span>

<span data-ttu-id="1d898-p161">Объект, содержащий массив строк, которые соответствуют регулярным выражениям, определяемым в XML-файле манифеста. Имя каждого массива равно соответствующему значению атрибута `RegExName` подходящего правила `ItemHasRegularExpressionMatch` или атрибута `FilterName` соответствующего правила `ItemHasKnownEntity`.</span><span class="sxs-lookup"><span data-stu-id="1d898-p161">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="1d898-1035">

<dt>Тип</dt>

</span><span class="sxs-lookup"><span data-stu-id="1d898-1035">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="1d898-1036">Object</span><span class="sxs-lookup"><span data-stu-id="1d898-1036">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="1d898-1037">Пример</span><span class="sxs-lookup"><span data-stu-id="1d898-1037">Example</span></span>

<span data-ttu-id="1d898-1038">В приведенном ниже примере показано, как получить доступ к массиву совпадений с элементами `fruits` и `veggies` правил активации регулярных выражений, указанными в манифесте.</span><span class="sxs-lookup"><span data-stu-id="1d898-1038">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```js
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veges = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="1d898-1039">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="1d898-1039">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="1d898-1040">Возвращает строковые значения в выбранном элементе, которые соответствуют именованному регулярному выражению, определенному в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="1d898-1040">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="1d898-1041">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="1d898-1041">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="1d898-1042">Метод `getRegExMatchesByName` возвращает строки, соответствующие регулярному выражению, которое определяется в элементе правила `ItemHasRegularExpressionMatch` в XML-файле манифеста, с использованием указанного значения элемента `RegExName`.</span><span class="sxs-lookup"><span data-stu-id="1d898-1042">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="1d898-p162">Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты.</span><span class="sxs-lookup"><span data-stu-id="1d898-p162">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1d898-1045">Параметры</span><span class="sxs-lookup"><span data-stu-id="1d898-1045">Parameters:</span></span>

|<span data-ttu-id="1d898-1046">Имя</span><span class="sxs-lookup"><span data-stu-id="1d898-1046">Name</span></span>|<span data-ttu-id="1d898-1047">Тип</span><span class="sxs-lookup"><span data-stu-id="1d898-1047">Type</span></span>|<span data-ttu-id="1d898-1048">Описание</span><span class="sxs-lookup"><span data-stu-id="1d898-1048">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="1d898-1049">String</span><span class="sxs-lookup"><span data-stu-id="1d898-1049">String</span></span>|<span data-ttu-id="1d898-1050">Имя элемента правила `ItemHasRegularExpressionMatch`, определяющее соответствующий фильтр.</span><span class="sxs-lookup"><span data-stu-id="1d898-1050">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1d898-1051">Требования</span><span class="sxs-lookup"><span data-stu-id="1d898-1051">Requirements</span></span>

|<span data-ttu-id="1d898-1052">Требование</span><span class="sxs-lookup"><span data-stu-id="1d898-1052">Requirement</span></span>|<span data-ttu-id="1d898-1053">Значение</span><span class="sxs-lookup"><span data-stu-id="1d898-1053">Value</span></span>|
|---|---|
|[<span data-ttu-id="1d898-1054">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="1d898-1054">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1d898-1055">1.0</span><span class="sxs-lookup"><span data-stu-id="1d898-1055">1.0</span></span>|
|[<span data-ttu-id="1d898-1056">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="1d898-1056">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1d898-1057">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1d898-1057">ReadItem</span></span>|
|[<span data-ttu-id="1d898-1058">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="1d898-1058">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="1d898-1059">Чтение</span><span class="sxs-lookup"><span data-stu-id="1d898-1059">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="1d898-1060">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="1d898-1060">Returns:</span></span>

<span data-ttu-id="1d898-1061">Массив строк, соответствующих регулярному выражению, определяемому в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="1d898-1061">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="1d898-1062">

<dt>Тип</dt>

</span><span class="sxs-lookup"><span data-stu-id="1d898-1062">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="1d898-1063">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="1d898-1063">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="1d898-1064">Пример</span><span class="sxs-lookup"><span data-stu-id="1d898-1064">Example</span></span>

```js
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="1d898-1065">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="1d898-1065">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="1d898-1066">Асинхронно возвращает данные, выбранные в теме или тексте сообщения.</span><span class="sxs-lookup"><span data-stu-id="1d898-1066">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="1d898-p163">Если выделенный фрагмент отсутствует, но курсор находится в тексте или теме, метод возвращает значение NULL для выбранных данных. Если выбраны не текст и не тема, метод возвращает ошибку `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="1d898-p163">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1d898-1069">Параметры</span><span class="sxs-lookup"><span data-stu-id="1d898-1069">Parameters:</span></span>

|<span data-ttu-id="1d898-1070">Имя</span><span class="sxs-lookup"><span data-stu-id="1d898-1070">Name</span></span>|<span data-ttu-id="1d898-1071">Тип</span><span class="sxs-lookup"><span data-stu-id="1d898-1071">Type</span></span>|<span data-ttu-id="1d898-1072">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="1d898-1072">Attributes</span></span>|<span data-ttu-id="1d898-1073">Описание</span><span class="sxs-lookup"><span data-stu-id="1d898-1073">Description</span></span>|
|---|---|---|---|
|`coercionType`|[<span data-ttu-id="1d898-1074">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="1d898-1074">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="1d898-p164">Запрашивает формат данных. Если задано значение Text, метод возвращает обычный текст как строку, удаляя все имеющиеся HTML-теги. Если задано значение HTML, метод возвращает выделенный текст (обычный текст или HTML).</span><span class="sxs-lookup"><span data-stu-id="1d898-p164">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`|<span data-ttu-id="1d898-1078">Object</span><span class="sxs-lookup"><span data-stu-id="1d898-1078">Object</span></span>|<span data-ttu-id="1d898-1079">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="1d898-1079">&lt;optional&gt;</span></span>|<span data-ttu-id="1d898-1080">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="1d898-1080">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="1d898-1081">Object</span><span class="sxs-lookup"><span data-stu-id="1d898-1081">Object</span></span>|<span data-ttu-id="1d898-1082">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="1d898-1082">&lt;optional&gt;</span></span>|<span data-ttu-id="1d898-1083">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="1d898-1083">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="1d898-1084">функция</span><span class="sxs-lookup"><span data-stu-id="1d898-1084">function</span></span>||<span data-ttu-id="1d898-1085">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="1d898-1085">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="1d898-1086">Чтобы получить доступ к выбранным данным из метода обратного вызова, вызовите `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="1d898-1086">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="1d898-1087">Для доступа к исходному свойству, представляющему собой источник выбранных данных, вызовите параметр `asyncResult.value.sourceProperty`, который может иметь значение `body` или `subject`.</span><span class="sxs-lookup"><span data-stu-id="1d898-1087">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1d898-1088">Требования</span><span class="sxs-lookup"><span data-stu-id="1d898-1088">Requirements</span></span>

|<span data-ttu-id="1d898-1089">Требование</span><span class="sxs-lookup"><span data-stu-id="1d898-1089">Requirement</span></span>|<span data-ttu-id="1d898-1090">Значение</span><span class="sxs-lookup"><span data-stu-id="1d898-1090">Value</span></span>|
|---|---|
|[<span data-ttu-id="1d898-1091">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="1d898-1091">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1d898-1092">1.2</span><span class="sxs-lookup"><span data-stu-id="1d898-1092">1.2</span></span>|
|[<span data-ttu-id="1d898-1093">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="1d898-1093">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1d898-1094">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="1d898-1094">ReadWriteItem</span></span>|
|[<span data-ttu-id="1d898-1095">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="1d898-1095">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="1d898-1096">Создание</span><span class="sxs-lookup"><span data-stu-id="1d898-1096">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="1d898-1097">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="1d898-1097">Returns:</span></span>

<span data-ttu-id="1d898-1098">Выбранные данные в виде строки с форматом, определенным в параметре `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="1d898-1098">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="1d898-1099">

<dt>Тип</dt>

</span><span class="sxs-lookup"><span data-stu-id="1d898-1099">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="1d898-1100">String</span><span class="sxs-lookup"><span data-stu-id="1d898-1100">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="1d898-1101">Пример</span><span class="sxs-lookup"><span data-stu-id="1d898-1101">Example</span></span>

```js
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

#### <a name="getselectedentities--entitiesjavascriptapioutlook17officeentities"></a><span data-ttu-id="1d898-1102">getSelectedEntities() → {[Entities](/javascript/api/outlook_1_7/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="1d898-1102">getSelectedEntities() → {[Entities](/javascript/api/outlook_1_7/office.entities)}</span></span>

<span data-ttu-id="1d898-p166">Возвращает сущности, найденные в выделенном совпадении, выбранном пользователем. Выделенные совпадения применяются к [контекстным надстройкам](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="1d898-p166">Gets the entities found in a highlighted match a user has selected. Highlighted matches apply to [contextual add-ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="1d898-1105">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="1d898-1105">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="1d898-1106">Требования</span><span class="sxs-lookup"><span data-stu-id="1d898-1106">Requirements</span></span>

|<span data-ttu-id="1d898-1107">Требование</span><span class="sxs-lookup"><span data-stu-id="1d898-1107">Requirement</span></span>|<span data-ttu-id="1d898-1108">Значение</span><span class="sxs-lookup"><span data-stu-id="1d898-1108">Value</span></span>|
|---|---|
|[<span data-ttu-id="1d898-1109">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="1d898-1109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1d898-1110">1.6</span><span class="sxs-lookup"><span data-stu-id="1d898-1110">1.6</span></span>|
|[<span data-ttu-id="1d898-1111">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="1d898-1111">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1d898-1112">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1d898-1112">ReadItem</span></span>|
|[<span data-ttu-id="1d898-1113">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="1d898-1113">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="1d898-1114">Чтение</span><span class="sxs-lookup"><span data-stu-id="1d898-1114">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="1d898-1115">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="1d898-1115">Returns:</span></span>

<span data-ttu-id="1d898-1116">Тип: [Entities](/javascript/api/outlook_1_7/office.entities)</span><span class="sxs-lookup"><span data-stu-id="1d898-1116">Type: [Entities](/javascript/api/outlook_1_7/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="1d898-1117">Пример</span><span class="sxs-lookup"><span data-stu-id="1d898-1117">Example</span></span>

<span data-ttu-id="1d898-1118">В приведенном ниже примере показано, как получить доступ к сущностям адресов в выделенном совпадении, выбранном пользователем.</span><span class="sxs-lookup"><span data-stu-id="1d898-1118">The following example accesses the addresses entities in the highlighted match selected by the user.</span></span>

```js
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

#### <a name="getselectedregexmatches--object"></a><span data-ttu-id="1d898-1119">getSelectedRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="1d898-1119">getSelectedRegExMatches() → {Object}</span></span>

<span data-ttu-id="1d898-p167">Возвращает строковые значения в выделенном совпадении, которые соответствуют регулярным выражениям, определенным в XML-файле манифеста. Выделенные совпадения применяются к [контекстным надстройкам](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="1d898-p167">Returns string values in a highlighted match that match the regular expressions defined in the manifest XML file. Highlighted matches apply to [contextual add-ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="1d898-1122">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="1d898-1122">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="1d898-p168">Метод `getSelectedRegExMatches` возвращает строки, соответствующие регулярному выражению, которое определяется в каждом элементе правила `ItemHasRegularExpressionMatch` или `ItemHasKnownEntity` в XML-файле манифеста. Для правила `ItemHasRegularExpressionMatch` соответствующую строку должно содержать свойство элемента, указанного этим правилом. Простой тип `PropertyName` определяет поддерживаемые свойства.</span><span class="sxs-lookup"><span data-stu-id="1d898-p168">The `getSelectedRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="1d898-1126">Например, рассмотрим манифест надстройки, который содержит указанный ниже элемент `Rule`.</span><span class="sxs-lookup"><span data-stu-id="1d898-1126">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="1d898-1127">Объект, возвращаемый методом `getRegExMatches`, будет содержать два свойства: `fruits` и `veggies`.</span><span class="sxs-lookup"><span data-stu-id="1d898-1127">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="1d898-p169">Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты. Лучше используйте метод [`Body.getAsync`](/javascript/api/outlook_1_7/office.body#getasync-coerciontype--options--callback-) для этого.</span><span class="sxs-lookup"><span data-stu-id="1d898-p169">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook_1_7/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="1d898-1131">Requirements</span><span class="sxs-lookup"><span data-stu-id="1d898-1131">Requirements</span></span>

|<span data-ttu-id="1d898-1132">Требование</span><span class="sxs-lookup"><span data-stu-id="1d898-1132">Requirement</span></span>|<span data-ttu-id="1d898-1133">Значение</span><span class="sxs-lookup"><span data-stu-id="1d898-1133">Value</span></span>|
|---|---|
|[<span data-ttu-id="1d898-1134">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="1d898-1134">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1d898-1135">1.6</span><span class="sxs-lookup"><span data-stu-id="1d898-1135">1.6</span></span>|
|[<span data-ttu-id="1d898-1136">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="1d898-1136">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1d898-1137">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1d898-1137">ReadItem</span></span>|
|[<span data-ttu-id="1d898-1138">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="1d898-1138">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="1d898-1139">Чтение</span><span class="sxs-lookup"><span data-stu-id="1d898-1139">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="1d898-1140">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="1d898-1140">Returns:</span></span>

<span data-ttu-id="1d898-p170">Объект, содержащий массив строк, которые соответствуют регулярным выражениям, определяемым в XML-файле манифеста. Имя каждого массива равно соответствующему значению атрибута `RegExName` подходящего правила `ItemHasRegularExpressionMatch` или атрибута `FilterName` соответствующего правила `ItemHasKnownEntity`.</span><span class="sxs-lookup"><span data-stu-id="1d898-p170">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

##### <a name="example"></a><span data-ttu-id="1d898-1143">Пример</span><span class="sxs-lookup"><span data-stu-id="1d898-1143">Example</span></span>

<span data-ttu-id="1d898-1144">В приведенном ниже примере показано, как получить доступ к массиву совпадений с элементами `fruits` и `veggies` правил активации регулярных выражений, указанными в манифесте.</span><span class="sxs-lookup"><span data-stu-id="1d898-1144">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```js
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="1d898-1145">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="1d898-1145">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="1d898-1146">Асинхронно загружает настраиваемые свойства для надстройки для выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="1d898-1146">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="1d898-p171">Настраиваемые свойства сохраняются в виде пар "ключ-значение" для каждого приложения и каждого элемента. Этот метод возвращает объект `CustomProperties` при обратном вызове, который предоставляет методы для доступа к настраиваемым свойствам, характерным для текущего элемента и текущей надстройки. Настраиваемые свойства не шифруются для элемента, поэтому этот способ хранения не является безопасным.</span><span class="sxs-lookup"><span data-stu-id="1d898-p171">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1d898-1150">Параметры</span><span class="sxs-lookup"><span data-stu-id="1d898-1150">Parameters:</span></span>

|<span data-ttu-id="1d898-1151">Имя</span><span class="sxs-lookup"><span data-stu-id="1d898-1151">Name</span></span>|<span data-ttu-id="1d898-1152">Тип</span><span class="sxs-lookup"><span data-stu-id="1d898-1152">Type</span></span>|<span data-ttu-id="1d898-1153">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="1d898-1153">Attributes</span></span>|<span data-ttu-id="1d898-1154">Описание</span><span class="sxs-lookup"><span data-stu-id="1d898-1154">Description</span></span>|
|---|---|---|---|
|`callback`|<span data-ttu-id="1d898-1155">function</span><span class="sxs-lookup"><span data-stu-id="1d898-1155">function</span></span>||<span data-ttu-id="1d898-1156">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="1d898-1156">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="1d898-1157">Настраиваемые свойства предоставляются в виде объекта [`CustomProperties`](/javascript/api/outlook_1_7/office.customproperties) в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="1d898-1157">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook_1_7/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="1d898-1158">Этот объект позволяет получить, задать и удалить настраиваемые свойства для элемента, а также сохранить изменения, внесенные в набор настраиваемых свойств, на сервере.</span><span class="sxs-lookup"><span data-stu-id="1d898-1158">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`|<span data-ttu-id="1d898-1159">Object</span><span class="sxs-lookup"><span data-stu-id="1d898-1159">Object</span></span>|<span data-ttu-id="1d898-1160">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="1d898-1160">&lt;optional&gt;</span></span>|<span data-ttu-id="1d898-1161">Разработчики могут указать любой объект, к которому необходимо получить доступ, в функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="1d898-1161">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="1d898-1162">Доступ к этому объекту можно получить с помощью свойства `asyncResult.asyncContext` в функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="1d898-1162">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1d898-1163">Требования</span><span class="sxs-lookup"><span data-stu-id="1d898-1163">Requirements</span></span>

|<span data-ttu-id="1d898-1164">Требование</span><span class="sxs-lookup"><span data-stu-id="1d898-1164">Requirement</span></span>|<span data-ttu-id="1d898-1165">Значение</span><span class="sxs-lookup"><span data-stu-id="1d898-1165">Value</span></span>|
|---|---|
|[<span data-ttu-id="1d898-1166">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="1d898-1166">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1d898-1167">1.0</span><span class="sxs-lookup"><span data-stu-id="1d898-1167">1.0</span></span>|
|[<span data-ttu-id="1d898-1168">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="1d898-1168">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1d898-1169">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1d898-1169">ReadItem</span></span>|
|[<span data-ttu-id="1d898-1170">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="1d898-1170">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="1d898-1171">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="1d898-1171">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="1d898-1172">Пример</span><span class="sxs-lookup"><span data-stu-id="1d898-1172">Example</span></span>

<span data-ttu-id="1d898-p174">Приведенный ниже пример кода показывает, как асинхронно загружать настраиваемые свойства, характерные для текущего элемента, с помощью метода `loadCustomPropertiesAsync`. Этот пример также показывает, как сохранять эти свойства на сервере с помощью метода `CustomProperties.saveAsync`. После загрузки настраиваемых свойств в этом примере кода метод `CustomProperties.get` используется для считывания настраиваемого свойства `myProp`, метод `CustomProperties.set` — для записи настраиваемого свойства `otherProp`, а метод `saveAsync` — для сохранения настраиваемых свойств.</span><span class="sxs-lookup"><span data-stu-id="1d898-p174">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

```js
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

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="1d898-1176">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="1d898-1176">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="1d898-1177">Удаляет вложение из сообщения или встречи.</span><span class="sxs-lookup"><span data-stu-id="1d898-1177">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="1d898-p175">Метод `removeAttachmentAsync` удаляет из элемента вложение с указанным идентификатором. Идентификатор вложения рекомендуется использовать для удаления вложения, только если оно добавлено тем же почтовым приложением в ходе текущего сеанса. В Outlook Web App и Outlook Web App для устройств идентификатор вложения действителен только в рамках одного сеанса. Сеанс завершается, когда пользователь закрывает приложение или начинает создавать элемент во встроенной форме, а затем переходит из формы в отдельное окно.</span><span class="sxs-lookup"><span data-stu-id="1d898-p175">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1d898-1182">Параметры</span><span class="sxs-lookup"><span data-stu-id="1d898-1182">Parameters:</span></span>

|<span data-ttu-id="1d898-1183">Имя</span><span class="sxs-lookup"><span data-stu-id="1d898-1183">Name</span></span>|<span data-ttu-id="1d898-1184">Тип</span><span class="sxs-lookup"><span data-stu-id="1d898-1184">Type</span></span>|<span data-ttu-id="1d898-1185">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="1d898-1185">Attributes</span></span>|<span data-ttu-id="1d898-1186">Описание</span><span class="sxs-lookup"><span data-stu-id="1d898-1186">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="1d898-1187">String</span><span class="sxs-lookup"><span data-stu-id="1d898-1187">String</span></span>||<span data-ttu-id="1d898-1188">Идентификатор удаляемого вложения.</span><span class="sxs-lookup"><span data-stu-id="1d898-1188">The identifier of the attachment to remove.</span></span>|
|`options`|<span data-ttu-id="1d898-1189">Object</span><span class="sxs-lookup"><span data-stu-id="1d898-1189">Object</span></span>|<span data-ttu-id="1d898-1190">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="1d898-1190">&lt;optional&gt;</span></span>|<span data-ttu-id="1d898-1191">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="1d898-1191">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="1d898-1192">Object</span><span class="sxs-lookup"><span data-stu-id="1d898-1192">Object</span></span>|<span data-ttu-id="1d898-1193">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="1d898-1193">&lt;optional&gt;</span></span>|<span data-ttu-id="1d898-1194">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="1d898-1194">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="1d898-1195">функция</span><span class="sxs-lookup"><span data-stu-id="1d898-1195">function</span></span>|<span data-ttu-id="1d898-1196">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="1d898-1196">&lt;optional&gt;</span></span>|<span data-ttu-id="1d898-1197">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="1d898-1197">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="1d898-1198">Если удалить вложение не удается, свойство `asyncResult.error` содержит код ошибки с указанием ее причины.</span><span class="sxs-lookup"><span data-stu-id="1d898-1198">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="1d898-1199">Ошибки</span><span class="sxs-lookup"><span data-stu-id="1d898-1199">Errors</span></span>

|<span data-ttu-id="1d898-1200">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="1d898-1200">Error code</span></span>|<span data-ttu-id="1d898-1201">Описание</span><span class="sxs-lookup"><span data-stu-id="1d898-1201">Description</span></span>|
|------------|-------------|
|`InvalidAttachmentId`|<span data-ttu-id="1d898-1202">Идентификатор вложения не существует.</span><span class="sxs-lookup"><span data-stu-id="1d898-1202">The attachment identifier does not exist.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1d898-1203">Требования</span><span class="sxs-lookup"><span data-stu-id="1d898-1203">Requirements</span></span>

|<span data-ttu-id="1d898-1204">Требование</span><span class="sxs-lookup"><span data-stu-id="1d898-1204">Requirement</span></span>|<span data-ttu-id="1d898-1205">Значение</span><span class="sxs-lookup"><span data-stu-id="1d898-1205">Value</span></span>|
|---|---|
|[<span data-ttu-id="1d898-1206">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="1d898-1206">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1d898-1207">1.1</span><span class="sxs-lookup"><span data-stu-id="1d898-1207">1.1</span></span>|
|[<span data-ttu-id="1d898-1208">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="1d898-1208">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1d898-1209">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="1d898-1209">ReadWriteItem</span></span>|
|[<span data-ttu-id="1d898-1210">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="1d898-1210">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="1d898-1211">Создание</span><span class="sxs-lookup"><span data-stu-id="1d898-1211">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="1d898-1212">Пример</span><span class="sxs-lookup"><span data-stu-id="1d898-1212">Example</span></span>

<span data-ttu-id="1d898-1213">Указанный ниже код удаляет вложение с идентификатором "0".</span><span class="sxs-lookup"><span data-stu-id="1d898-1213">The following code removes an attachment with an identifier of '0'.</span></span>

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

####  <a name="removehandlerasynceventtype-options-callback"></a><span data-ttu-id="1d898-1214">removeHandlerAsync(eventType, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="1d898-1214">removeHandlerAsync(eventType, [options], [callback])</span></span>

<span data-ttu-id="1d898-1215">Удаляет обработчиков для поддерживаемого типа события.</span><span class="sxs-lookup"><span data-stu-id="1d898-1215">Removes the event handlers for a supported event type.</span></span>

<span data-ttu-id="1d898-1216">Сейчас поддерживаются следующие типы событий: `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged` и `Office.EventType.RecurrenceChanged`.</span><span class="sxs-lookup"><span data-stu-id="1d898-1216">Currently the supported event types are `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`</span></span>

##### <a name="parameters"></a><span data-ttu-id="1d898-1217">Параметры</span><span class="sxs-lookup"><span data-stu-id="1d898-1217">Parameters:</span></span>

| <span data-ttu-id="1d898-1218">Имя</span><span class="sxs-lookup"><span data-stu-id="1d898-1218">Name</span></span> | <span data-ttu-id="1d898-1219">Тип</span><span class="sxs-lookup"><span data-stu-id="1d898-1219">Type</span></span> | <span data-ttu-id="1d898-1220">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="1d898-1220">Attributes</span></span> | <span data-ttu-id="1d898-1221">Описание</span><span class="sxs-lookup"><span data-stu-id="1d898-1221">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="1d898-1222">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="1d898-1222">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="1d898-1223">Событие, которое должно вызвать обработчик.</span><span class="sxs-lookup"><span data-stu-id="1d898-1223">The event that should invoke the handler.</span></span> |
| `options` | <span data-ttu-id="1d898-1224">Object</span><span class="sxs-lookup"><span data-stu-id="1d898-1224">Object</span></span> | <span data-ttu-id="1d898-1225">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="1d898-1225">&lt;optional&gt;</span></span> | <span data-ttu-id="1d898-1226">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="1d898-1226">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="1d898-1227">Object</span><span class="sxs-lookup"><span data-stu-id="1d898-1227">Object</span></span> | <span data-ttu-id="1d898-1228">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="1d898-1228">&lt;optional&gt;</span></span> | <span data-ttu-id="1d898-1229">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="1d898-1229">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="1d898-1230">функция</span><span class="sxs-lookup"><span data-stu-id="1d898-1230">function</span></span>| <span data-ttu-id="1d898-1231">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="1d898-1231">&lt;optional&gt;</span></span>|<span data-ttu-id="1d898-1232">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="1d898-1232">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1d898-1233">Требования</span><span class="sxs-lookup"><span data-stu-id="1d898-1233">Requirements</span></span>

|<span data-ttu-id="1d898-1234">Требование</span><span class="sxs-lookup"><span data-stu-id="1d898-1234">Requirement</span></span>| <span data-ttu-id="1d898-1235">Значение</span><span class="sxs-lookup"><span data-stu-id="1d898-1235">Value</span></span>|
|---|---|
|[<span data-ttu-id="1d898-1236">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="1d898-1236">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1d898-1237">1.7</span><span class="sxs-lookup"><span data-stu-id="1d898-1237">1.7</span></span> |
|[<span data-ttu-id="1d898-1238">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="1d898-1238">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1d898-1239">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1d898-1239">ReadItem</span></span> |
|[<span data-ttu-id="1d898-1240">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="1d898-1240">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1d898-1241">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="1d898-1241">Compose or read</span></span> |

##### <a name="example"></a><span data-ttu-id="1d898-1242">Пример</span><span class="sxs-lookup"><span data-stu-id="1d898-1242">Example</span></span>

```js
Office.initialize = function (reason) {
  $(document).ready(function () {
    Office.context.mailbox.item.removeHandlerAsync(Office.EventType.RecurrenceChanged, function (result) {
      if (result.status === Office.AsyncResultStatus.Failed) {
        // Handle error
      }
    });
  });
};
```

####  <a name="saveasyncoptions-callback"></a><span data-ttu-id="1d898-1243">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="1d898-1243">saveAsync([options], callback)</span></span>

<span data-ttu-id="1d898-1244">Асинхронно сохраняет элемент.</span><span class="sxs-lookup"><span data-stu-id="1d898-1244">Asynchronously saves an item.</span></span>

<span data-ttu-id="1d898-p176">При вызове этот метод сохраняет текущее сообщение в виде черновика и возвращает идентификатор элемента с помощью метода обратного вызова. В Outlook Web App или интерактивном режиме Outlook этот элемент сохраняется на сервере. В Outlook в режиме кэширования этот элемент сохраняется в локальном кэше.</span><span class="sxs-lookup"><span data-stu-id="1d898-p176">When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="1d898-1248">Если в режиме создания надстройка вызывает для элемента метод `saveAsync`, чтобы получить параметр `itemId` для использования с EWS или REST API, необходимо помнить, что синхронизация элемента с сервером может занять много времени, если Outlook работает в режиме кэширования данных.</span><span class="sxs-lookup"><span data-stu-id="1d898-1248">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="1d898-1249">До окончания синхронизации элемента применение параметра `itemId` будет приводить к ошибке.</span><span class="sxs-lookup"><span data-stu-id="1d898-1249">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="1d898-p178">Если метод `saveAsync` вызывается для встречи в режиме создания, она сохраняется как обычная встреча в календаре пользователя, а не как черновик. При сохранении новой встречи приглашения не отправляются. При сохранении существующей встречи уведомления отправляются добавленным или удаленным участникам.</span><span class="sxs-lookup"><span data-stu-id="1d898-p178">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="1d898-1253">Следующие клиенты отличаются другим поведением `saveAsync` в отношении встреч в режиме создания:</span><span class="sxs-lookup"><span data-stu-id="1d898-1253">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="1d898-1254">Outlook для Mac не поддерживает `saveAsync` для собраний в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="1d898-1254">Mac Outlook does not support `saveAsync` on a meeting in compose mode.</span></span> <span data-ttu-id="1d898-1255">При вызове `saveAsync` для собрания в Outlook для Mac возвращается ошибка.</span><span class="sxs-lookup"><span data-stu-id="1d898-1255">Calling `saveAsync` on a meeting in Mac Outlook will return an error.</span></span>
> - <span data-ttu-id="1d898-1256">Outlook в Интернете всегда отправляет приглашение или обновление при вызове `saveAsync` для встречи в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="1d898-1256">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1d898-1257">Параметры:</span><span class="sxs-lookup"><span data-stu-id="1d898-1257">Parameters:</span></span>

|<span data-ttu-id="1d898-1258">Имя</span><span class="sxs-lookup"><span data-stu-id="1d898-1258">Name</span></span>|<span data-ttu-id="1d898-1259">Тип</span><span class="sxs-lookup"><span data-stu-id="1d898-1259">Type</span></span>|<span data-ttu-id="1d898-1260">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="1d898-1260">Attributes</span></span>|<span data-ttu-id="1d898-1261">Описание</span><span class="sxs-lookup"><span data-stu-id="1d898-1261">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="1d898-1262">Object</span><span class="sxs-lookup"><span data-stu-id="1d898-1262">Object</span></span>|<span data-ttu-id="1d898-1263">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="1d898-1263">&lt;optional&gt;</span></span>|<span data-ttu-id="1d898-1264">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="1d898-1264">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="1d898-1265">Object</span><span class="sxs-lookup"><span data-stu-id="1d898-1265">Object</span></span>|<span data-ttu-id="1d898-1266">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="1d898-1266">&lt;optional&gt;</span></span>|<span data-ttu-id="1d898-1267">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="1d898-1267">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="1d898-1268">функция</span><span class="sxs-lookup"><span data-stu-id="1d898-1268">function</span></span>||<span data-ttu-id="1d898-1269">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="1d898-1269">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="1d898-1270">После успешного выполнения идентификатор элемента будет предоставлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="1d898-1270">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1d898-1271">Требования</span><span class="sxs-lookup"><span data-stu-id="1d898-1271">Requirements</span></span>

|<span data-ttu-id="1d898-1272">Требование</span><span class="sxs-lookup"><span data-stu-id="1d898-1272">Requirement</span></span>|<span data-ttu-id="1d898-1273">Значение</span><span class="sxs-lookup"><span data-stu-id="1d898-1273">Value</span></span>|
|---|---|
|[<span data-ttu-id="1d898-1274">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="1d898-1274">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1d898-1275">1.3</span><span class="sxs-lookup"><span data-stu-id="1d898-1275">1.3</span></span>|
|[<span data-ttu-id="1d898-1276">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="1d898-1276">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1d898-1277">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="1d898-1277">ReadWriteItem</span></span>|
|[<span data-ttu-id="1d898-1278">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="1d898-1278">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="1d898-1279">Создание</span><span class="sxs-lookup"><span data-stu-id="1d898-1279">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="1d898-1280">Примеры</span><span class="sxs-lookup"><span data-stu-id="1d898-1280">Examples</span></span>

```js
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result
  });
```

<span data-ttu-id="1d898-p180">Ниже приведен пример параметра `result`, переданного функции обратного вызова. Свойство `value` содержит идентификатор элемента.</span><span class="sxs-lookup"><span data-stu-id="1d898-p180">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

####  <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="1d898-1283">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="1d898-1283">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="1d898-1284">Асинхронно вставляет данные в текст или тему сообщения.</span><span class="sxs-lookup"><span data-stu-id="1d898-1284">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="1d898-p181">Метод `setSelectedDataAsync` вставляет указанную строку в местоположение курсора в теме или тексте элемента либо, если текст выделен в редакторе, он заменяет выделенный текст. Если курсор находится вне текста или темы элемента, возвращается ошибка. После вставки курсор помещается в конец вставленного содержимого.</span><span class="sxs-lookup"><span data-stu-id="1d898-p181">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1d898-1288">Параметры:</span><span class="sxs-lookup"><span data-stu-id="1d898-1288">Parameters:</span></span>

|<span data-ttu-id="1d898-1289">Имя</span><span class="sxs-lookup"><span data-stu-id="1d898-1289">Name</span></span>|<span data-ttu-id="1d898-1290">Тип</span><span class="sxs-lookup"><span data-stu-id="1d898-1290">Type</span></span>|<span data-ttu-id="1d898-1291">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="1d898-1291">Attributes</span></span>|<span data-ttu-id="1d898-1292">Описание</span><span class="sxs-lookup"><span data-stu-id="1d898-1292">Description</span></span>|
|---|---|---|---|
|`data`|<span data-ttu-id="1d898-1293">String</span><span class="sxs-lookup"><span data-stu-id="1d898-1293">String</span></span>||<span data-ttu-id="1d898-p182">Вставляемые данные. Объем данных не должен превышать 1 000 000 символов. Если передано больше 1 000 000 символов, возвращается исключение `ArgumentOutOfRange`.</span><span class="sxs-lookup"><span data-stu-id="1d898-p182">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`|<span data-ttu-id="1d898-1297">Object</span><span class="sxs-lookup"><span data-stu-id="1d898-1297">Object</span></span>|<span data-ttu-id="1d898-1298">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="1d898-1298">&lt;optional&gt;</span></span>|<span data-ttu-id="1d898-1299">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="1d898-1299">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="1d898-1300">Object</span><span class="sxs-lookup"><span data-stu-id="1d898-1300">Object</span></span>|<span data-ttu-id="1d898-1301">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="1d898-1301">&lt;optional&gt;</span></span>|<span data-ttu-id="1d898-1302">В методе обратного вызова разработчики могут указать любой объект, к которому необходимо получить доступ.</span><span class="sxs-lookup"><span data-stu-id="1d898-1302">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="1d898-1303">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="1d898-1303">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="1d898-1304">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="1d898-1304">&lt;optional&gt;</span></span>|<span data-ttu-id="1d898-p183">Если задано значение `text`, текущий стиль применяется в Outlook Web App и Outlook. Если поле представляет собой редактор HTML, вставляются только текстовые данные, даже если они имеют формат HTML.</span><span class="sxs-lookup"><span data-stu-id="1d898-p183">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="1d898-p184">Если задано значение `html` и поле (не тема) поддерживает HTML, в Outlook Web App применяется текущий стиль, а в Outlook — стиль по умолчанию. Если поле является текстовым, возвращается ошибка `InvalidDataFormat`.</span><span class="sxs-lookup"><span data-stu-id="1d898-p184">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="1d898-1309">Если свойство `coercionType` не задано, результат зависит от поля: если поле имеет формат HTML, используется текст в формате HTML, а если поле текстовое, применяется обычный текст.</span><span class="sxs-lookup"><span data-stu-id="1d898-1309">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`|<span data-ttu-id="1d898-1310">функция</span><span class="sxs-lookup"><span data-stu-id="1d898-1310">function</span></span>||<span data-ttu-id="1d898-1311">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="1d898-1311">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1d898-1312">Требования</span><span class="sxs-lookup"><span data-stu-id="1d898-1312">Requirements</span></span>

|<span data-ttu-id="1d898-1313">Требование</span><span class="sxs-lookup"><span data-stu-id="1d898-1313">Requirement</span></span>|<span data-ttu-id="1d898-1314">Значение</span><span class="sxs-lookup"><span data-stu-id="1d898-1314">Value</span></span>|
|---|---|
|[<span data-ttu-id="1d898-1315">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="1d898-1315">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1d898-1316">1.2</span><span class="sxs-lookup"><span data-stu-id="1d898-1316">1.2</span></span>|
|[<span data-ttu-id="1d898-1317">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="1d898-1317">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1d898-1318">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="1d898-1318">ReadWriteItem</span></span>|
|[<span data-ttu-id="1d898-1319">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="1d898-1319">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="1d898-1320">Создание</span><span class="sxs-lookup"><span data-stu-id="1d898-1320">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="1d898-1321">Пример</span><span class="sxs-lookup"><span data-stu-id="1d898-1321">Example</span></span>

```js
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
