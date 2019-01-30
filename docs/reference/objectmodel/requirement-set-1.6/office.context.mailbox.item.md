---
title: Office.Context.Mailbox.Item - требование задать 1.6
description: ''
ms.date: 12/18/2018
localization_priority: Normal
ms.openlocfilehash: 23f27a2949ddcdaa17ffe3f4711002d47d699458
ms.sourcegitcommit: d1aa7201820176ed986b9f00bb9c88e055906c77
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 01/23/2019
ms.locfileid: "29387319"
---
# <a name="item"></a><span data-ttu-id="37cd3-102">item</span><span class="sxs-lookup"><span data-stu-id="37cd3-102">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="37cd3-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="37cd3-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="37cd3-p101">Пространство имен `item` используется для доступа к выбранному в данный момент сообщению, приглашению на собрание или описанию встречи. Вы можете определить тип пространства имен `item` с помощью свойства [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook16officemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="37cd3-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook16officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="37cd3-106">Requirements</span><span class="sxs-lookup"><span data-stu-id="37cd3-106">Requirements</span></span>

|<span data-ttu-id="37cd3-107">Требование</span><span class="sxs-lookup"><span data-stu-id="37cd3-107">Requirement</span></span>| <span data-ttu-id="37cd3-108">Значение</span><span class="sxs-lookup"><span data-stu-id="37cd3-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="37cd3-109">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="37cd3-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="37cd3-110">1.0</span><span class="sxs-lookup"><span data-stu-id="37cd3-110">1.0</span></span>|
|[<span data-ttu-id="37cd3-111">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="37cd3-111">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="37cd3-112">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="37cd3-112">Restricted</span></span>|
|[<span data-ttu-id="37cd3-113">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="37cd3-113">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="37cd3-114">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="37cd3-114">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="37cd3-115">Элементы и методы</span><span class="sxs-lookup"><span data-stu-id="37cd3-115">Members and methods</span></span>

| <span data-ttu-id="37cd3-116">Элемент</span><span class="sxs-lookup"><span data-stu-id="37cd3-116">Member</span></span> | <span data-ttu-id="37cd3-117">Тип</span><span class="sxs-lookup"><span data-stu-id="37cd3-117">Type</span></span> |
|--------|------|
| [<span data-ttu-id="37cd3-118">attachments</span><span class="sxs-lookup"><span data-stu-id="37cd3-118">attachments</span></span>](#attachments-arrayattachmentdetailsjavascriptapioutlook16officeattachmentdetails) | <span data-ttu-id="37cd3-119">Элемент</span><span class="sxs-lookup"><span data-stu-id="37cd3-119">Member</span></span> |
| [<span data-ttu-id="37cd3-120">bcc</span><span class="sxs-lookup"><span data-stu-id="37cd3-120">bcc</span></span>](#bcc-recipientsjavascriptapioutlook16officerecipients) | <span data-ttu-id="37cd3-121">Элемент</span><span class="sxs-lookup"><span data-stu-id="37cd3-121">Member</span></span> |
| [<span data-ttu-id="37cd3-122">body</span><span class="sxs-lookup"><span data-stu-id="37cd3-122">body</span></span>](#body-bodyjavascriptapioutlook16officebody) | <span data-ttu-id="37cd3-123">Элемент</span><span class="sxs-lookup"><span data-stu-id="37cd3-123">Member</span></span> |
| [<span data-ttu-id="37cd3-124">cc</span><span class="sxs-lookup"><span data-stu-id="37cd3-124">cc</span></span>](#cc-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients) | <span data-ttu-id="37cd3-125">Элемент</span><span class="sxs-lookup"><span data-stu-id="37cd3-125">Member</span></span> |
| [<span data-ttu-id="37cd3-126">conversationId</span><span class="sxs-lookup"><span data-stu-id="37cd3-126">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="37cd3-127">Элемент</span><span class="sxs-lookup"><span data-stu-id="37cd3-127">Member</span></span> |
| [<span data-ttu-id="37cd3-128">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="37cd3-128">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="37cd3-129">Элемент</span><span class="sxs-lookup"><span data-stu-id="37cd3-129">Member</span></span> |
| [<span data-ttu-id="37cd3-130">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="37cd3-130">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="37cd3-131">Элемент</span><span class="sxs-lookup"><span data-stu-id="37cd3-131">Member</span></span> |
| [<span data-ttu-id="37cd3-132">end</span><span class="sxs-lookup"><span data-stu-id="37cd3-132">end</span></span>](#end-datetimejavascriptapioutlook16officetime) | <span data-ttu-id="37cd3-133">Элемент</span><span class="sxs-lookup"><span data-stu-id="37cd3-133">Member</span></span> |
| [<span data-ttu-id="37cd3-134">from</span><span class="sxs-lookup"><span data-stu-id="37cd3-134">from</span></span>](#from-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails) | <span data-ttu-id="37cd3-135">Элемент</span><span class="sxs-lookup"><span data-stu-id="37cd3-135">Member</span></span> |
| [<span data-ttu-id="37cd3-136">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="37cd3-136">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="37cd3-137">Элемент</span><span class="sxs-lookup"><span data-stu-id="37cd3-137">Member</span></span> |
| [<span data-ttu-id="37cd3-138">itemClass</span><span class="sxs-lookup"><span data-stu-id="37cd3-138">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="37cd3-139">Элемент</span><span class="sxs-lookup"><span data-stu-id="37cd3-139">Member</span></span> |
| [<span data-ttu-id="37cd3-140">itemId</span><span class="sxs-lookup"><span data-stu-id="37cd3-140">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="37cd3-141">Элемент</span><span class="sxs-lookup"><span data-stu-id="37cd3-141">Member</span></span> |
| [<span data-ttu-id="37cd3-142">itemType</span><span class="sxs-lookup"><span data-stu-id="37cd3-142">itemType</span></span>](#itemtype-officemailboxenumsitemtypejavascriptapioutlook16officemailboxenumsitemtype) | <span data-ttu-id="37cd3-143">Элемент</span><span class="sxs-lookup"><span data-stu-id="37cd3-143">Member</span></span> |
| [<span data-ttu-id="37cd3-144">location</span><span class="sxs-lookup"><span data-stu-id="37cd3-144">location</span></span>](#location-stringlocationjavascriptapioutlook16officelocation) | <span data-ttu-id="37cd3-145">Элемент</span><span class="sxs-lookup"><span data-stu-id="37cd3-145">Member</span></span> |
| [<span data-ttu-id="37cd3-146">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="37cd3-146">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="37cd3-147">Элемент</span><span class="sxs-lookup"><span data-stu-id="37cd3-147">Member</span></span> |
| [<span data-ttu-id="37cd3-148">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="37cd3-148">notificationMessages</span></span>](#notificationmessages-notificationmessagesjavascriptapioutlook16officenotificationmessages) | <span data-ttu-id="37cd3-149">Элемент</span><span class="sxs-lookup"><span data-stu-id="37cd3-149">Member</span></span> |
| [<span data-ttu-id="37cd3-150">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="37cd3-150">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients) | <span data-ttu-id="37cd3-151">Элемент</span><span class="sxs-lookup"><span data-stu-id="37cd3-151">Member</span></span> |
| [<span data-ttu-id="37cd3-152">organizer</span><span class="sxs-lookup"><span data-stu-id="37cd3-152">organizer</span></span>](#organizer-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails) | <span data-ttu-id="37cd3-153">Элемент</span><span class="sxs-lookup"><span data-stu-id="37cd3-153">Member</span></span> |
| [<span data-ttu-id="37cd3-154">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="37cd3-154">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients) | <span data-ttu-id="37cd3-155">Элемент</span><span class="sxs-lookup"><span data-stu-id="37cd3-155">Member</span></span> |
| [<span data-ttu-id="37cd3-156">sender</span><span class="sxs-lookup"><span data-stu-id="37cd3-156">sender</span></span>](#sender-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails) | <span data-ttu-id="37cd3-157">Элемент</span><span class="sxs-lookup"><span data-stu-id="37cd3-157">Member</span></span> |
| [<span data-ttu-id="37cd3-158">start</span><span class="sxs-lookup"><span data-stu-id="37cd3-158">start</span></span>](#start-datetimejavascriptapioutlook16officetime) | <span data-ttu-id="37cd3-159">Элемент</span><span class="sxs-lookup"><span data-stu-id="37cd3-159">Member</span></span> |
| [<span data-ttu-id="37cd3-160">subject</span><span class="sxs-lookup"><span data-stu-id="37cd3-160">subject</span></span>](#subject-stringsubjectjavascriptapioutlook16officesubject) | <span data-ttu-id="37cd3-161">Элемент</span><span class="sxs-lookup"><span data-stu-id="37cd3-161">Member</span></span> |
| [<span data-ttu-id="37cd3-162">to</span><span class="sxs-lookup"><span data-stu-id="37cd3-162">to</span></span>](#to-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients) | <span data-ttu-id="37cd3-163">Элемент</span><span class="sxs-lookup"><span data-stu-id="37cd3-163">Member</span></span> |
| [<span data-ttu-id="37cd3-164">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="37cd3-164">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="37cd3-165">Метод</span><span class="sxs-lookup"><span data-stu-id="37cd3-165">Method</span></span> |
| [<span data-ttu-id="37cd3-166">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="37cd3-166">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="37cd3-167">Метод</span><span class="sxs-lookup"><span data-stu-id="37cd3-167">Method</span></span> |
| [<span data-ttu-id="37cd3-168">close</span><span class="sxs-lookup"><span data-stu-id="37cd3-168">close</span></span>](#close) | <span data-ttu-id="37cd3-169">Метод</span><span class="sxs-lookup"><span data-stu-id="37cd3-169">Method</span></span> |
| [<span data-ttu-id="37cd3-170">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="37cd3-170">displayReplyAllForm</span></span>](#displayreplyallformformdata) | <span data-ttu-id="37cd3-171">Метод</span><span class="sxs-lookup"><span data-stu-id="37cd3-171">Method</span></span> |
| [<span data-ttu-id="37cd3-172">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="37cd3-172">displayReplyForm</span></span>](#displayreplyformformdata) | <span data-ttu-id="37cd3-173">Метод</span><span class="sxs-lookup"><span data-stu-id="37cd3-173">Method</span></span> |
| [<span data-ttu-id="37cd3-174">getEntities</span><span class="sxs-lookup"><span data-stu-id="37cd3-174">getEntities</span></span>](#getentities--entitiesjavascriptapioutlook16officeentities) | <span data-ttu-id="37cd3-175">Метод</span><span class="sxs-lookup"><span data-stu-id="37cd3-175">Method</span></span> |
| [<span data-ttu-id="37cd3-176">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="37cd3-176">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook16officecontactmeetingsuggestionjavascriptapioutlook16officemeetingsuggestionphonenumberjavascriptapioutlook16officephonenumbertasksuggestionjavascriptapioutlook16officetasksuggestion) | <span data-ttu-id="37cd3-177">Метод</span><span class="sxs-lookup"><span data-stu-id="37cd3-177">Method</span></span> |
| [<span data-ttu-id="37cd3-178">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="37cd3-178">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook16officecontactmeetingsuggestionjavascriptapioutlook16officemeetingsuggestionphonenumberjavascriptapioutlook16officephonenumbertasksuggestionjavascriptapioutlook16officetasksuggestion) | <span data-ttu-id="37cd3-179">Метод</span><span class="sxs-lookup"><span data-stu-id="37cd3-179">Method</span></span> |
| [<span data-ttu-id="37cd3-180">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="37cd3-180">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="37cd3-181">Метод</span><span class="sxs-lookup"><span data-stu-id="37cd3-181">Method</span></span> |
| [<span data-ttu-id="37cd3-182">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="37cd3-182">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="37cd3-183">Метод</span><span class="sxs-lookup"><span data-stu-id="37cd3-183">Method</span></span> |
| [<span data-ttu-id="37cd3-184">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="37cd3-184">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="37cd3-185">Метод</span><span class="sxs-lookup"><span data-stu-id="37cd3-185">Method</span></span> |
| [<span data-ttu-id="37cd3-186">getSelectedEntities</span><span class="sxs-lookup"><span data-stu-id="37cd3-186">getSelectedEntities</span></span>](#getselectedentities--entitiesjavascriptapioutlook16officeentities) | <span data-ttu-id="37cd3-187">Метод</span><span class="sxs-lookup"><span data-stu-id="37cd3-187">Method</span></span> |
| [<span data-ttu-id="37cd3-188">getSelectedRegExMatches</span><span class="sxs-lookup"><span data-stu-id="37cd3-188">getSelectedRegExMatches</span></span>](#getselectedregexmatches--object) | <span data-ttu-id="37cd3-189">Метод</span><span class="sxs-lookup"><span data-stu-id="37cd3-189">Method</span></span> |
| [<span data-ttu-id="37cd3-190">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="37cd3-190">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="37cd3-191">Метод</span><span class="sxs-lookup"><span data-stu-id="37cd3-191">Method</span></span> |
| [<span data-ttu-id="37cd3-192">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="37cd3-192">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="37cd3-193">Метод</span><span class="sxs-lookup"><span data-stu-id="37cd3-193">Method</span></span> |
| [<span data-ttu-id="37cd3-194">saveAsync</span><span class="sxs-lookup"><span data-stu-id="37cd3-194">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="37cd3-195">Метод</span><span class="sxs-lookup"><span data-stu-id="37cd3-195">Method</span></span> |
| [<span data-ttu-id="37cd3-196">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="37cd3-196">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="37cd3-197">Метод</span><span class="sxs-lookup"><span data-stu-id="37cd3-197">Method</span></span> |

### <a name="example"></a><span data-ttu-id="37cd3-198">Пример</span><span class="sxs-lookup"><span data-stu-id="37cd3-198">Example</span></span>

<span data-ttu-id="37cd3-199">В примере кода JavaScript, приведенном ниже, показано, как получить доступ к свойству `subject` текущего элемента в Outlook.</span><span class="sxs-lookup"><span data-stu-id="37cd3-199">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="37cd3-200">Элементы</span><span class="sxs-lookup"><span data-stu-id="37cd3-200">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook16officeattachmentdetails"></a><span data-ttu-id="37cd3-201">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_6/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="37cd3-201">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_6/office.attachmentdetails)></span></span>

<span data-ttu-id="37cd3-p102">Получает массив вложений для элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="37cd3-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="37cd3-204">Outlook блокирует определенные типы файлов из-за потенциальных проблем с безопасностью, поэтому они не возвращаются.</span><span class="sxs-lookup"><span data-stu-id="37cd3-204">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="37cd3-205">Дополнительные сведения см. в статье [Блокировка вложений в Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="37cd3-205">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="37cd3-206">Тип:</span><span class="sxs-lookup"><span data-stu-id="37cd3-206">Type:</span></span>

*   <span data-ttu-id="37cd3-207">Array.<[AttachmentDetails](/javascript/api/outlook_1_6/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="37cd3-207">Array.<[AttachmentDetails](/javascript/api/outlook_1_6/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="37cd3-208">Требования</span><span class="sxs-lookup"><span data-stu-id="37cd3-208">Requirements</span></span>

|<span data-ttu-id="37cd3-209">Требование</span><span class="sxs-lookup"><span data-stu-id="37cd3-209">Requirement</span></span>| <span data-ttu-id="37cd3-210">Значение</span><span class="sxs-lookup"><span data-stu-id="37cd3-210">Value</span></span>|
|---|---|
|[<span data-ttu-id="37cd3-211">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="37cd3-211">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="37cd3-212">1.0</span><span class="sxs-lookup"><span data-stu-id="37cd3-212">1.0</span></span>|
|[<span data-ttu-id="37cd3-213">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="37cd3-213">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="37cd3-214">ReadItem</span><span class="sxs-lookup"><span data-stu-id="37cd3-214">ReadItem</span></span>|
|[<span data-ttu-id="37cd3-215">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="37cd3-215">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="37cd3-216">Чтение</span><span class="sxs-lookup"><span data-stu-id="37cd3-216">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="37cd3-217">Пример</span><span class="sxs-lookup"><span data-stu-id="37cd3-217">Example</span></span>

<span data-ttu-id="37cd3-218">С помощью приведенного ниже кода можно создать HTML-строку с подробными сведениями обо всех вложениях для текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="37cd3-218">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

####  <a name="bcc-recipientsjavascriptapioutlook16officerecipients"></a><span data-ttu-id="37cd3-219">bcc :[Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="37cd3-219">bcc :[Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

<span data-ttu-id="37cd3-220">Получает объект, который предоставляет методы для получения или обновления получателей скрытой копии сообщения.</span><span class="sxs-lookup"><span data-stu-id="37cd3-220">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="37cd3-221">Только в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="37cd3-221">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="37cd3-222">Тип:</span><span class="sxs-lookup"><span data-stu-id="37cd3-222">Type:</span></span>

*   [<span data-ttu-id="37cd3-223">Recipients</span><span class="sxs-lookup"><span data-stu-id="37cd3-223">Recipients</span></span>](/javascript/api/outlook_1_6/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="37cd3-224">Требования</span><span class="sxs-lookup"><span data-stu-id="37cd3-224">Requirements</span></span>

|<span data-ttu-id="37cd3-225">Требование</span><span class="sxs-lookup"><span data-stu-id="37cd3-225">Requirement</span></span>| <span data-ttu-id="37cd3-226">Значение</span><span class="sxs-lookup"><span data-stu-id="37cd3-226">Value</span></span>|
|---|---|
|[<span data-ttu-id="37cd3-227">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="37cd3-227">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="37cd3-228">1.1</span><span class="sxs-lookup"><span data-stu-id="37cd3-228">1.1</span></span>|
|[<span data-ttu-id="37cd3-229">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="37cd3-229">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="37cd3-230">ReadItem</span><span class="sxs-lookup"><span data-stu-id="37cd3-230">ReadItem</span></span>|
|[<span data-ttu-id="37cd3-231">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="37cd3-231">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="37cd3-232">Создание</span><span class="sxs-lookup"><span data-stu-id="37cd3-232">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="37cd3-233">Пример</span><span class="sxs-lookup"><span data-stu-id="37cd3-233">Example</span></span>

```js
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlook16officebody"></a><span data-ttu-id="37cd3-234">body :[Body](/javascript/api/outlook_1_6/office.body)</span><span class="sxs-lookup"><span data-stu-id="37cd3-234">body :[Body](/javascript/api/outlook_1_6/office.body)</span></span>

<span data-ttu-id="37cd3-235">Получает объект, предоставляющий методы для работы с основным текстом элемента.</span><span class="sxs-lookup"><span data-stu-id="37cd3-235">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="37cd3-236">Тип:</span><span class="sxs-lookup"><span data-stu-id="37cd3-236">Type:</span></span>

*   [<span data-ttu-id="37cd3-237">Body</span><span class="sxs-lookup"><span data-stu-id="37cd3-237">Body</span></span>](/javascript/api/outlook_1_6/office.body)

##### <a name="requirements"></a><span data-ttu-id="37cd3-238">Требования</span><span class="sxs-lookup"><span data-stu-id="37cd3-238">Requirements</span></span>

|<span data-ttu-id="37cd3-239">Требование</span><span class="sxs-lookup"><span data-stu-id="37cd3-239">Requirement</span></span>| <span data-ttu-id="37cd3-240">Значение</span><span class="sxs-lookup"><span data-stu-id="37cd3-240">Value</span></span>|
|---|---|
|[<span data-ttu-id="37cd3-241">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="37cd3-241">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="37cd3-242">1.1</span><span class="sxs-lookup"><span data-stu-id="37cd3-242">1.1</span></span>|
|[<span data-ttu-id="37cd3-243">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="37cd3-243">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="37cd3-244">ReadItem</span><span class="sxs-lookup"><span data-stu-id="37cd3-244">ReadItem</span></span>|
|[<span data-ttu-id="37cd3-245">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="37cd3-245">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="37cd3-246">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="37cd3-246">Compose or read</span></span>|

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients"></a><span data-ttu-id="37cd3-247">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="37cd3-247">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

<span data-ttu-id="37cd3-248">Предоставляет доступ к получателям копии сообщения.</span><span class="sxs-lookup"><span data-stu-id="37cd3-248">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="37cd3-249">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="37cd3-249">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="37cd3-250">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="37cd3-250">Read mode</span></span>

<span data-ttu-id="37cd3-p106">Свойство `cc` возвращает массив, который содержит объект `EmailAddressDetails` для каждого получателя, указанного в строке **Копия** сообщения. Коллекция может включать не более 100 элементов.</span><span class="sxs-lookup"><span data-stu-id="37cd3-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="37cd3-253">Режим создания</span><span class="sxs-lookup"><span data-stu-id="37cd3-253">Compose mode</span></span>

<span data-ttu-id="37cd3-254">Свойство `cc` возвращает объект `Recipients`, предоставляющий методы для получения или обновления получателей, которые указаны в строке **Копия** сообщения.</span><span class="sxs-lookup"><span data-stu-id="37cd3-254">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="37cd3-255">Тип:</span><span class="sxs-lookup"><span data-stu-id="37cd3-255">Type:</span></span>

*   <span data-ttu-id="37cd3-256">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="37cd3-256">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="37cd3-257">Требования</span><span class="sxs-lookup"><span data-stu-id="37cd3-257">Requirements</span></span>

|<span data-ttu-id="37cd3-258">Требование</span><span class="sxs-lookup"><span data-stu-id="37cd3-258">Requirement</span></span>| <span data-ttu-id="37cd3-259">Значение</span><span class="sxs-lookup"><span data-stu-id="37cd3-259">Value</span></span>|
|---|---|
|[<span data-ttu-id="37cd3-260">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="37cd3-260">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="37cd3-261">1.0</span><span class="sxs-lookup"><span data-stu-id="37cd3-261">1.0</span></span>|
|[<span data-ttu-id="37cd3-262">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="37cd3-262">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="37cd3-263">ReadItem</span><span class="sxs-lookup"><span data-stu-id="37cd3-263">ReadItem</span></span>|
|[<span data-ttu-id="37cd3-264">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="37cd3-264">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="37cd3-265">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="37cd3-265">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="37cd3-266">Пример</span><span class="sxs-lookup"><span data-stu-id="37cd3-266">Example</span></span>

```js
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="37cd3-267">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="37cd3-267">(nullable) conversationId :String</span></span>

<span data-ttu-id="37cd3-268">Получает идентификатор разговора по электронной почте, содержащего конкретное сообщение.</span><span class="sxs-lookup"><span data-stu-id="37cd3-268">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="37cd3-p107">Вы можете получить целочисленное значение этого свойства, если ваше почтовое приложение активируется в формах просмотра или формах создания ответов. Если пользователь изменит тему ответа, после его отправки идентификатор беседы будет изменен, и полученное ранее значение будет недействительным.</span><span class="sxs-lookup"><span data-stu-id="37cd3-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="37cd3-p108">Это свойство имеет значение NULL для нового элемента в форме создания. Свойство `conversationId` вернет значение, если пользователь задаст тему и сохранит элемент.</span><span class="sxs-lookup"><span data-stu-id="37cd3-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="37cd3-273">Тип:</span><span class="sxs-lookup"><span data-stu-id="37cd3-273">Type:</span></span>

*   <span data-ttu-id="37cd3-274">String</span><span class="sxs-lookup"><span data-stu-id="37cd3-274">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="37cd3-275">Требования</span><span class="sxs-lookup"><span data-stu-id="37cd3-275">Requirements</span></span>

|<span data-ttu-id="37cd3-276">Требование</span><span class="sxs-lookup"><span data-stu-id="37cd3-276">Requirement</span></span>| <span data-ttu-id="37cd3-277">Значение</span><span class="sxs-lookup"><span data-stu-id="37cd3-277">Value</span></span>|
|---|---|
|[<span data-ttu-id="37cd3-278">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="37cd3-278">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="37cd3-279">1.0</span><span class="sxs-lookup"><span data-stu-id="37cd3-279">1.0</span></span>|
|[<span data-ttu-id="37cd3-280">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="37cd3-280">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="37cd3-281">ReadItem</span><span class="sxs-lookup"><span data-stu-id="37cd3-281">ReadItem</span></span>|
|[<span data-ttu-id="37cd3-282">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="37cd3-282">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="37cd3-283">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="37cd3-283">Compose or read</span></span>|

#### <a name="datetimecreated-date"></a><span data-ttu-id="37cd3-284">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="37cd3-284">dateTimeCreated :Date</span></span>

<span data-ttu-id="37cd3-p109">Получает дату и время создания элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="37cd3-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="37cd3-287">Тип:</span><span class="sxs-lookup"><span data-stu-id="37cd3-287">Type:</span></span>

*   <span data-ttu-id="37cd3-288">Date</span><span class="sxs-lookup"><span data-stu-id="37cd3-288">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="37cd3-289">Требования</span><span class="sxs-lookup"><span data-stu-id="37cd3-289">Requirements</span></span>

|<span data-ttu-id="37cd3-290">Требование</span><span class="sxs-lookup"><span data-stu-id="37cd3-290">Requirement</span></span>| <span data-ttu-id="37cd3-291">Значение</span><span class="sxs-lookup"><span data-stu-id="37cd3-291">Value</span></span>|
|---|---|
|[<span data-ttu-id="37cd3-292">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="37cd3-292">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="37cd3-293">1.0</span><span class="sxs-lookup"><span data-stu-id="37cd3-293">1.0</span></span>|
|[<span data-ttu-id="37cd3-294">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="37cd3-294">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="37cd3-295">ReadItem</span><span class="sxs-lookup"><span data-stu-id="37cd3-295">ReadItem</span></span>|
|[<span data-ttu-id="37cd3-296">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="37cd3-296">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="37cd3-297">Чтение</span><span class="sxs-lookup"><span data-stu-id="37cd3-297">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="37cd3-298">Пример</span><span class="sxs-lookup"><span data-stu-id="37cd3-298">Example</span></span>

```js
var created = Office.context.mailbox.item.dateTimeCreated;
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="37cd3-299">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="37cd3-299">dateTimeModified :Date</span></span>

<span data-ttu-id="37cd3-p110">Получает дату и время последнего изменения элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="37cd3-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="37cd3-302">Этот элемент не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="37cd3-302">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="37cd3-303">Тип:</span><span class="sxs-lookup"><span data-stu-id="37cd3-303">Type:</span></span>

*   <span data-ttu-id="37cd3-304">Date</span><span class="sxs-lookup"><span data-stu-id="37cd3-304">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="37cd3-305">Требования</span><span class="sxs-lookup"><span data-stu-id="37cd3-305">Requirements</span></span>

|<span data-ttu-id="37cd3-306">Требование</span><span class="sxs-lookup"><span data-stu-id="37cd3-306">Requirement</span></span>| <span data-ttu-id="37cd3-307">Значение</span><span class="sxs-lookup"><span data-stu-id="37cd3-307">Value</span></span>|
|---|---|
|[<span data-ttu-id="37cd3-308">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="37cd3-308">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="37cd3-309">1.0</span><span class="sxs-lookup"><span data-stu-id="37cd3-309">1.0</span></span>|
|[<span data-ttu-id="37cd3-310">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="37cd3-310">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="37cd3-311">ReadItem</span><span class="sxs-lookup"><span data-stu-id="37cd3-311">ReadItem</span></span>|
|[<span data-ttu-id="37cd3-312">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="37cd3-312">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="37cd3-313">Чтение</span><span class="sxs-lookup"><span data-stu-id="37cd3-313">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="37cd3-314">Пример</span><span class="sxs-lookup"><span data-stu-id="37cd3-314">Example</span></span>

```js
var modified = Office.context.mailbox.item.dateTimeModified;
```

####  <a name="end-datetimejavascriptapioutlook16officetime"></a><span data-ttu-id="37cd3-315">end :Date|[Time](/javascript/api/outlook_1_6/office.time)</span><span class="sxs-lookup"><span data-stu-id="37cd3-315">end :Date|[Time](/javascript/api/outlook_1_6/office.time)</span></span>

<span data-ttu-id="37cd3-316">Получает или задает дату и время окончания встречи.</span><span class="sxs-lookup"><span data-stu-id="37cd3-316">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="37cd3-p111">Свойство `end` представлено в виде значения даты и времени в формате UTC. Преобразовать значение свойства end в местные значения даты и времени клиента можно с помощью метода [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook16officelocalclienttime).</span><span class="sxs-lookup"><span data-stu-id="37cd3-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook16officelocalclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="37cd3-319">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="37cd3-319">Read mode</span></span>

<span data-ttu-id="37cd3-320">Свойство `end` возвращает объект `Date`.</span><span class="sxs-lookup"><span data-stu-id="37cd3-320">The `end` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="37cd3-321">Режим создания</span><span class="sxs-lookup"><span data-stu-id="37cd3-321">Compose mode</span></span>

<span data-ttu-id="37cd3-322">Свойство `end` возвращает объект `Time`.</span><span class="sxs-lookup"><span data-stu-id="37cd3-322">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="37cd3-323">Если вы задаете время окончания с помощью метода [`Time.setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-), необходимо использовать метод [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) для преобразования местного времени на клиенте в формат UTC для сервера.</span><span class="sxs-lookup"><span data-stu-id="37cd3-323">When you use the [`Time.setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="37cd3-324">Тип:</span><span class="sxs-lookup"><span data-stu-id="37cd3-324">Type:</span></span>

*   <span data-ttu-id="37cd3-325">Date | [Time](/javascript/api/outlook_1_6/office.time)</span><span class="sxs-lookup"><span data-stu-id="37cd3-325">Date | [Time](/javascript/api/outlook_1_6/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="37cd3-326">Требования</span><span class="sxs-lookup"><span data-stu-id="37cd3-326">Requirements</span></span>

|<span data-ttu-id="37cd3-327">Требование</span><span class="sxs-lookup"><span data-stu-id="37cd3-327">Requirement</span></span>| <span data-ttu-id="37cd3-328">Значение</span><span class="sxs-lookup"><span data-stu-id="37cd3-328">Value</span></span>|
|---|---|
|[<span data-ttu-id="37cd3-329">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="37cd3-329">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="37cd3-330">1.0</span><span class="sxs-lookup"><span data-stu-id="37cd3-330">1.0</span></span>|
|[<span data-ttu-id="37cd3-331">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="37cd3-331">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="37cd3-332">ReadItem</span><span class="sxs-lookup"><span data-stu-id="37cd3-332">ReadItem</span></span>|
|[<span data-ttu-id="37cd3-333">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="37cd3-333">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="37cd3-334">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="37cd3-334">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="37cd3-335">Пример</span><span class="sxs-lookup"><span data-stu-id="37cd3-335">Example</span></span>

<span data-ttu-id="37cd3-336">В примере ниже показано, как с помощью метода [`setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) объекта `Time` задать время окончания встречи в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="37cd3-336">The following example sets the end time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

#### <a name="from-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails"></a><span data-ttu-id="37cd3-337">from :[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="37cd3-337">from :[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)</span></span>

<span data-ttu-id="37cd3-p112">Получает электронный адрес отправителя сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="37cd3-p112">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="37cd3-p113">Свойства `from` и [`sender`](#sender-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails) представляют одно лицо, если сообщение не отправлено представителем. В противном случае свойство `from` представляет лицо, делегировавшее полномочия, а свойство sender — представителя.</span><span class="sxs-lookup"><span data-stu-id="37cd3-p113">The `from` and [`sender`](#sender-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="37cd3-342">Свойству `recipientType`, принадлежащему объекту `EmailAddressDetails` в свойстве `from`, задано значение `undefined`.</span><span class="sxs-lookup"><span data-stu-id="37cd3-342">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="37cd3-343">Тип:</span><span class="sxs-lookup"><span data-stu-id="37cd3-343">Type:</span></span>

*   [<span data-ttu-id="37cd3-344">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="37cd3-344">EmailAddressDetails</span></span>](/javascript/api/outlook_1_6/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="37cd3-345">Требования</span><span class="sxs-lookup"><span data-stu-id="37cd3-345">Requirements</span></span>

|<span data-ttu-id="37cd3-346">Требование</span><span class="sxs-lookup"><span data-stu-id="37cd3-346">Requirement</span></span>| <span data-ttu-id="37cd3-347">Значение</span><span class="sxs-lookup"><span data-stu-id="37cd3-347">Value</span></span>|
|---|---|
|[<span data-ttu-id="37cd3-348">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="37cd3-348">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="37cd3-349">1.0</span><span class="sxs-lookup"><span data-stu-id="37cd3-349">1.0</span></span>|
|[<span data-ttu-id="37cd3-350">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="37cd3-350">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="37cd3-351">ReadItem</span><span class="sxs-lookup"><span data-stu-id="37cd3-351">ReadItem</span></span>|
|[<span data-ttu-id="37cd3-352">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="37cd3-352">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="37cd3-353">Чтение</span><span class="sxs-lookup"><span data-stu-id="37cd3-353">Read</span></span>|

#### <a name="internetmessageid-string"></a><span data-ttu-id="37cd3-354">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="37cd3-354">internetMessageId :String</span></span>

<span data-ttu-id="37cd3-p114">Получает идентификатор интернет-сообщения для электронного сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="37cd3-p114">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="37cd3-357">Тип:</span><span class="sxs-lookup"><span data-stu-id="37cd3-357">Type:</span></span>

*   <span data-ttu-id="37cd3-358">String</span><span class="sxs-lookup"><span data-stu-id="37cd3-358">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="37cd3-359">Требования</span><span class="sxs-lookup"><span data-stu-id="37cd3-359">Requirements</span></span>

|<span data-ttu-id="37cd3-360">Требование</span><span class="sxs-lookup"><span data-stu-id="37cd3-360">Requirement</span></span>| <span data-ttu-id="37cd3-361">Значение</span><span class="sxs-lookup"><span data-stu-id="37cd3-361">Value</span></span>|
|---|---|
|[<span data-ttu-id="37cd3-362">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="37cd3-362">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="37cd3-363">1.0</span><span class="sxs-lookup"><span data-stu-id="37cd3-363">1.0</span></span>|
|[<span data-ttu-id="37cd3-364">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="37cd3-364">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="37cd3-365">ReadItem</span><span class="sxs-lookup"><span data-stu-id="37cd3-365">ReadItem</span></span>|
|[<span data-ttu-id="37cd3-366">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="37cd3-366">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="37cd3-367">Чтение</span><span class="sxs-lookup"><span data-stu-id="37cd3-367">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="37cd3-368">Пример</span><span class="sxs-lookup"><span data-stu-id="37cd3-368">Example</span></span>

```js
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="37cd3-369">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="37cd3-369">itemClass :String</span></span>

<span data-ttu-id="37cd3-p115">Получает класс элемента веб-служб Exchange для выбранного элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="37cd3-p115">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="37cd3-p116">Свойство `itemClass` указывает класс сообщения выбранного элемента. Ниже приводятся классы сообщения по умолчанию для элемента сообщения или встречи.</span><span class="sxs-lookup"><span data-stu-id="37cd3-p116">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="37cd3-374">Тип</span><span class="sxs-lookup"><span data-stu-id="37cd3-374">Type</span></span> | <span data-ttu-id="37cd3-375">Описание</span><span class="sxs-lookup"><span data-stu-id="37cd3-375">Description</span></span> | <span data-ttu-id="37cd3-376">Класс элемента</span><span class="sxs-lookup"><span data-stu-id="37cd3-376">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="37cd3-377">Элементы встречи</span><span class="sxs-lookup"><span data-stu-id="37cd3-377">Appointment items</span></span> | <span data-ttu-id="37cd3-378">Это элементы календаря для класса элемента `IPM.Appointment` или `IPM.Appointment.Occurence`.</span><span class="sxs-lookup"><span data-stu-id="37cd3-378">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurence` |
| <span data-ttu-id="37cd3-379">Элементы сообщения</span><span class="sxs-lookup"><span data-stu-id="37cd3-379">Message items</span></span> | <span data-ttu-id="37cd3-380">Сюда входят электронные сообщения, для которых по умолчанию задан класс сообщения `IPM.Note`, а также приглашения на собрания, ответы на них и уведомления об их отмене, использующие `IPM.Schedule.Meeting` в качестве базового класса сообщения.</span><span class="sxs-lookup"><span data-stu-id="37cd3-380">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="37cd3-381">Можно создавать настраиваемые классы сообщения, расширяющие классы сообщения по умолчанию, например настраиваемый класс сообщения о встрече `IPM.Appointment.Contoso`.</span><span class="sxs-lookup"><span data-stu-id="37cd3-381">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="37cd3-382">Тип:</span><span class="sxs-lookup"><span data-stu-id="37cd3-382">Type:</span></span>

*   <span data-ttu-id="37cd3-383">String</span><span class="sxs-lookup"><span data-stu-id="37cd3-383">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="37cd3-384">Требования</span><span class="sxs-lookup"><span data-stu-id="37cd3-384">Requirements</span></span>

|<span data-ttu-id="37cd3-385">Требование</span><span class="sxs-lookup"><span data-stu-id="37cd3-385">Requirement</span></span>| <span data-ttu-id="37cd3-386">Значение</span><span class="sxs-lookup"><span data-stu-id="37cd3-386">Value</span></span>|
|---|---|
|[<span data-ttu-id="37cd3-387">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="37cd3-387">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="37cd3-388">1.0</span><span class="sxs-lookup"><span data-stu-id="37cd3-388">1.0</span></span>|
|[<span data-ttu-id="37cd3-389">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="37cd3-389">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="37cd3-390">ReadItem</span><span class="sxs-lookup"><span data-stu-id="37cd3-390">ReadItem</span></span>|
|[<span data-ttu-id="37cd3-391">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="37cd3-391">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="37cd3-392">Чтение</span><span class="sxs-lookup"><span data-stu-id="37cd3-392">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="37cd3-393">Пример</span><span class="sxs-lookup"><span data-stu-id="37cd3-393">Example</span></span>

```js
var itemClass = Office.context.mailbox.item.itemClass;
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="37cd3-394">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="37cd3-394">(nullable) itemId :String</span></span>

<span data-ttu-id="37cd3-p117">Получает идентификатор элемента веб-служб Exchange для текущего элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="37cd3-p117">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="37cd3-397">Идентификатор, возвращаемый свойством `itemId`, совпадает с идентификатором элемента веб-служб Exchange.</span><span class="sxs-lookup"><span data-stu-id="37cd3-397">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="37cd3-398">Свойство `itemId` не совпадает с идентификатором записи Outlook, а также идентификатором, который используется REST API Outlook.</span><span class="sxs-lookup"><span data-stu-id="37cd3-398">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="37cd3-399">Прежде чем совершать вызовы REST API, используя это значение, его необходимо преобразовать с помощью [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="37cd3-399">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="37cd3-400">Дополнительные сведения см. в статье [Использование REST API Outlook из надстройки Outlook](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="37cd3-400">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="37cd3-p119">Свойство `itemId` недоступно в режиме создания. Если требуется идентификатор элемента, с помощью метода [`saveAsync`](#saveasyncoptions-callback) можно сохранить элемент в хранилище. При этом в параметре [`AsyncResult.value`](/javascript/api/office/office.asyncresult) функции обратного вызова возвращается идентификатор элемента.</span><span class="sxs-lookup"><span data-stu-id="37cd3-p119">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="37cd3-403">Тип:</span><span class="sxs-lookup"><span data-stu-id="37cd3-403">Type:</span></span>

*   <span data-ttu-id="37cd3-404">String</span><span class="sxs-lookup"><span data-stu-id="37cd3-404">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="37cd3-405">Требования</span><span class="sxs-lookup"><span data-stu-id="37cd3-405">Requirements</span></span>

|<span data-ttu-id="37cd3-406">Требование</span><span class="sxs-lookup"><span data-stu-id="37cd3-406">Requirement</span></span>| <span data-ttu-id="37cd3-407">Значение</span><span class="sxs-lookup"><span data-stu-id="37cd3-407">Value</span></span>|
|---|---|
|[<span data-ttu-id="37cd3-408">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="37cd3-408">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="37cd3-409">1.0</span><span class="sxs-lookup"><span data-stu-id="37cd3-409">1.0</span></span>|
|[<span data-ttu-id="37cd3-410">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="37cd3-410">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="37cd3-411">ReadItem</span><span class="sxs-lookup"><span data-stu-id="37cd3-411">ReadItem</span></span>|
|[<span data-ttu-id="37cd3-412">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="37cd3-412">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="37cd3-413">Чтение</span><span class="sxs-lookup"><span data-stu-id="37cd3-413">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="37cd3-414">Пример</span><span class="sxs-lookup"><span data-stu-id="37cd3-414">Example</span></span>

<span data-ttu-id="37cd3-p120">Указанный ниже код проверяет наличие идентификатора элемента. Если свойство `itemId` возвращает значение `null` или `undefined`, элемент будет сохранен в хранилище, а из асинхронного результата будет получен идентификатор элемента.</span><span class="sxs-lookup"><span data-stu-id="37cd3-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```js
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result){
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook16officemailboxenumsitemtype"></a><span data-ttu-id="37cd3-417">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_6/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="37cd3-417">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_6/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="37cd3-418">Получает тип элемента, который представляет экземпляр.</span><span class="sxs-lookup"><span data-stu-id="37cd3-418">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="37cd3-419">Свойство `itemType` возвращает одно из значений перечисления `ItemType`, которое указывает, является ли экземпляр объекта `item` сообщением или встречей.</span><span class="sxs-lookup"><span data-stu-id="37cd3-419">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="37cd3-420">Тип:</span><span class="sxs-lookup"><span data-stu-id="37cd3-420">Type:</span></span>

*   [<span data-ttu-id="37cd3-421">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="37cd3-421">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook_1_6/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="37cd3-422">Требования</span><span class="sxs-lookup"><span data-stu-id="37cd3-422">Requirements</span></span>

|<span data-ttu-id="37cd3-423">Требование</span><span class="sxs-lookup"><span data-stu-id="37cd3-423">Requirement</span></span>| <span data-ttu-id="37cd3-424">Значение</span><span class="sxs-lookup"><span data-stu-id="37cd3-424">Value</span></span>|
|---|---|
|[<span data-ttu-id="37cd3-425">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="37cd3-425">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="37cd3-426">1.0</span><span class="sxs-lookup"><span data-stu-id="37cd3-426">1.0</span></span>|
|[<span data-ttu-id="37cd3-427">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="37cd3-427">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="37cd3-428">ReadItem</span><span class="sxs-lookup"><span data-stu-id="37cd3-428">ReadItem</span></span>|
|[<span data-ttu-id="37cd3-429">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="37cd3-429">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="37cd3-430">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="37cd3-430">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="37cd3-431">Пример</span><span class="sxs-lookup"><span data-stu-id="37cd3-431">Example</span></span>

```js
if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Message)
  // do something
else
  // do something else
```

####  <a name="location-stringlocationjavascriptapioutlook16officelocation"></a><span data-ttu-id="37cd3-432">location :String|[Location](/javascript/api/outlook_1_6/office.location)</span><span class="sxs-lookup"><span data-stu-id="37cd3-432">location :String|[Location](/javascript/api/outlook_1_6/office.location)</span></span>

<span data-ttu-id="37cd3-433">Получает или задает место встречи.</span><span class="sxs-lookup"><span data-stu-id="37cd3-433">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="37cd3-434">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="37cd3-434">Read mode</span></span>

<span data-ttu-id="37cd3-435">Свойство `location` возвращает строку, содержащую сведения о месте встречи.</span><span class="sxs-lookup"><span data-stu-id="37cd3-435">The `location` property returns a string that contains the location of the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="37cd3-436">Режим создания</span><span class="sxs-lookup"><span data-stu-id="37cd3-436">Compose mode</span></span>

<span data-ttu-id="37cd3-437">Свойство `location` возвращает объект `Location`, предоставляющий методы, которые используются для получения и задания места встречи.</span><span class="sxs-lookup"><span data-stu-id="37cd3-437">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="37cd3-438">Тип:</span><span class="sxs-lookup"><span data-stu-id="37cd3-438">Type:</span></span>

*   <span data-ttu-id="37cd3-439">String | [Location](/javascript/api/outlook_1_6/office.location)</span><span class="sxs-lookup"><span data-stu-id="37cd3-439">String | [Location](/javascript/api/outlook_1_6/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="37cd3-440">Требования</span><span class="sxs-lookup"><span data-stu-id="37cd3-440">Requirements</span></span>

|<span data-ttu-id="37cd3-441">Требование</span><span class="sxs-lookup"><span data-stu-id="37cd3-441">Requirement</span></span>| <span data-ttu-id="37cd3-442">Значение</span><span class="sxs-lookup"><span data-stu-id="37cd3-442">Value</span></span>|
|---|---|
|[<span data-ttu-id="37cd3-443">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="37cd3-443">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="37cd3-444">1.0</span><span class="sxs-lookup"><span data-stu-id="37cd3-444">1.0</span></span>|
|[<span data-ttu-id="37cd3-445">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="37cd3-445">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="37cd3-446">ReadItem</span><span class="sxs-lookup"><span data-stu-id="37cd3-446">ReadItem</span></span>|
|[<span data-ttu-id="37cd3-447">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="37cd3-447">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="37cd3-448">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="37cd3-448">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="37cd3-449">Пример</span><span class="sxs-lookup"><span data-stu-id="37cd3-449">Example</span></span>

```js
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

#### <a name="normalizedsubject-string"></a><span data-ttu-id="37cd3-450">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="37cd3-450">normalizedSubject :String</span></span>

<span data-ttu-id="37cd3-p121">Получает тему элемента со всеми удаленными префиксами (включая `RE:` и `FWD:`). Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="37cd3-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="37cd3-p122">Свойство normalizedSubject получает тему элемента со стандартными префиксами (такими как `RE:` и `FW:`), добавляемыми почтовыми программами. Для получения темы элемента с неизмененными префиксами используйте свойство [`subject`](#subject-stringsubjectjavascriptapioutlook16officesubject).</span><span class="sxs-lookup"><span data-stu-id="37cd3-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubjectjavascriptapioutlook16officesubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="37cd3-455">Тип:</span><span class="sxs-lookup"><span data-stu-id="37cd3-455">Type:</span></span>

*   <span data-ttu-id="37cd3-456">String</span><span class="sxs-lookup"><span data-stu-id="37cd3-456">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="37cd3-457">Требования</span><span class="sxs-lookup"><span data-stu-id="37cd3-457">Requirements</span></span>

|<span data-ttu-id="37cd3-458">Требование</span><span class="sxs-lookup"><span data-stu-id="37cd3-458">Requirement</span></span>| <span data-ttu-id="37cd3-459">Значение</span><span class="sxs-lookup"><span data-stu-id="37cd3-459">Value</span></span>|
|---|---|
|[<span data-ttu-id="37cd3-460">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="37cd3-460">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="37cd3-461">1.0</span><span class="sxs-lookup"><span data-stu-id="37cd3-461">1.0</span></span>|
|[<span data-ttu-id="37cd3-462">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="37cd3-462">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="37cd3-463">ReadItem</span><span class="sxs-lookup"><span data-stu-id="37cd3-463">ReadItem</span></span>|
|[<span data-ttu-id="37cd3-464">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="37cd3-464">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="37cd3-465">Чтение</span><span class="sxs-lookup"><span data-stu-id="37cd3-465">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="37cd3-466">Пример</span><span class="sxs-lookup"><span data-stu-id="37cd3-466">Example</span></span>

```js
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
```

####  <a name="notificationmessages-notificationmessagesjavascriptapioutlook16officenotificationmessages"></a><span data-ttu-id="37cd3-467">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_6/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="37cd3-467">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_6/office.notificationmessages)</span></span>

<span data-ttu-id="37cd3-468">Получает сообщения уведомления для элемента.</span><span class="sxs-lookup"><span data-stu-id="37cd3-468">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="37cd3-469">Тип:</span><span class="sxs-lookup"><span data-stu-id="37cd3-469">Type:</span></span>

*   [<span data-ttu-id="37cd3-470">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="37cd3-470">NotificationMessages</span></span>](/javascript/api/outlook_1_6/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="37cd3-471">Требования</span><span class="sxs-lookup"><span data-stu-id="37cd3-471">Requirements</span></span>

|<span data-ttu-id="37cd3-472">Требование</span><span class="sxs-lookup"><span data-stu-id="37cd3-472">Requirement</span></span>| <span data-ttu-id="37cd3-473">Значение</span><span class="sxs-lookup"><span data-stu-id="37cd3-473">Value</span></span>|
|---|---|
|[<span data-ttu-id="37cd3-474">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="37cd3-474">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="37cd3-475">1.3</span><span class="sxs-lookup"><span data-stu-id="37cd3-475">1.3</span></span>|
|[<span data-ttu-id="37cd3-476">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="37cd3-476">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="37cd3-477">ReadItem</span><span class="sxs-lookup"><span data-stu-id="37cd3-477">ReadItem</span></span>|
|[<span data-ttu-id="37cd3-478">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="37cd3-478">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="37cd3-479">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="37cd3-479">Compose or read</span></span>|

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients"></a><span data-ttu-id="37cd3-480">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="37cd3-480">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

<span data-ttu-id="37cd3-481">Предоставляет доступ к необязательным участникам события.</span><span class="sxs-lookup"><span data-stu-id="37cd3-481">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="37cd3-482">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="37cd3-482">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="37cd3-483">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="37cd3-483">Read mode</span></span>

<span data-ttu-id="37cd3-484">Свойство `optionalAttendees` возвращает массив, содержащий объект `EmailAddressDetails` для каждого необязательного участника собрания.</span><span class="sxs-lookup"><span data-stu-id="37cd3-484">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="37cd3-485">Режим создания</span><span class="sxs-lookup"><span data-stu-id="37cd3-485">Compose mode</span></span>

<span data-ttu-id="37cd3-486">Свойство `optionalAttendees` возвращает объект `Recipients`, который предоставляет методы для получения или обновления необязательных участников собрания.</span><span class="sxs-lookup"><span data-stu-id="37cd3-486">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="37cd3-487">Тип:</span><span class="sxs-lookup"><span data-stu-id="37cd3-487">Type:</span></span>

*   <span data-ttu-id="37cd3-488">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="37cd3-488">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="37cd3-489">Требования</span><span class="sxs-lookup"><span data-stu-id="37cd3-489">Requirements</span></span>

|<span data-ttu-id="37cd3-490">Требование</span><span class="sxs-lookup"><span data-stu-id="37cd3-490">Requirement</span></span>| <span data-ttu-id="37cd3-491">Значение</span><span class="sxs-lookup"><span data-stu-id="37cd3-491">Value</span></span>|
|---|---|
|[<span data-ttu-id="37cd3-492">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="37cd3-492">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="37cd3-493">1.0</span><span class="sxs-lookup"><span data-stu-id="37cd3-493">1.0</span></span>|
|[<span data-ttu-id="37cd3-494">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="37cd3-494">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="37cd3-495">ReadItem</span><span class="sxs-lookup"><span data-stu-id="37cd3-495">ReadItem</span></span>|
|[<span data-ttu-id="37cd3-496">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="37cd3-496">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="37cd3-497">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="37cd3-497">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="37cd3-498">Пример</span><span class="sxs-lookup"><span data-stu-id="37cd3-498">Example</span></span>

```js
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails"></a><span data-ttu-id="37cd3-499">organizer :[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="37cd3-499">organizer :[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)</span></span>

<span data-ttu-id="37cd3-p124">Получает электронный адрес организатора указанного собрания. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="37cd3-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="37cd3-502">Тип:</span><span class="sxs-lookup"><span data-stu-id="37cd3-502">Type:</span></span>

*   [<span data-ttu-id="37cd3-503">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="37cd3-503">EmailAddressDetails</span></span>](/javascript/api/outlook_1_6/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="37cd3-504">Требования</span><span class="sxs-lookup"><span data-stu-id="37cd3-504">Requirements</span></span>

|<span data-ttu-id="37cd3-505">Требование</span><span class="sxs-lookup"><span data-stu-id="37cd3-505">Requirement</span></span>| <span data-ttu-id="37cd3-506">Значение</span><span class="sxs-lookup"><span data-stu-id="37cd3-506">Value</span></span>|
|---|---|
|[<span data-ttu-id="37cd3-507">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="37cd3-507">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="37cd3-508">1.0</span><span class="sxs-lookup"><span data-stu-id="37cd3-508">1.0</span></span>|
|[<span data-ttu-id="37cd3-509">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="37cd3-509">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="37cd3-510">ReadItem</span><span class="sxs-lookup"><span data-stu-id="37cd3-510">ReadItem</span></span>|
|[<span data-ttu-id="37cd3-511">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="37cd3-511">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="37cd3-512">Чтение</span><span class="sxs-lookup"><span data-stu-id="37cd3-512">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="37cd3-513">Пример</span><span class="sxs-lookup"><span data-stu-id="37cd3-513">Example</span></span>

```js
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
```

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients"></a><span data-ttu-id="37cd3-514">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="37cd3-514">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

<span data-ttu-id="37cd3-515">Предоставляет доступ к обязательным участникам события.</span><span class="sxs-lookup"><span data-stu-id="37cd3-515">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="37cd3-516">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="37cd3-516">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="37cd3-517">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="37cd3-517">Read mode</span></span>

<span data-ttu-id="37cd3-518">Свойство `requiredAttendees` возвращает массив, содержащий объект `EmailAddressDetails` для каждого обязательного участника собрания.</span><span class="sxs-lookup"><span data-stu-id="37cd3-518">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="37cd3-519">Режим создания</span><span class="sxs-lookup"><span data-stu-id="37cd3-519">Compose mode</span></span>

<span data-ttu-id="37cd3-520">Свойство `requiredAttendees` возвращает объект `Recipients`, предоставляющий методы, с помощью которых можно получить или обновить сведения об обязательных участниках собрания.</span><span class="sxs-lookup"><span data-stu-id="37cd3-520">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="37cd3-521">Тип:</span><span class="sxs-lookup"><span data-stu-id="37cd3-521">Type:</span></span>

*   <span data-ttu-id="37cd3-522">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="37cd3-522">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="37cd3-523">Требования</span><span class="sxs-lookup"><span data-stu-id="37cd3-523">Requirements</span></span>

|<span data-ttu-id="37cd3-524">Требование</span><span class="sxs-lookup"><span data-stu-id="37cd3-524">Requirement</span></span>| <span data-ttu-id="37cd3-525">Значение</span><span class="sxs-lookup"><span data-stu-id="37cd3-525">Value</span></span>|
|---|---|
|[<span data-ttu-id="37cd3-526">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="37cd3-526">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="37cd3-527">1.0</span><span class="sxs-lookup"><span data-stu-id="37cd3-527">1.0</span></span>|
|[<span data-ttu-id="37cd3-528">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="37cd3-528">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="37cd3-529">ReadItem</span><span class="sxs-lookup"><span data-stu-id="37cd3-529">ReadItem</span></span>|
|[<span data-ttu-id="37cd3-530">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="37cd3-530">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="37cd3-531">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="37cd3-531">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="37cd3-532">Пример</span><span class="sxs-lookup"><span data-stu-id="37cd3-532">Example</span></span>

```js
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
}
```

#### <a name="sender-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails"></a><span data-ttu-id="37cd3-533">sender :[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="37cd3-533">sender :[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)</span></span>

<span data-ttu-id="37cd3-p126">Получает электронный адрес отправителя электронного сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="37cd3-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="37cd3-p127">Свойства [`from`](#from-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails) и `sender` представляют одно лицо, если сообщение не отправлено представителем. В противном случае свойство `from` представляет лицо, делегировавшее полномочия, а свойство sender — представителя.</span><span class="sxs-lookup"><span data-stu-id="37cd3-p127">The [`from`](#from-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="37cd3-538">Свойству `recipientType`, принадлежащему объекту `EmailAddressDetails` в свойстве `sender`, задано значение `undefined`.</span><span class="sxs-lookup"><span data-stu-id="37cd3-538">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="37cd3-539">Тип:</span><span class="sxs-lookup"><span data-stu-id="37cd3-539">Type:</span></span>

*   [<span data-ttu-id="37cd3-540">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="37cd3-540">EmailAddressDetails</span></span>](/javascript/api/outlook_1_6/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="37cd3-541">Требования</span><span class="sxs-lookup"><span data-stu-id="37cd3-541">Requirements</span></span>

|<span data-ttu-id="37cd3-542">Требование</span><span class="sxs-lookup"><span data-stu-id="37cd3-542">Requirement</span></span>| <span data-ttu-id="37cd3-543">Значение</span><span class="sxs-lookup"><span data-stu-id="37cd3-543">Value</span></span>|
|---|---|
|[<span data-ttu-id="37cd3-544">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="37cd3-544">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="37cd3-545">1.0</span><span class="sxs-lookup"><span data-stu-id="37cd3-545">1.0</span></span>|
|[<span data-ttu-id="37cd3-546">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="37cd3-546">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="37cd3-547">ReadItem</span><span class="sxs-lookup"><span data-stu-id="37cd3-547">ReadItem</span></span>|
|[<span data-ttu-id="37cd3-548">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="37cd3-548">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="37cd3-549">Чтение</span><span class="sxs-lookup"><span data-stu-id="37cd3-549">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="37cd3-550">Пример</span><span class="sxs-lookup"><span data-stu-id="37cd3-550">Example</span></span>

```js
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
```

####  <a name="start-datetimejavascriptapioutlook16officetime"></a><span data-ttu-id="37cd3-551">start :Date|[Time](/javascript/api/outlook_1_6/office.time)</span><span class="sxs-lookup"><span data-stu-id="37cd3-551">start :Date|[Time](/javascript/api/outlook_1_6/office.time)</span></span>

<span data-ttu-id="37cd3-552">Получает или задает дату и время начала встречи.</span><span class="sxs-lookup"><span data-stu-id="37cd3-552">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="37cd3-p128">Свойство `start` представлено в виде значения даты и времени в формате UTC. Это значение можно преобразовать в местные значения даты и времени клиента с помощью метода [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook16officelocalclienttime).</span><span class="sxs-lookup"><span data-stu-id="37cd3-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook16officelocalclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="37cd3-555">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="37cd3-555">Read mode</span></span>

<span data-ttu-id="37cd3-556">Свойство `start` возвращает объект `Date`.</span><span class="sxs-lookup"><span data-stu-id="37cd3-556">The `start` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="37cd3-557">Режим создания</span><span class="sxs-lookup"><span data-stu-id="37cd3-557">Compose mode</span></span>

<span data-ttu-id="37cd3-558">Свойство `start` возвращает объект `Time`.</span><span class="sxs-lookup"><span data-stu-id="37cd3-558">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="37cd3-559">Если вы задаете время начала с помощью метода [`Time.setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-), необходимо использовать метод [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) для преобразования местного времени на клиенте в формат UTC для сервера.</span><span class="sxs-lookup"><span data-stu-id="37cd3-559">When you use the [`Time.setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="37cd3-560">Тип:</span><span class="sxs-lookup"><span data-stu-id="37cd3-560">Type:</span></span>

*   <span data-ttu-id="37cd3-561">Date | [Time](/javascript/api/outlook_1_6/office.time)</span><span class="sxs-lookup"><span data-stu-id="37cd3-561">Date | [Time](/javascript/api/outlook_1_6/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="37cd3-562">Требования</span><span class="sxs-lookup"><span data-stu-id="37cd3-562">Requirements</span></span>

|<span data-ttu-id="37cd3-563">Требование</span><span class="sxs-lookup"><span data-stu-id="37cd3-563">Requirement</span></span>| <span data-ttu-id="37cd3-564">Значение</span><span class="sxs-lookup"><span data-stu-id="37cd3-564">Value</span></span>|
|---|---|
|[<span data-ttu-id="37cd3-565">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="37cd3-565">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="37cd3-566">1.0</span><span class="sxs-lookup"><span data-stu-id="37cd3-566">1.0</span></span>|
|[<span data-ttu-id="37cd3-567">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="37cd3-567">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="37cd3-568">ReadItem</span><span class="sxs-lookup"><span data-stu-id="37cd3-568">ReadItem</span></span>|
|[<span data-ttu-id="37cd3-569">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="37cd3-569">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="37cd3-570">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="37cd3-570">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="37cd3-571">Пример</span><span class="sxs-lookup"><span data-stu-id="37cd3-571">Example</span></span>

<span data-ttu-id="37cd3-572">В примере ниже с помощью метода [`setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) объекта `Time` задается время начала встречи в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="37cd3-572">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

####  <a name="subject-stringsubjectjavascriptapioutlook16officesubject"></a><span data-ttu-id="37cd3-573">subject :String|[Subject](/javascript/api/outlook_1_6/office.subject)</span><span class="sxs-lookup"><span data-stu-id="37cd3-573">subject :String|[Subject](/javascript/api/outlook_1_6/office.subject)</span></span>

<span data-ttu-id="37cd3-574">Получает или задает описание, которое отображается в поле темы элемента.</span><span class="sxs-lookup"><span data-stu-id="37cd3-574">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="37cd3-575">Свойство `subject` получает или задает всю тему элемента для отправки с почтового сервера.</span><span class="sxs-lookup"><span data-stu-id="37cd3-575">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="37cd3-576">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="37cd3-576">Read mode</span></span>

<span data-ttu-id="37cd3-p129">Свойство `subject` возвращает строку. С помощью свойства [`normalizedSubject`](#normalizedsubject-string) можно получить тему без начальных префиксов, таких как `RE:` и `FW:`.</span><span class="sxs-lookup"><span data-stu-id="37cd3-p129">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```js
var subject = Office.context.mailbox.item.subject;
```

##### <a name="compose-mode"></a><span data-ttu-id="37cd3-579">Режим создания</span><span class="sxs-lookup"><span data-stu-id="37cd3-579">Compose mode</span></span>

<span data-ttu-id="37cd3-580">Свойство `subject` возвращает объект `Subject`, который предоставляет методы для получения и задания темы.</span><span class="sxs-lookup"><span data-stu-id="37cd3-580">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```js
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="37cd3-581">Тип:</span><span class="sxs-lookup"><span data-stu-id="37cd3-581">Type:</span></span>

*   <span data-ttu-id="37cd3-582">String | [Subject](/javascript/api/outlook_1_6/office.subject)</span><span class="sxs-lookup"><span data-stu-id="37cd3-582">String | [Subject](/javascript/api/outlook_1_6/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="37cd3-583">Требования</span><span class="sxs-lookup"><span data-stu-id="37cd3-583">Requirements</span></span>

|<span data-ttu-id="37cd3-584">Требование</span><span class="sxs-lookup"><span data-stu-id="37cd3-584">Requirement</span></span>| <span data-ttu-id="37cd3-585">Значение</span><span class="sxs-lookup"><span data-stu-id="37cd3-585">Value</span></span>|
|---|---|
|[<span data-ttu-id="37cd3-586">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="37cd3-586">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="37cd3-587">1.0</span><span class="sxs-lookup"><span data-stu-id="37cd3-587">1.0</span></span>|
|[<span data-ttu-id="37cd3-588">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="37cd3-588">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="37cd3-589">ReadItem</span><span class="sxs-lookup"><span data-stu-id="37cd3-589">ReadItem</span></span>|
|[<span data-ttu-id="37cd3-590">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="37cd3-590">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="37cd3-591">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="37cd3-591">Compose or read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients"></a><span data-ttu-id="37cd3-592">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="37cd3-592">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

<span data-ttu-id="37cd3-593">Предоставляет доступ к получателям, указанным в строке **Кому** сообщения.</span><span class="sxs-lookup"><span data-stu-id="37cd3-593">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="37cd3-594">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="37cd3-594">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="37cd3-595">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="37cd3-595">Read mode</span></span>

<span data-ttu-id="37cd3-p131">Свойство `to` возвращает массив, содержащий объект `EmailAddressDetails` для каждого получателя в строке **Кому** сообщения. Коллекция может включать не более 100 элементов.</span><span class="sxs-lookup"><span data-stu-id="37cd3-p131">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="37cd3-598">Режим создания</span><span class="sxs-lookup"><span data-stu-id="37cd3-598">Compose mode</span></span>

<span data-ttu-id="37cd3-599">Свойство `to` возвращает объект `Recipients`, предоставляющий методы для получения или обновления получателей, которые указаны в строке **Кому** сообщения.</span><span class="sxs-lookup"><span data-stu-id="37cd3-599">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="37cd3-600">Тип:</span><span class="sxs-lookup"><span data-stu-id="37cd3-600">Type:</span></span>

*   <span data-ttu-id="37cd3-601">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="37cd3-601">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="37cd3-602">Требования</span><span class="sxs-lookup"><span data-stu-id="37cd3-602">Requirements</span></span>

|<span data-ttu-id="37cd3-603">Требование</span><span class="sxs-lookup"><span data-stu-id="37cd3-603">Requirement</span></span>| <span data-ttu-id="37cd3-604">Значение</span><span class="sxs-lookup"><span data-stu-id="37cd3-604">Value</span></span>|
|---|---|
|[<span data-ttu-id="37cd3-605">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="37cd3-605">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="37cd3-606">1.0</span><span class="sxs-lookup"><span data-stu-id="37cd3-606">1.0</span></span>|
|[<span data-ttu-id="37cd3-607">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="37cd3-607">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="37cd3-608">ReadItem</span><span class="sxs-lookup"><span data-stu-id="37cd3-608">ReadItem</span></span>|
|[<span data-ttu-id="37cd3-609">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="37cd3-609">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="37cd3-610">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="37cd3-610">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="37cd3-611">Пример</span><span class="sxs-lookup"><span data-stu-id="37cd3-611">Example</span></span>

```js
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

### <a name="methods"></a><span data-ttu-id="37cd3-612">Методы</span><span class="sxs-lookup"><span data-stu-id="37cd3-612">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="37cd3-613">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="37cd3-613">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="37cd3-614">Добавляет файл в сообщение или встречу в качестве вложения.</span><span class="sxs-lookup"><span data-stu-id="37cd3-614">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="37cd3-615">Метод `addFileAttachmentAsync` передает файл по указанному универсальному коду ресурса (URI) и вкладывает его в элемент в форме создания.</span><span class="sxs-lookup"><span data-stu-id="37cd3-615">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="37cd3-616">Идентификатор можно последовательно использовать с методом [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback), чтобы удалить вложение, добавленное во время текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="37cd3-616">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="37cd3-617">Параметры</span><span class="sxs-lookup"><span data-stu-id="37cd3-617">Parameters:</span></span>

|<span data-ttu-id="37cd3-618">Имя</span><span class="sxs-lookup"><span data-stu-id="37cd3-618">Name</span></span>| <span data-ttu-id="37cd3-619">Тип</span><span class="sxs-lookup"><span data-stu-id="37cd3-619">Type</span></span>| <span data-ttu-id="37cd3-620">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="37cd3-620">Attributes</span></span>| <span data-ttu-id="37cd3-621">Описание</span><span class="sxs-lookup"><span data-stu-id="37cd3-621">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="37cd3-622">String</span><span class="sxs-lookup"><span data-stu-id="37cd3-622">String</span></span>||<span data-ttu-id="37cd3-p132">Универсальный код ресурса (URI), представляющий расположение файла, который нужно вложить в сообщение или встречу. Максимальная длина — 2048 символов.</span><span class="sxs-lookup"><span data-stu-id="37cd3-p132">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="37cd3-625">String</span><span class="sxs-lookup"><span data-stu-id="37cd3-625">String</span></span>||<span data-ttu-id="37cd3-p133">Имя вложения, которое отображается при передаче вложения. Максимальная длина — 255 символов.</span><span class="sxs-lookup"><span data-stu-id="37cd3-p133">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="37cd3-628">Object</span><span class="sxs-lookup"><span data-stu-id="37cd3-628">Object</span></span>| <span data-ttu-id="37cd3-629">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="37cd3-629">&lt;optional&gt;</span></span>|<span data-ttu-id="37cd3-630">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="37cd3-630">An object literal that contains one or more of the following properties.</span></span>|
| `options.asyncContext` | <span data-ttu-id="37cd3-631">Object</span><span class="sxs-lookup"><span data-stu-id="37cd3-631">Object</span></span> | <span data-ttu-id="37cd3-632">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="37cd3-632">&lt;optional&gt;</span></span> | <span data-ttu-id="37cd3-633">В методе обратного вызова разработчики могут указать любой объект, к которому необходимо получить доступ.</span><span class="sxs-lookup"><span data-stu-id="37cd3-633">Developers can provide any object they wish to access in the callback method.</span></span> |
| `options.isInline` | <span data-ttu-id="37cd3-634">Boolean</span><span class="sxs-lookup"><span data-stu-id="37cd3-634">Boolean</span></span> | <span data-ttu-id="37cd3-635">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="37cd3-635">&lt;optional&gt;</span></span> | <span data-ttu-id="37cd3-636">Значение `true` указывает на то, что вложение будет встроено в текст сообщения и не должно отображаться в списке вложений.</span><span class="sxs-lookup"><span data-stu-id="37cd3-636">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
|`callback`| <span data-ttu-id="37cd3-637">function</span><span class="sxs-lookup"><span data-stu-id="37cd3-637">function</span></span>| <span data-ttu-id="37cd3-638">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="37cd3-638">&lt;optional&gt;</span></span>|<span data-ttu-id="37cd3-639">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="37cd3-639">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="37cd3-640">После успешного выполнения идентификатор вложения будет представлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="37cd3-640">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="37cd3-641">Если добавить вложение не удастся, объект `asyncResult` будет содержать объект `Error` с описанием ошибки.</span><span class="sxs-lookup"><span data-stu-id="37cd3-641">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="37cd3-642">Ошибки</span><span class="sxs-lookup"><span data-stu-id="37cd3-642">Errors</span></span>

| <span data-ttu-id="37cd3-643">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="37cd3-643">Error code</span></span> | <span data-ttu-id="37cd3-644">Описание</span><span class="sxs-lookup"><span data-stu-id="37cd3-644">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="37cd3-645">Вложение превышает максимальный размер.</span><span class="sxs-lookup"><span data-stu-id="37cd3-645">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="37cd3-646">Расширение вложения не поддерживается.</span><span class="sxs-lookup"><span data-stu-id="37cd3-646">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="37cd3-647">Сообщение или встреча содержат слишком много вложений.</span><span class="sxs-lookup"><span data-stu-id="37cd3-647">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="37cd3-648">Требования</span><span class="sxs-lookup"><span data-stu-id="37cd3-648">Requirements</span></span>

|<span data-ttu-id="37cd3-649">Требование</span><span class="sxs-lookup"><span data-stu-id="37cd3-649">Requirement</span></span>| <span data-ttu-id="37cd3-650">Значение</span><span class="sxs-lookup"><span data-stu-id="37cd3-650">Value</span></span>|
|---|---|
|[<span data-ttu-id="37cd3-651">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="37cd3-651">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="37cd3-652">1.1</span><span class="sxs-lookup"><span data-stu-id="37cd3-652">1.1</span></span>|
|[<span data-ttu-id="37cd3-653">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="37cd3-653">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="37cd3-654">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="37cd3-654">ReadWriteItem</span></span>|
|[<span data-ttu-id="37cd3-655">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="37cd3-655">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="37cd3-656">Создание</span><span class="sxs-lookup"><span data-stu-id="37cd3-656">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="37cd3-657">Примеры</span><span class="sxs-lookup"><span data-stu-id="37cd3-657">Examples</span></span>

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

<span data-ttu-id="37cd3-658">В приведенном ниже примере файл изображения добавляется в качестве встроенного вложения, а ссылка на вложение добавляется в текст сообщения.</span><span class="sxs-lookup"><span data-stu-id="37cd3-658">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

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

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="37cd3-659">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="37cd3-659">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="37cd3-660">Добавляет к сообщению элемент Exchange, например сообщение, в виде вложения.</span><span class="sxs-lookup"><span data-stu-id="37cd3-660">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="37cd3-p134">С помощью метода `addItemAttachmentAsync` можно в элемент формы создания вложить элемент с указанным идентификатором Exchange. Если указать метод обратного вызова, то этот метод вызывается с помощью параметра `asyncResult`, который содержит идентификатор вложения или код, указывающий на ошибки, которые произошли при вложении элемента. При необходимости можно использовать параметр `options` для передачи сведений о состоянии методу обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="37cd3-p134">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="37cd3-664">Идентификатор можно использовать с методом [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback), чтобы удалить вложение, добавленное во время текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="37cd3-664">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="37cd3-665">Если ваша надстройка Office выполняется в Outlook Web App, метод `addItemAttachmentAsync` обеспечивает вложение элементов в элементы, отличные от редактируемого. Однако это действие не рекомендуем выполнять, так как оно не поддерживается.</span><span class="sxs-lookup"><span data-stu-id="37cd3-665">If your Office Add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="37cd3-666">Параметры:</span><span class="sxs-lookup"><span data-stu-id="37cd3-666">Parameters:</span></span>

|<span data-ttu-id="37cd3-667">Имя</span><span class="sxs-lookup"><span data-stu-id="37cd3-667">Name</span></span>| <span data-ttu-id="37cd3-668">Тип</span><span class="sxs-lookup"><span data-stu-id="37cd3-668">Type</span></span>| <span data-ttu-id="37cd3-669">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="37cd3-669">Attributes</span></span>| <span data-ttu-id="37cd3-670">Описание</span><span class="sxs-lookup"><span data-stu-id="37cd3-670">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="37cd3-671">String</span><span class="sxs-lookup"><span data-stu-id="37cd3-671">String</span></span>||<span data-ttu-id="37cd3-p135">Идентификатор Exchange для вкладываемого элемента. Максимальная длина — 100 символов.</span><span class="sxs-lookup"><span data-stu-id="37cd3-p135">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="37cd3-674">String</span><span class="sxs-lookup"><span data-stu-id="37cd3-674">String</span></span>||<span data-ttu-id="37cd3-p136">Тема вкладываемого элемента. Максимальная длина — 255 символов.</span><span class="sxs-lookup"><span data-stu-id="37cd3-p136">The sujbect of the item to be attached. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="37cd3-677">Object</span><span class="sxs-lookup"><span data-stu-id="37cd3-677">Object</span></span>| <span data-ttu-id="37cd3-678">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="37cd3-678">&lt;optional&gt;</span></span>|<span data-ttu-id="37cd3-679">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="37cd3-679">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="37cd3-680">Object</span><span class="sxs-lookup"><span data-stu-id="37cd3-680">Object</span></span>| <span data-ttu-id="37cd3-681">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="37cd3-681">&lt;optional&gt;</span></span>|<span data-ttu-id="37cd3-682">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="37cd3-682">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="37cd3-683">функция</span><span class="sxs-lookup"><span data-stu-id="37cd3-683">function</span></span>| <span data-ttu-id="37cd3-684">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="37cd3-684">&lt;optional&gt;</span></span>|<span data-ttu-id="37cd3-685">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="37cd3-685">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="37cd3-686">После успешного выполнения идентификатор вложения будет представлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="37cd3-686">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="37cd3-687">Если добавить вложение не удастся, объект `asyncResult` будет содержать объект `Error` с описанием ошибки.</span><span class="sxs-lookup"><span data-stu-id="37cd3-687">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="37cd3-688">Ошибки</span><span class="sxs-lookup"><span data-stu-id="37cd3-688">Errors</span></span>

| <span data-ttu-id="37cd3-689">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="37cd3-689">Error code</span></span> | <span data-ttu-id="37cd3-690">Описание</span><span class="sxs-lookup"><span data-stu-id="37cd3-690">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="37cd3-691">Сообщение или встреча содержат слишком много вложений.</span><span class="sxs-lookup"><span data-stu-id="37cd3-691">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="37cd3-692">Требования</span><span class="sxs-lookup"><span data-stu-id="37cd3-692">Requirements</span></span>

|<span data-ttu-id="37cd3-693">Требование</span><span class="sxs-lookup"><span data-stu-id="37cd3-693">Requirement</span></span>| <span data-ttu-id="37cd3-694">Значение</span><span class="sxs-lookup"><span data-stu-id="37cd3-694">Value</span></span>|
|---|---|
|[<span data-ttu-id="37cd3-695">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="37cd3-695">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="37cd3-696">1.1</span><span class="sxs-lookup"><span data-stu-id="37cd3-696">1.1</span></span>|
|[<span data-ttu-id="37cd3-697">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="37cd3-697">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="37cd3-698">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="37cd3-698">ReadWriteItem</span></span>|
|[<span data-ttu-id="37cd3-699">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="37cd3-699">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="37cd3-700">Создание</span><span class="sxs-lookup"><span data-stu-id="37cd3-700">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="37cd3-701">Пример</span><span class="sxs-lookup"><span data-stu-id="37cd3-701">Example</span></span>

<span data-ttu-id="37cd3-702">В следующем примере существующий элемент Outlook добавляется в виде вложения с именем `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="37cd3-702">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

####  <a name="close"></a><span data-ttu-id="37cd3-703">close()</span><span class="sxs-lookup"><span data-stu-id="37cd3-703">close()</span></span>

<span data-ttu-id="37cd3-704">Закрывает текущий создаваемый элемент.</span><span class="sxs-lookup"><span data-stu-id="37cd3-704">Closes the current item that is being composed.</span></span>

<span data-ttu-id="37cd3-p137">Работа метода `close` зависит от текущего состояния создаваемого элемента. Если элемент содержит несохраненные изменения, клиент предложит пользователю сохранить или отклонить их либо отменить действие закрытия.</span><span class="sxs-lookup"><span data-stu-id="37cd3-p137">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="37cd3-707">Если в Outlook в Интернете элемент представляет собой встречу, ранее сохраненную с помощью метода `saveAsync`, пользователю предлагается сохранить, отклонить или отменить действие, даже если с момента последнего сохранения элемента не вносились какие-либо изменения.</span><span class="sxs-lookup"><span data-stu-id="37cd3-707">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="37cd3-708">Если в клиенте Outlook для настольных ПК сообщение представляет собой ответ в тексте, метод `close` не работает.</span><span class="sxs-lookup"><span data-stu-id="37cd3-708">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="37cd3-709">Требования</span><span class="sxs-lookup"><span data-stu-id="37cd3-709">Requirements</span></span>

|<span data-ttu-id="37cd3-710">Требование</span><span class="sxs-lookup"><span data-stu-id="37cd3-710">Requirement</span></span>| <span data-ttu-id="37cd3-711">Значение</span><span class="sxs-lookup"><span data-stu-id="37cd3-711">Value</span></span>|
|---|---|
|[<span data-ttu-id="37cd3-712">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="37cd3-712">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="37cd3-713">1.3</span><span class="sxs-lookup"><span data-stu-id="37cd3-713">1.3</span></span>|
|[<span data-ttu-id="37cd3-714">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="37cd3-714">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="37cd3-715">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="37cd3-715">Restricted</span></span>|
|[<span data-ttu-id="37cd3-716">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="37cd3-716">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="37cd3-717">Создание</span><span class="sxs-lookup"><span data-stu-id="37cd3-717">Compose</span></span>|

#### <a name="displayreplyallformformdata"></a><span data-ttu-id="37cd3-718">displayReplyAllForm(formData)</span><span class="sxs-lookup"><span data-stu-id="37cd3-718">displayReplyAllForm(formData)</span></span>

<span data-ttu-id="37cd3-719">Отображает форму ответа, включающую отправителя и всех получателей выбранного сообщения или организатора и всех участников выбранной встречи.</span><span class="sxs-lookup"><span data-stu-id="37cd3-719">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="37cd3-720">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="37cd3-720">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="37cd3-721">В Outlook Web App форма ответа отображается в виде всплывающей формы в представлении с 3 либо 1 или 2 колонками.</span><span class="sxs-lookup"><span data-stu-id="37cd3-721">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="37cd3-722">Если любой строковый параметр превышает указанные для него ограничения, `displayReplyAllForm` возвращает исключение.</span><span class="sxs-lookup"><span data-stu-id="37cd3-722">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="37cd3-p138">Если в параметре `formData.attachments` указаны вложения, Outlook и Outlook Web App пытаются скачать их и вложить в форму ответа. Если какие-либо вложения добавить не удается, в форме отображается сообщение об ошибке. Если сообщения об ошибках не предусмотрены, то они не отображаются.</span><span class="sxs-lookup"><span data-stu-id="37cd3-p138">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="37cd3-726">Параметры</span><span class="sxs-lookup"><span data-stu-id="37cd3-726">Parameters:</span></span>

| <span data-ttu-id="37cd3-727">Имя</span><span class="sxs-lookup"><span data-stu-id="37cd3-727">Name</span></span> | <span data-ttu-id="37cd3-728">Тип</span><span class="sxs-lookup"><span data-stu-id="37cd3-728">Type</span></span> | <span data-ttu-id="37cd3-729">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="37cd3-729">Attributes</span></span> | <span data-ttu-id="37cd3-730">Описание</span><span class="sxs-lookup"><span data-stu-id="37cd3-730">Description</span></span> |
|---|---|---|---|
|`formData`| <span data-ttu-id="37cd3-731">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="37cd3-731">String &#124; Object</span></span>| |<span data-ttu-id="37cd3-p139">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="37cd3-p139">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="37cd3-734">**ИЛИ**</span><span class="sxs-lookup"><span data-stu-id="37cd3-734">**OR**</span></span><br/><span data-ttu-id="37cd3-p140">Объект, который содержит текст или данные вложения и функцию обратного вызова. Ниже представлено определение этого объекта.</span><span class="sxs-lookup"><span data-stu-id="37cd3-p140">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="37cd3-737">String</span><span class="sxs-lookup"><span data-stu-id="37cd3-737">String</span></span> | <span data-ttu-id="37cd3-738">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="37cd3-738">&lt;optional&gt;</span></span> | <span data-ttu-id="37cd3-p141">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="37cd3-p141">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="37cd3-741">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="37cd3-741">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="37cd3-742">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="37cd3-742">&lt;optional&gt;</span></span> | <span data-ttu-id="37cd3-743">Массив объектов JSON, представляющих собой вложенные файлы или элементы.</span><span class="sxs-lookup"><span data-stu-id="37cd3-743">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="37cd3-744">String</span><span class="sxs-lookup"><span data-stu-id="37cd3-744">String</span></span> | | <span data-ttu-id="37cd3-p142">Указывает тип вложения. Допустимые значения: `file` для вложенного файла и `item` для вложенного элемента.</span><span class="sxs-lookup"><span data-stu-id="37cd3-p142">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="37cd3-747">String</span><span class="sxs-lookup"><span data-stu-id="37cd3-747">String</span></span> | | <span data-ttu-id="37cd3-748">Строка, содержащая имя вложения, длиной до 255 символов.</span><span class="sxs-lookup"><span data-stu-id="37cd3-748">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="37cd3-749">String</span><span class="sxs-lookup"><span data-stu-id="37cd3-749">String</span></span> | | <span data-ttu-id="37cd3-p143">Используется, только если свойству `type` задано значение `file`. URI расположения файла.</span><span class="sxs-lookup"><span data-stu-id="37cd3-p143">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.isInline` | <span data-ttu-id="37cd3-752">Boolean</span><span class="sxs-lookup"><span data-stu-id="37cd3-752">Boolean</span></span> | | <span data-ttu-id="37cd3-p144">Используется, только если свойству `type` задано значение `file`. Значение `true` указывает на то, что вложение будет встроено в текст сообщения и не должно отображаться в списке вложений.</span><span class="sxs-lookup"><span data-stu-id="37cd3-p144">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="37cd3-755">String</span><span class="sxs-lookup"><span data-stu-id="37cd3-755">String</span></span> | | <span data-ttu-id="37cd3-p145">Используется, только если свойству `type` задано значение `item`. Идентификатор вложения EWS. Это строка длиной до 100 символов.</span><span class="sxs-lookup"><span data-stu-id="37cd3-p145">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="37cd3-759">function</span><span class="sxs-lookup"><span data-stu-id="37cd3-759">function</span></span> | <span data-ttu-id="37cd3-760">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="37cd3-760">&lt;optional&gt;</span></span> | <span data-ttu-id="37cd3-761">По завершении работы метода функция, переданная параметру `callback`, вызывается с помощью одного параметра `asyncResult`, представляющего собой объект [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="37cd3-761">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="37cd3-762">Требования</span><span class="sxs-lookup"><span data-stu-id="37cd3-762">Requirements</span></span>

|<span data-ttu-id="37cd3-763">Требование</span><span class="sxs-lookup"><span data-stu-id="37cd3-763">Requirement</span></span>| <span data-ttu-id="37cd3-764">Значение</span><span class="sxs-lookup"><span data-stu-id="37cd3-764">Value</span></span>|
|---|---|
|[<span data-ttu-id="37cd3-765">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="37cd3-765">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="37cd3-766">1.0</span><span class="sxs-lookup"><span data-stu-id="37cd3-766">1.0</span></span>|
|[<span data-ttu-id="37cd3-767">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="37cd3-767">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="37cd3-768">ReadItem</span><span class="sxs-lookup"><span data-stu-id="37cd3-768">ReadItem</span></span>|
|[<span data-ttu-id="37cd3-769">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="37cd3-769">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="37cd3-770">Чтение</span><span class="sxs-lookup"><span data-stu-id="37cd3-770">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="37cd3-771">Примеры</span><span class="sxs-lookup"><span data-stu-id="37cd3-771">Examples</span></span>

<span data-ttu-id="37cd3-772">Приведенный ниже код передает строку в функцию `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="37cd3-772">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="37cd3-773">Ответ с пустым текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="37cd3-773">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="37cd3-774">Ответ только с текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="37cd3-774">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="37cd3-775">Ответ с текстом сообщения и вложенным файлом.</span><span class="sxs-lookup"><span data-stu-id="37cd3-775">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="37cd3-776">Ответ с текстом сообщения и вложенным элементом.</span><span class="sxs-lookup"><span data-stu-id="37cd3-776">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="37cd3-777">Ответ с текстом сообщения, вложенным файлом, вложенным элементом и обратным вызовом.</span><span class="sxs-lookup"><span data-stu-id="37cd3-777">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata"></a><span data-ttu-id="37cd3-778">displayReplyForm(formData)</span><span class="sxs-lookup"><span data-stu-id="37cd3-778">displayReplyForm(formData)</span></span>

<span data-ttu-id="37cd3-779">Отображает форму ответа, включающую только отправителя выбранного сообщения или организатора выбранной встречи.</span><span class="sxs-lookup"><span data-stu-id="37cd3-779">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="37cd3-780">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="37cd3-780">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="37cd3-781">В Outlook Web App форма ответа отображается в виде всплывающей формы в представлении с 3 либо 1 или 2 колонками.</span><span class="sxs-lookup"><span data-stu-id="37cd3-781">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="37cd3-782">Если любой строковый параметр превышает указанные для него ограничения, `displayReplyForm` возвращает исключение.</span><span class="sxs-lookup"><span data-stu-id="37cd3-782">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="37cd3-p146">Если в параметре `formData.attachments` указаны вложения, Outlook и Outlook Web App пытаются скачать их и вложить в форму ответа. Если какие-либо вложения добавить не удается, в форме отображается сообщение об ошибке. Если сообщения об ошибках не предусмотрены, то они не отображаются.</span><span class="sxs-lookup"><span data-stu-id="37cd3-p146">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="37cd3-786">Параметры</span><span class="sxs-lookup"><span data-stu-id="37cd3-786">Parameters:</span></span>

| <span data-ttu-id="37cd3-787">Имя</span><span class="sxs-lookup"><span data-stu-id="37cd3-787">Name</span></span> | <span data-ttu-id="37cd3-788">Тип</span><span class="sxs-lookup"><span data-stu-id="37cd3-788">Type</span></span> | <span data-ttu-id="37cd3-789">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="37cd3-789">Attributes</span></span> | <span data-ttu-id="37cd3-790">Описание</span><span class="sxs-lookup"><span data-stu-id="37cd3-790">Description</span></span> |
|---|---|---|---|
|`formData`| <span data-ttu-id="37cd3-791">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="37cd3-791">String &#124; Object</span></span>| | <span data-ttu-id="37cd3-p147">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="37cd3-p147">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="37cd3-794">**ИЛИ**</span><span class="sxs-lookup"><span data-stu-id="37cd3-794">**OR**</span></span><br/><span data-ttu-id="37cd3-p148">Объект, который содержит текст или данные вложения и функцию обратного вызова. Ниже представлено определение этого объекта.</span><span class="sxs-lookup"><span data-stu-id="37cd3-p148">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="37cd3-797">String</span><span class="sxs-lookup"><span data-stu-id="37cd3-797">String</span></span> | <span data-ttu-id="37cd3-798">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="37cd3-798">&lt;optional&gt;</span></span> | <span data-ttu-id="37cd3-p149">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="37cd3-p149">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="37cd3-801">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="37cd3-801">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="37cd3-802">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="37cd3-802">&lt;optional&gt;</span></span> | <span data-ttu-id="37cd3-803">Массив объектов JSON, представляющих собой вложенные файлы или элементы.</span><span class="sxs-lookup"><span data-stu-id="37cd3-803">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="37cd3-804">String</span><span class="sxs-lookup"><span data-stu-id="37cd3-804">String</span></span> | | <span data-ttu-id="37cd3-p150">Указывает тип вложения. Допустимые значения: `file` для вложенного файла и `item` для вложенного элемента.</span><span class="sxs-lookup"><span data-stu-id="37cd3-p150">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="37cd3-807">String</span><span class="sxs-lookup"><span data-stu-id="37cd3-807">String</span></span> | | <span data-ttu-id="37cd3-808">Строка, содержащая имя вложения, длиной до 255 символов.</span><span class="sxs-lookup"><span data-stu-id="37cd3-808">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="37cd3-809">String</span><span class="sxs-lookup"><span data-stu-id="37cd3-809">String</span></span> | | <span data-ttu-id="37cd3-p151">Используется, только если свойству `type` задано значение `file`. URI расположения файла.</span><span class="sxs-lookup"><span data-stu-id="37cd3-p151">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.isInline` | <span data-ttu-id="37cd3-812">Boolean</span><span class="sxs-lookup"><span data-stu-id="37cd3-812">Boolean</span></span> | | <span data-ttu-id="37cd3-p152">Используется, только если свойству `type` задано значение `file`. Значение `true` указывает на то, что вложение будет встроено в текст сообщения и не должно отображаться в списке вложений.</span><span class="sxs-lookup"><span data-stu-id="37cd3-p152">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="37cd3-815">String</span><span class="sxs-lookup"><span data-stu-id="37cd3-815">String</span></span> | | <span data-ttu-id="37cd3-p153">Используется, только если свойству `type` задано значение `item`. Идентификатор вложения EWS. Это строка длиной до 100 символов.</span><span class="sxs-lookup"><span data-stu-id="37cd3-p153">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="37cd3-819">function</span><span class="sxs-lookup"><span data-stu-id="37cd3-819">function</span></span> | <span data-ttu-id="37cd3-820">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="37cd3-820">&lt;optional&gt;</span></span> | <span data-ttu-id="37cd3-821">По завершении работы метода функция, переданная параметру `callback`, вызывается с помощью одного параметра `asyncResult`, представляющего собой объект [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="37cd3-821">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="37cd3-822">Требования</span><span class="sxs-lookup"><span data-stu-id="37cd3-822">Requirements</span></span>

|<span data-ttu-id="37cd3-823">Требование</span><span class="sxs-lookup"><span data-stu-id="37cd3-823">Requirement</span></span>| <span data-ttu-id="37cd3-824">Значение</span><span class="sxs-lookup"><span data-stu-id="37cd3-824">Value</span></span>|
|---|---|
|[<span data-ttu-id="37cd3-825">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="37cd3-825">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="37cd3-826">1.0</span><span class="sxs-lookup"><span data-stu-id="37cd3-826">1.0</span></span>|
|[<span data-ttu-id="37cd3-827">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="37cd3-827">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="37cd3-828">ReadItem</span><span class="sxs-lookup"><span data-stu-id="37cd3-828">ReadItem</span></span>|
|[<span data-ttu-id="37cd3-829">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="37cd3-829">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="37cd3-830">Чтение</span><span class="sxs-lookup"><span data-stu-id="37cd3-830">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="37cd3-831">Примеры</span><span class="sxs-lookup"><span data-stu-id="37cd3-831">Examples</span></span>

<span data-ttu-id="37cd3-832">Приведенный ниже код передает строку в функцию `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="37cd3-832">The following code passes a string to the `displayReplyForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="37cd3-833">Ответ с пустым текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="37cd3-833">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="37cd3-834">Ответ только с текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="37cd3-834">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="37cd3-835">Ответ с текстом сообщения и вложенным файлом.</span><span class="sxs-lookup"><span data-stu-id="37cd3-835">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="37cd3-836">Ответ с текстом сообщения и вложенным элементом.</span><span class="sxs-lookup"><span data-stu-id="37cd3-836">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="37cd3-837">Ответ с текстом сообщения, вложенным файлом, вложенным элементом и обратным вызовом.</span><span class="sxs-lookup"><span data-stu-id="37cd3-837">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlook16officeentities"></a><span data-ttu-id="37cd3-838">getEntities() → {[Entities](/javascript/api/outlook_1_6/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="37cd3-838">getEntities() → {[Entities](/javascript/api/outlook_1_6/office.entities)}</span></span>

<span data-ttu-id="37cd3-839">Получает сущности, обнаруженные в теле выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="37cd3-839">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="37cd3-840">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="37cd3-840">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="37cd3-841">Требования</span><span class="sxs-lookup"><span data-stu-id="37cd3-841">Requirements</span></span>

|<span data-ttu-id="37cd3-842">Требование</span><span class="sxs-lookup"><span data-stu-id="37cd3-842">Requirement</span></span>| <span data-ttu-id="37cd3-843">Значение</span><span class="sxs-lookup"><span data-stu-id="37cd3-843">Value</span></span>|
|---|---|
|[<span data-ttu-id="37cd3-844">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="37cd3-844">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="37cd3-845">1.0</span><span class="sxs-lookup"><span data-stu-id="37cd3-845">1.0</span></span>|
|[<span data-ttu-id="37cd3-846">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="37cd3-846">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="37cd3-847">ReadItem</span><span class="sxs-lookup"><span data-stu-id="37cd3-847">ReadItem</span></span>|
|[<span data-ttu-id="37cd3-848">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="37cd3-848">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="37cd3-849">Чтение</span><span class="sxs-lookup"><span data-stu-id="37cd3-849">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="37cd3-850">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="37cd3-850">Returns:</span></span>

<span data-ttu-id="37cd3-851">Тип: [Entities](/javascript/api/outlook_1_6/office.entities)</span><span class="sxs-lookup"><span data-stu-id="37cd3-851">Type: [Entities](/javascript/api/outlook_1_6/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="37cd3-852">Пример</span><span class="sxs-lookup"><span data-stu-id="37cd3-852">Example</span></span>

<span data-ttu-id="37cd3-853">Ниже приведен пример получения доступа к сущностям контактов в теле текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="37cd3-853">The following example accesses the contacts entities in the current item's body.</span></span>

```js
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook16officecontactmeetingsuggestionjavascriptapioutlook16officemeetingsuggestionphonenumberjavascriptapioutlook16officephonenumbertasksuggestionjavascriptapioutlook16officetasksuggestion"></a><span data-ttu-id="37cd3-854">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="37cd3-854">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))>}</span></span>

<span data-ttu-id="37cd3-855">Получает массив всех сущностей указанного типа, обнаруженных в теле выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="37cd3-855">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="37cd3-856">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="37cd3-856">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="37cd3-857">Параметры</span><span class="sxs-lookup"><span data-stu-id="37cd3-857">Parameters:</span></span>

|<span data-ttu-id="37cd3-858">Имя</span><span class="sxs-lookup"><span data-stu-id="37cd3-858">Name</span></span>| <span data-ttu-id="37cd3-859">Тип</span><span class="sxs-lookup"><span data-stu-id="37cd3-859">Type</span></span>| <span data-ttu-id="37cd3-860">Описание</span><span class="sxs-lookup"><span data-stu-id="37cd3-860">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="37cd3-861">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="37cd3-861">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook_1_6/office.mailboxenums.entitytype)|<span data-ttu-id="37cd3-862">Одно из значений перечисления EntityType.</span><span class="sxs-lookup"><span data-stu-id="37cd3-862">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="37cd3-863">Требования</span><span class="sxs-lookup"><span data-stu-id="37cd3-863">Requirements</span></span>

|<span data-ttu-id="37cd3-864">Требование</span><span class="sxs-lookup"><span data-stu-id="37cd3-864">Requirement</span></span>| <span data-ttu-id="37cd3-865">Значение</span><span class="sxs-lookup"><span data-stu-id="37cd3-865">Value</span></span>|
|---|---|
|[<span data-ttu-id="37cd3-866">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="37cd3-866">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="37cd3-867">1.0</span><span class="sxs-lookup"><span data-stu-id="37cd3-867">1.0</span></span>|
|[<span data-ttu-id="37cd3-868">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="37cd3-868">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="37cd3-869">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="37cd3-869">Restricted</span></span>|
|[<span data-ttu-id="37cd3-870">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="37cd3-870">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="37cd3-871">Чтение</span><span class="sxs-lookup"><span data-stu-id="37cd3-871">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="37cd3-872">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="37cd3-872">Returns:</span></span>

<span data-ttu-id="37cd3-873">Если значение, переданное в `entityType`, не является допустимым членом перечисления `EntityType`, метод возвращает значение NULL.</span><span class="sxs-lookup"><span data-stu-id="37cd3-873">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="37cd3-874">Если в теле элемента отсутствуют сущности указанного типа, метод возвращает пустой массив.</span><span class="sxs-lookup"><span data-stu-id="37cd3-874">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="37cd3-875">В противном случае тип объектов в возвращаемом массиве зависит от типа сущности, запрошенной в параметре `entityType`.</span><span class="sxs-lookup"><span data-stu-id="37cd3-875">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="37cd3-876">Хотя минимальный уровень разрешений для использования этого метода — **Restricted**, для некоторых типов сущностей требуется доступ на уровне **ReadItem**, как указано в приведенной ниже таблице.</span><span class="sxs-lookup"><span data-stu-id="37cd3-876">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="37cd3-877">Значение параметра `entityType`</span><span class="sxs-lookup"><span data-stu-id="37cd3-877">Value of `entityType`</span></span> | <span data-ttu-id="37cd3-878">Тип объектов в возвращаемом массиве</span><span class="sxs-lookup"><span data-stu-id="37cd3-878">Type of objects in returned array</span></span> | <span data-ttu-id="37cd3-879">Необходимый уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="37cd3-879">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="37cd3-880">String</span><span class="sxs-lookup"><span data-stu-id="37cd3-880">String</span></span> | <span data-ttu-id="37cd3-881">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="37cd3-881">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="37cd3-882">Contact</span><span class="sxs-lookup"><span data-stu-id="37cd3-882">Contact</span></span> | <span data-ttu-id="37cd3-883">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="37cd3-883">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="37cd3-884">String</span><span class="sxs-lookup"><span data-stu-id="37cd3-884">String</span></span> | <span data-ttu-id="37cd3-885">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="37cd3-885">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="37cd3-886">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="37cd3-886">MeetingSuggestion</span></span> | <span data-ttu-id="37cd3-887">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="37cd3-887">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="37cd3-888">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="37cd3-888">PhoneNumber</span></span> | <span data-ttu-id="37cd3-889">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="37cd3-889">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="37cd3-890">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="37cd3-890">TaskSuggestion</span></span> | <span data-ttu-id="37cd3-891">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="37cd3-891">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="37cd3-892">String</span><span class="sxs-lookup"><span data-stu-id="37cd3-892">String</span></span> | <span data-ttu-id="37cd3-893">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="37cd3-893">**Restricted**</span></span> |

<span data-ttu-id="37cd3-894">Тип: Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="37cd3-894">Type: Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="37cd3-895">Пример</span><span class="sxs-lookup"><span data-stu-id="37cd3-895">Example</span></span>

<span data-ttu-id="37cd3-896">В примере ниже показано, как получить доступ к массиву строк, которые представляют собой почтовые адреса в теле текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="37cd3-896">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook16officecontactmeetingsuggestionjavascriptapioutlook16officemeetingsuggestionphonenumberjavascriptapioutlook16officephonenumbertasksuggestionjavascriptapioutlook16officetasksuggestion"></a><span data-ttu-id="37cd3-897">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="37cd3-897">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))>}</span></span>

<span data-ttu-id="37cd3-898">Возвращает известные сущности в выбранном элементе, которые проходят через именованный фильтр, определяемый в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="37cd3-898">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="37cd3-899">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="37cd3-899">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="37cd3-900">Метод `getFilteredEntitiesByName` возвращает сущности, соответствующие регулярному выражению, которое определяется в элементе правила [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) в XML-файле манифеста, с использованием указанного значения элемента `FilterName`.</span><span class="sxs-lookup"><span data-stu-id="37cd3-900">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="37cd3-901">Параметры</span><span class="sxs-lookup"><span data-stu-id="37cd3-901">Parameters:</span></span>

|<span data-ttu-id="37cd3-902">Имя</span><span class="sxs-lookup"><span data-stu-id="37cd3-902">Name</span></span>| <span data-ttu-id="37cd3-903">Тип</span><span class="sxs-lookup"><span data-stu-id="37cd3-903">Type</span></span>| <span data-ttu-id="37cd3-904">Описание</span><span class="sxs-lookup"><span data-stu-id="37cd3-904">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="37cd3-905">String</span><span class="sxs-lookup"><span data-stu-id="37cd3-905">String</span></span>|<span data-ttu-id="37cd3-906">Имя элемента правила `ItemHasKnownEntity`, определяющее соответствующий фильтр.</span><span class="sxs-lookup"><span data-stu-id="37cd3-906">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="37cd3-907">Требования</span><span class="sxs-lookup"><span data-stu-id="37cd3-907">Requirements</span></span>

|<span data-ttu-id="37cd3-908">Требование</span><span class="sxs-lookup"><span data-stu-id="37cd3-908">Requirement</span></span>| <span data-ttu-id="37cd3-909">Значение</span><span class="sxs-lookup"><span data-stu-id="37cd3-909">Value</span></span>|
|---|---|
|[<span data-ttu-id="37cd3-910">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="37cd3-910">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="37cd3-911">1.0</span><span class="sxs-lookup"><span data-stu-id="37cd3-911">1.0</span></span>|
|[<span data-ttu-id="37cd3-912">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="37cd3-912">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="37cd3-913">ReadItem</span><span class="sxs-lookup"><span data-stu-id="37cd3-913">ReadItem</span></span>|
|[<span data-ttu-id="37cd3-914">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="37cd3-914">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="37cd3-915">Чтение</span><span class="sxs-lookup"><span data-stu-id="37cd3-915">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="37cd3-916">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="37cd3-916">Returns:</span></span>

<span data-ttu-id="37cd3-p155">Если в манифесте нет элемента `ItemHasKnownEntity` со значением `FilterName`, соответствующим параметру `name`, метод возвращает `null`. Если параметр `name` соответствует элементу `ItemHasKnownEntity` в манифесте, но при этом в текущем элементе нет соответствующих сущностей, метод возвращает пустой массив.</span><span class="sxs-lookup"><span data-stu-id="37cd3-p155">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="37cd3-919">Тип: Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="37cd3-919">Type: Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))></span></span>

#### <a name="getregexmatches--object"></a><span data-ttu-id="37cd3-920">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="37cd3-920">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="37cd3-921">Возвращает строковые значения в выбранном элементе, которые соответствуют регулярным выражениям, определенным в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="37cd3-921">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="37cd3-922">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="37cd3-922">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="37cd3-p156">Метод `getRegExMatches` возвращает строки, соответствующие регулярному выражению, которое определяется в каждом элементе правила `ItemHasRegularExpressionMatch` или `ItemHasKnownEntity` в XML-файле манифеста. Для правила `ItemHasRegularExpressionMatch` соответствующую строку должно содержать свойство элемента, указанного этим правилом. Простой тип `PropertyName` определяет поддерживаемые свойства.</span><span class="sxs-lookup"><span data-stu-id="37cd3-p156">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="37cd3-926">Например, рассмотрим манифест надстройки, который содержит указанный ниже элемент `Rule`.</span><span class="sxs-lookup"><span data-stu-id="37cd3-926">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="37cd3-927">Объект, возвращаемый методом `getRegExMatches`, будет содержать два свойства: `fruits` и `veggies`.</span><span class="sxs-lookup"><span data-stu-id="37cd3-927">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="37cd3-p157">Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты. Лучше используйте метод [`Body.getAsync`](/javascript/api/outlook_1_6/office.body#getasync-coerciontype--options--callback-) для этого.</span><span class="sxs-lookup"><span data-stu-id="37cd3-p157">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook_1_6/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="37cd3-931">Requirements</span><span class="sxs-lookup"><span data-stu-id="37cd3-931">Requirements</span></span>

|<span data-ttu-id="37cd3-932">Требование</span><span class="sxs-lookup"><span data-stu-id="37cd3-932">Requirement</span></span>| <span data-ttu-id="37cd3-933">Значение</span><span class="sxs-lookup"><span data-stu-id="37cd3-933">Value</span></span>|
|---|---|
|[<span data-ttu-id="37cd3-934">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="37cd3-934">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="37cd3-935">1.0</span><span class="sxs-lookup"><span data-stu-id="37cd3-935">1.0</span></span>|
|[<span data-ttu-id="37cd3-936">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="37cd3-936">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="37cd3-937">ReadItem</span><span class="sxs-lookup"><span data-stu-id="37cd3-937">ReadItem</span></span>|
|[<span data-ttu-id="37cd3-938">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="37cd3-938">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="37cd3-939">Чтение</span><span class="sxs-lookup"><span data-stu-id="37cd3-939">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="37cd3-940">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="37cd3-940">Returns:</span></span>

<span data-ttu-id="37cd3-p158">Объект, содержащий массив строк, которые соответствуют регулярным выражениям, определяемым в XML-файле манифеста. Имя каждого массива равно соответствующему значению атрибута `RegExName` подходящего правила `ItemHasRegularExpressionMatch` или атрибута `FilterName` соответствующего правила `ItemHasKnownEntity`.</span><span class="sxs-lookup"><span data-stu-id="37cd3-p158">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="37cd3-943">

<dt>Тип</dt>

</span><span class="sxs-lookup"><span data-stu-id="37cd3-943">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="37cd3-944">Object</span><span class="sxs-lookup"><span data-stu-id="37cd3-944">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="37cd3-945">Пример</span><span class="sxs-lookup"><span data-stu-id="37cd3-945">Example</span></span>

<span data-ttu-id="37cd3-946">В приведенном ниже примере показано, как получить доступ к массиву совпадений с элементами `fruits` и `veggies` правил активации регулярных выражений, указанными в манифесте.</span><span class="sxs-lookup"><span data-stu-id="37cd3-946">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```js
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veges = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="37cd3-947">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="37cd3-947">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="37cd3-948">Возвращает строковые значения в выбранном элементе, которые соответствуют именованному регулярному выражению, определенному в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="37cd3-948">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="37cd3-949">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="37cd3-949">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="37cd3-950">Метод `getRegExMatchesByName` возвращает строки, соответствующие регулярному выражению, которое определяется в элементе правила `ItemHasRegularExpressionMatch` в XML-файле манифеста, с использованием указанного значения элемента `RegExName`.</span><span class="sxs-lookup"><span data-stu-id="37cd3-950">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="37cd3-p159">Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты.</span><span class="sxs-lookup"><span data-stu-id="37cd3-p159">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="37cd3-953">Параметры</span><span class="sxs-lookup"><span data-stu-id="37cd3-953">Parameters:</span></span>

|<span data-ttu-id="37cd3-954">Имя</span><span class="sxs-lookup"><span data-stu-id="37cd3-954">Name</span></span>| <span data-ttu-id="37cd3-955">Тип</span><span class="sxs-lookup"><span data-stu-id="37cd3-955">Type</span></span>| <span data-ttu-id="37cd3-956">Описание</span><span class="sxs-lookup"><span data-stu-id="37cd3-956">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="37cd3-957">String</span><span class="sxs-lookup"><span data-stu-id="37cd3-957">String</span></span>|<span data-ttu-id="37cd3-958">Имя элемента правила `ItemHasRegularExpressionMatch`, определяющее соответствующий фильтр.</span><span class="sxs-lookup"><span data-stu-id="37cd3-958">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="37cd3-959">Требования</span><span class="sxs-lookup"><span data-stu-id="37cd3-959">Requirements</span></span>

|<span data-ttu-id="37cd3-960">Требование</span><span class="sxs-lookup"><span data-stu-id="37cd3-960">Requirement</span></span>| <span data-ttu-id="37cd3-961">Значение</span><span class="sxs-lookup"><span data-stu-id="37cd3-961">Value</span></span>|
|---|---|
|[<span data-ttu-id="37cd3-962">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="37cd3-962">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="37cd3-963">1.0</span><span class="sxs-lookup"><span data-stu-id="37cd3-963">1.0</span></span>|
|[<span data-ttu-id="37cd3-964">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="37cd3-964">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="37cd3-965">ReadItem</span><span class="sxs-lookup"><span data-stu-id="37cd3-965">ReadItem</span></span>|
|[<span data-ttu-id="37cd3-966">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="37cd3-966">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="37cd3-967">Чтение</span><span class="sxs-lookup"><span data-stu-id="37cd3-967">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="37cd3-968">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="37cd3-968">Returns:</span></span>

<span data-ttu-id="37cd3-969">Массив строк, соответствующих регулярному выражению, определяемому в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="37cd3-969">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="37cd3-970">

<dt>Тип</dt>

</span><span class="sxs-lookup"><span data-stu-id="37cd3-970">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="37cd3-971">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="37cd3-971">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="37cd3-972">Пример</span><span class="sxs-lookup"><span data-stu-id="37cd3-972">Example</span></span>

```js
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="37cd3-973">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="37cd3-973">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="37cd3-974">Асинхронно возвращает данные, выбранные в теме или тексте сообщения.</span><span class="sxs-lookup"><span data-stu-id="37cd3-974">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="37cd3-p160">Если выделенный фрагмент отсутствует, но курсор находится в тексте или теме, метод возвращает значение NULL для выбранных данных. Если выбраны не текст и не тема, метод возвращает ошибку `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="37cd3-p160">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="37cd3-977">Параметры</span><span class="sxs-lookup"><span data-stu-id="37cd3-977">Parameters:</span></span>

|<span data-ttu-id="37cd3-978">Имя</span><span class="sxs-lookup"><span data-stu-id="37cd3-978">Name</span></span>| <span data-ttu-id="37cd3-979">Тип</span><span class="sxs-lookup"><span data-stu-id="37cd3-979">Type</span></span>| <span data-ttu-id="37cd3-980">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="37cd3-980">Attributes</span></span>| <span data-ttu-id="37cd3-981">Описание</span><span class="sxs-lookup"><span data-stu-id="37cd3-981">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="37cd3-982">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="37cd3-982">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="37cd3-p161">Запрашивает формат данных. Если задано значение Text, метод возвращает обычный текст как строку, удаляя все имеющиеся HTML-теги. Если задано значение HTML, метод возвращает выделенный текст (обычный текст или HTML).</span><span class="sxs-lookup"><span data-stu-id="37cd3-p161">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="37cd3-986">Object</span><span class="sxs-lookup"><span data-stu-id="37cd3-986">Object</span></span>| <span data-ttu-id="37cd3-987">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="37cd3-987">&lt;optional&gt;</span></span>|<span data-ttu-id="37cd3-988">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="37cd3-988">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="37cd3-989">Object</span><span class="sxs-lookup"><span data-stu-id="37cd3-989">Object</span></span>| <span data-ttu-id="37cd3-990">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="37cd3-990">&lt;optional&gt;</span></span>|<span data-ttu-id="37cd3-991">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="37cd3-991">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="37cd3-992">функция</span><span class="sxs-lookup"><span data-stu-id="37cd3-992">function</span></span>||<span data-ttu-id="37cd3-993">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="37cd3-993">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="37cd3-994">Чтобы получить доступ к выбранным данным из метода обратного вызова, вызовите `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="37cd3-994">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="37cd3-995">Для доступа к исходному свойству, представляющему собой источник выбранных данных, вызовите параметр `asyncResult.value.sourceProperty`, который может иметь значение `body` или `subject`.</span><span class="sxs-lookup"><span data-stu-id="37cd3-995">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="37cd3-996">Требования</span><span class="sxs-lookup"><span data-stu-id="37cd3-996">Requirements</span></span>

|<span data-ttu-id="37cd3-997">Требование</span><span class="sxs-lookup"><span data-stu-id="37cd3-997">Requirement</span></span>| <span data-ttu-id="37cd3-998">Значение</span><span class="sxs-lookup"><span data-stu-id="37cd3-998">Value</span></span>|
|---|---|
|[<span data-ttu-id="37cd3-999">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="37cd3-999">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="37cd3-1000">1.2</span><span class="sxs-lookup"><span data-stu-id="37cd3-1000">1.2</span></span>|
|[<span data-ttu-id="37cd3-1001">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="37cd3-1001">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="37cd3-1002">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="37cd3-1002">ReadWriteItem</span></span>|
|[<span data-ttu-id="37cd3-1003">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="37cd3-1003">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="37cd3-1004">Создание</span><span class="sxs-lookup"><span data-stu-id="37cd3-1004">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="37cd3-1005">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="37cd3-1005">Returns:</span></span>

<span data-ttu-id="37cd3-1006">Выбранные данные в виде строки с форматом, определенным в параметре `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="37cd3-1006">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="37cd3-1007">

<dt>Тип</dt>

</span><span class="sxs-lookup"><span data-stu-id="37cd3-1007">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="37cd3-1008">String</span><span class="sxs-lookup"><span data-stu-id="37cd3-1008">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="37cd3-1009">Пример</span><span class="sxs-lookup"><span data-stu-id="37cd3-1009">Example</span></span>

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

#### <a name="getselectedentities--entitiesjavascriptapioutlook16officeentities"></a><span data-ttu-id="37cd3-1010">getSelectedEntities() → {[Entities](/javascript/api/outlook_1_6/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="37cd3-1010">getSelectedEntities() → {[Entities](/javascript/api/outlook_1_6/office.entities)}</span></span>

<span data-ttu-id="37cd3-p163">Возвращает сущности, найденные в выделенном совпадении, выбранном пользователем. Выделенные совпадения применяются к [контекстным надстройкам](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="37cd3-p163">Gets the entities found in a highlighted match a user has selected. Highlighted matches apply to [contextual add-ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="37cd3-1013">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="37cd3-1013">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="37cd3-1014">Требования</span><span class="sxs-lookup"><span data-stu-id="37cd3-1014">Requirements</span></span>

|<span data-ttu-id="37cd3-1015">Требование</span><span class="sxs-lookup"><span data-stu-id="37cd3-1015">Requirement</span></span>| <span data-ttu-id="37cd3-1016">Значение</span><span class="sxs-lookup"><span data-stu-id="37cd3-1016">Value</span></span>|
|---|---|
|[<span data-ttu-id="37cd3-1017">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="37cd3-1017">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="37cd3-1018">1.6</span><span class="sxs-lookup"><span data-stu-id="37cd3-1018">1.6</span></span> |
|[<span data-ttu-id="37cd3-1019">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="37cd3-1019">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="37cd3-1020">ReadItem</span><span class="sxs-lookup"><span data-stu-id="37cd3-1020">ReadItem</span></span>|
|[<span data-ttu-id="37cd3-1021">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="37cd3-1021">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="37cd3-1022">Чтение</span><span class="sxs-lookup"><span data-stu-id="37cd3-1022">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="37cd3-1023">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="37cd3-1023">Returns:</span></span>

<span data-ttu-id="37cd3-1024">Тип: [Entities](/javascript/api/outlook_1_6/office.entities)</span><span class="sxs-lookup"><span data-stu-id="37cd3-1024">Type: [Entities](/javascript/api/outlook_1_6/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="37cd3-1025">Пример</span><span class="sxs-lookup"><span data-stu-id="37cd3-1025">Example</span></span>

<span data-ttu-id="37cd3-1026">В приведенном ниже примере показано, как получить доступ к сущностям адресов в выделенном совпадении, выбранном пользователем.</span><span class="sxs-lookup"><span data-stu-id="37cd3-1026">The following example accesses the addresses entities in the highlighted match selected by the user.</span></span>

```js
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

#### <a name="getselectedregexmatches--object"></a><span data-ttu-id="37cd3-1027">getSelectedRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="37cd3-1027">getSelectedRegExMatches() → {Object}</span></span>

<span data-ttu-id="37cd3-p164">Возвращает строковые значения в выделенном совпадении, которые соответствуют регулярным выражениям, определенным в XML-файле манифеста. Выделенные совпадения применяются к [контекстным надстройкам](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="37cd3-p164">Returns string values in a highlighted match that match the regular expressions defined in the manifest XML file. Highlighted matches apply to [contextual add-ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="37cd3-1030">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="37cd3-1030">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="37cd3-p165">Метод `getSelectedRegExMatches` возвращает строки, соответствующие регулярному выражению, которое определяется в каждом элементе правила `ItemHasRegularExpressionMatch` или `ItemHasKnownEntity` в XML-файле манифеста. Для правила `ItemHasRegularExpressionMatch` соответствующую строку должно содержать свойство элемента, указанного этим правилом. Простой тип `PropertyName` определяет поддерживаемые свойства.</span><span class="sxs-lookup"><span data-stu-id="37cd3-p165">The `getSelectedRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="37cd3-1034">Например, рассмотрим манифест надстройки, который содержит указанный ниже элемент `Rule`.</span><span class="sxs-lookup"><span data-stu-id="37cd3-1034">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="37cd3-1035">Объект, возвращаемый методом `getRegExMatches`, будет содержать два свойства: `fruits` и `veggies`.</span><span class="sxs-lookup"><span data-stu-id="37cd3-1035">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="37cd3-p166">Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты. Лучше используйте метод [`Body.getAsync`](/javascript/api/outlook_1_6/office.body#getasync-coerciontype--options--callback-) для этого.</span><span class="sxs-lookup"><span data-stu-id="37cd3-p166">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook_1_6/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="37cd3-1039">Requirements</span><span class="sxs-lookup"><span data-stu-id="37cd3-1039">Requirements</span></span>

|<span data-ttu-id="37cd3-1040">Требование</span><span class="sxs-lookup"><span data-stu-id="37cd3-1040">Requirement</span></span>| <span data-ttu-id="37cd3-1041">Значение</span><span class="sxs-lookup"><span data-stu-id="37cd3-1041">Value</span></span>|
|---|---|
|[<span data-ttu-id="37cd3-1042">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="37cd3-1042">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="37cd3-1043">1.6</span><span class="sxs-lookup"><span data-stu-id="37cd3-1043">1.6</span></span> |
|[<span data-ttu-id="37cd3-1044">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="37cd3-1044">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="37cd3-1045">ReadItem</span><span class="sxs-lookup"><span data-stu-id="37cd3-1045">ReadItem</span></span>|
|[<span data-ttu-id="37cd3-1046">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="37cd3-1046">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="37cd3-1047">Чтение</span><span class="sxs-lookup"><span data-stu-id="37cd3-1047">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="37cd3-1048">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="37cd3-1048">Returns:</span></span>

<span data-ttu-id="37cd3-p167">Объект, содержащий массив строк, которые соответствуют регулярным выражениям, определяемым в XML-файле манифеста. Имя каждого массива равно соответствующему значению атрибута `RegExName` подходящего правила `ItemHasRegularExpressionMatch` или атрибута `FilterName` соответствующего правила `ItemHasKnownEntity`.</span><span class="sxs-lookup"><span data-stu-id="37cd3-p167">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

##### <a name="example"></a><span data-ttu-id="37cd3-1051">Пример</span><span class="sxs-lookup"><span data-stu-id="37cd3-1051">Example</span></span>

<span data-ttu-id="37cd3-1052">В приведенном ниже примере показано, как получить доступ к массиву совпадений с элементами `fruits` и `veggies` правил активации регулярных выражений, указанными в манифесте.</span><span class="sxs-lookup"><span data-stu-id="37cd3-1052">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```javascript
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="37cd3-1053">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="37cd3-1053">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="37cd3-1054">Асинхронно загружает настраиваемые свойства для надстройки для выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="37cd3-1054">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="37cd3-p168">Настраиваемые свойства сохраняются в виде пар "ключ-значение" для каждого приложения и каждого элемента. Этот метод возвращает объект `CustomProperties` при обратном вызове, который предоставляет методы для доступа к настраиваемым свойствам, характерным для текущего элемента и текущей надстройки. Настраиваемые свойства не шифруются для элемента, поэтому этот способ хранения не является безопасным.</span><span class="sxs-lookup"><span data-stu-id="37cd3-p168">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="37cd3-1058">Параметры</span><span class="sxs-lookup"><span data-stu-id="37cd3-1058">Parameters:</span></span>

|<span data-ttu-id="37cd3-1059">Имя</span><span class="sxs-lookup"><span data-stu-id="37cd3-1059">Name</span></span>| <span data-ttu-id="37cd3-1060">Тип</span><span class="sxs-lookup"><span data-stu-id="37cd3-1060">Type</span></span>| <span data-ttu-id="37cd3-1061">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="37cd3-1061">Attributes</span></span>| <span data-ttu-id="37cd3-1062">Описание</span><span class="sxs-lookup"><span data-stu-id="37cd3-1062">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="37cd3-1063">function</span><span class="sxs-lookup"><span data-stu-id="37cd3-1063">function</span></span>||<span data-ttu-id="37cd3-1064">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="37cd3-1064">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="37cd3-1065">Настраиваемые свойства предоставляются в виде объекта [`CustomProperties`](/javascript/api/outlook_1_6/office.customproperties) в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="37cd3-1065">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook_1_6/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="37cd3-1066">Этот объект позволяет получить, задать и удалить настраиваемые свойства для элемента, а также сохранить изменения, внесенные в набор настраиваемых свойств, на сервере.</span><span class="sxs-lookup"><span data-stu-id="37cd3-1066">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="37cd3-1067">Object</span><span class="sxs-lookup"><span data-stu-id="37cd3-1067">Object</span></span>| <span data-ttu-id="37cd3-1068">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="37cd3-1068">&lt;optional&gt;</span></span>|<span data-ttu-id="37cd3-1069">Разработчики могут указать любой объект, к которому необходимо получить доступ, в функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="37cd3-1069">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="37cd3-1070">Доступ к этому объекту можно получить с помощью свойства `asyncResult.asyncContext` в функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="37cd3-1070">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="37cd3-1071">Требования</span><span class="sxs-lookup"><span data-stu-id="37cd3-1071">Requirements</span></span>

|<span data-ttu-id="37cd3-1072">Требование</span><span class="sxs-lookup"><span data-stu-id="37cd3-1072">Requirement</span></span>| <span data-ttu-id="37cd3-1073">Значение</span><span class="sxs-lookup"><span data-stu-id="37cd3-1073">Value</span></span>|
|---|---|
|[<span data-ttu-id="37cd3-1074">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="37cd3-1074">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="37cd3-1075">1.0</span><span class="sxs-lookup"><span data-stu-id="37cd3-1075">1.0</span></span>|
|[<span data-ttu-id="37cd3-1076">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="37cd3-1076">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="37cd3-1077">ReadItem</span><span class="sxs-lookup"><span data-stu-id="37cd3-1077">ReadItem</span></span>|
|[<span data-ttu-id="37cd3-1078">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="37cd3-1078">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="37cd3-1079">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="37cd3-1079">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="37cd3-1080">Пример</span><span class="sxs-lookup"><span data-stu-id="37cd3-1080">Example</span></span>

<span data-ttu-id="37cd3-p171">Приведенный ниже пример кода показывает, как асинхронно загружать настраиваемые свойства, характерные для текущего элемента, с помощью метода `loadCustomPropertiesAsync`. Этот пример также показывает, как сохранять эти свойства на сервере с помощью метода `CustomProperties.saveAsync`. После загрузки настраиваемых свойств в этом примере кода метод `CustomProperties.get` используется для считывания настраиваемого свойства `myProp`, метод `CustomProperties.set` — для записи настраиваемого свойства `otherProp`, а метод `saveAsync` — для сохранения настраиваемых свойств.</span><span class="sxs-lookup"><span data-stu-id="37cd3-p171">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="37cd3-1084">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="37cd3-1084">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="37cd3-1085">Удаляет вложение из сообщения или встречи.</span><span class="sxs-lookup"><span data-stu-id="37cd3-1085">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="37cd3-p172">Метод `removeAttachmentAsync` удаляет из элемента вложение с указанным идентификатором. Идентификатор вложения рекомендуется использовать для удаления вложения, только если оно добавлено тем же почтовым приложением в ходе текущего сеанса. В Outlook Web App и Outlook Web App для устройств идентификатор вложения действителен только в рамках одного сеанса. Сеанс завершается, когда пользователь закрывает приложение или начинает создавать элемент во встроенной форме, а затем переходит из формы в отдельное окно.</span><span class="sxs-lookup"><span data-stu-id="37cd3-p172">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="37cd3-1090">Параметры</span><span class="sxs-lookup"><span data-stu-id="37cd3-1090">Parameters:</span></span>

|<span data-ttu-id="37cd3-1091">Имя</span><span class="sxs-lookup"><span data-stu-id="37cd3-1091">Name</span></span>| <span data-ttu-id="37cd3-1092">Тип</span><span class="sxs-lookup"><span data-stu-id="37cd3-1092">Type</span></span>| <span data-ttu-id="37cd3-1093">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="37cd3-1093">Attributes</span></span>| <span data-ttu-id="37cd3-1094">Описание</span><span class="sxs-lookup"><span data-stu-id="37cd3-1094">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="37cd3-1095">String</span><span class="sxs-lookup"><span data-stu-id="37cd3-1095">String</span></span>||<span data-ttu-id="37cd3-1096">Идентификатор удаляемого вложения.</span><span class="sxs-lookup"><span data-stu-id="37cd3-1096">The identifier of the attachment to remove.</span></span>|
|`options`| <span data-ttu-id="37cd3-1097">Object</span><span class="sxs-lookup"><span data-stu-id="37cd3-1097">Object</span></span>| <span data-ttu-id="37cd3-1098">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="37cd3-1098">&lt;optional&gt;</span></span>|<span data-ttu-id="37cd3-1099">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="37cd3-1099">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="37cd3-1100">Object</span><span class="sxs-lookup"><span data-stu-id="37cd3-1100">Object</span></span>| <span data-ttu-id="37cd3-1101">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="37cd3-1101">&lt;optional&gt;</span></span>|<span data-ttu-id="37cd3-1102">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="37cd3-1102">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="37cd3-1103">функция</span><span class="sxs-lookup"><span data-stu-id="37cd3-1103">function</span></span>| <span data-ttu-id="37cd3-1104">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="37cd3-1104">&lt;optional&gt;</span></span>|<span data-ttu-id="37cd3-1105">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="37cd3-1105">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="37cd3-1106">Если удалить вложение не удается, свойство `asyncResult.error` содержит код ошибки с указанием ее причины.</span><span class="sxs-lookup"><span data-stu-id="37cd3-1106">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="37cd3-1107">Ошибки</span><span class="sxs-lookup"><span data-stu-id="37cd3-1107">Errors</span></span>

| <span data-ttu-id="37cd3-1108">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="37cd3-1108">Error code</span></span> | <span data-ttu-id="37cd3-1109">Описание</span><span class="sxs-lookup"><span data-stu-id="37cd3-1109">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="37cd3-1110">Идентификатор вложения не существует.</span><span class="sxs-lookup"><span data-stu-id="37cd3-1110">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="37cd3-1111">Требования</span><span class="sxs-lookup"><span data-stu-id="37cd3-1111">Requirements</span></span>

|<span data-ttu-id="37cd3-1112">Требование</span><span class="sxs-lookup"><span data-stu-id="37cd3-1112">Requirement</span></span>| <span data-ttu-id="37cd3-1113">Значение</span><span class="sxs-lookup"><span data-stu-id="37cd3-1113">Value</span></span>|
|---|---|
|[<span data-ttu-id="37cd3-1114">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="37cd3-1114">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="37cd3-1115">1.1</span><span class="sxs-lookup"><span data-stu-id="37cd3-1115">1.1</span></span>|
|[<span data-ttu-id="37cd3-1116">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="37cd3-1116">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="37cd3-1117">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="37cd3-1117">ReadWriteItem</span></span>|
|[<span data-ttu-id="37cd3-1118">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="37cd3-1118">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="37cd3-1119">Создание</span><span class="sxs-lookup"><span data-stu-id="37cd3-1119">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="37cd3-1120">Пример</span><span class="sxs-lookup"><span data-stu-id="37cd3-1120">Example</span></span>

<span data-ttu-id="37cd3-1121">Указанный ниже код удаляет вложение с идентификатором "0".</span><span class="sxs-lookup"><span data-stu-id="37cd3-1121">The following code removes an attachment with an identifier of '0'.</span></span>

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

####  <a name="saveasyncoptions-callback"></a><span data-ttu-id="37cd3-1122">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="37cd3-1122">saveAsync([options], callback)</span></span>

<span data-ttu-id="37cd3-1123">Асинхронно сохраняет элемент.</span><span class="sxs-lookup"><span data-stu-id="37cd3-1123">Asynchronously saves an item.</span></span>

<span data-ttu-id="37cd3-p173">При вызове этот метод сохраняет текущее сообщение в виде черновика и возвращает идентификатор элемента с помощью метода обратного вызова. В Outlook Web App или интерактивном режиме Outlook этот элемент сохраняется на сервере. В Outlook в режиме кэширования этот элемент сохраняется в локальном кэше.</span><span class="sxs-lookup"><span data-stu-id="37cd3-p173">When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="37cd3-1127">Если в режиме создания надстройка вызывает для элемента метод `saveAsync`, чтобы получить параметр `itemId` для использования с EWS или REST API, необходимо помнить, что синхронизация элемента с сервером может занять много времени, если Outlook работает в режиме кэширования данных.</span><span class="sxs-lookup"><span data-stu-id="37cd3-1127">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="37cd3-1128">До окончания синхронизации элемента применение параметра `itemId` будет приводить к ошибке.</span><span class="sxs-lookup"><span data-stu-id="37cd3-1128">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="37cd3-p175">Если метод `saveAsync` вызывается для встречи в режиме создания, она сохраняется как обычная встреча в календаре пользователя, а не как черновик. При сохранении новой встречи приглашения не отправляются. При сохранении существующей встречи уведомления отправляются добавленным или удаленным участникам.</span><span class="sxs-lookup"><span data-stu-id="37cd3-p175">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="37cd3-1132">Следующие клиенты отличаются другим поведением `saveAsync` в отношении встреч в режиме создания:</span><span class="sxs-lookup"><span data-stu-id="37cd3-1132">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="37cd3-1133">Outlook для Mac не поддерживает `saveAsync` для собраний в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="37cd3-1133">Mac Outlook does not support `saveAsync` on a meeting in compose mode.</span></span> <span data-ttu-id="37cd3-1134">При вызове `saveAsync` для собрания в Outlook для Mac возвращается ошибка.</span><span class="sxs-lookup"><span data-stu-id="37cd3-1134">Calling `saveAsync` on a meeting in Mac Outlook will return an error.</span></span>
> - <span data-ttu-id="37cd3-1135">Outlook в Интернете всегда отправляет приглашение или обновление при вызове `saveAsync` для встречи в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="37cd3-1135">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="37cd3-1136">Параметры:</span><span class="sxs-lookup"><span data-stu-id="37cd3-1136">Parameters:</span></span>

|<span data-ttu-id="37cd3-1137">Имя</span><span class="sxs-lookup"><span data-stu-id="37cd3-1137">Name</span></span>| <span data-ttu-id="37cd3-1138">Тип</span><span class="sxs-lookup"><span data-stu-id="37cd3-1138">Type</span></span>| <span data-ttu-id="37cd3-1139">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="37cd3-1139">Attributes</span></span>| <span data-ttu-id="37cd3-1140">Описание</span><span class="sxs-lookup"><span data-stu-id="37cd3-1140">Description</span></span>|
|---|---|---|---|
|`options`| <span data-ttu-id="37cd3-1141">Object</span><span class="sxs-lookup"><span data-stu-id="37cd3-1141">Object</span></span>| <span data-ttu-id="37cd3-1142">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="37cd3-1142">&lt;optional&gt;</span></span>|<span data-ttu-id="37cd3-1143">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="37cd3-1143">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="37cd3-1144">Object</span><span class="sxs-lookup"><span data-stu-id="37cd3-1144">Object</span></span>| <span data-ttu-id="37cd3-1145">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="37cd3-1145">&lt;optional&gt;</span></span>|<span data-ttu-id="37cd3-1146">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="37cd3-1146">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="37cd3-1147">функция</span><span class="sxs-lookup"><span data-stu-id="37cd3-1147">function</span></span>||<span data-ttu-id="37cd3-1148">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="37cd3-1148">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="37cd3-1149">После успешного выполнения идентификатор элемента будет предоставлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="37cd3-1149">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="37cd3-1150">Требования</span><span class="sxs-lookup"><span data-stu-id="37cd3-1150">Requirements</span></span>

|<span data-ttu-id="37cd3-1151">Требование</span><span class="sxs-lookup"><span data-stu-id="37cd3-1151">Requirement</span></span>| <span data-ttu-id="37cd3-1152">Значение</span><span class="sxs-lookup"><span data-stu-id="37cd3-1152">Value</span></span>|
|---|---|
|[<span data-ttu-id="37cd3-1153">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="37cd3-1153">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="37cd3-1154">1.3</span><span class="sxs-lookup"><span data-stu-id="37cd3-1154">1.3</span></span>|
|[<span data-ttu-id="37cd3-1155">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="37cd3-1155">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="37cd3-1156">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="37cd3-1156">ReadWriteItem</span></span>|
|[<span data-ttu-id="37cd3-1157">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="37cd3-1157">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="37cd3-1158">Создание</span><span class="sxs-lookup"><span data-stu-id="37cd3-1158">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="37cd3-1159">Примеры</span><span class="sxs-lookup"><span data-stu-id="37cd3-1159">Examples</span></span>

```js
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result
  });
```

<span data-ttu-id="37cd3-p177">Ниже приведен пример параметра `result`, переданного функции обратного вызова. Свойство `value` содержит идентификатор элемента.</span><span class="sxs-lookup"><span data-stu-id="37cd3-p177">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

####  <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="37cd3-1162">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="37cd3-1162">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="37cd3-1163">Асинхронно вставляет данные в текст или тему сообщения.</span><span class="sxs-lookup"><span data-stu-id="37cd3-1163">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="37cd3-p178">Метод `setSelectedDataAsync` вставляет указанную строку в местоположение курсора в теме или тексте элемента либо, если текст выделен в редакторе, он заменяет выделенный текст. Если курсор находится вне текста или темы элемента, возвращается ошибка. После вставки курсор помещается в конец вставленного содержимого.</span><span class="sxs-lookup"><span data-stu-id="37cd3-p178">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="37cd3-1167">Параметры:</span><span class="sxs-lookup"><span data-stu-id="37cd3-1167">Parameters:</span></span>

|<span data-ttu-id="37cd3-1168">Имя</span><span class="sxs-lookup"><span data-stu-id="37cd3-1168">Name</span></span>| <span data-ttu-id="37cd3-1169">Тип</span><span class="sxs-lookup"><span data-stu-id="37cd3-1169">Type</span></span>| <span data-ttu-id="37cd3-1170">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="37cd3-1170">Attributes</span></span>| <span data-ttu-id="37cd3-1171">Описание</span><span class="sxs-lookup"><span data-stu-id="37cd3-1171">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="37cd3-1172">String</span><span class="sxs-lookup"><span data-stu-id="37cd3-1172">String</span></span>||<span data-ttu-id="37cd3-p179">Вставляемые данные. Объем данных не должен превышать 1 000 000 символов. Если передано больше 1 000 000 символов, возвращается исключение `ArgumentOutOfRange`.</span><span class="sxs-lookup"><span data-stu-id="37cd3-p179">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="37cd3-1176">Object</span><span class="sxs-lookup"><span data-stu-id="37cd3-1176">Object</span></span>| <span data-ttu-id="37cd3-1177">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="37cd3-1177">&lt;optional&gt;</span></span>|<span data-ttu-id="37cd3-1178">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="37cd3-1178">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="37cd3-1179">Object</span><span class="sxs-lookup"><span data-stu-id="37cd3-1179">Object</span></span>| <span data-ttu-id="37cd3-1180">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="37cd3-1180">&lt;optional&gt;</span></span>|<span data-ttu-id="37cd3-1181">В методе обратного вызова разработчики могут указать любой объект, к которому необходимо получить доступ.</span><span class="sxs-lookup"><span data-stu-id="37cd3-1181">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`| [<span data-ttu-id="37cd3-1182">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="37cd3-1182">Office.CoercionType</span></span>](office.md#coerciontype-string)| <span data-ttu-id="37cd3-1183">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="37cd3-1183">&lt;optional&gt;</span></span>|<span data-ttu-id="37cd3-p180">Если задано значение `text`, текущий стиль применяется в Outlook Web App и Outlook. Если поле представляет собой редактор HTML, вставляются только текстовые данные, даже если они имеют формат HTML.</span><span class="sxs-lookup"><span data-stu-id="37cd3-p180">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="37cd3-p181">Если задано значение `html` и поле (не тема) поддерживает HTML, в Outlook Web App применяется текущий стиль, а в Outlook — стиль по умолчанию. Если поле является текстовым, возвращается ошибка `InvalidDataFormat`.</span><span class="sxs-lookup"><span data-stu-id="37cd3-p181">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="37cd3-1188">Если свойство `coercionType` не задано, результат зависит от поля: если поле имеет формат HTML, используется текст в формате HTML, а если поле текстовое, применяется обычный текст.</span><span class="sxs-lookup"><span data-stu-id="37cd3-1188">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="37cd3-1189">функция</span><span class="sxs-lookup"><span data-stu-id="37cd3-1189">function</span></span>||<span data-ttu-id="37cd3-1190">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="37cd3-1190">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="37cd3-1191">Требования</span><span class="sxs-lookup"><span data-stu-id="37cd3-1191">Requirements</span></span>

|<span data-ttu-id="37cd3-1192">Требование</span><span class="sxs-lookup"><span data-stu-id="37cd3-1192">Requirement</span></span>| <span data-ttu-id="37cd3-1193">Значение</span><span class="sxs-lookup"><span data-stu-id="37cd3-1193">Value</span></span>|
|---|---|
|[<span data-ttu-id="37cd3-1194">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="37cd3-1194">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="37cd3-1195">1.2</span><span class="sxs-lookup"><span data-stu-id="37cd3-1195">1.2</span></span>|
|[<span data-ttu-id="37cd3-1196">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="37cd3-1196">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="37cd3-1197">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="37cd3-1197">ReadWriteItem</span></span>|
|[<span data-ttu-id="37cd3-1198">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="37cd3-1198">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="37cd3-1199">Создание</span><span class="sxs-lookup"><span data-stu-id="37cd3-1199">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="37cd3-1200">Пример</span><span class="sxs-lookup"><span data-stu-id="37cd3-1200">Example</span></span>

```js
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
