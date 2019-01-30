---
title: Office.context.mailbox.item — набор обязательных элементов 1.5
description: ''
ms.date: 12/18/2018
localization_priority: Priority
ms.openlocfilehash: 48bc1291e7aa6d8e335c07d16ddd74e6e9455f0d
ms.sourcegitcommit: d1aa7201820176ed986b9f00bb9c88e055906c77
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 01/23/2019
ms.locfileid: "29389573"
---
# <a name="item"></a><span data-ttu-id="5b02e-102">item</span><span class="sxs-lookup"><span data-stu-id="5b02e-102">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="5b02e-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="5b02e-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="5b02e-p101">Пространство имен `item` используется для доступа к выбранному в данный момент сообщению, приглашению на собрание или описанию встречи. Вы можете определить тип пространства имен `item` с помощью свойства [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook15officemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="5b02e-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook15officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="5b02e-106">Requirements</span><span class="sxs-lookup"><span data-stu-id="5b02e-106">Requirements</span></span>

|<span data-ttu-id="5b02e-107">Требование</span><span class="sxs-lookup"><span data-stu-id="5b02e-107">Requirement</span></span>| <span data-ttu-id="5b02e-108">Значение</span><span class="sxs-lookup"><span data-stu-id="5b02e-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="5b02e-109">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5b02e-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5b02e-110">1.0</span><span class="sxs-lookup"><span data-stu-id="5b02e-110">1.0</span></span>|
|[<span data-ttu-id="5b02e-111">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5b02e-111">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5b02e-112">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="5b02e-112">Restricted</span></span>|
|[<span data-ttu-id="5b02e-113">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5b02e-113">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5b02e-114">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="5b02e-114">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="5b02e-115">Элементы и методы</span><span class="sxs-lookup"><span data-stu-id="5b02e-115">Members and methods</span></span>

| <span data-ttu-id="5b02e-116">Элемент</span><span class="sxs-lookup"><span data-stu-id="5b02e-116">Member</span></span> | <span data-ttu-id="5b02e-117">Тип</span><span class="sxs-lookup"><span data-stu-id="5b02e-117">Type</span></span> |
|--------|------|
| [<span data-ttu-id="5b02e-118">attachments</span><span class="sxs-lookup"><span data-stu-id="5b02e-118">attachments</span></span>](#attachments-arrayattachmentdetailsjavascriptapioutlook15officeattachmentdetails) | <span data-ttu-id="5b02e-119">Элемент</span><span class="sxs-lookup"><span data-stu-id="5b02e-119">Member</span></span> |
| [<span data-ttu-id="5b02e-120">bcc</span><span class="sxs-lookup"><span data-stu-id="5b02e-120">bcc</span></span>](#bcc-recipientsjavascriptapioutlook15officerecipients) | <span data-ttu-id="5b02e-121">Элемент</span><span class="sxs-lookup"><span data-stu-id="5b02e-121">Member</span></span> |
| [<span data-ttu-id="5b02e-122">body</span><span class="sxs-lookup"><span data-stu-id="5b02e-122">body</span></span>](#body-bodyjavascriptapioutlook15officebody) | <span data-ttu-id="5b02e-123">Элемент</span><span class="sxs-lookup"><span data-stu-id="5b02e-123">Member</span></span> |
| [<span data-ttu-id="5b02e-124">cc</span><span class="sxs-lookup"><span data-stu-id="5b02e-124">cc</span></span>](#cc-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients) | <span data-ttu-id="5b02e-125">Элемент</span><span class="sxs-lookup"><span data-stu-id="5b02e-125">Member</span></span> |
| [<span data-ttu-id="5b02e-126">conversationId</span><span class="sxs-lookup"><span data-stu-id="5b02e-126">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="5b02e-127">Элемент</span><span class="sxs-lookup"><span data-stu-id="5b02e-127">Member</span></span> |
| [<span data-ttu-id="5b02e-128">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="5b02e-128">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="5b02e-129">Элемент</span><span class="sxs-lookup"><span data-stu-id="5b02e-129">Member</span></span> |
| [<span data-ttu-id="5b02e-130">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="5b02e-130">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="5b02e-131">Элемент</span><span class="sxs-lookup"><span data-stu-id="5b02e-131">Member</span></span> |
| [<span data-ttu-id="5b02e-132">end</span><span class="sxs-lookup"><span data-stu-id="5b02e-132">end</span></span>](#end-datetimejavascriptapioutlook15officetime) | <span data-ttu-id="5b02e-133">Элемент</span><span class="sxs-lookup"><span data-stu-id="5b02e-133">Member</span></span> |
| [<span data-ttu-id="5b02e-134">from</span><span class="sxs-lookup"><span data-stu-id="5b02e-134">from</span></span>](#from-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails) | <span data-ttu-id="5b02e-135">Элемент</span><span class="sxs-lookup"><span data-stu-id="5b02e-135">Member</span></span> |
| [<span data-ttu-id="5b02e-136">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="5b02e-136">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="5b02e-137">Элемент</span><span class="sxs-lookup"><span data-stu-id="5b02e-137">Member</span></span> |
| [<span data-ttu-id="5b02e-138">itemClass</span><span class="sxs-lookup"><span data-stu-id="5b02e-138">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="5b02e-139">Элемент</span><span class="sxs-lookup"><span data-stu-id="5b02e-139">Member</span></span> |
| [<span data-ttu-id="5b02e-140">itemId</span><span class="sxs-lookup"><span data-stu-id="5b02e-140">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="5b02e-141">Элемент</span><span class="sxs-lookup"><span data-stu-id="5b02e-141">Member</span></span> |
| [<span data-ttu-id="5b02e-142">itemType</span><span class="sxs-lookup"><span data-stu-id="5b02e-142">itemType</span></span>](#itemtype-officemailboxenumsitemtypejavascriptapioutlook15officemailboxenumsitemtype) | <span data-ttu-id="5b02e-143">Элемент</span><span class="sxs-lookup"><span data-stu-id="5b02e-143">Member</span></span> |
| [<span data-ttu-id="5b02e-144">location</span><span class="sxs-lookup"><span data-stu-id="5b02e-144">location</span></span>](#location-stringlocationjavascriptapioutlook15officelocation) | <span data-ttu-id="5b02e-145">Элемент</span><span class="sxs-lookup"><span data-stu-id="5b02e-145">Member</span></span> |
| [<span data-ttu-id="5b02e-146">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="5b02e-146">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="5b02e-147">Элемент</span><span class="sxs-lookup"><span data-stu-id="5b02e-147">Member</span></span> |
| [<span data-ttu-id="5b02e-148">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="5b02e-148">notificationMessages</span></span>](#notificationmessages-notificationmessagesjavascriptapioutlook15officenotificationmessages) | <span data-ttu-id="5b02e-149">Элемент</span><span class="sxs-lookup"><span data-stu-id="5b02e-149">Member</span></span> |
| [<span data-ttu-id="5b02e-150">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="5b02e-150">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients) | <span data-ttu-id="5b02e-151">Элемент</span><span class="sxs-lookup"><span data-stu-id="5b02e-151">Member</span></span> |
| [<span data-ttu-id="5b02e-152">organizer</span><span class="sxs-lookup"><span data-stu-id="5b02e-152">organizer</span></span>](#organizer-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails) | <span data-ttu-id="5b02e-153">Элемент</span><span class="sxs-lookup"><span data-stu-id="5b02e-153">Member</span></span> |
| [<span data-ttu-id="5b02e-154">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="5b02e-154">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients) | <span data-ttu-id="5b02e-155">Member</span><span class="sxs-lookup"><span data-stu-id="5b02e-155">Member</span></span> |
| [<span data-ttu-id="5b02e-156">sender</span><span class="sxs-lookup"><span data-stu-id="5b02e-156">sender</span></span>](#sender-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails) | <span data-ttu-id="5b02e-157">Элемент</span><span class="sxs-lookup"><span data-stu-id="5b02e-157">Member</span></span> |
| [<span data-ttu-id="5b02e-158">start</span><span class="sxs-lookup"><span data-stu-id="5b02e-158">start</span></span>](#start-datetimejavascriptapioutlook15officetime) | <span data-ttu-id="5b02e-159">Элемент</span><span class="sxs-lookup"><span data-stu-id="5b02e-159">Member</span></span> |
| [<span data-ttu-id="5b02e-160">subject</span><span class="sxs-lookup"><span data-stu-id="5b02e-160">subject</span></span>](#subject-stringsubjectjavascriptapioutlook15officesubject) | <span data-ttu-id="5b02e-161">Элемент</span><span class="sxs-lookup"><span data-stu-id="5b02e-161">Member</span></span> |
| [<span data-ttu-id="5b02e-162">to</span><span class="sxs-lookup"><span data-stu-id="5b02e-162">to</span></span>](#to-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients) | <span data-ttu-id="5b02e-163">Элемент</span><span class="sxs-lookup"><span data-stu-id="5b02e-163">Member</span></span> |
| [<span data-ttu-id="5b02e-164">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="5b02e-164">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="5b02e-165">Метод</span><span class="sxs-lookup"><span data-stu-id="5b02e-165">Method</span></span> |
| [<span data-ttu-id="5b02e-166">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="5b02e-166">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="5b02e-167">Метод</span><span class="sxs-lookup"><span data-stu-id="5b02e-167">Method</span></span> |
| [<span data-ttu-id="5b02e-168">close</span><span class="sxs-lookup"><span data-stu-id="5b02e-168">close</span></span>](#close) | <span data-ttu-id="5b02e-169">Метод</span><span class="sxs-lookup"><span data-stu-id="5b02e-169">Method</span></span> |
| [<span data-ttu-id="5b02e-170">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="5b02e-170">displayReplyAllForm</span></span>](#displayreplyallformformdata) | <span data-ttu-id="5b02e-171">Метод</span><span class="sxs-lookup"><span data-stu-id="5b02e-171">Method</span></span> |
| [<span data-ttu-id="5b02e-172">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="5b02e-172">displayReplyForm</span></span>](#displayreplyformformdata) | <span data-ttu-id="5b02e-173">Метод</span><span class="sxs-lookup"><span data-stu-id="5b02e-173">Method</span></span> |
| [<span data-ttu-id="5b02e-174">getEntities</span><span class="sxs-lookup"><span data-stu-id="5b02e-174">getEntities</span></span>](#getentities--entitiesjavascriptapioutlook15officeentities) | <span data-ttu-id="5b02e-175">Метод</span><span class="sxs-lookup"><span data-stu-id="5b02e-175">Method</span></span> |
| [<span data-ttu-id="5b02e-176">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="5b02e-176">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook15officecontactmeetingsuggestionjavascriptapioutlook15officemeetingsuggestionphonenumberjavascriptapioutlook15officephonenumbertasksuggestionjavascriptapioutlook15officetasksuggestion) | <span data-ttu-id="5b02e-177">Метод</span><span class="sxs-lookup"><span data-stu-id="5b02e-177">Method</span></span> |
| [<span data-ttu-id="5b02e-178">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="5b02e-178">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook15officecontactmeetingsuggestionjavascriptapioutlook15officemeetingsuggestionphonenumberjavascriptapioutlook15officephonenumbertasksuggestionjavascriptapioutlook15officetasksuggestion) | <span data-ttu-id="5b02e-179">Метод</span><span class="sxs-lookup"><span data-stu-id="5b02e-179">Method</span></span> |
| [<span data-ttu-id="5b02e-180">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="5b02e-180">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="5b02e-181">Метод</span><span class="sxs-lookup"><span data-stu-id="5b02e-181">Method</span></span> |
| [<span data-ttu-id="5b02e-182">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="5b02e-182">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="5b02e-183">Метод</span><span class="sxs-lookup"><span data-stu-id="5b02e-183">Method</span></span> |
| [<span data-ttu-id="5b02e-184">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="5b02e-184">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="5b02e-185">Метод</span><span class="sxs-lookup"><span data-stu-id="5b02e-185">Method</span></span> |
| [<span data-ttu-id="5b02e-186">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="5b02e-186">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="5b02e-187">Метод</span><span class="sxs-lookup"><span data-stu-id="5b02e-187">Method</span></span> |
| [<span data-ttu-id="5b02e-188">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="5b02e-188">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="5b02e-189">Метод</span><span class="sxs-lookup"><span data-stu-id="5b02e-189">Method</span></span> |
| [<span data-ttu-id="5b02e-190">saveAsync</span><span class="sxs-lookup"><span data-stu-id="5b02e-190">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="5b02e-191">Метод</span><span class="sxs-lookup"><span data-stu-id="5b02e-191">Method</span></span> |
| [<span data-ttu-id="5b02e-192">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="5b02e-192">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="5b02e-193">Метод</span><span class="sxs-lookup"><span data-stu-id="5b02e-193">Method</span></span> |

### <a name="example"></a><span data-ttu-id="5b02e-194">Пример</span><span class="sxs-lookup"><span data-stu-id="5b02e-194">Example</span></span>

<span data-ttu-id="5b02e-195">В примере кода JavaScript, приведенном ниже, показано, как получить доступ к свойству `subject` текущего элемента в Outlook.</span><span class="sxs-lookup"><span data-stu-id="5b02e-195">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="5b02e-196">Элементы</span><span class="sxs-lookup"><span data-stu-id="5b02e-196">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook15officeattachmentdetails"></a><span data-ttu-id="5b02e-197">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_5/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="5b02e-197">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_5/office.attachmentdetails)></span></span>

<span data-ttu-id="5b02e-p102">Получает массив вложений для элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="5b02e-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="5b02e-200">Outlook блокирует определенные типы файлов из-за потенциальных проблем с безопасностью, поэтому они не возвращаются.</span><span class="sxs-lookup"><span data-stu-id="5b02e-200">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="5b02e-201">Дополнительные сведения см. в статье [Блокировка вложений в Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="5b02e-201">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="5b02e-202">Тип:</span><span class="sxs-lookup"><span data-stu-id="5b02e-202">Type:</span></span>

*   <span data-ttu-id="5b02e-203">Array.<[AttachmentDetails](/javascript/api/outlook_1_5/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="5b02e-203">Array.<[AttachmentDetails](/javascript/api/outlook_1_5/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="5b02e-204">Требования</span><span class="sxs-lookup"><span data-stu-id="5b02e-204">Requirements</span></span>

|<span data-ttu-id="5b02e-205">Требование</span><span class="sxs-lookup"><span data-stu-id="5b02e-205">Requirement</span></span>| <span data-ttu-id="5b02e-206">Значение</span><span class="sxs-lookup"><span data-stu-id="5b02e-206">Value</span></span>|
|---|---|
|[<span data-ttu-id="5b02e-207">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5b02e-207">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5b02e-208">1.0</span><span class="sxs-lookup"><span data-stu-id="5b02e-208">1.0</span></span>|
|[<span data-ttu-id="5b02e-209">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5b02e-209">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5b02e-210">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5b02e-210">ReadItem</span></span>|
|[<span data-ttu-id="5b02e-211">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5b02e-211">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5b02e-212">Чтение</span><span class="sxs-lookup"><span data-stu-id="5b02e-212">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="5b02e-213">Пример</span><span class="sxs-lookup"><span data-stu-id="5b02e-213">Example</span></span>

<span data-ttu-id="5b02e-214">С помощью приведенного ниже кода можно создать HTML-строку с подробными сведениями обо всех вложениях для текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="5b02e-214">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

####  <a name="bcc-recipientsjavascriptapioutlook15officerecipients"></a><span data-ttu-id="5b02e-215">bcc :[Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="5b02e-215">bcc :[Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

<span data-ttu-id="5b02e-216">Получает объект, который предоставляет методы для получения или обновления получателей скрытой копии сообщения.</span><span class="sxs-lookup"><span data-stu-id="5b02e-216">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="5b02e-217">Только в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="5b02e-217">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="5b02e-218">Тип:</span><span class="sxs-lookup"><span data-stu-id="5b02e-218">Type:</span></span>

*   [<span data-ttu-id="5b02e-219">Recipients</span><span class="sxs-lookup"><span data-stu-id="5b02e-219">Recipients</span></span>](/javascript/api/outlook_1_5/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="5b02e-220">Требования</span><span class="sxs-lookup"><span data-stu-id="5b02e-220">Requirements</span></span>

|<span data-ttu-id="5b02e-221">Требование</span><span class="sxs-lookup"><span data-stu-id="5b02e-221">Requirement</span></span>| <span data-ttu-id="5b02e-222">Значение</span><span class="sxs-lookup"><span data-stu-id="5b02e-222">Value</span></span>|
|---|---|
|[<span data-ttu-id="5b02e-223">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5b02e-223">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5b02e-224">1.1</span><span class="sxs-lookup"><span data-stu-id="5b02e-224">1.1</span></span>|
|[<span data-ttu-id="5b02e-225">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5b02e-225">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5b02e-226">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5b02e-226">ReadItem</span></span>|
|[<span data-ttu-id="5b02e-227">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5b02e-227">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5b02e-228">Создание</span><span class="sxs-lookup"><span data-stu-id="5b02e-228">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="5b02e-229">Пример</span><span class="sxs-lookup"><span data-stu-id="5b02e-229">Example</span></span>

```js
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlook15officebody"></a><span data-ttu-id="5b02e-230">body :[Body](/javascript/api/outlook_1_5/office.body)</span><span class="sxs-lookup"><span data-stu-id="5b02e-230">body :[Body](/javascript/api/outlook_1_5/office.body)</span></span>

<span data-ttu-id="5b02e-231">Получает объект, предоставляющий методы для работы с основным текстом элемента.</span><span class="sxs-lookup"><span data-stu-id="5b02e-231">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="5b02e-232">Тип:</span><span class="sxs-lookup"><span data-stu-id="5b02e-232">Type:</span></span>

*   [<span data-ttu-id="5b02e-233">Body</span><span class="sxs-lookup"><span data-stu-id="5b02e-233">Body</span></span>](/javascript/api/outlook_1_5/office.body)

##### <a name="requirements"></a><span data-ttu-id="5b02e-234">Требования</span><span class="sxs-lookup"><span data-stu-id="5b02e-234">Requirements</span></span>

|<span data-ttu-id="5b02e-235">Требование</span><span class="sxs-lookup"><span data-stu-id="5b02e-235">Requirement</span></span>| <span data-ttu-id="5b02e-236">Значение</span><span class="sxs-lookup"><span data-stu-id="5b02e-236">Value</span></span>|
|---|---|
|[<span data-ttu-id="5b02e-237">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5b02e-237">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5b02e-238">1.1</span><span class="sxs-lookup"><span data-stu-id="5b02e-238">1.1</span></span>|
|[<span data-ttu-id="5b02e-239">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5b02e-239">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5b02e-240">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5b02e-240">ReadItem</span></span>|
|[<span data-ttu-id="5b02e-241">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5b02e-241">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5b02e-242">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="5b02e-242">Compose or read</span></span>|

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients"></a><span data-ttu-id="5b02e-243">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="5b02e-243">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

<span data-ttu-id="5b02e-244">Предоставляет доступ к получателям копии сообщения.</span><span class="sxs-lookup"><span data-stu-id="5b02e-244">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="5b02e-245">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="5b02e-245">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="5b02e-246">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="5b02e-246">Read mode</span></span>

<span data-ttu-id="5b02e-p106">Свойство `cc` возвращает массив, который содержит объект `EmailAddressDetails` для каждого получателя, указанного в строке **Копия** сообщения. Коллекция может включать не более 100 элементов.</span><span class="sxs-lookup"><span data-stu-id="5b02e-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="5b02e-249">Режим создания</span><span class="sxs-lookup"><span data-stu-id="5b02e-249">Compose mode</span></span>

<span data-ttu-id="5b02e-250">Свойство `cc` возвращает объект `Recipients`, предоставляющий методы для получения или обновления получателей, которые указаны в строке **Копия** сообщения.</span><span class="sxs-lookup"><span data-stu-id="5b02e-250">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="5b02e-251">Тип:</span><span class="sxs-lookup"><span data-stu-id="5b02e-251">Type:</span></span>

*   <span data-ttu-id="5b02e-252">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="5b02e-252">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="5b02e-253">Требования</span><span class="sxs-lookup"><span data-stu-id="5b02e-253">Requirements</span></span>

|<span data-ttu-id="5b02e-254">Требование</span><span class="sxs-lookup"><span data-stu-id="5b02e-254">Requirement</span></span>| <span data-ttu-id="5b02e-255">Значение</span><span class="sxs-lookup"><span data-stu-id="5b02e-255">Value</span></span>|
|---|---|
|[<span data-ttu-id="5b02e-256">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5b02e-256">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5b02e-257">1.0</span><span class="sxs-lookup"><span data-stu-id="5b02e-257">1.0</span></span>|
|[<span data-ttu-id="5b02e-258">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5b02e-258">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5b02e-259">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5b02e-259">ReadItem</span></span>|
|[<span data-ttu-id="5b02e-260">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5b02e-260">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5b02e-261">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="5b02e-261">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="5b02e-262">Пример</span><span class="sxs-lookup"><span data-stu-id="5b02e-262">Example</span></span>

```js
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="5b02e-263">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="5b02e-263">(nullable) conversationId :String</span></span>

<span data-ttu-id="5b02e-264">Получает идентификатор разговора по электронной почте, содержащего конкретное сообщение.</span><span class="sxs-lookup"><span data-stu-id="5b02e-264">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="5b02e-p107">Вы можете получить целочисленное значение этого свойства, если ваше почтовое приложение активируется в формах просмотра или формах создания ответов. Если пользователь изменит тему ответа, после его отправки идентификатор беседы будет изменен, и полученное ранее значение будет недействительным.</span><span class="sxs-lookup"><span data-stu-id="5b02e-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="5b02e-p108">Это свойство имеет значение NULL для нового элемента в форме создания. Свойство `conversationId` вернет значение, если пользователь задаст тему и сохранит элемент.</span><span class="sxs-lookup"><span data-stu-id="5b02e-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="5b02e-269">Тип:</span><span class="sxs-lookup"><span data-stu-id="5b02e-269">Type:</span></span>

*   <span data-ttu-id="5b02e-270">String</span><span class="sxs-lookup"><span data-stu-id="5b02e-270">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="5b02e-271">Требования</span><span class="sxs-lookup"><span data-stu-id="5b02e-271">Requirements</span></span>

|<span data-ttu-id="5b02e-272">Требование</span><span class="sxs-lookup"><span data-stu-id="5b02e-272">Requirement</span></span>| <span data-ttu-id="5b02e-273">Значение</span><span class="sxs-lookup"><span data-stu-id="5b02e-273">Value</span></span>|
|---|---|
|[<span data-ttu-id="5b02e-274">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5b02e-274">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5b02e-275">1.0</span><span class="sxs-lookup"><span data-stu-id="5b02e-275">1.0</span></span>|
|[<span data-ttu-id="5b02e-276">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5b02e-276">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5b02e-277">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5b02e-277">ReadItem</span></span>|
|[<span data-ttu-id="5b02e-278">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5b02e-278">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5b02e-279">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="5b02e-279">Compose or read</span></span>|

#### <a name="datetimecreated-date"></a><span data-ttu-id="5b02e-280">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="5b02e-280">dateTimeCreated :Date</span></span>

<span data-ttu-id="5b02e-p109">Получает дату и время создания элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="5b02e-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="5b02e-283">Тип:</span><span class="sxs-lookup"><span data-stu-id="5b02e-283">Type:</span></span>

*   <span data-ttu-id="5b02e-284">Date</span><span class="sxs-lookup"><span data-stu-id="5b02e-284">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="5b02e-285">Требования</span><span class="sxs-lookup"><span data-stu-id="5b02e-285">Requirements</span></span>

|<span data-ttu-id="5b02e-286">Требование</span><span class="sxs-lookup"><span data-stu-id="5b02e-286">Requirement</span></span>| <span data-ttu-id="5b02e-287">Значение</span><span class="sxs-lookup"><span data-stu-id="5b02e-287">Value</span></span>|
|---|---|
|[<span data-ttu-id="5b02e-288">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5b02e-288">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5b02e-289">1.0</span><span class="sxs-lookup"><span data-stu-id="5b02e-289">1.0</span></span>|
|[<span data-ttu-id="5b02e-290">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5b02e-290">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5b02e-291">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5b02e-291">ReadItem</span></span>|
|[<span data-ttu-id="5b02e-292">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5b02e-292">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5b02e-293">Чтение</span><span class="sxs-lookup"><span data-stu-id="5b02e-293">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="5b02e-294">Пример</span><span class="sxs-lookup"><span data-stu-id="5b02e-294">Example</span></span>

```js
var created = Office.context.mailbox.item.dateTimeCreated;
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="5b02e-295">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="5b02e-295">dateTimeModified :Date</span></span>

<span data-ttu-id="5b02e-p110">Получает дату и время последнего изменения элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="5b02e-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="5b02e-298">Этот элемент не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="5b02e-298">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="5b02e-299">Тип:</span><span class="sxs-lookup"><span data-stu-id="5b02e-299">Type:</span></span>

*   <span data-ttu-id="5b02e-300">Date</span><span class="sxs-lookup"><span data-stu-id="5b02e-300">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="5b02e-301">Требования</span><span class="sxs-lookup"><span data-stu-id="5b02e-301">Requirements</span></span>

|<span data-ttu-id="5b02e-302">Требование</span><span class="sxs-lookup"><span data-stu-id="5b02e-302">Requirement</span></span>| <span data-ttu-id="5b02e-303">Значение</span><span class="sxs-lookup"><span data-stu-id="5b02e-303">Value</span></span>|
|---|---|
|[<span data-ttu-id="5b02e-304">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5b02e-304">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5b02e-305">1.0</span><span class="sxs-lookup"><span data-stu-id="5b02e-305">1.0</span></span>|
|[<span data-ttu-id="5b02e-306">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5b02e-306">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5b02e-307">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5b02e-307">ReadItem</span></span>|
|[<span data-ttu-id="5b02e-308">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5b02e-308">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5b02e-309">Чтение</span><span class="sxs-lookup"><span data-stu-id="5b02e-309">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="5b02e-310">Пример</span><span class="sxs-lookup"><span data-stu-id="5b02e-310">Example</span></span>

```js
var modified = Office.context.mailbox.item.dateTimeModified;
```

####  <a name="end-datetimejavascriptapioutlook15officetime"></a><span data-ttu-id="5b02e-311">end :Date|[Time](/javascript/api/outlook_1_5/office.time)</span><span class="sxs-lookup"><span data-stu-id="5b02e-311">end :Date|[Time](/javascript/api/outlook_1_5/office.time)</span></span>

<span data-ttu-id="5b02e-312">Получает или задает дату и время окончания встречи.</span><span class="sxs-lookup"><span data-stu-id="5b02e-312">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="5b02e-p111">Свойство `end` представлено в виде значения даты и времени в формате UTC. Преобразовать значение свойства end в местные значения даты и времени клиента можно с помощью метода [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook15officelocalclienttime).</span><span class="sxs-lookup"><span data-stu-id="5b02e-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook15officelocalclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="5b02e-315">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="5b02e-315">Read mode</span></span>

<span data-ttu-id="5b02e-316">Свойство `end` возвращает объект `Date`.</span><span class="sxs-lookup"><span data-stu-id="5b02e-316">The `end` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="5b02e-317">Режим создания</span><span class="sxs-lookup"><span data-stu-id="5b02e-317">Compose mode</span></span>

<span data-ttu-id="5b02e-318">Свойство `end` возвращает объект `Time`.</span><span class="sxs-lookup"><span data-stu-id="5b02e-318">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="5b02e-319">Если вы задаете время окончания с помощью метода [`Time.setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-), необходимо использовать метод [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) для преобразования местного времени на клиенте в формат UTC для сервера.</span><span class="sxs-lookup"><span data-stu-id="5b02e-319">When you use the [`Time.setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="5b02e-320">Тип:</span><span class="sxs-lookup"><span data-stu-id="5b02e-320">Type:</span></span>

*   <span data-ttu-id="5b02e-321">Date | [Time](/javascript/api/outlook_1_5/office.time)</span><span class="sxs-lookup"><span data-stu-id="5b02e-321">Date | [Time](/javascript/api/outlook_1_5/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="5b02e-322">Требования</span><span class="sxs-lookup"><span data-stu-id="5b02e-322">Requirements</span></span>

|<span data-ttu-id="5b02e-323">Требование</span><span class="sxs-lookup"><span data-stu-id="5b02e-323">Requirement</span></span>| <span data-ttu-id="5b02e-324">Значение</span><span class="sxs-lookup"><span data-stu-id="5b02e-324">Value</span></span>|
|---|---|
|[<span data-ttu-id="5b02e-325">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5b02e-325">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5b02e-326">1.0</span><span class="sxs-lookup"><span data-stu-id="5b02e-326">1.0</span></span>|
|[<span data-ttu-id="5b02e-327">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5b02e-327">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5b02e-328">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5b02e-328">ReadItem</span></span>|
|[<span data-ttu-id="5b02e-329">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5b02e-329">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5b02e-330">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="5b02e-330">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="5b02e-331">Пример</span><span class="sxs-lookup"><span data-stu-id="5b02e-331">Example</span></span>

<span data-ttu-id="5b02e-332">В примере ниже показано, как с помощью метода [`setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) объекта `Time` задать время окончания встречи в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="5b02e-332">The following example sets the end time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

#### <a name="from-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails"></a><span data-ttu-id="5b02e-333">from :[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="5b02e-333">from :[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)</span></span>

<span data-ttu-id="5b02e-p112">Получает электронный адрес отправителя сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="5b02e-p112">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="5b02e-p113">Свойства `from` и [`sender`](#sender-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails) представляют одно лицо, если сообщение не отправлено представителем. В противном случае свойство `from` представляет лицо, делегировавшее полномочия, а свойство sender — представителя.</span><span class="sxs-lookup"><span data-stu-id="5b02e-p113">The `from` and [`sender`](#sender-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="5b02e-338">Свойству `recipientType`, принадлежащему объекту `EmailAddressDetails` в свойстве `from`, задано значение `undefined`.</span><span class="sxs-lookup"><span data-stu-id="5b02e-338">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="5b02e-339">Тип:</span><span class="sxs-lookup"><span data-stu-id="5b02e-339">Type:</span></span>

*   [<span data-ttu-id="5b02e-340">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="5b02e-340">EmailAddressDetails</span></span>](/javascript/api/outlook_1_5/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="5b02e-341">Требования</span><span class="sxs-lookup"><span data-stu-id="5b02e-341">Requirements</span></span>

|<span data-ttu-id="5b02e-342">Требование</span><span class="sxs-lookup"><span data-stu-id="5b02e-342">Requirement</span></span>| <span data-ttu-id="5b02e-343">Значение</span><span class="sxs-lookup"><span data-stu-id="5b02e-343">Value</span></span>|
|---|---|
|[<span data-ttu-id="5b02e-344">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5b02e-344">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5b02e-345">1.0</span><span class="sxs-lookup"><span data-stu-id="5b02e-345">1.0</span></span>|
|[<span data-ttu-id="5b02e-346">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5b02e-346">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5b02e-347">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5b02e-347">ReadItem</span></span>|
|[<span data-ttu-id="5b02e-348">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5b02e-348">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5b02e-349">Чтение</span><span class="sxs-lookup"><span data-stu-id="5b02e-349">Read</span></span>|

#### <a name="internetmessageid-string"></a><span data-ttu-id="5b02e-350">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="5b02e-350">internetMessageId :String</span></span>

<span data-ttu-id="5b02e-p114">Получает идентификатор интернет-сообщения для электронного сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="5b02e-p114">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="5b02e-353">Тип:</span><span class="sxs-lookup"><span data-stu-id="5b02e-353">Type:</span></span>

*   <span data-ttu-id="5b02e-354">String</span><span class="sxs-lookup"><span data-stu-id="5b02e-354">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="5b02e-355">Требования</span><span class="sxs-lookup"><span data-stu-id="5b02e-355">Requirements</span></span>

|<span data-ttu-id="5b02e-356">Требование</span><span class="sxs-lookup"><span data-stu-id="5b02e-356">Requirement</span></span>| <span data-ttu-id="5b02e-357">Значение</span><span class="sxs-lookup"><span data-stu-id="5b02e-357">Value</span></span>|
|---|---|
|[<span data-ttu-id="5b02e-358">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5b02e-358">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5b02e-359">1.0</span><span class="sxs-lookup"><span data-stu-id="5b02e-359">1.0</span></span>|
|[<span data-ttu-id="5b02e-360">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5b02e-360">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5b02e-361">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5b02e-361">ReadItem</span></span>|
|[<span data-ttu-id="5b02e-362">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5b02e-362">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5b02e-363">Чтение</span><span class="sxs-lookup"><span data-stu-id="5b02e-363">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="5b02e-364">Пример</span><span class="sxs-lookup"><span data-stu-id="5b02e-364">Example</span></span>

```js
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="5b02e-365">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="5b02e-365">itemClass :String</span></span>

<span data-ttu-id="5b02e-p115">Получает класс элемента веб-служб Exchange для выбранного элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="5b02e-p115">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="5b02e-p116">Свойство `itemClass` указывает класс сообщения выбранного элемента. Ниже приводятся классы сообщения по умолчанию для элемента сообщения или встречи.</span><span class="sxs-lookup"><span data-stu-id="5b02e-p116">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="5b02e-370">Тип</span><span class="sxs-lookup"><span data-stu-id="5b02e-370">Type</span></span> | <span data-ttu-id="5b02e-371">Описание</span><span class="sxs-lookup"><span data-stu-id="5b02e-371">Description</span></span> | <span data-ttu-id="5b02e-372">Класс элемента</span><span class="sxs-lookup"><span data-stu-id="5b02e-372">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="5b02e-373">Элементы встречи</span><span class="sxs-lookup"><span data-stu-id="5b02e-373">Appointment items</span></span> | <span data-ttu-id="5b02e-374">Это элементы календаря для класса элемента `IPM.Appointment` или `IPM.Appointment.Occurence`.</span><span class="sxs-lookup"><span data-stu-id="5b02e-374">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurence` |
| <span data-ttu-id="5b02e-375">Элементы сообщения</span><span class="sxs-lookup"><span data-stu-id="5b02e-375">Message items</span></span> | <span data-ttu-id="5b02e-376">Сюда входят электронные сообщения, для которых по умолчанию задан класс сообщения `IPM.Note`, а также приглашения на собрания, ответы на них и уведомления об их отмене, использующие `IPM.Schedule.Meeting` в качестве базового класса сообщения.</span><span class="sxs-lookup"><span data-stu-id="5b02e-376">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="5b02e-377">Можно создавать настраиваемые классы сообщения, расширяющие классы сообщения по умолчанию, например настраиваемый класс сообщения о встрече `IPM.Appointment.Contoso`.</span><span class="sxs-lookup"><span data-stu-id="5b02e-377">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="5b02e-378">Тип:</span><span class="sxs-lookup"><span data-stu-id="5b02e-378">Type:</span></span>

*   <span data-ttu-id="5b02e-379">String</span><span class="sxs-lookup"><span data-stu-id="5b02e-379">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="5b02e-380">Требования</span><span class="sxs-lookup"><span data-stu-id="5b02e-380">Requirements</span></span>

|<span data-ttu-id="5b02e-381">Требование</span><span class="sxs-lookup"><span data-stu-id="5b02e-381">Requirement</span></span>| <span data-ttu-id="5b02e-382">Значение</span><span class="sxs-lookup"><span data-stu-id="5b02e-382">Value</span></span>|
|---|---|
|[<span data-ttu-id="5b02e-383">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5b02e-383">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5b02e-384">1.0</span><span class="sxs-lookup"><span data-stu-id="5b02e-384">1.0</span></span>|
|[<span data-ttu-id="5b02e-385">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5b02e-385">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5b02e-386">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5b02e-386">ReadItem</span></span>|
|[<span data-ttu-id="5b02e-387">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5b02e-387">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5b02e-388">Чтение</span><span class="sxs-lookup"><span data-stu-id="5b02e-388">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="5b02e-389">Пример</span><span class="sxs-lookup"><span data-stu-id="5b02e-389">Example</span></span>

```js
var itemClass = Office.context.mailbox.item.itemClass;
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="5b02e-390">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="5b02e-390">(nullable) itemId :String</span></span>

<span data-ttu-id="5b02e-p117">Получает идентификатор элемента веб-служб Exchange для текущего элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="5b02e-p117">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="5b02e-393">Идентификатор, возвращаемый свойством `itemId`, совпадает с идентификатором элемента веб-служб Exchange.</span><span class="sxs-lookup"><span data-stu-id="5b02e-393">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="5b02e-394">Свойство `itemId` не совпадает с идентификатором записи Outlook, а также идентификатором, который используется REST API Outlook.</span><span class="sxs-lookup"><span data-stu-id="5b02e-394">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="5b02e-395">Прежде чем совершать вызовы REST API, используя это значение, его необходимо преобразовать с помощью [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="5b02e-395">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="5b02e-396">Дополнительные сведения см. в статье [Использование REST API Outlook из надстройки Outlook](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="5b02e-396">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="5b02e-p119">Свойство `itemId` недоступно в режиме создания. Если требуется идентификатор элемента, с помощью метода [`saveAsync`](#saveasyncoptions-callback) можно сохранить элемент в хранилище. При этом в параметре [`AsyncResult.value`](/javascript/api/office/office.asyncresult) функции обратного вызова возвращается идентификатор элемента.</span><span class="sxs-lookup"><span data-stu-id="5b02e-p119">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="5b02e-399">Тип:</span><span class="sxs-lookup"><span data-stu-id="5b02e-399">Type:</span></span>

*   <span data-ttu-id="5b02e-400">String</span><span class="sxs-lookup"><span data-stu-id="5b02e-400">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="5b02e-401">Требования</span><span class="sxs-lookup"><span data-stu-id="5b02e-401">Requirements</span></span>

|<span data-ttu-id="5b02e-402">Требование</span><span class="sxs-lookup"><span data-stu-id="5b02e-402">Requirement</span></span>| <span data-ttu-id="5b02e-403">Значение</span><span class="sxs-lookup"><span data-stu-id="5b02e-403">Value</span></span>|
|---|---|
|[<span data-ttu-id="5b02e-404">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5b02e-404">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5b02e-405">1.0</span><span class="sxs-lookup"><span data-stu-id="5b02e-405">1.0</span></span>|
|[<span data-ttu-id="5b02e-406">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5b02e-406">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5b02e-407">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5b02e-407">ReadItem</span></span>|
|[<span data-ttu-id="5b02e-408">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5b02e-408">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5b02e-409">Чтение</span><span class="sxs-lookup"><span data-stu-id="5b02e-409">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="5b02e-410">Пример</span><span class="sxs-lookup"><span data-stu-id="5b02e-410">Example</span></span>

<span data-ttu-id="5b02e-p120">Указанный ниже код проверяет наличие идентификатора элемента. Если свойство `itemId` возвращает значение `null` или `undefined`, элемент будет сохранен в хранилище, а из асинхронного результата будет получен идентификатор элемента.</span><span class="sxs-lookup"><span data-stu-id="5b02e-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```js
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result){
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook15officemailboxenumsitemtype"></a><span data-ttu-id="5b02e-413">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_5/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="5b02e-413">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_5/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="5b02e-414">Получает тип элемента, который представляет экземпляр.</span><span class="sxs-lookup"><span data-stu-id="5b02e-414">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="5b02e-415">Свойство `itemType` возвращает одно из значений перечисления `ItemType`, которое указывает, является ли экземпляр объекта `item` сообщением или встречей.</span><span class="sxs-lookup"><span data-stu-id="5b02e-415">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="5b02e-416">Тип:</span><span class="sxs-lookup"><span data-stu-id="5b02e-416">Type:</span></span>

*   [<span data-ttu-id="5b02e-417">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="5b02e-417">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook_1_5/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="5b02e-418">Требования</span><span class="sxs-lookup"><span data-stu-id="5b02e-418">Requirements</span></span>

|<span data-ttu-id="5b02e-419">Требование</span><span class="sxs-lookup"><span data-stu-id="5b02e-419">Requirement</span></span>| <span data-ttu-id="5b02e-420">Значение</span><span class="sxs-lookup"><span data-stu-id="5b02e-420">Value</span></span>|
|---|---|
|[<span data-ttu-id="5b02e-421">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5b02e-421">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5b02e-422">1.0</span><span class="sxs-lookup"><span data-stu-id="5b02e-422">1.0</span></span>|
|[<span data-ttu-id="5b02e-423">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5b02e-423">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5b02e-424">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5b02e-424">ReadItem</span></span>|
|[<span data-ttu-id="5b02e-425">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5b02e-425">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5b02e-426">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="5b02e-426">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="5b02e-427">Пример</span><span class="sxs-lookup"><span data-stu-id="5b02e-427">Example</span></span>

```js
if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Message)
  // do something
else
  // do something else
```

####  <a name="location-stringlocationjavascriptapioutlook15officelocation"></a><span data-ttu-id="5b02e-428">location :String|[Location](/javascript/api/outlook_1_5/office.location)</span><span class="sxs-lookup"><span data-stu-id="5b02e-428">location :String|[Location](/javascript/api/outlook_1_5/office.location)</span></span>

<span data-ttu-id="5b02e-429">Получает или задает место встречи.</span><span class="sxs-lookup"><span data-stu-id="5b02e-429">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="5b02e-430">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="5b02e-430">Read mode</span></span>

<span data-ttu-id="5b02e-431">Свойство `location` возвращает строку, содержащую сведения о месте встречи.</span><span class="sxs-lookup"><span data-stu-id="5b02e-431">The `location` property returns a string that contains the location of the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="5b02e-432">Режим создания</span><span class="sxs-lookup"><span data-stu-id="5b02e-432">Compose mode</span></span>

<span data-ttu-id="5b02e-433">Свойство `location` возвращает объект `Location`, предоставляющий методы, которые используются для получения и задания места встречи.</span><span class="sxs-lookup"><span data-stu-id="5b02e-433">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="5b02e-434">Тип:</span><span class="sxs-lookup"><span data-stu-id="5b02e-434">Type:</span></span>

*   <span data-ttu-id="5b02e-435">String | [Location](/javascript/api/outlook_1_5/office.location)</span><span class="sxs-lookup"><span data-stu-id="5b02e-435">String | [Location](/javascript/api/outlook_1_5/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="5b02e-436">Требования</span><span class="sxs-lookup"><span data-stu-id="5b02e-436">Requirements</span></span>

|<span data-ttu-id="5b02e-437">Требование</span><span class="sxs-lookup"><span data-stu-id="5b02e-437">Requirement</span></span>| <span data-ttu-id="5b02e-438">Значение</span><span class="sxs-lookup"><span data-stu-id="5b02e-438">Value</span></span>|
|---|---|
|[<span data-ttu-id="5b02e-439">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5b02e-439">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5b02e-440">1.0</span><span class="sxs-lookup"><span data-stu-id="5b02e-440">1.0</span></span>|
|[<span data-ttu-id="5b02e-441">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5b02e-441">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5b02e-442">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5b02e-442">ReadItem</span></span>|
|[<span data-ttu-id="5b02e-443">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5b02e-443">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5b02e-444">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="5b02e-444">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="5b02e-445">Пример</span><span class="sxs-lookup"><span data-stu-id="5b02e-445">Example</span></span>

```js
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

#### <a name="normalizedsubject-string"></a><span data-ttu-id="5b02e-446">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="5b02e-446">normalizedSubject :String</span></span>

<span data-ttu-id="5b02e-p121">Получает тему элемента со всеми удаленными префиксами (включая `RE:` и `FWD:`). Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="5b02e-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="5b02e-p122">Свойство normalizedSubject получает тему элемента со стандартными префиксами (такими как `RE:` и `FW:`), добавляемыми почтовыми программами. Для получения темы элемента с неизмененными префиксами используйте свойство [`subject`](#subject-stringsubjectjavascriptapioutlook15officesubject).</span><span class="sxs-lookup"><span data-stu-id="5b02e-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubjectjavascriptapioutlook15officesubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="5b02e-451">Тип:</span><span class="sxs-lookup"><span data-stu-id="5b02e-451">Type:</span></span>

*   <span data-ttu-id="5b02e-452">String</span><span class="sxs-lookup"><span data-stu-id="5b02e-452">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="5b02e-453">Требования</span><span class="sxs-lookup"><span data-stu-id="5b02e-453">Requirements</span></span>

|<span data-ttu-id="5b02e-454">Требование</span><span class="sxs-lookup"><span data-stu-id="5b02e-454">Requirement</span></span>| <span data-ttu-id="5b02e-455">Значение</span><span class="sxs-lookup"><span data-stu-id="5b02e-455">Value</span></span>|
|---|---|
|[<span data-ttu-id="5b02e-456">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5b02e-456">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5b02e-457">1.0</span><span class="sxs-lookup"><span data-stu-id="5b02e-457">1.0</span></span>|
|[<span data-ttu-id="5b02e-458">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5b02e-458">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5b02e-459">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5b02e-459">ReadItem</span></span>|
|[<span data-ttu-id="5b02e-460">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5b02e-460">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5b02e-461">Чтение</span><span class="sxs-lookup"><span data-stu-id="5b02e-461">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="5b02e-462">Пример</span><span class="sxs-lookup"><span data-stu-id="5b02e-462">Example</span></span>

```js
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
```

####  <a name="notificationmessages-notificationmessagesjavascriptapioutlook15officenotificationmessages"></a><span data-ttu-id="5b02e-463">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_5/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="5b02e-463">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_5/office.notificationmessages)</span></span>

<span data-ttu-id="5b02e-464">Получает сообщения уведомления для элемента.</span><span class="sxs-lookup"><span data-stu-id="5b02e-464">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="5b02e-465">Тип:</span><span class="sxs-lookup"><span data-stu-id="5b02e-465">Type:</span></span>

*   [<span data-ttu-id="5b02e-466">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="5b02e-466">NotificationMessages</span></span>](/javascript/api/outlook_1_5/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="5b02e-467">Требования</span><span class="sxs-lookup"><span data-stu-id="5b02e-467">Requirements</span></span>

|<span data-ttu-id="5b02e-468">Требование</span><span class="sxs-lookup"><span data-stu-id="5b02e-468">Requirement</span></span>| <span data-ttu-id="5b02e-469">Значение</span><span class="sxs-lookup"><span data-stu-id="5b02e-469">Value</span></span>|
|---|---|
|[<span data-ttu-id="5b02e-470">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="5b02e-470">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5b02e-471">1.3</span><span class="sxs-lookup"><span data-stu-id="5b02e-471">1.3</span></span>|
|[<span data-ttu-id="5b02e-472">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5b02e-472">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5b02e-473">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5b02e-473">ReadItem</span></span>|
|[<span data-ttu-id="5b02e-474">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5b02e-474">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5b02e-475">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="5b02e-475">Compose or read</span></span>|

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients"></a><span data-ttu-id="5b02e-476">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="5b02e-476">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

<span data-ttu-id="5b02e-477">Предоставляет доступ к необязательным участникам события.</span><span class="sxs-lookup"><span data-stu-id="5b02e-477">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="5b02e-478">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="5b02e-478">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="5b02e-479">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="5b02e-479">Read mode</span></span>

<span data-ttu-id="5b02e-480">Свойство `optionalAttendees` возвращает массив, содержащий объект `EmailAddressDetails` для каждого необязательного участника собрания.</span><span class="sxs-lookup"><span data-stu-id="5b02e-480">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="5b02e-481">Режим создания</span><span class="sxs-lookup"><span data-stu-id="5b02e-481">Compose mode</span></span>

<span data-ttu-id="5b02e-482">Свойство `optionalAttendees` возвращает объект `Recipients`, который предоставляет методы для получения или обновления необязательных участников собрания.</span><span class="sxs-lookup"><span data-stu-id="5b02e-482">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="5b02e-483">Тип:</span><span class="sxs-lookup"><span data-stu-id="5b02e-483">Type:</span></span>

*   <span data-ttu-id="5b02e-484">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="5b02e-484">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="5b02e-485">Требования</span><span class="sxs-lookup"><span data-stu-id="5b02e-485">Requirements</span></span>

|<span data-ttu-id="5b02e-486">Требование</span><span class="sxs-lookup"><span data-stu-id="5b02e-486">Requirement</span></span>| <span data-ttu-id="5b02e-487">Значение</span><span class="sxs-lookup"><span data-stu-id="5b02e-487">Value</span></span>|
|---|---|
|[<span data-ttu-id="5b02e-488">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5b02e-488">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5b02e-489">1.0</span><span class="sxs-lookup"><span data-stu-id="5b02e-489">1.0</span></span>|
|[<span data-ttu-id="5b02e-490">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5b02e-490">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5b02e-491">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5b02e-491">ReadItem</span></span>|
|[<span data-ttu-id="5b02e-492">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5b02e-492">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5b02e-493">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="5b02e-493">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="5b02e-494">Пример</span><span class="sxs-lookup"><span data-stu-id="5b02e-494">Example</span></span>

```js
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails"></a><span data-ttu-id="5b02e-495">organizer :[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="5b02e-495">organizer :[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)</span></span>

<span data-ttu-id="5b02e-p124">Получает электронный адрес организатора указанного собрания. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="5b02e-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="5b02e-498">Тип:</span><span class="sxs-lookup"><span data-stu-id="5b02e-498">Type:</span></span>

*   [<span data-ttu-id="5b02e-499">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="5b02e-499">EmailAddressDetails</span></span>](/javascript/api/outlook_1_5/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="5b02e-500">Требования</span><span class="sxs-lookup"><span data-stu-id="5b02e-500">Requirements</span></span>

|<span data-ttu-id="5b02e-501">Требование</span><span class="sxs-lookup"><span data-stu-id="5b02e-501">Requirement</span></span>| <span data-ttu-id="5b02e-502">Значение</span><span class="sxs-lookup"><span data-stu-id="5b02e-502">Value</span></span>|
|---|---|
|[<span data-ttu-id="5b02e-503">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5b02e-503">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5b02e-504">1.0</span><span class="sxs-lookup"><span data-stu-id="5b02e-504">1.0</span></span>|
|[<span data-ttu-id="5b02e-505">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5b02e-505">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5b02e-506">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5b02e-506">ReadItem</span></span>|
|[<span data-ttu-id="5b02e-507">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5b02e-507">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5b02e-508">Чтение</span><span class="sxs-lookup"><span data-stu-id="5b02e-508">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="5b02e-509">Пример</span><span class="sxs-lookup"><span data-stu-id="5b02e-509">Example</span></span>

```js
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
```

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients"></a><span data-ttu-id="5b02e-510">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="5b02e-510">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

<span data-ttu-id="5b02e-511">Предоставляет доступ к обязательным участникам события.</span><span class="sxs-lookup"><span data-stu-id="5b02e-511">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="5b02e-512">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="5b02e-512">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="5b02e-513">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="5b02e-513">Read mode</span></span>

<span data-ttu-id="5b02e-514">Свойство `requiredAttendees` возвращает массив, содержащий объект `EmailAddressDetails` для каждого обязательного участника собрания.</span><span class="sxs-lookup"><span data-stu-id="5b02e-514">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="5b02e-515">Режим создания</span><span class="sxs-lookup"><span data-stu-id="5b02e-515">Compose mode</span></span>

<span data-ttu-id="5b02e-516">Свойство `requiredAttendees` возвращает объект `Recipients`, предоставляющий методы, с помощью которых можно получить или обновить сведения об обязательных участниках собрания.</span><span class="sxs-lookup"><span data-stu-id="5b02e-516">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="5b02e-517">Тип:</span><span class="sxs-lookup"><span data-stu-id="5b02e-517">Type:</span></span>

*   <span data-ttu-id="5b02e-518">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="5b02e-518">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="5b02e-519">Требования</span><span class="sxs-lookup"><span data-stu-id="5b02e-519">Requirements</span></span>

|<span data-ttu-id="5b02e-520">Требование</span><span class="sxs-lookup"><span data-stu-id="5b02e-520">Requirement</span></span>| <span data-ttu-id="5b02e-521">Значение</span><span class="sxs-lookup"><span data-stu-id="5b02e-521">Value</span></span>|
|---|---|
|[<span data-ttu-id="5b02e-522">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5b02e-522">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5b02e-523">1.0</span><span class="sxs-lookup"><span data-stu-id="5b02e-523">1.0</span></span>|
|[<span data-ttu-id="5b02e-524">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5b02e-524">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5b02e-525">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5b02e-525">ReadItem</span></span>|
|[<span data-ttu-id="5b02e-526">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5b02e-526">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5b02e-527">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="5b02e-527">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="5b02e-528">Пример</span><span class="sxs-lookup"><span data-stu-id="5b02e-528">Example</span></span>

```js
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
}
```

#### <a name="sender-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails"></a><span data-ttu-id="5b02e-529">sender :[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="5b02e-529">sender :[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)</span></span>

<span data-ttu-id="5b02e-p126">Получает электронный адрес отправителя электронного сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="5b02e-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="5b02e-p127">Свойства [`from`](#from-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails) и `sender` представляют одно лицо, если сообщение не отправлено представителем. В противном случае свойство `from` представляет лицо, делегировавшее полномочия, а свойство sender — представителя.</span><span class="sxs-lookup"><span data-stu-id="5b02e-p127">The [`from`](#from-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="5b02e-534">Свойству `recipientType`, принадлежащему объекту `EmailAddressDetails` в свойстве `sender`, задано значение `undefined`.</span><span class="sxs-lookup"><span data-stu-id="5b02e-534">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="5b02e-535">Тип:</span><span class="sxs-lookup"><span data-stu-id="5b02e-535">Type:</span></span>

*   [<span data-ttu-id="5b02e-536">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="5b02e-536">EmailAddressDetails</span></span>](/javascript/api/outlook_1_5/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="5b02e-537">Требования</span><span class="sxs-lookup"><span data-stu-id="5b02e-537">Requirements</span></span>

|<span data-ttu-id="5b02e-538">Требование</span><span class="sxs-lookup"><span data-stu-id="5b02e-538">Requirement</span></span>| <span data-ttu-id="5b02e-539">Значение</span><span class="sxs-lookup"><span data-stu-id="5b02e-539">Value</span></span>|
|---|---|
|[<span data-ttu-id="5b02e-540">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5b02e-540">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5b02e-541">1.0</span><span class="sxs-lookup"><span data-stu-id="5b02e-541">1.0</span></span>|
|[<span data-ttu-id="5b02e-542">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5b02e-542">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5b02e-543">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5b02e-543">ReadItem</span></span>|
|[<span data-ttu-id="5b02e-544">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5b02e-544">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5b02e-545">Чтение</span><span class="sxs-lookup"><span data-stu-id="5b02e-545">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="5b02e-546">Пример</span><span class="sxs-lookup"><span data-stu-id="5b02e-546">Example</span></span>

```js
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
```

####  <a name="start-datetimejavascriptapioutlook15officetime"></a><span data-ttu-id="5b02e-547">start :Date|[Time](/javascript/api/outlook_1_5/office.time)</span><span class="sxs-lookup"><span data-stu-id="5b02e-547">start :Date|[Time](/javascript/api/outlook_1_5/office.time)</span></span>

<span data-ttu-id="5b02e-548">Получает или задает дату и время начала встречи.</span><span class="sxs-lookup"><span data-stu-id="5b02e-548">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="5b02e-p128">Свойство `start` представлено в виде значения даты и времени в формате UTC. Это значение можно преобразовать в местные значения даты и времени клиента с помощью метода [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook15officelocalclienttime).</span><span class="sxs-lookup"><span data-stu-id="5b02e-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook15officelocalclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="5b02e-551">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="5b02e-551">Read mode</span></span>

<span data-ttu-id="5b02e-552">Свойство `start` возвращает объект `Date`.</span><span class="sxs-lookup"><span data-stu-id="5b02e-552">The `start` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="5b02e-553">Режим создания</span><span class="sxs-lookup"><span data-stu-id="5b02e-553">Compose mode</span></span>

<span data-ttu-id="5b02e-554">Свойство `start` возвращает объект `Time`.</span><span class="sxs-lookup"><span data-stu-id="5b02e-554">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="5b02e-555">Если вы задаете время начала с помощью метода [`Time.setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-), необходимо использовать метод [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) для преобразования местного времени на клиенте в формат UTC для сервера.</span><span class="sxs-lookup"><span data-stu-id="5b02e-555">When you use the [`Time.setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="5b02e-556">Тип:</span><span class="sxs-lookup"><span data-stu-id="5b02e-556">Type:</span></span>

*   <span data-ttu-id="5b02e-557">Date | [Time](/javascript/api/outlook_1_5/office.time)</span><span class="sxs-lookup"><span data-stu-id="5b02e-557">Date | [Time](/javascript/api/outlook_1_5/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="5b02e-558">Требования</span><span class="sxs-lookup"><span data-stu-id="5b02e-558">Requirements</span></span>

|<span data-ttu-id="5b02e-559">Требование</span><span class="sxs-lookup"><span data-stu-id="5b02e-559">Requirement</span></span>| <span data-ttu-id="5b02e-560">Значение</span><span class="sxs-lookup"><span data-stu-id="5b02e-560">Value</span></span>|
|---|---|
|[<span data-ttu-id="5b02e-561">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5b02e-561">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5b02e-562">1.0</span><span class="sxs-lookup"><span data-stu-id="5b02e-562">1.0</span></span>|
|[<span data-ttu-id="5b02e-563">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5b02e-563">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5b02e-564">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5b02e-564">ReadItem</span></span>|
|[<span data-ttu-id="5b02e-565">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5b02e-565">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5b02e-566">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="5b02e-566">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="5b02e-567">Пример</span><span class="sxs-lookup"><span data-stu-id="5b02e-567">Example</span></span>

<span data-ttu-id="5b02e-568">В примере ниже с помощью метода [`setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) объекта `Time` задается время начала встречи в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="5b02e-568">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

####  <a name="subject-stringsubjectjavascriptapioutlook15officesubject"></a><span data-ttu-id="5b02e-569">subject :String|[Subject](/javascript/api/outlook_1_5/office.subject)</span><span class="sxs-lookup"><span data-stu-id="5b02e-569">subject :String|[Subject](/javascript/api/outlook_1_5/office.subject)</span></span>

<span data-ttu-id="5b02e-570">Получает или задает описание, которое отображается в поле темы элемента.</span><span class="sxs-lookup"><span data-stu-id="5b02e-570">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="5b02e-571">Свойство `subject` получает или задает всю тему элемента для отправки с почтового сервера.</span><span class="sxs-lookup"><span data-stu-id="5b02e-571">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="5b02e-572">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="5b02e-572">Read mode</span></span>

<span data-ttu-id="5b02e-p129">Свойство `subject` возвращает строку. С помощью свойства [`normalizedSubject`](#normalizedsubject-string) можно получить тему без начальных префиксов, таких как `RE:` и `FW:`.</span><span class="sxs-lookup"><span data-stu-id="5b02e-p129">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```js
var subject = Office.context.mailbox.item.subject;
```

##### <a name="compose-mode"></a><span data-ttu-id="5b02e-575">Режим создания</span><span class="sxs-lookup"><span data-stu-id="5b02e-575">Compose mode</span></span>

<span data-ttu-id="5b02e-576">Свойство `subject` возвращает объект `Subject`, который предоставляет методы для получения и задания темы.</span><span class="sxs-lookup"><span data-stu-id="5b02e-576">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```js
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="5b02e-577">Тип:</span><span class="sxs-lookup"><span data-stu-id="5b02e-577">Type:</span></span>

*   <span data-ttu-id="5b02e-578">String | [Subject](/javascript/api/outlook_1_5/office.subject)</span><span class="sxs-lookup"><span data-stu-id="5b02e-578">String | [Subject](/javascript/api/outlook_1_5/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="5b02e-579">Требования</span><span class="sxs-lookup"><span data-stu-id="5b02e-579">Requirements</span></span>

|<span data-ttu-id="5b02e-580">Требование</span><span class="sxs-lookup"><span data-stu-id="5b02e-580">Requirement</span></span>| <span data-ttu-id="5b02e-581">Значение</span><span class="sxs-lookup"><span data-stu-id="5b02e-581">Value</span></span>|
|---|---|
|[<span data-ttu-id="5b02e-582">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5b02e-582">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5b02e-583">1.0</span><span class="sxs-lookup"><span data-stu-id="5b02e-583">1.0</span></span>|
|[<span data-ttu-id="5b02e-584">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5b02e-584">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5b02e-585">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5b02e-585">ReadItem</span></span>|
|[<span data-ttu-id="5b02e-586">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5b02e-586">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5b02e-587">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="5b02e-587">Compose or read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients"></a><span data-ttu-id="5b02e-588">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="5b02e-588">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

<span data-ttu-id="5b02e-589">Предоставляет доступ к получателям, указанным в строке **Кому** сообщения.</span><span class="sxs-lookup"><span data-stu-id="5b02e-589">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="5b02e-590">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="5b02e-590">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="5b02e-591">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="5b02e-591">Read mode</span></span>

<span data-ttu-id="5b02e-p131">Свойство `to` возвращает массив, содержащий объект `EmailAddressDetails` для каждого получателя в строке **Кому** сообщения. Коллекция может включать не более 100 элементов.</span><span class="sxs-lookup"><span data-stu-id="5b02e-p131">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="5b02e-594">Режим создания</span><span class="sxs-lookup"><span data-stu-id="5b02e-594">Compose mode</span></span>

<span data-ttu-id="5b02e-595">Свойство `to` возвращает объект `Recipients`, предоставляющий методы для получения или обновления получателей, которые указаны в строке **Кому** сообщения.</span><span class="sxs-lookup"><span data-stu-id="5b02e-595">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="5b02e-596">Тип:</span><span class="sxs-lookup"><span data-stu-id="5b02e-596">Type:</span></span>

*   <span data-ttu-id="5b02e-597">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="5b02e-597">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="5b02e-598">Требования</span><span class="sxs-lookup"><span data-stu-id="5b02e-598">Requirements</span></span>

|<span data-ttu-id="5b02e-599">Требование</span><span class="sxs-lookup"><span data-stu-id="5b02e-599">Requirement</span></span>| <span data-ttu-id="5b02e-600">Значение</span><span class="sxs-lookup"><span data-stu-id="5b02e-600">Value</span></span>|
|---|---|
|[<span data-ttu-id="5b02e-601">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5b02e-601">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5b02e-602">1.0</span><span class="sxs-lookup"><span data-stu-id="5b02e-602">1.0</span></span>|
|[<span data-ttu-id="5b02e-603">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5b02e-603">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5b02e-604">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5b02e-604">ReadItem</span></span>|
|[<span data-ttu-id="5b02e-605">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5b02e-605">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5b02e-606">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="5b02e-606">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="5b02e-607">Пример</span><span class="sxs-lookup"><span data-stu-id="5b02e-607">Example</span></span>

```js
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

### <a name="methods"></a><span data-ttu-id="5b02e-608">Методы</span><span class="sxs-lookup"><span data-stu-id="5b02e-608">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="5b02e-609">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="5b02e-609">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="5b02e-610">Добавляет файл в сообщение или встречу в качестве вложения.</span><span class="sxs-lookup"><span data-stu-id="5b02e-610">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="5b02e-611">Метод `addFileAttachmentAsync` передает файл по указанному универсальному коду ресурса (URI) и вкладывает его в элемент в форме создания.</span><span class="sxs-lookup"><span data-stu-id="5b02e-611">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="5b02e-612">Идентификатор можно последовательно использовать с методом [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback), чтобы удалить вложение, добавленное во время текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="5b02e-612">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="5b02e-613">Параметры</span><span class="sxs-lookup"><span data-stu-id="5b02e-613">Parameters:</span></span>

|<span data-ttu-id="5b02e-614">Имя</span><span class="sxs-lookup"><span data-stu-id="5b02e-614">Name</span></span>| <span data-ttu-id="5b02e-615">Тип</span><span class="sxs-lookup"><span data-stu-id="5b02e-615">Type</span></span>| <span data-ttu-id="5b02e-616">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="5b02e-616">Attributes</span></span>| <span data-ttu-id="5b02e-617">Описание</span><span class="sxs-lookup"><span data-stu-id="5b02e-617">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="5b02e-618">String</span><span class="sxs-lookup"><span data-stu-id="5b02e-618">String</span></span>||<span data-ttu-id="5b02e-p132">Универсальный код ресурса (URI), представляющий расположение файла, который нужно вложить в сообщение или встречу. Максимальная длина — 2048 символов.</span><span class="sxs-lookup"><span data-stu-id="5b02e-p132">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="5b02e-621">String</span><span class="sxs-lookup"><span data-stu-id="5b02e-621">String</span></span>||<span data-ttu-id="5b02e-p133">Имя вложения, которое отображается при передаче вложения. Максимальная длина — 255 символов.</span><span class="sxs-lookup"><span data-stu-id="5b02e-p133">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="5b02e-624">Object</span><span class="sxs-lookup"><span data-stu-id="5b02e-624">Object</span></span>| <span data-ttu-id="5b02e-625">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="5b02e-625">&lt;optional&gt;</span></span>|<span data-ttu-id="5b02e-626">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="5b02e-626">An object literal that contains one or more of the following properties.</span></span>|
| `options.asyncContext` | <span data-ttu-id="5b02e-627">Object</span><span class="sxs-lookup"><span data-stu-id="5b02e-627">Object</span></span> | <span data-ttu-id="5b02e-628">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="5b02e-628">&lt;optional&gt;</span></span> | <span data-ttu-id="5b02e-629">В методе обратного вызова разработчики могут указать любой объект, к которому необходимо получить доступ.</span><span class="sxs-lookup"><span data-stu-id="5b02e-629">Developers can provide any object they wish to access in the callback method.</span></span> |
| `options.isInline` | <span data-ttu-id="5b02e-630">Boolean</span><span class="sxs-lookup"><span data-stu-id="5b02e-630">Boolean</span></span> | <span data-ttu-id="5b02e-631">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="5b02e-631">&lt;optional&gt;</span></span> | <span data-ttu-id="5b02e-632">Значение `true` указывает на то, что вложение будет встроено в текст сообщения и не должно отображаться в списке вложений.</span><span class="sxs-lookup"><span data-stu-id="5b02e-632">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
|`callback`| <span data-ttu-id="5b02e-633">function</span><span class="sxs-lookup"><span data-stu-id="5b02e-633">function</span></span>| <span data-ttu-id="5b02e-634">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="5b02e-634">&lt;optional&gt;</span></span>|<span data-ttu-id="5b02e-635">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="5b02e-635">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="5b02e-636">После успешного выполнения идентификатор вложения будет представлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="5b02e-636">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="5b02e-637">Если добавить вложение не удастся, объект `asyncResult` будет содержать объект `Error` с описанием ошибки.</span><span class="sxs-lookup"><span data-stu-id="5b02e-637">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="5b02e-638">Ошибки</span><span class="sxs-lookup"><span data-stu-id="5b02e-638">Errors</span></span>

| <span data-ttu-id="5b02e-639">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="5b02e-639">Error code</span></span> | <span data-ttu-id="5b02e-640">Описание</span><span class="sxs-lookup"><span data-stu-id="5b02e-640">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="5b02e-641">Вложение превышает максимальный размер.</span><span class="sxs-lookup"><span data-stu-id="5b02e-641">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="5b02e-642">Расширение вложения не поддерживается.</span><span class="sxs-lookup"><span data-stu-id="5b02e-642">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="5b02e-643">Сообщение или встреча содержат слишком много вложений.</span><span class="sxs-lookup"><span data-stu-id="5b02e-643">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="5b02e-644">Требования</span><span class="sxs-lookup"><span data-stu-id="5b02e-644">Requirements</span></span>

|<span data-ttu-id="5b02e-645">Требование</span><span class="sxs-lookup"><span data-stu-id="5b02e-645">Requirement</span></span>| <span data-ttu-id="5b02e-646">Значение</span><span class="sxs-lookup"><span data-stu-id="5b02e-646">Value</span></span>|
|---|---|
|[<span data-ttu-id="5b02e-647">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5b02e-647">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5b02e-648">1.1</span><span class="sxs-lookup"><span data-stu-id="5b02e-648">1.1</span></span>|
|[<span data-ttu-id="5b02e-649">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5b02e-649">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5b02e-650">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="5b02e-650">ReadWriteItem</span></span>|
|[<span data-ttu-id="5b02e-651">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5b02e-651">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5b02e-652">Создание</span><span class="sxs-lookup"><span data-stu-id="5b02e-652">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="5b02e-653">Примеры</span><span class="sxs-lookup"><span data-stu-id="5b02e-653">Examples</span></span>

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

<span data-ttu-id="5b02e-654">В приведенном ниже примере файл изображения добавляется в качестве встроенного вложения, а ссылка на вложение добавляется в текст сообщения.</span><span class="sxs-lookup"><span data-stu-id="5b02e-654">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

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

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="5b02e-655">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="5b02e-655">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="5b02e-656">Добавляет к сообщению элемент Exchange, например сообщение, в виде вложения.</span><span class="sxs-lookup"><span data-stu-id="5b02e-656">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="5b02e-p134">С помощью метода `addItemAttachmentAsync` можно в элемент формы создания вложить элемент с указанным идентификатором Exchange. Если указать метод обратного вызова, то этот метод вызывается с помощью параметра `asyncResult`, который содержит идентификатор вложения или код, указывающий на ошибки, которые произошли при вложении элемента. При необходимости можно использовать параметр `options` для передачи сведений о состоянии методу обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="5b02e-p134">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="5b02e-660">Идентификатор можно использовать с методом [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback), чтобы удалить вложение, добавленное во время текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="5b02e-660">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="5b02e-661">Если ваша надстройка Office выполняется в Outlook Web App, метод `addItemAttachmentAsync` обеспечивает вложение элементов в элементы, отличные от редактируемого. Однако это действие не рекомендуем выполнять, так как оно не поддерживается.</span><span class="sxs-lookup"><span data-stu-id="5b02e-661">If your Office Add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="5b02e-662">Параметры:</span><span class="sxs-lookup"><span data-stu-id="5b02e-662">Parameters:</span></span>

|<span data-ttu-id="5b02e-663">Имя</span><span class="sxs-lookup"><span data-stu-id="5b02e-663">Name</span></span>| <span data-ttu-id="5b02e-664">Тип</span><span class="sxs-lookup"><span data-stu-id="5b02e-664">Type</span></span>| <span data-ttu-id="5b02e-665">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="5b02e-665">Attributes</span></span>| <span data-ttu-id="5b02e-666">Описание</span><span class="sxs-lookup"><span data-stu-id="5b02e-666">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="5b02e-667">String</span><span class="sxs-lookup"><span data-stu-id="5b02e-667">String</span></span>||<span data-ttu-id="5b02e-p135">Идентификатор Exchange для вкладываемого элемента. Максимальная длина — 100 символов.</span><span class="sxs-lookup"><span data-stu-id="5b02e-p135">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="5b02e-670">String</span><span class="sxs-lookup"><span data-stu-id="5b02e-670">String</span></span>||<span data-ttu-id="5b02e-p136">Тема вкладываемого элемента. Максимальная длина — 255 символов.</span><span class="sxs-lookup"><span data-stu-id="5b02e-p136">The sujbect of the item to be attached. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="5b02e-673">Object</span><span class="sxs-lookup"><span data-stu-id="5b02e-673">Object</span></span>| <span data-ttu-id="5b02e-674">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="5b02e-674">&lt;optional&gt;</span></span>|<span data-ttu-id="5b02e-675">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="5b02e-675">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="5b02e-676">Object</span><span class="sxs-lookup"><span data-stu-id="5b02e-676">Object</span></span>| <span data-ttu-id="5b02e-677">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="5b02e-677">&lt;optional&gt;</span></span>|<span data-ttu-id="5b02e-678">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="5b02e-678">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="5b02e-679">функция</span><span class="sxs-lookup"><span data-stu-id="5b02e-679">function</span></span>| <span data-ttu-id="5b02e-680">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="5b02e-680">&lt;optional&gt;</span></span>|<span data-ttu-id="5b02e-681">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="5b02e-681">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="5b02e-682">После успешного выполнения идентификатор вложения будет представлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="5b02e-682">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="5b02e-683">Если добавить вложение не удастся, объект `asyncResult` будет содержать объект `Error` с описанием ошибки.</span><span class="sxs-lookup"><span data-stu-id="5b02e-683">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="5b02e-684">Ошибки</span><span class="sxs-lookup"><span data-stu-id="5b02e-684">Errors</span></span>

| <span data-ttu-id="5b02e-685">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="5b02e-685">Error code</span></span> | <span data-ttu-id="5b02e-686">Описание</span><span class="sxs-lookup"><span data-stu-id="5b02e-686">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="5b02e-687">Сообщение или встреча содержат слишком много вложений.</span><span class="sxs-lookup"><span data-stu-id="5b02e-687">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="5b02e-688">Требования</span><span class="sxs-lookup"><span data-stu-id="5b02e-688">Requirements</span></span>

|<span data-ttu-id="5b02e-689">Требование</span><span class="sxs-lookup"><span data-stu-id="5b02e-689">Requirement</span></span>| <span data-ttu-id="5b02e-690">Значение</span><span class="sxs-lookup"><span data-stu-id="5b02e-690">Value</span></span>|
|---|---|
|[<span data-ttu-id="5b02e-691">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5b02e-691">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5b02e-692">1.1</span><span class="sxs-lookup"><span data-stu-id="5b02e-692">1.1</span></span>|
|[<span data-ttu-id="5b02e-693">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5b02e-693">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5b02e-694">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="5b02e-694">ReadWriteItem</span></span>|
|[<span data-ttu-id="5b02e-695">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5b02e-695">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5b02e-696">Создание</span><span class="sxs-lookup"><span data-stu-id="5b02e-696">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="5b02e-697">Пример</span><span class="sxs-lookup"><span data-stu-id="5b02e-697">Example</span></span>

<span data-ttu-id="5b02e-698">В следующем примере существующий элемент Outlook добавляется в виде вложения с именем `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="5b02e-698">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

####  <a name="close"></a><span data-ttu-id="5b02e-699">close()</span><span class="sxs-lookup"><span data-stu-id="5b02e-699">close()</span></span>

<span data-ttu-id="5b02e-700">Закрывает текущий создаваемый элемент.</span><span class="sxs-lookup"><span data-stu-id="5b02e-700">Closes the current item that is being composed.</span></span>

<span data-ttu-id="5b02e-p137">Работа метода `close` зависит от текущего состояния создаваемого элемента. Если элемент содержит несохраненные изменения, клиент предложит пользователю сохранить или отклонить их либо отменить действие закрытия.</span><span class="sxs-lookup"><span data-stu-id="5b02e-p137">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="5b02e-703">Если в Outlook в Интернете элемент представляет собой встречу, ранее сохраненную с помощью метода `saveAsync`, пользователю предлагается сохранить, отклонить или отменить действие, даже если с момента последнего сохранения элемента не вносились какие-либо изменения.</span><span class="sxs-lookup"><span data-stu-id="5b02e-703">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="5b02e-704">Если в клиенте Outlook для настольных ПК сообщение представляет собой ответ в тексте, метод `close` не работает.</span><span class="sxs-lookup"><span data-stu-id="5b02e-704">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="5b02e-705">Требования</span><span class="sxs-lookup"><span data-stu-id="5b02e-705">Requirements</span></span>

|<span data-ttu-id="5b02e-706">Требование</span><span class="sxs-lookup"><span data-stu-id="5b02e-706">Requirement</span></span>| <span data-ttu-id="5b02e-707">Значение</span><span class="sxs-lookup"><span data-stu-id="5b02e-707">Value</span></span>|
|---|---|
|[<span data-ttu-id="5b02e-708">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="5b02e-708">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5b02e-709">1.3</span><span class="sxs-lookup"><span data-stu-id="5b02e-709">1.3</span></span>|
|[<span data-ttu-id="5b02e-710">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5b02e-710">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5b02e-711">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="5b02e-711">Restricted</span></span>|
|[<span data-ttu-id="5b02e-712">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5b02e-712">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5b02e-713">Создание</span><span class="sxs-lookup"><span data-stu-id="5b02e-713">Compose</span></span>|

#### <a name="displayreplyallformformdata"></a><span data-ttu-id="5b02e-714">displayReplyAllForm(formData)</span><span class="sxs-lookup"><span data-stu-id="5b02e-714">displayReplyAllForm(formData)</span></span>

<span data-ttu-id="5b02e-715">Отображает форму ответа, включающую отправителя и всех получателей выбранного сообщения или организатора и всех участников выбранной встречи.</span><span class="sxs-lookup"><span data-stu-id="5b02e-715">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="5b02e-716">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="5b02e-716">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="5b02e-717">В Outlook Web App форма ответа отображается в виде всплывающей формы в представлении с 3 либо 1 или 2 колонками.</span><span class="sxs-lookup"><span data-stu-id="5b02e-717">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="5b02e-718">Если любой строковый параметр превышает указанные для него ограничения, `displayReplyAllForm` возвращает исключение.</span><span class="sxs-lookup"><span data-stu-id="5b02e-718">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="5b02e-p138">Если в параметре `formData.attachments` указаны вложения, Outlook и Outlook Web App пытаются скачать их и вложить в форму ответа. Если какие-либо вложения добавить не удается, в форме отображается сообщение об ошибке. Если сообщения об ошибках не предусмотрены, то они не отображаются.</span><span class="sxs-lookup"><span data-stu-id="5b02e-p138">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="5b02e-722">Параметры</span><span class="sxs-lookup"><span data-stu-id="5b02e-722">Parameters:</span></span>

| <span data-ttu-id="5b02e-723">Имя</span><span class="sxs-lookup"><span data-stu-id="5b02e-723">Name</span></span> | <span data-ttu-id="5b02e-724">Тип</span><span class="sxs-lookup"><span data-stu-id="5b02e-724">Type</span></span> | <span data-ttu-id="5b02e-725">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="5b02e-725">Attributes</span></span> | <span data-ttu-id="5b02e-726">Описание</span><span class="sxs-lookup"><span data-stu-id="5b02e-726">Description</span></span> |
|---|---|---|---|
|`formData`| <span data-ttu-id="5b02e-727">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="5b02e-727">String &#124; Object</span></span>| |<span data-ttu-id="5b02e-p139">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="5b02e-p139">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="5b02e-730">**ИЛИ**</span><span class="sxs-lookup"><span data-stu-id="5b02e-730">**OR**</span></span><br/><span data-ttu-id="5b02e-p140">Объект, который содержит текст или данные вложения и функцию обратного вызова. Ниже представлено определение этого объекта.</span><span class="sxs-lookup"><span data-stu-id="5b02e-p140">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="5b02e-733">String</span><span class="sxs-lookup"><span data-stu-id="5b02e-733">String</span></span> | <span data-ttu-id="5b02e-734">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="5b02e-734">&lt;optional&gt;</span></span> | <span data-ttu-id="5b02e-p141">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="5b02e-p141">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="5b02e-737">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="5b02e-737">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="5b02e-738">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="5b02e-738">&lt;optional&gt;</span></span> | <span data-ttu-id="5b02e-739">Массив объектов JSON, представляющих собой вложенные файлы или элементы.</span><span class="sxs-lookup"><span data-stu-id="5b02e-739">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="5b02e-740">String</span><span class="sxs-lookup"><span data-stu-id="5b02e-740">String</span></span> | | <span data-ttu-id="5b02e-p142">Указывает тип вложения. Допустимые значения: `file` для вложенного файла и `item` для вложенного элемента.</span><span class="sxs-lookup"><span data-stu-id="5b02e-p142">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="5b02e-743">String</span><span class="sxs-lookup"><span data-stu-id="5b02e-743">String</span></span> | | <span data-ttu-id="5b02e-744">Строка, содержащая имя вложения, длиной до 255 символов.</span><span class="sxs-lookup"><span data-stu-id="5b02e-744">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="5b02e-745">String</span><span class="sxs-lookup"><span data-stu-id="5b02e-745">String</span></span> | | <span data-ttu-id="5b02e-p143">Используется, только если свойству `type` задано значение `file`. URI расположения файла.</span><span class="sxs-lookup"><span data-stu-id="5b02e-p143">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.isInline` | <span data-ttu-id="5b02e-748">Логический</span><span class="sxs-lookup"><span data-stu-id="5b02e-748">Boolean</span></span> | | <span data-ttu-id="5b02e-p144">Используется, только если свойству `type` задано значение `file`. Значение `true` указывает на то, что вложение будет встроено в текст сообщения и не должно отображаться в списке вложений.</span><span class="sxs-lookup"><span data-stu-id="5b02e-p144">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="5b02e-751">String</span><span class="sxs-lookup"><span data-stu-id="5b02e-751">String</span></span> | | <span data-ttu-id="5b02e-p145">Используется, только если свойству `type` задано значение `item`. Идентификатор вложения EWS. Это строка длиной до 100 символов.</span><span class="sxs-lookup"><span data-stu-id="5b02e-p145">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="5b02e-755">function</span><span class="sxs-lookup"><span data-stu-id="5b02e-755">function</span></span> | <span data-ttu-id="5b02e-756">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="5b02e-756">&lt;optional&gt;</span></span> | <span data-ttu-id="5b02e-757">По завершении работы метода функция, переданная параметру `callback`, вызывается с помощью одного параметра `asyncResult`, представляющего собой объект [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="5b02e-757">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="5b02e-758">Требования</span><span class="sxs-lookup"><span data-stu-id="5b02e-758">Requirements</span></span>

|<span data-ttu-id="5b02e-759">Требование</span><span class="sxs-lookup"><span data-stu-id="5b02e-759">Requirement</span></span>| <span data-ttu-id="5b02e-760">Значение</span><span class="sxs-lookup"><span data-stu-id="5b02e-760">Value</span></span>|
|---|---|
|[<span data-ttu-id="5b02e-761">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5b02e-761">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5b02e-762">1.0</span><span class="sxs-lookup"><span data-stu-id="5b02e-762">1.0</span></span>|
|[<span data-ttu-id="5b02e-763">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5b02e-763">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5b02e-764">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5b02e-764">ReadItem</span></span>|
|[<span data-ttu-id="5b02e-765">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5b02e-765">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5b02e-766">Чтение</span><span class="sxs-lookup"><span data-stu-id="5b02e-766">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="5b02e-767">Примеры</span><span class="sxs-lookup"><span data-stu-id="5b02e-767">Examples</span></span>

<span data-ttu-id="5b02e-768">Приведенный ниже код передает строку в функцию `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="5b02e-768">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="5b02e-769">Ответ с пустым текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="5b02e-769">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="5b02e-770">Ответ только с текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="5b02e-770">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="5b02e-771">Ответ с текстом сообщения и вложенным файлом.</span><span class="sxs-lookup"><span data-stu-id="5b02e-771">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="5b02e-772">Ответ с текстом сообщения и вложенным элементом.</span><span class="sxs-lookup"><span data-stu-id="5b02e-772">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="5b02e-773">Ответ с текстом сообщения, вложенным файлом, вложенным элементом и обратным вызовом.</span><span class="sxs-lookup"><span data-stu-id="5b02e-773">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata"></a><span data-ttu-id="5b02e-774">displayReplyForm(formData)</span><span class="sxs-lookup"><span data-stu-id="5b02e-774">displayReplyForm(formData)</span></span>

<span data-ttu-id="5b02e-775">Отображает форму ответа, включающую только отправителя выбранного сообщения или организатора выбранной встречи.</span><span class="sxs-lookup"><span data-stu-id="5b02e-775">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="5b02e-776">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="5b02e-776">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="5b02e-777">В Outlook Web App форма ответа отображается в виде всплывающей формы в представлении с 3 либо 1 или 2 колонками.</span><span class="sxs-lookup"><span data-stu-id="5b02e-777">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="5b02e-778">Если любой строковый параметр превышает указанные для него ограничения, `displayReplyForm` возвращает исключение.</span><span class="sxs-lookup"><span data-stu-id="5b02e-778">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="5b02e-p146">Если в параметре `formData.attachments` указаны вложения, Outlook и Outlook Web App пытаются скачать их и вложить в форму ответа. Если какие-либо вложения добавить не удается, в форме отображается сообщение об ошибке. Если сообщения об ошибках не предусмотрены, то они не отображаются.</span><span class="sxs-lookup"><span data-stu-id="5b02e-p146">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="5b02e-782">Параметры</span><span class="sxs-lookup"><span data-stu-id="5b02e-782">Parameters:</span></span>

| <span data-ttu-id="5b02e-783">Имя</span><span class="sxs-lookup"><span data-stu-id="5b02e-783">Name</span></span> | <span data-ttu-id="5b02e-784">Тип</span><span class="sxs-lookup"><span data-stu-id="5b02e-784">Type</span></span> | <span data-ttu-id="5b02e-785">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="5b02e-785">Attributes</span></span> | <span data-ttu-id="5b02e-786">Описание</span><span class="sxs-lookup"><span data-stu-id="5b02e-786">Description</span></span> |
|---|---|---|---|
|`formData`| <span data-ttu-id="5b02e-787">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="5b02e-787">String &#124; Object</span></span>| | <span data-ttu-id="5b02e-p147">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="5b02e-p147">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="5b02e-790">**ИЛИ**</span><span class="sxs-lookup"><span data-stu-id="5b02e-790">**OR**</span></span><br/><span data-ttu-id="5b02e-p148">Объект, который содержит текст или данные вложения и функцию обратного вызова. Ниже представлено определение этого объекта.</span><span class="sxs-lookup"><span data-stu-id="5b02e-p148">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="5b02e-793">String</span><span class="sxs-lookup"><span data-stu-id="5b02e-793">String</span></span> | <span data-ttu-id="5b02e-794">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="5b02e-794">&lt;optional&gt;</span></span> | <span data-ttu-id="5b02e-p149">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="5b02e-p149">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="5b02e-797">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="5b02e-797">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="5b02e-798">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="5b02e-798">&lt;optional&gt;</span></span> | <span data-ttu-id="5b02e-799">Массив объектов JSON, представляющих собой вложенные файлы или элементы.</span><span class="sxs-lookup"><span data-stu-id="5b02e-799">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="5b02e-800">String</span><span class="sxs-lookup"><span data-stu-id="5b02e-800">String</span></span> | | <span data-ttu-id="5b02e-p150">Указывает тип вложения. Допустимые значения: `file` для вложенного файла и `item` для вложенного элемента.</span><span class="sxs-lookup"><span data-stu-id="5b02e-p150">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="5b02e-803">String</span><span class="sxs-lookup"><span data-stu-id="5b02e-803">String</span></span> | | <span data-ttu-id="5b02e-804">Строка, содержащая имя вложения, длиной до 255 символов.</span><span class="sxs-lookup"><span data-stu-id="5b02e-804">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="5b02e-805">String</span><span class="sxs-lookup"><span data-stu-id="5b02e-805">String</span></span> | | <span data-ttu-id="5b02e-p151">Используется, только если свойству `type` задано значение `file`. URI расположения файла.</span><span class="sxs-lookup"><span data-stu-id="5b02e-p151">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.isInline` | <span data-ttu-id="5b02e-808">Логический</span><span class="sxs-lookup"><span data-stu-id="5b02e-808">Boolean</span></span> | | <span data-ttu-id="5b02e-p152">Используется, только если свойству `type` задано значение `file`. Значение `true` указывает на то, что вложение будет встроено в текст сообщения и не должно отображаться в списке вложений.</span><span class="sxs-lookup"><span data-stu-id="5b02e-p152">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="5b02e-811">String</span><span class="sxs-lookup"><span data-stu-id="5b02e-811">String</span></span> | | <span data-ttu-id="5b02e-p153">Используется, только если свойству `type` задано значение `item`. Идентификатор вложения EWS. Это строка длиной до 100 символов.</span><span class="sxs-lookup"><span data-stu-id="5b02e-p153">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="5b02e-815">function</span><span class="sxs-lookup"><span data-stu-id="5b02e-815">function</span></span> | <span data-ttu-id="5b02e-816">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="5b02e-816">&lt;optional&gt;</span></span> | <span data-ttu-id="5b02e-817">По завершении работы метода функция, переданная параметру `callback`, вызывается с помощью одного параметра `asyncResult`, представляющего собой объект [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="5b02e-817">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="5b02e-818">Требования</span><span class="sxs-lookup"><span data-stu-id="5b02e-818">Requirements</span></span>

|<span data-ttu-id="5b02e-819">Требование</span><span class="sxs-lookup"><span data-stu-id="5b02e-819">Requirement</span></span>| <span data-ttu-id="5b02e-820">Значение</span><span class="sxs-lookup"><span data-stu-id="5b02e-820">Value</span></span>|
|---|---|
|[<span data-ttu-id="5b02e-821">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5b02e-821">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5b02e-822">1.0</span><span class="sxs-lookup"><span data-stu-id="5b02e-822">1.0</span></span>|
|[<span data-ttu-id="5b02e-823">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5b02e-823">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5b02e-824">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5b02e-824">ReadItem</span></span>|
|[<span data-ttu-id="5b02e-825">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5b02e-825">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5b02e-826">Чтение</span><span class="sxs-lookup"><span data-stu-id="5b02e-826">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="5b02e-827">Примеры</span><span class="sxs-lookup"><span data-stu-id="5b02e-827">Examples</span></span>

<span data-ttu-id="5b02e-828">Приведенный ниже код передает строку в функцию `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="5b02e-828">The following code passes a string to the `displayReplyForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="5b02e-829">Ответ с пустым текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="5b02e-829">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="5b02e-830">Ответ только с текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="5b02e-830">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="5b02e-831">Ответ с текстом сообщения и вложенным файлом.</span><span class="sxs-lookup"><span data-stu-id="5b02e-831">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="5b02e-832">Ответ с текстом сообщения и вложенным элементом.</span><span class="sxs-lookup"><span data-stu-id="5b02e-832">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="5b02e-833">Ответ с текстом сообщения, вложенным файлом, вложенным элементом и обратным вызовом.</span><span class="sxs-lookup"><span data-stu-id="5b02e-833">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlook15officeentities"></a><span data-ttu-id="5b02e-834">getEntities() → {[Entities](/javascript/api/outlook_1_5/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="5b02e-834">getEntities() → {[Entities](/javascript/api/outlook_1_5/office.entities)}</span></span>

<span data-ttu-id="5b02e-835">Получает сущности, обнаруженные в теле выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="5b02e-835">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="5b02e-836">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="5b02e-836">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="5b02e-837">Требования</span><span class="sxs-lookup"><span data-stu-id="5b02e-837">Requirements</span></span>

|<span data-ttu-id="5b02e-838">Требование</span><span class="sxs-lookup"><span data-stu-id="5b02e-838">Requirement</span></span>| <span data-ttu-id="5b02e-839">Значение</span><span class="sxs-lookup"><span data-stu-id="5b02e-839">Value</span></span>|
|---|---|
|[<span data-ttu-id="5b02e-840">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5b02e-840">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5b02e-841">1.0</span><span class="sxs-lookup"><span data-stu-id="5b02e-841">1.0</span></span>|
|[<span data-ttu-id="5b02e-842">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5b02e-842">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5b02e-843">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5b02e-843">ReadItem</span></span>|
|[<span data-ttu-id="5b02e-844">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5b02e-844">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5b02e-845">Чтение</span><span class="sxs-lookup"><span data-stu-id="5b02e-845">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="5b02e-846">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="5b02e-846">Returns:</span></span>

<span data-ttu-id="5b02e-847">Тип: [Entities](/javascript/api/outlook_1_5/office.entities)</span><span class="sxs-lookup"><span data-stu-id="5b02e-847">Type: [Entities](/javascript/api/outlook_1_5/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="5b02e-848">Пример</span><span class="sxs-lookup"><span data-stu-id="5b02e-848">Example</span></span>

<span data-ttu-id="5b02e-849">Ниже приведен пример получения доступа к сущностям контактов в теле текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="5b02e-849">The following example accesses the contacts entities in the current item's body.</span></span>

```js
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook15officecontactmeetingsuggestionjavascriptapioutlook15officemeetingsuggestionphonenumberjavascriptapioutlook15officephonenumbertasksuggestionjavascriptapioutlook15officetasksuggestion"></a><span data-ttu-id="5b02e-850">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="5b02e-850">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))>}</span></span>

<span data-ttu-id="5b02e-851">Получает массив всех сущностей указанного типа, обнаруженных в теле выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="5b02e-851">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="5b02e-852">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="5b02e-852">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="5b02e-853">Параметры</span><span class="sxs-lookup"><span data-stu-id="5b02e-853">Parameters:</span></span>

|<span data-ttu-id="5b02e-854">Имя</span><span class="sxs-lookup"><span data-stu-id="5b02e-854">Name</span></span>| <span data-ttu-id="5b02e-855">Тип</span><span class="sxs-lookup"><span data-stu-id="5b02e-855">Type</span></span>| <span data-ttu-id="5b02e-856">Описание</span><span class="sxs-lookup"><span data-stu-id="5b02e-856">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="5b02e-857">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="5b02e-857">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook_1_5/office.mailboxenums.entitytype)|<span data-ttu-id="5b02e-858">Одно из значений перечисления EntityType.</span><span class="sxs-lookup"><span data-stu-id="5b02e-858">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="5b02e-859">Требования</span><span class="sxs-lookup"><span data-stu-id="5b02e-859">Requirements</span></span>

|<span data-ttu-id="5b02e-860">Требование</span><span class="sxs-lookup"><span data-stu-id="5b02e-860">Requirement</span></span>| <span data-ttu-id="5b02e-861">Значение</span><span class="sxs-lookup"><span data-stu-id="5b02e-861">Value</span></span>|
|---|---|
|[<span data-ttu-id="5b02e-862">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5b02e-862">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5b02e-863">1.0</span><span class="sxs-lookup"><span data-stu-id="5b02e-863">1.0</span></span>|
|[<span data-ttu-id="5b02e-864">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5b02e-864">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5b02e-865">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="5b02e-865">Restricted</span></span>|
|[<span data-ttu-id="5b02e-866">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5b02e-866">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5b02e-867">Чтение</span><span class="sxs-lookup"><span data-stu-id="5b02e-867">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="5b02e-868">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="5b02e-868">Returns:</span></span>

<span data-ttu-id="5b02e-869">Если значение, переданное в `entityType`, не является допустимым членом перечисления `EntityType`, метод возвращает значение NULL.</span><span class="sxs-lookup"><span data-stu-id="5b02e-869">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="5b02e-870">Если в теле элемента отсутствуют сущности указанного типа, метод возвращает пустой массив.</span><span class="sxs-lookup"><span data-stu-id="5b02e-870">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="5b02e-871">В противном случае тип объектов в возвращаемом массиве зависит от типа сущности, запрошенной в параметре `entityType`.</span><span class="sxs-lookup"><span data-stu-id="5b02e-871">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="5b02e-872">Хотя минимальный уровень разрешений для использования этого метода — **Restricted**, для некоторых типов сущностей требуется доступ на уровне **ReadItem**, как указано в приведенной ниже таблице.</span><span class="sxs-lookup"><span data-stu-id="5b02e-872">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="5b02e-873">Значение параметра `entityType`</span><span class="sxs-lookup"><span data-stu-id="5b02e-873">Value of `entityType`</span></span> | <span data-ttu-id="5b02e-874">Тип объектов в возвращаемом массиве</span><span class="sxs-lookup"><span data-stu-id="5b02e-874">Type of objects in returned array</span></span> | <span data-ttu-id="5b02e-875">Необходимый уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5b02e-875">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="5b02e-876">String</span><span class="sxs-lookup"><span data-stu-id="5b02e-876">String</span></span> | <span data-ttu-id="5b02e-877">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="5b02e-877">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="5b02e-878">Contact</span><span class="sxs-lookup"><span data-stu-id="5b02e-878">Contact</span></span> | <span data-ttu-id="5b02e-879">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="5b02e-879">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="5b02e-880">String</span><span class="sxs-lookup"><span data-stu-id="5b02e-880">String</span></span> | <span data-ttu-id="5b02e-881">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="5b02e-881">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="5b02e-882">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="5b02e-882">MeetingSuggestion</span></span> | <span data-ttu-id="5b02e-883">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="5b02e-883">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="5b02e-884">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="5b02e-884">PhoneNumber</span></span> | <span data-ttu-id="5b02e-885">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="5b02e-885">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="5b02e-886">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="5b02e-886">TaskSuggestion</span></span> | <span data-ttu-id="5b02e-887">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="5b02e-887">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="5b02e-888">String</span><span class="sxs-lookup"><span data-stu-id="5b02e-888">String</span></span> | <span data-ttu-id="5b02e-889">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="5b02e-889">**Restricted**</span></span> |

<span data-ttu-id="5b02e-890">Тип: Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="5b02e-890">Type: Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="5b02e-891">Пример</span><span class="sxs-lookup"><span data-stu-id="5b02e-891">Example</span></span>

<span data-ttu-id="5b02e-892">В примере ниже показано, как получить доступ к массиву строк, которые представляют собой почтовые адреса в теле текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="5b02e-892">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook15officecontactmeetingsuggestionjavascriptapioutlook15officemeetingsuggestionphonenumberjavascriptapioutlook15officephonenumbertasksuggestionjavascriptapioutlook15officetasksuggestion"></a><span data-ttu-id="5b02e-893">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="5b02e-893">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))>}</span></span>

<span data-ttu-id="5b02e-894">Возвращает известные сущности в выбранном элементе, которые проходят через именованный фильтр, определяемый в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="5b02e-894">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="5b02e-895">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="5b02e-895">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="5b02e-896">Метод `getFilteredEntitiesByName` возвращает сущности, соответствующие регулярному выражению, которое определяется в элементе правила [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) в XML-файле манифеста, с использованием указанного значения элемента `FilterName`.</span><span class="sxs-lookup"><span data-stu-id="5b02e-896">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="5b02e-897">Параметры</span><span class="sxs-lookup"><span data-stu-id="5b02e-897">Parameters:</span></span>

|<span data-ttu-id="5b02e-898">Имя</span><span class="sxs-lookup"><span data-stu-id="5b02e-898">Name</span></span>| <span data-ttu-id="5b02e-899">Тип</span><span class="sxs-lookup"><span data-stu-id="5b02e-899">Type</span></span>| <span data-ttu-id="5b02e-900">Описание</span><span class="sxs-lookup"><span data-stu-id="5b02e-900">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="5b02e-901">String</span><span class="sxs-lookup"><span data-stu-id="5b02e-901">String</span></span>|<span data-ttu-id="5b02e-902">Имя элемента правила `ItemHasKnownEntity`, определяющее соответствующий фильтр.</span><span class="sxs-lookup"><span data-stu-id="5b02e-902">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="5b02e-903">Требования</span><span class="sxs-lookup"><span data-stu-id="5b02e-903">Requirements</span></span>

|<span data-ttu-id="5b02e-904">Требование</span><span class="sxs-lookup"><span data-stu-id="5b02e-904">Requirement</span></span>| <span data-ttu-id="5b02e-905">Значение</span><span class="sxs-lookup"><span data-stu-id="5b02e-905">Value</span></span>|
|---|---|
|[<span data-ttu-id="5b02e-906">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5b02e-906">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5b02e-907">1.0</span><span class="sxs-lookup"><span data-stu-id="5b02e-907">1.0</span></span>|
|[<span data-ttu-id="5b02e-908">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5b02e-908">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5b02e-909">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5b02e-909">ReadItem</span></span>|
|[<span data-ttu-id="5b02e-910">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5b02e-910">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5b02e-911">Чтение</span><span class="sxs-lookup"><span data-stu-id="5b02e-911">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="5b02e-912">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="5b02e-912">Returns:</span></span>

<span data-ttu-id="5b02e-p155">Если в манифесте нет элемента `ItemHasKnownEntity` со значением `FilterName`, соответствующим параметру `name`, метод возвращает `null`. Если параметр `name` соответствует элементу `ItemHasKnownEntity` в манифесте, но при этом в текущем элементе нет соответствующих сущностей, метод возвращает пустой массив.</span><span class="sxs-lookup"><span data-stu-id="5b02e-p155">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="5b02e-915">Тип: Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="5b02e-915">Type: Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))></span></span>

#### <a name="getregexmatches--object"></a><span data-ttu-id="5b02e-916">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="5b02e-916">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="5b02e-917">Возвращает строковые значения в выбранном элементе, которые соответствуют регулярным выражениям, определенным в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="5b02e-917">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="5b02e-918">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="5b02e-918">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="5b02e-p156">Метод `getRegExMatches` возвращает строки, соответствующие регулярному выражению, которое определяется в каждом элементе правила `ItemHasRegularExpressionMatch` или `ItemHasKnownEntity` в XML-файле манифеста. Для правила `ItemHasRegularExpressionMatch` соответствующую строку должно содержать свойство элемента, указанного этим правилом. Простой тип `PropertyName` определяет поддерживаемые свойства.</span><span class="sxs-lookup"><span data-stu-id="5b02e-p156">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="5b02e-922">Например, рассмотрим манифест надстройки, который содержит указанный ниже элемент `Rule`.</span><span class="sxs-lookup"><span data-stu-id="5b02e-922">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="5b02e-923">Объект, возвращаемый методом `getRegExMatches`, будет содержать два свойства: `fruits` и `veggies`.</span><span class="sxs-lookup"><span data-stu-id="5b02e-923">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="5b02e-p157">Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты. Лучше используйте метод [`Body.getAsync`](/javascript/api/outlook_1_5/office.body#getasync-coerciontype--options--callback-) для этого.</span><span class="sxs-lookup"><span data-stu-id="5b02e-p157">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook_1_5/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="5b02e-927">Requirements</span><span class="sxs-lookup"><span data-stu-id="5b02e-927">Requirements</span></span>

|<span data-ttu-id="5b02e-928">Требование</span><span class="sxs-lookup"><span data-stu-id="5b02e-928">Requirement</span></span>| <span data-ttu-id="5b02e-929">Значение</span><span class="sxs-lookup"><span data-stu-id="5b02e-929">Value</span></span>|
|---|---|
|[<span data-ttu-id="5b02e-930">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5b02e-930">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5b02e-931">1.0</span><span class="sxs-lookup"><span data-stu-id="5b02e-931">1.0</span></span>|
|[<span data-ttu-id="5b02e-932">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5b02e-932">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5b02e-933">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5b02e-933">ReadItem</span></span>|
|[<span data-ttu-id="5b02e-934">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5b02e-934">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5b02e-935">Чтение</span><span class="sxs-lookup"><span data-stu-id="5b02e-935">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="5b02e-936">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="5b02e-936">Returns:</span></span>

<span data-ttu-id="5b02e-p158">Объект, содержащий массив строк, которые соответствуют регулярным выражениям, определяемым в XML-файле манифеста. Имя каждого массива равно соответствующему значению атрибута `RegExName` подходящего правила `ItemHasRegularExpressionMatch` или атрибута `FilterName` соответствующего правила `ItemHasKnownEntity`.</span><span class="sxs-lookup"><span data-stu-id="5b02e-p158">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="5b02e-939">

<dt>Тип</dt>

</span><span class="sxs-lookup"><span data-stu-id="5b02e-939">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="5b02e-940">Object</span><span class="sxs-lookup"><span data-stu-id="5b02e-940">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="5b02e-941">Пример</span><span class="sxs-lookup"><span data-stu-id="5b02e-941">Example</span></span>

<span data-ttu-id="5b02e-942">В примере ниже показано, как получить доступ к массиву совпадений для <rule>элементов регулярного выражения `fruits` и `veggies`, которые указаны в манифесте</rule>.</span><span class="sxs-lookup"><span data-stu-id="5b02e-942">The following example shows how to access the array of matches for the regular expression <rule>elements `fruits` and `veggies`, which are specified in the manifest.</rule></span></span>

```js
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veges = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="5b02e-943">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="5b02e-943">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="5b02e-944">Возвращает строковые значения в выбранном элементе, которые соответствуют именованному регулярному выражению, определенному в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="5b02e-944">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="5b02e-945">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="5b02e-945">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="5b02e-946">Метод `getRegExMatchesByName` возвращает строки, соответствующие регулярному выражению, которое определяется в элементе правила `ItemHasRegularExpressionMatch` в XML-файле манифеста, с использованием указанного значения элемента `RegExName`.</span><span class="sxs-lookup"><span data-stu-id="5b02e-946">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="5b02e-p159">Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты.</span><span class="sxs-lookup"><span data-stu-id="5b02e-p159">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="5b02e-949">Параметры</span><span class="sxs-lookup"><span data-stu-id="5b02e-949">Parameters:</span></span>

|<span data-ttu-id="5b02e-950">Имя</span><span class="sxs-lookup"><span data-stu-id="5b02e-950">Name</span></span>| <span data-ttu-id="5b02e-951">Тип</span><span class="sxs-lookup"><span data-stu-id="5b02e-951">Type</span></span>| <span data-ttu-id="5b02e-952">Описание</span><span class="sxs-lookup"><span data-stu-id="5b02e-952">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="5b02e-953">String</span><span class="sxs-lookup"><span data-stu-id="5b02e-953">String</span></span>|<span data-ttu-id="5b02e-954">Имя элемента правила `ItemHasRegularExpressionMatch`, определяющее соответствующий фильтр.</span><span class="sxs-lookup"><span data-stu-id="5b02e-954">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="5b02e-955">Требования</span><span class="sxs-lookup"><span data-stu-id="5b02e-955">Requirements</span></span>

|<span data-ttu-id="5b02e-956">Требование</span><span class="sxs-lookup"><span data-stu-id="5b02e-956">Requirement</span></span>| <span data-ttu-id="5b02e-957">Значение</span><span class="sxs-lookup"><span data-stu-id="5b02e-957">Value</span></span>|
|---|---|
|[<span data-ttu-id="5b02e-958">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5b02e-958">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5b02e-959">1.0</span><span class="sxs-lookup"><span data-stu-id="5b02e-959">1.0</span></span>|
|[<span data-ttu-id="5b02e-960">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5b02e-960">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5b02e-961">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5b02e-961">ReadItem</span></span>|
|[<span data-ttu-id="5b02e-962">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5b02e-962">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5b02e-963">Чтение</span><span class="sxs-lookup"><span data-stu-id="5b02e-963">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="5b02e-964">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="5b02e-964">Returns:</span></span>

<span data-ttu-id="5b02e-965">Массив строк, соответствующих регулярному выражению, определяемому в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="5b02e-965">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="5b02e-966">

<dt>Тип</dt>

</span><span class="sxs-lookup"><span data-stu-id="5b02e-966">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="5b02e-967">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="5b02e-967">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="5b02e-968">Пример</span><span class="sxs-lookup"><span data-stu-id="5b02e-968">Example</span></span>

```js
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="5b02e-969">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="5b02e-969">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="5b02e-970">Асинхронно возвращает данные, выбранные в теме или тексте сообщения.</span><span class="sxs-lookup"><span data-stu-id="5b02e-970">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="5b02e-p160">Если выделенный фрагмент отсутствует, но курсор находится в тексте или теме, метод возвращает значение NULL для выбранных данных. Если выбраны не текст и не тема, метод возвращает ошибку `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="5b02e-p160">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="5b02e-973">Параметры</span><span class="sxs-lookup"><span data-stu-id="5b02e-973">Parameters:</span></span>

|<span data-ttu-id="5b02e-974">Имя</span><span class="sxs-lookup"><span data-stu-id="5b02e-974">Name</span></span>| <span data-ttu-id="5b02e-975">Тип</span><span class="sxs-lookup"><span data-stu-id="5b02e-975">Type</span></span>| <span data-ttu-id="5b02e-976">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="5b02e-976">Attributes</span></span>| <span data-ttu-id="5b02e-977">Описание</span><span class="sxs-lookup"><span data-stu-id="5b02e-977">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="5b02e-978">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="5b02e-978">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="5b02e-p161">Запрашивает формат данных. Если задано значение Text, метод возвращает обычный текст как строку, удаляя все имеющиеся HTML-теги. Если задано значение HTML, метод возвращает выделенный текст (обычный текст или HTML).</span><span class="sxs-lookup"><span data-stu-id="5b02e-p161">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="5b02e-982">Object</span><span class="sxs-lookup"><span data-stu-id="5b02e-982">Object</span></span>| <span data-ttu-id="5b02e-983">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="5b02e-983">&lt;optional&gt;</span></span>|<span data-ttu-id="5b02e-984">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="5b02e-984">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="5b02e-985">Object</span><span class="sxs-lookup"><span data-stu-id="5b02e-985">Object</span></span>| <span data-ttu-id="5b02e-986">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="5b02e-986">&lt;optional&gt;</span></span>|<span data-ttu-id="5b02e-987">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="5b02e-987">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="5b02e-988">функция</span><span class="sxs-lookup"><span data-stu-id="5b02e-988">function</span></span>||<span data-ttu-id="5b02e-989">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="5b02e-989">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="5b02e-990">Чтобы получить доступ к выбранным данным из метода обратного вызова, вызовите `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="5b02e-990">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="5b02e-991">Для доступа к исходному свойству, представляющему собой источник выбранных данных, вызовите параметр `asyncResult.value.sourceProperty`, который может иметь значение `body` или `subject`.</span><span class="sxs-lookup"><span data-stu-id="5b02e-991">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="5b02e-992">Требования</span><span class="sxs-lookup"><span data-stu-id="5b02e-992">Requirements</span></span>

|<span data-ttu-id="5b02e-993">Требование</span><span class="sxs-lookup"><span data-stu-id="5b02e-993">Requirement</span></span>| <span data-ttu-id="5b02e-994">Значение</span><span class="sxs-lookup"><span data-stu-id="5b02e-994">Value</span></span>|
|---|---|
|[<span data-ttu-id="5b02e-995">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="5b02e-995">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5b02e-996">1.2</span><span class="sxs-lookup"><span data-stu-id="5b02e-996">1.2</span></span>|
|[<span data-ttu-id="5b02e-997">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5b02e-997">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5b02e-998">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="5b02e-998">ReadWriteItem</span></span>|
|[<span data-ttu-id="5b02e-999">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5b02e-999">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5b02e-1000">Создание</span><span class="sxs-lookup"><span data-stu-id="5b02e-1000">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="5b02e-1001">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="5b02e-1001">Returns:</span></span>

<span data-ttu-id="5b02e-1002">Выбранные данные в виде строки с форматом, определенным в параметре `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="5b02e-1002">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="5b02e-1003">

<dt>Тип</dt>

</span><span class="sxs-lookup"><span data-stu-id="5b02e-1003">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="5b02e-1004">String</span><span class="sxs-lookup"><span data-stu-id="5b02e-1004">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="5b02e-1005">Пример</span><span class="sxs-lookup"><span data-stu-id="5b02e-1005">Example</span></span>

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

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="5b02e-1006">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="5b02e-1006">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="5b02e-1007">Асинхронно загружает настраиваемые свойства для надстройки для выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="5b02e-1007">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="5b02e-p163">Настраиваемые свойства сохраняются в виде пар "ключ-значение" для каждого приложения и каждого элемента. Этот метод возвращает объект `CustomProperties` при обратном вызове, который предоставляет методы для доступа к настраиваемым свойствам, характерным для текущего элемента и текущей надстройки. Настраиваемые свойства не шифруются для элемента, поэтому этот способ хранения не является безопасным.</span><span class="sxs-lookup"><span data-stu-id="5b02e-p163">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="5b02e-1011">Параметры</span><span class="sxs-lookup"><span data-stu-id="5b02e-1011">Parameters:</span></span>

|<span data-ttu-id="5b02e-1012">Имя</span><span class="sxs-lookup"><span data-stu-id="5b02e-1012">Name</span></span>| <span data-ttu-id="5b02e-1013">Тип</span><span class="sxs-lookup"><span data-stu-id="5b02e-1013">Type</span></span>| <span data-ttu-id="5b02e-1014">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="5b02e-1014">Attributes</span></span>| <span data-ttu-id="5b02e-1015">Описание</span><span class="sxs-lookup"><span data-stu-id="5b02e-1015">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="5b02e-1016">function</span><span class="sxs-lookup"><span data-stu-id="5b02e-1016">function</span></span>||<span data-ttu-id="5b02e-1017">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="5b02e-1017">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="5b02e-1018">Настраиваемые свойства предоставляются в виде объекта [`CustomProperties`](/javascript/api/outlook_1_5/office.customproperties) в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="5b02e-1018">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook_1_5/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="5b02e-1019">Этот объект позволяет получить, задать и удалить настраиваемые свойства для элемента, а также сохранить изменения, внесенные в набор настраиваемых свойств, на сервере.</span><span class="sxs-lookup"><span data-stu-id="5b02e-1019">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="5b02e-1020">Объект</span><span class="sxs-lookup"><span data-stu-id="5b02e-1020">Object</span></span>| <span data-ttu-id="5b02e-1021">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="5b02e-1021">&lt;optional&gt;</span></span>|<span data-ttu-id="5b02e-1022">Разработчики могут указать любой объект, к которому необходимо получить доступ, в функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="5b02e-1022">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="5b02e-1023">Доступ к этому объекту можно получить с помощью свойства `asyncResult.asyncContext` в функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="5b02e-1023">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="5b02e-1024">Требования</span><span class="sxs-lookup"><span data-stu-id="5b02e-1024">Requirements</span></span>

|<span data-ttu-id="5b02e-1025">Требование</span><span class="sxs-lookup"><span data-stu-id="5b02e-1025">Requirement</span></span>| <span data-ttu-id="5b02e-1026">Значение</span><span class="sxs-lookup"><span data-stu-id="5b02e-1026">Value</span></span>|
|---|---|
|[<span data-ttu-id="5b02e-1027">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5b02e-1027">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5b02e-1028">1.0</span><span class="sxs-lookup"><span data-stu-id="5b02e-1028">1.0</span></span>|
|[<span data-ttu-id="5b02e-1029">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5b02e-1029">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5b02e-1030">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5b02e-1030">ReadItem</span></span>|
|[<span data-ttu-id="5b02e-1031">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5b02e-1031">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5b02e-1032">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="5b02e-1032">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="5b02e-1033">Пример</span><span class="sxs-lookup"><span data-stu-id="5b02e-1033">Example</span></span>

<span data-ttu-id="5b02e-p166">Приведенный ниже пример кода показывает, как асинхронно загружать настраиваемые свойства, характерные для текущего элемента, с помощью метода `loadCustomPropertiesAsync`. Этот пример также показывает, как сохранять эти свойства на сервере с помощью метода `CustomProperties.saveAsync`. После загрузки настраиваемых свойств в этом примере кода метод `CustomProperties.get` используется для считывания настраиваемого свойства `myProp`, метод `CustomProperties.set` — для записи настраиваемого свойства `otherProp`, а метод `saveAsync` — для сохранения настраиваемых свойств.</span><span class="sxs-lookup"><span data-stu-id="5b02e-p166">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="5b02e-1037">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="5b02e-1037">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="5b02e-1038">Удаляет вложение из сообщения или встречи.</span><span class="sxs-lookup"><span data-stu-id="5b02e-1038">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="5b02e-p167">Метод `removeAttachmentAsync` удаляет из элемента вложение с указанным идентификатором. Идентификатор вложения рекомендуется использовать для удаления вложения, только если оно добавлено тем же почтовым приложением в ходе текущего сеанса. В Outlook Web App и Outlook Web App для устройств идентификатор вложения действителен только в рамках одного сеанса. Сеанс завершается, когда пользователь закрывает приложение или начинает создавать элемент во встроенной форме, а затем переходит из формы в отдельное окно.</span><span class="sxs-lookup"><span data-stu-id="5b02e-p167">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="5b02e-1043">Параметры</span><span class="sxs-lookup"><span data-stu-id="5b02e-1043">Parameters:</span></span>

|<span data-ttu-id="5b02e-1044">Имя</span><span class="sxs-lookup"><span data-stu-id="5b02e-1044">Name</span></span>| <span data-ttu-id="5b02e-1045">Тип</span><span class="sxs-lookup"><span data-stu-id="5b02e-1045">Type</span></span>| <span data-ttu-id="5b02e-1046">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="5b02e-1046">Attributes</span></span>| <span data-ttu-id="5b02e-1047">Описание</span><span class="sxs-lookup"><span data-stu-id="5b02e-1047">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="5b02e-1048">String</span><span class="sxs-lookup"><span data-stu-id="5b02e-1048">String</span></span>||<span data-ttu-id="5b02e-1049">Идентификатор удаляемого вложения.</span><span class="sxs-lookup"><span data-stu-id="5b02e-1049">The identifier of the attachment to remove.</span></span>|
|`options`| <span data-ttu-id="5b02e-1050">Object</span><span class="sxs-lookup"><span data-stu-id="5b02e-1050">Object</span></span>| <span data-ttu-id="5b02e-1051">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="5b02e-1051">&lt;optional&gt;</span></span>|<span data-ttu-id="5b02e-1052">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="5b02e-1052">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="5b02e-1053">Object</span><span class="sxs-lookup"><span data-stu-id="5b02e-1053">Object</span></span>| <span data-ttu-id="5b02e-1054">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="5b02e-1054">&lt;optional&gt;</span></span>|<span data-ttu-id="5b02e-1055">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="5b02e-1055">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="5b02e-1056">функция</span><span class="sxs-lookup"><span data-stu-id="5b02e-1056">function</span></span>| <span data-ttu-id="5b02e-1057">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="5b02e-1057">&lt;optional&gt;</span></span>|<span data-ttu-id="5b02e-1058">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="5b02e-1058">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="5b02e-1059">Если удалить вложение не удается, свойство `asyncResult.error` содержит код ошибки с указанием ее причины.</span><span class="sxs-lookup"><span data-stu-id="5b02e-1059">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="5b02e-1060">Ошибки</span><span class="sxs-lookup"><span data-stu-id="5b02e-1060">Errors</span></span>

| <span data-ttu-id="5b02e-1061">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="5b02e-1061">Error code</span></span> | <span data-ttu-id="5b02e-1062">Описание</span><span class="sxs-lookup"><span data-stu-id="5b02e-1062">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="5b02e-1063">Идентификатор вложения не существует.</span><span class="sxs-lookup"><span data-stu-id="5b02e-1063">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="5b02e-1064">Требования</span><span class="sxs-lookup"><span data-stu-id="5b02e-1064">Requirements</span></span>

|<span data-ttu-id="5b02e-1065">Требование</span><span class="sxs-lookup"><span data-stu-id="5b02e-1065">Requirement</span></span>| <span data-ttu-id="5b02e-1066">Значение</span><span class="sxs-lookup"><span data-stu-id="5b02e-1066">Value</span></span>|
|---|---|
|[<span data-ttu-id="5b02e-1067">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5b02e-1067">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5b02e-1068">1.1</span><span class="sxs-lookup"><span data-stu-id="5b02e-1068">1.1</span></span>|
|[<span data-ttu-id="5b02e-1069">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5b02e-1069">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5b02e-1070">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="5b02e-1070">ReadWriteItem</span></span>|
|[<span data-ttu-id="5b02e-1071">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5b02e-1071">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5b02e-1072">Создание</span><span class="sxs-lookup"><span data-stu-id="5b02e-1072">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="5b02e-1073">Пример</span><span class="sxs-lookup"><span data-stu-id="5b02e-1073">Example</span></span>

<span data-ttu-id="5b02e-1074">Указанный ниже код удаляет вложение с идентификатором "0".</span><span class="sxs-lookup"><span data-stu-id="5b02e-1074">The following code removes an attachment with an identifier of '0'.</span></span>

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

####  <a name="saveasyncoptions-callback"></a><span data-ttu-id="5b02e-1075">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="5b02e-1075">saveAsync([options], callback)</span></span>

<span data-ttu-id="5b02e-1076">Асинхронно сохраняет элемент.</span><span class="sxs-lookup"><span data-stu-id="5b02e-1076">Asynchronously saves an item.</span></span>

<span data-ttu-id="5b02e-p168">При вызове этот метод сохраняет текущее сообщение в виде черновика и возвращает идентификатор элемента с помощью метода обратного вызова. В Outlook Web App или интерактивном режиме Outlook этот элемент сохраняется на сервере. В Outlook в режиме кэширования этот элемент сохраняется в локальном кэше.</span><span class="sxs-lookup"><span data-stu-id="5b02e-p168">When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="5b02e-1080">Если в режиме создания надстройка вызывает для элемента метод `saveAsync`, чтобы получить параметр `itemId` для использования с EWS или REST API, необходимо помнить, что синхронизация элемента с сервером может занять много времени, если Outlook работает в режиме кэширования данных.</span><span class="sxs-lookup"><span data-stu-id="5b02e-1080">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="5b02e-1081">До окончания синхронизации элемента применение параметра `itemId` будет приводить к ошибке.</span><span class="sxs-lookup"><span data-stu-id="5b02e-1081">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="5b02e-p170">Если метод `saveAsync` вызывается для встречи в режиме создания, она сохраняется как обычная встреча в календаре пользователя, а не как черновик. При сохранении новой встречи приглашения не отправляются. При сохранении существующей встречи уведомления отправляются добавленным или удаленным участникам.</span><span class="sxs-lookup"><span data-stu-id="5b02e-p170">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="5b02e-1085">Следующие клиенты отличаются другим поведением `saveAsync` в отношении встреч в режиме создания:</span><span class="sxs-lookup"><span data-stu-id="5b02e-1085">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="5b02e-1086">Outlook для Mac не поддерживает `saveAsync` для собраний в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="5b02e-1086">Mac Outlook does not support `saveAsync` on a meeting in compose mode.</span></span> <span data-ttu-id="5b02e-1087">При вызове `saveAsync` для собрания в Outlook для Mac возвращается ошибка.</span><span class="sxs-lookup"><span data-stu-id="5b02e-1087">Calling `saveAsync` on a meeting in Mac Outlook will return an error.</span></span>
> - <span data-ttu-id="5b02e-1088">Outlook в Интернете всегда отправляет приглашение или обновление при вызове `saveAsync` для встречи в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="5b02e-1088">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="5b02e-1089">Параметры:</span><span class="sxs-lookup"><span data-stu-id="5b02e-1089">Parameters:</span></span>

|<span data-ttu-id="5b02e-1090">Имя</span><span class="sxs-lookup"><span data-stu-id="5b02e-1090">Name</span></span>| <span data-ttu-id="5b02e-1091">Тип</span><span class="sxs-lookup"><span data-stu-id="5b02e-1091">Type</span></span>| <span data-ttu-id="5b02e-1092">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="5b02e-1092">Attributes</span></span>| <span data-ttu-id="5b02e-1093">Описание</span><span class="sxs-lookup"><span data-stu-id="5b02e-1093">Description</span></span>|
|---|---|---|---|
|`options`| <span data-ttu-id="5b02e-1094">Object</span><span class="sxs-lookup"><span data-stu-id="5b02e-1094">Object</span></span>| <span data-ttu-id="5b02e-1095">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="5b02e-1095">&lt;optional&gt;</span></span>|<span data-ttu-id="5b02e-1096">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="5b02e-1096">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="5b02e-1097">Object</span><span class="sxs-lookup"><span data-stu-id="5b02e-1097">Object</span></span>| <span data-ttu-id="5b02e-1098">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="5b02e-1098">&lt;optional&gt;</span></span>|<span data-ttu-id="5b02e-1099">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="5b02e-1099">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="5b02e-1100">функция</span><span class="sxs-lookup"><span data-stu-id="5b02e-1100">function</span></span>||<span data-ttu-id="5b02e-1101">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="5b02e-1101">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="5b02e-1102">После успешного выполнения идентификатор элемента будет предоставлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="5b02e-1102">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="5b02e-1103">Требования</span><span class="sxs-lookup"><span data-stu-id="5b02e-1103">Requirements</span></span>

|<span data-ttu-id="5b02e-1104">Требование</span><span class="sxs-lookup"><span data-stu-id="5b02e-1104">Requirement</span></span>| <span data-ttu-id="5b02e-1105">Значение</span><span class="sxs-lookup"><span data-stu-id="5b02e-1105">Value</span></span>|
|---|---|
|[<span data-ttu-id="5b02e-1106">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="5b02e-1106">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5b02e-1107">1.3</span><span class="sxs-lookup"><span data-stu-id="5b02e-1107">1.3</span></span>|
|[<span data-ttu-id="5b02e-1108">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5b02e-1108">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5b02e-1109">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="5b02e-1109">ReadWriteItem</span></span>|
|[<span data-ttu-id="5b02e-1110">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5b02e-1110">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5b02e-1111">Создание</span><span class="sxs-lookup"><span data-stu-id="5b02e-1111">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="5b02e-1112">Примеры</span><span class="sxs-lookup"><span data-stu-id="5b02e-1112">Examples</span></span>

```js
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result
  });
```

<span data-ttu-id="5b02e-p172">Ниже приведен пример параметра `result`, переданного функции обратного вызова. Свойство `value` содержит идентификатор элемента.</span><span class="sxs-lookup"><span data-stu-id="5b02e-p172">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

####  <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="5b02e-1115">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="5b02e-1115">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="5b02e-1116">Асинхронно вставляет данные в текст или тему сообщения.</span><span class="sxs-lookup"><span data-stu-id="5b02e-1116">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="5b02e-p173">Метод `setSelectedDataAsync` вставляет указанную строку в местоположение курсора в теме или тексте элемента либо, если текст выделен в редакторе, он заменяет выделенный текст. Если курсор находится вне текста или темы элемента, возвращается ошибка. После вставки курсор помещается в конец вставленного содержимого.</span><span class="sxs-lookup"><span data-stu-id="5b02e-p173">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="5b02e-1120">Параметры:</span><span class="sxs-lookup"><span data-stu-id="5b02e-1120">Parameters:</span></span>

|<span data-ttu-id="5b02e-1121">Имя</span><span class="sxs-lookup"><span data-stu-id="5b02e-1121">Name</span></span>| <span data-ttu-id="5b02e-1122">Тип</span><span class="sxs-lookup"><span data-stu-id="5b02e-1122">Type</span></span>| <span data-ttu-id="5b02e-1123">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="5b02e-1123">Attributes</span></span>| <span data-ttu-id="5b02e-1124">Описание</span><span class="sxs-lookup"><span data-stu-id="5b02e-1124">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="5b02e-1125">String</span><span class="sxs-lookup"><span data-stu-id="5b02e-1125">String</span></span>||<span data-ttu-id="5b02e-p174">Вставляемые данные. Объем данных не должен превышать 1 000 000 символов. Если передано больше 1 000 000 символов, возвращается исключение `ArgumentOutOfRange`.</span><span class="sxs-lookup"><span data-stu-id="5b02e-p174">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="5b02e-1129">Object</span><span class="sxs-lookup"><span data-stu-id="5b02e-1129">Object</span></span>| <span data-ttu-id="5b02e-1130">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="5b02e-1130">&lt;optional&gt;</span></span>|<span data-ttu-id="5b02e-1131">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="5b02e-1131">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="5b02e-1132">Object</span><span class="sxs-lookup"><span data-stu-id="5b02e-1132">Object</span></span>| <span data-ttu-id="5b02e-1133">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="5b02e-1133">&lt;optional&gt;</span></span>|<span data-ttu-id="5b02e-1134">В методе обратного вызова разработчики могут указать любой объект, к которому необходимо получить доступ.</span><span class="sxs-lookup"><span data-stu-id="5b02e-1134">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`| [<span data-ttu-id="5b02e-1135">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="5b02e-1135">Office.CoercionType</span></span>](office.md#coerciontype-string)| <span data-ttu-id="5b02e-1136">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="5b02e-1136">&lt;optional&gt;</span></span>|<span data-ttu-id="5b02e-p175">Если задано значение `text`, текущий стиль применяется в Outlook Web App и Outlook. Если поле представляет собой редактор HTML, вставляются только текстовые данные, даже если они имеют формат HTML.</span><span class="sxs-lookup"><span data-stu-id="5b02e-p175">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="5b02e-p176">Если задано значение `html` и поле (не тема) поддерживает HTML, в Outlook Web App применяется текущий стиль, а в Outlook — стиль по умолчанию. Если поле является текстовым, возвращается ошибка `InvalidDataFormat`.</span><span class="sxs-lookup"><span data-stu-id="5b02e-p176">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="5b02e-1141">Если свойство `coercionType` не задано, результат зависит от поля: если поле имеет формат HTML, используется текст в формате HTML, а если поле текстовое, применяется обычный текст.</span><span class="sxs-lookup"><span data-stu-id="5b02e-1141">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="5b02e-1142">функция</span><span class="sxs-lookup"><span data-stu-id="5b02e-1142">function</span></span>||<span data-ttu-id="5b02e-1143">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="5b02e-1143">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="5b02e-1144">Требования</span><span class="sxs-lookup"><span data-stu-id="5b02e-1144">Requirements</span></span>

|<span data-ttu-id="5b02e-1145">Требование</span><span class="sxs-lookup"><span data-stu-id="5b02e-1145">Requirement</span></span>| <span data-ttu-id="5b02e-1146">Значение</span><span class="sxs-lookup"><span data-stu-id="5b02e-1146">Value</span></span>|
|---|---|
|[<span data-ttu-id="5b02e-1147">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="5b02e-1147">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5b02e-1148">1.2</span><span class="sxs-lookup"><span data-stu-id="5b02e-1148">1.2</span></span>|
|[<span data-ttu-id="5b02e-1149">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5b02e-1149">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5b02e-1150">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="5b02e-1150">ReadWriteItem</span></span>|
|[<span data-ttu-id="5b02e-1151">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5b02e-1151">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5b02e-1152">Создание</span><span class="sxs-lookup"><span data-stu-id="5b02e-1152">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="5b02e-1153">Пример</span><span class="sxs-lookup"><span data-stu-id="5b02e-1153">Example</span></span>

```js
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
