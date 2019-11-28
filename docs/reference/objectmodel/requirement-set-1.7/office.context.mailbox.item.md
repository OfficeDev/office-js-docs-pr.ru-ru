---
title: Office. Context. Mailbox. Item — набор требований 1,7
description: ''
ms.date: 11/25/2019
localization_priority: Normal
ms.openlocfilehash: d400765293449899eb2e26f3d87128bc88b70000
ms.sourcegitcommit: 05a883a7fd89136301ce35aabc57638e9f563288
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 11/27/2019
ms.locfileid: "39629681"
---
# <a name="item"></a><span data-ttu-id="a0efc-102">item</span><span class="sxs-lookup"><span data-stu-id="a0efc-102">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="a0efc-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="a0efc-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="a0efc-p101">Пространство имен `item` используется для доступа к выбранному в данный момент сообщению, приглашению на собрание или описанию встречи. Вы можете определить тип пространства имен `item` с помощью свойства [itemType](#itemtype-officemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="a0efc-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="a0efc-106">Requirements</span><span class="sxs-lookup"><span data-stu-id="a0efc-106">Requirements</span></span>

|<span data-ttu-id="a0efc-107">Требование</span><span class="sxs-lookup"><span data-stu-id="a0efc-107">Requirement</span></span>|<span data-ttu-id="a0efc-108">Значение</span><span class="sxs-lookup"><span data-stu-id="a0efc-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="a0efc-109">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="a0efc-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a0efc-110">1.0</span><span class="sxs-lookup"><span data-stu-id="a0efc-110">1.0</span></span>|
|[<span data-ttu-id="a0efc-111">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="a0efc-111">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a0efc-112">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="a0efc-112">Restricted</span></span>|
|[<span data-ttu-id="a0efc-113">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="a0efc-113">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="a0efc-114">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="a0efc-114">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="a0efc-115">Элементы и методы</span><span class="sxs-lookup"><span data-stu-id="a0efc-115">Members and methods</span></span>

| <span data-ttu-id="a0efc-116">Элемент</span><span class="sxs-lookup"><span data-stu-id="a0efc-116">Member</span></span> | <span data-ttu-id="a0efc-117">Тип</span><span class="sxs-lookup"><span data-stu-id="a0efc-117">Type</span></span> |
|--------|------|
| [<span data-ttu-id="a0efc-118">attachments</span><span class="sxs-lookup"><span data-stu-id="a0efc-118">attachments</span></span>](#attachments-arrayattachmentdetails) | <span data-ttu-id="a0efc-119">Элемент</span><span class="sxs-lookup"><span data-stu-id="a0efc-119">Member</span></span> |
| [<span data-ttu-id="a0efc-120">bcc</span><span class="sxs-lookup"><span data-stu-id="a0efc-120">bcc</span></span>](#bcc-recipients) | <span data-ttu-id="a0efc-121">Элемент</span><span class="sxs-lookup"><span data-stu-id="a0efc-121">Member</span></span> |
| [<span data-ttu-id="a0efc-122">body</span><span class="sxs-lookup"><span data-stu-id="a0efc-122">body</span></span>](#body-body) | <span data-ttu-id="a0efc-123">Элемент</span><span class="sxs-lookup"><span data-stu-id="a0efc-123">Member</span></span> |
| [<span data-ttu-id="a0efc-124">cc</span><span class="sxs-lookup"><span data-stu-id="a0efc-124">cc</span></span>](#cc-arrayemailaddressdetailsrecipients) | <span data-ttu-id="a0efc-125">Элемент</span><span class="sxs-lookup"><span data-stu-id="a0efc-125">Member</span></span> |
| [<span data-ttu-id="a0efc-126">conversationId</span><span class="sxs-lookup"><span data-stu-id="a0efc-126">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="a0efc-127">Элемент</span><span class="sxs-lookup"><span data-stu-id="a0efc-127">Member</span></span> |
| [<span data-ttu-id="a0efc-128">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="a0efc-128">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="a0efc-129">Элемент</span><span class="sxs-lookup"><span data-stu-id="a0efc-129">Member</span></span> |
| [<span data-ttu-id="a0efc-130">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="a0efc-130">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="a0efc-131">Элемент</span><span class="sxs-lookup"><span data-stu-id="a0efc-131">Member</span></span> |
| [<span data-ttu-id="a0efc-132">end</span><span class="sxs-lookup"><span data-stu-id="a0efc-132">end</span></span>](#end-datetime) | <span data-ttu-id="a0efc-133">Элемент</span><span class="sxs-lookup"><span data-stu-id="a0efc-133">Member</span></span> |
| [<span data-ttu-id="a0efc-134">from</span><span class="sxs-lookup"><span data-stu-id="a0efc-134">from</span></span>](#from-emailaddressdetailsfrom) | <span data-ttu-id="a0efc-135">Элемент</span><span class="sxs-lookup"><span data-stu-id="a0efc-135">Member</span></span> |
| [<span data-ttu-id="a0efc-136">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="a0efc-136">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="a0efc-137">Элемент</span><span class="sxs-lookup"><span data-stu-id="a0efc-137">Member</span></span> |
| [<span data-ttu-id="a0efc-138">itemClass</span><span class="sxs-lookup"><span data-stu-id="a0efc-138">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="a0efc-139">Элемент</span><span class="sxs-lookup"><span data-stu-id="a0efc-139">Member</span></span> |
| [<span data-ttu-id="a0efc-140">itemId</span><span class="sxs-lookup"><span data-stu-id="a0efc-140">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="a0efc-141">Элемент</span><span class="sxs-lookup"><span data-stu-id="a0efc-141">Member</span></span> |
| [<span data-ttu-id="a0efc-142">itemType</span><span class="sxs-lookup"><span data-stu-id="a0efc-142">itemType</span></span>](#itemtype-officemailboxenumsitemtype) | <span data-ttu-id="a0efc-143">Элемент</span><span class="sxs-lookup"><span data-stu-id="a0efc-143">Member</span></span> |
| [<span data-ttu-id="a0efc-144">location</span><span class="sxs-lookup"><span data-stu-id="a0efc-144">location</span></span>](#location-stringlocation) | <span data-ttu-id="a0efc-145">Элемент</span><span class="sxs-lookup"><span data-stu-id="a0efc-145">Member</span></span> |
| [<span data-ttu-id="a0efc-146">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="a0efc-146">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="a0efc-147">Элемент</span><span class="sxs-lookup"><span data-stu-id="a0efc-147">Member</span></span> |
| [<span data-ttu-id="a0efc-148">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="a0efc-148">notificationMessages</span></span>](#notificationmessages-notificationmessages) | <span data-ttu-id="a0efc-149">Элемент</span><span class="sxs-lookup"><span data-stu-id="a0efc-149">Member</span></span> |
| [<span data-ttu-id="a0efc-150">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="a0efc-150">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="a0efc-151">Элемент</span><span class="sxs-lookup"><span data-stu-id="a0efc-151">Member</span></span> |
| [<span data-ttu-id="a0efc-152">organizer</span><span class="sxs-lookup"><span data-stu-id="a0efc-152">organizer</span></span>](#organizer-emailaddressdetailsorganizer) | <span data-ttu-id="a0efc-153">Элемент</span><span class="sxs-lookup"><span data-stu-id="a0efc-153">Member</span></span> |
| [<span data-ttu-id="a0efc-154">recurrence</span><span class="sxs-lookup"><span data-stu-id="a0efc-154">recurrence</span></span>](#nullable-recurrence-recurrence) | <span data-ttu-id="a0efc-155">Member</span><span class="sxs-lookup"><span data-stu-id="a0efc-155">Member</span></span> |
| [<span data-ttu-id="a0efc-156">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="a0efc-156">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="a0efc-157">Элемент</span><span class="sxs-lookup"><span data-stu-id="a0efc-157">Member</span></span> |
| [<span data-ttu-id="a0efc-158">sender</span><span class="sxs-lookup"><span data-stu-id="a0efc-158">sender</span></span>](#sender-emailaddressdetails) | <span data-ttu-id="a0efc-159">Элемент</span><span class="sxs-lookup"><span data-stu-id="a0efc-159">Member</span></span> |
| [<span data-ttu-id="a0efc-160">seriesId</span><span class="sxs-lookup"><span data-stu-id="a0efc-160">seriesId</span></span>](#nullable-seriesid-string) | <span data-ttu-id="a0efc-161">Элемент</span><span class="sxs-lookup"><span data-stu-id="a0efc-161">Member</span></span> |
| [<span data-ttu-id="a0efc-162">start</span><span class="sxs-lookup"><span data-stu-id="a0efc-162">start</span></span>](#start-datetime) | <span data-ttu-id="a0efc-163">Элемент</span><span class="sxs-lookup"><span data-stu-id="a0efc-163">Member</span></span> |
| [<span data-ttu-id="a0efc-164">subject</span><span class="sxs-lookup"><span data-stu-id="a0efc-164">subject</span></span>](#subject-stringsubject) | <span data-ttu-id="a0efc-165">Элемент</span><span class="sxs-lookup"><span data-stu-id="a0efc-165">Member</span></span> |
| [<span data-ttu-id="a0efc-166">to</span><span class="sxs-lookup"><span data-stu-id="a0efc-166">to</span></span>](#to-arrayemailaddressdetailsrecipients) | <span data-ttu-id="a0efc-167">Элемент</span><span class="sxs-lookup"><span data-stu-id="a0efc-167">Member</span></span> |
| [<span data-ttu-id="a0efc-168">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="a0efc-168">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="a0efc-169">Метод</span><span class="sxs-lookup"><span data-stu-id="a0efc-169">Method</span></span> |
| [<span data-ttu-id="a0efc-170">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="a0efc-170">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="a0efc-171">Метод</span><span class="sxs-lookup"><span data-stu-id="a0efc-171">Method</span></span> |
| [<span data-ttu-id="a0efc-172">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="a0efc-172">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="a0efc-173">Метод</span><span class="sxs-lookup"><span data-stu-id="a0efc-173">Method</span></span> |
| [<span data-ttu-id="a0efc-174">close</span><span class="sxs-lookup"><span data-stu-id="a0efc-174">close</span></span>](#close) | <span data-ttu-id="a0efc-175">Метод</span><span class="sxs-lookup"><span data-stu-id="a0efc-175">Method</span></span> |
| [<span data-ttu-id="a0efc-176">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="a0efc-176">displayReplyAllForm</span></span>](#displayreplyallformformdata-callback) | <span data-ttu-id="a0efc-177">Метод</span><span class="sxs-lookup"><span data-stu-id="a0efc-177">Method</span></span> |
| [<span data-ttu-id="a0efc-178">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="a0efc-178">displayReplyForm</span></span>](#displayreplyformformdata-callback) | <span data-ttu-id="a0efc-179">Метод</span><span class="sxs-lookup"><span data-stu-id="a0efc-179">Method</span></span> |
| [<span data-ttu-id="a0efc-180">getEntities</span><span class="sxs-lookup"><span data-stu-id="a0efc-180">getEntities</span></span>](#getentities--entities) | <span data-ttu-id="a0efc-181">Метод</span><span class="sxs-lookup"><span data-stu-id="a0efc-181">Method</span></span> |
| [<span data-ttu-id="a0efc-182">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="a0efc-182">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="a0efc-183">Метод</span><span class="sxs-lookup"><span data-stu-id="a0efc-183">Method</span></span> |
| [<span data-ttu-id="a0efc-184">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="a0efc-184">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="a0efc-185">Метод</span><span class="sxs-lookup"><span data-stu-id="a0efc-185">Method</span></span> |
| [<span data-ttu-id="a0efc-186">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="a0efc-186">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="a0efc-187">Метод</span><span class="sxs-lookup"><span data-stu-id="a0efc-187">Method</span></span> |
| [<span data-ttu-id="a0efc-188">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="a0efc-188">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="a0efc-189">Метод</span><span class="sxs-lookup"><span data-stu-id="a0efc-189">Method</span></span> |
| [<span data-ttu-id="a0efc-190">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="a0efc-190">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="a0efc-191">Метод</span><span class="sxs-lookup"><span data-stu-id="a0efc-191">Method</span></span> |
| [<span data-ttu-id="a0efc-192">жетселектедентитиес</span><span class="sxs-lookup"><span data-stu-id="a0efc-192">getSelectedEntities</span></span>](#getselectedentities--entities) | <span data-ttu-id="a0efc-193">Метод</span><span class="sxs-lookup"><span data-stu-id="a0efc-193">Method</span></span> |
| [<span data-ttu-id="a0efc-194">жетселектедрежексматчес</span><span class="sxs-lookup"><span data-stu-id="a0efc-194">getSelectedRegExMatches</span></span>](#getselectedregexmatches--object) | <span data-ttu-id="a0efc-195">Метод</span><span class="sxs-lookup"><span data-stu-id="a0efc-195">Method</span></span> |
| [<span data-ttu-id="a0efc-196">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="a0efc-196">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="a0efc-197">Метод</span><span class="sxs-lookup"><span data-stu-id="a0efc-197">Method</span></span> |
| [<span data-ttu-id="a0efc-198">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="a0efc-198">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="a0efc-199">Метод</span><span class="sxs-lookup"><span data-stu-id="a0efc-199">Method</span></span> |
| [<span data-ttu-id="a0efc-200">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="a0efc-200">removeHandlerAsync</span></span>](#removehandlerasynceventtype-options-callback) | <span data-ttu-id="a0efc-201">Метод</span><span class="sxs-lookup"><span data-stu-id="a0efc-201">Method</span></span> |
| [<span data-ttu-id="a0efc-202">saveAsync</span><span class="sxs-lookup"><span data-stu-id="a0efc-202">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="a0efc-203">Метод</span><span class="sxs-lookup"><span data-stu-id="a0efc-203">Method</span></span> |
| [<span data-ttu-id="a0efc-204">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="a0efc-204">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="a0efc-205">Метод</span><span class="sxs-lookup"><span data-stu-id="a0efc-205">Method</span></span> |

### <a name="example"></a><span data-ttu-id="a0efc-206">Пример</span><span class="sxs-lookup"><span data-stu-id="a0efc-206">Example</span></span>

<span data-ttu-id="a0efc-207">В примере кода JavaScript, приведенном ниже, показано, как получить доступ к свойству `subject` текущего элемента в Outlook.</span><span class="sxs-lookup"><span data-stu-id="a0efc-207">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="a0efc-208">Members</span><span class="sxs-lookup"><span data-stu-id="a0efc-208">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetailsviewoutlook-js-17"></a><span data-ttu-id="a0efc-209">attachments: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.7)></span><span class="sxs-lookup"><span data-stu-id="a0efc-209">attachments: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.7)></span></span>

<span data-ttu-id="a0efc-p102">Получает массив вложений для элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="a0efc-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="a0efc-212">Outlook блокирует определенные типы файлов из-за потенциальных проблем с безопасностью, поэтому они не возвращаются.</span><span class="sxs-lookup"><span data-stu-id="a0efc-212">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="a0efc-213">Дополнительные сведения см. в статье [Блокировка вложений в Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="a0efc-213">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="a0efc-214">Тип</span><span class="sxs-lookup"><span data-stu-id="a0efc-214">Type</span></span>

*   <span data-ttu-id="a0efc-215">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.7)></span><span class="sxs-lookup"><span data-stu-id="a0efc-215">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.7)></span></span>

##### <a name="requirements"></a><span data-ttu-id="a0efc-216">Requirements</span><span class="sxs-lookup"><span data-stu-id="a0efc-216">Requirements</span></span>

|<span data-ttu-id="a0efc-217">Требование</span><span class="sxs-lookup"><span data-stu-id="a0efc-217">Requirement</span></span>|<span data-ttu-id="a0efc-218">Значение</span><span class="sxs-lookup"><span data-stu-id="a0efc-218">Value</span></span>|
|---|---|
|[<span data-ttu-id="a0efc-219">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="a0efc-219">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a0efc-220">1.0</span><span class="sxs-lookup"><span data-stu-id="a0efc-220">1.0</span></span>|
|[<span data-ttu-id="a0efc-221">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="a0efc-221">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a0efc-222">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a0efc-222">ReadItem</span></span>|
|[<span data-ttu-id="a0efc-223">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="a0efc-223">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="a0efc-224">Чтение</span><span class="sxs-lookup"><span data-stu-id="a0efc-224">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a0efc-225">Пример</span><span class="sxs-lookup"><span data-stu-id="a0efc-225">Example</span></span>

<span data-ttu-id="a0efc-226">С помощью приведенного ниже кода можно создать HTML-строку с подробными сведениями обо всех вложениях для текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="a0efc-226">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

#### <a name="bcc-recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-17"></a><span data-ttu-id="a0efc-227">bcc: [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="a0efc-227">bcc: [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span></span>

<span data-ttu-id="a0efc-228">Получает объект, который предоставляет методы для получения или обновления получателей скрытой копии сообщения.</span><span class="sxs-lookup"><span data-stu-id="a0efc-228">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="a0efc-229">Только в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="a0efc-229">Compose mode only.</span></span>

<span data-ttu-id="a0efc-230">Коллекция может включать не более 100 элементов по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="a0efc-230">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="a0efc-231">Однако для компьютеров с Windows и Mac OS действуют указанные ниже ограничения.</span><span class="sxs-lookup"><span data-stu-id="a0efc-231">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="a0efc-232">Максимальное количество элементов — 500.</span><span class="sxs-lookup"><span data-stu-id="a0efc-232">Get 500 members maximum.</span></span>
- <span data-ttu-id="a0efc-233">Установить ограничение количества элементов на вызов — не более 100, общего количества — не более 500.</span><span class="sxs-lookup"><span data-stu-id="a0efc-233">Set a maximum of 100 members per call, up to 500 members total.</span></span>

##### <a name="type"></a><span data-ttu-id="a0efc-234">Тип</span><span class="sxs-lookup"><span data-stu-id="a0efc-234">Type</span></span>

*   [<span data-ttu-id="a0efc-235">Получатели</span><span class="sxs-lookup"><span data-stu-id="a0efc-235">Recipients</span></span>](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)

##### <a name="requirements"></a><span data-ttu-id="a0efc-236">Requirements</span><span class="sxs-lookup"><span data-stu-id="a0efc-236">Requirements</span></span>

|<span data-ttu-id="a0efc-237">Требование</span><span class="sxs-lookup"><span data-stu-id="a0efc-237">Requirement</span></span>|<span data-ttu-id="a0efc-238">Значение</span><span class="sxs-lookup"><span data-stu-id="a0efc-238">Value</span></span>|
|---|---|
|[<span data-ttu-id="a0efc-239">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="a0efc-239">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a0efc-240">1.1</span><span class="sxs-lookup"><span data-stu-id="a0efc-240">1.1</span></span>|
|[<span data-ttu-id="a0efc-241">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="a0efc-241">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a0efc-242">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a0efc-242">ReadItem</span></span>|
|[<span data-ttu-id="a0efc-243">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="a0efc-243">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="a0efc-244">Создание</span><span class="sxs-lookup"><span data-stu-id="a0efc-244">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="a0efc-245">Пример</span><span class="sxs-lookup"><span data-stu-id="a0efc-245">Example</span></span>

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

#### <a name="body-bodyjavascriptapioutlookofficebodyviewoutlook-js-17"></a><span data-ttu-id="a0efc-246">body: [Body](/javascript/api/outlook/office.body?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="a0efc-246">body: [Body](/javascript/api/outlook/office.body?view=outlook-js-1.7)</span></span>

<span data-ttu-id="a0efc-247">Получает объект, предоставляющий методы для работы с основным текстом элемента.</span><span class="sxs-lookup"><span data-stu-id="a0efc-247">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="a0efc-248">Тип</span><span class="sxs-lookup"><span data-stu-id="a0efc-248">Type</span></span>

*   [<span data-ttu-id="a0efc-249">Body</span><span class="sxs-lookup"><span data-stu-id="a0efc-249">Body</span></span>](/javascript/api/outlook/office.body?view=outlook-js-1.7)

##### <a name="requirements"></a><span data-ttu-id="a0efc-250">Требования</span><span class="sxs-lookup"><span data-stu-id="a0efc-250">Requirements</span></span>

|<span data-ttu-id="a0efc-251">Требование</span><span class="sxs-lookup"><span data-stu-id="a0efc-251">Requirement</span></span>|<span data-ttu-id="a0efc-252">Значение</span><span class="sxs-lookup"><span data-stu-id="a0efc-252">Value</span></span>|
|---|---|
|[<span data-ttu-id="a0efc-253">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="a0efc-253">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a0efc-254">1.1</span><span class="sxs-lookup"><span data-stu-id="a0efc-254">1.1</span></span>|
|[<span data-ttu-id="a0efc-255">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="a0efc-255">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a0efc-256">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a0efc-256">ReadItem</span></span>|
|[<span data-ttu-id="a0efc-257">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="a0efc-257">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="a0efc-258">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="a0efc-258">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a0efc-259">Пример</span><span class="sxs-lookup"><span data-stu-id="a0efc-259">Example</span></span>

<span data-ttu-id="a0efc-260">В этом примере возвращается текст сообщения в формате обычного текста.</span><span class="sxs-lookup"><span data-stu-id="a0efc-260">This example gets the body of the message in plain text.</span></span>

```js
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="a0efc-261">Ниже приведен пример итогового параметра, переданного функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="a0efc-261">The following is an example of the result parameter passed to the callback function.</span></span>

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

#### <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-17recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-17"></a><span data-ttu-id="a0efc-262">cc: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="a0efc-262">cc: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span></span>

<span data-ttu-id="a0efc-263">Предоставляет доступ к получателям копии сообщения.</span><span class="sxs-lookup"><span data-stu-id="a0efc-263">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="a0efc-264">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="a0efc-264">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="a0efc-265">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="a0efc-265">Read mode</span></span>

<span data-ttu-id="a0efc-266">Свойство `cc` возвращает массив, который содержит объект `EmailAddressDetails` для каждого получателя, указанного в строке **Копия** сообщения.</span><span class="sxs-lookup"><span data-stu-id="a0efc-266">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message.</span></span> <span data-ttu-id="a0efc-267">Коллекция может включать не более 100 элементов по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="a0efc-267">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="a0efc-268">Однако на компьютерах с Windows и Mac OS может быть до 500 элементов.</span><span class="sxs-lookup"><span data-stu-id="a0efc-268">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="a0efc-269">Режим создания</span><span class="sxs-lookup"><span data-stu-id="a0efc-269">Compose mode</span></span>

<span data-ttu-id="a0efc-270">Свойство `cc` возвращает объект `Recipients`, предоставляющий методы для получения или обновления получателей, которые указаны в строке **Копия** сообщения.</span><span class="sxs-lookup"><span data-stu-id="a0efc-270">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span> <span data-ttu-id="a0efc-271">Коллекция может включать не более 100 элементов по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="a0efc-271">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="a0efc-272">Однако для компьютеров с Windows и Mac OS действуют указанные ниже ограничения.</span><span class="sxs-lookup"><span data-stu-id="a0efc-272">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="a0efc-273">Максимальное количество элементов — 500.</span><span class="sxs-lookup"><span data-stu-id="a0efc-273">Get 500 members maximum.</span></span>
- <span data-ttu-id="a0efc-274">Установить ограничение количества элементов на вызов — не более 100, общего количества — не более 500.</span><span class="sxs-lookup"><span data-stu-id="a0efc-274">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="a0efc-275">Тип</span><span class="sxs-lookup"><span data-stu-id="a0efc-275">Type</span></span>

*   <span data-ttu-id="a0efc-276">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="a0efc-276">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span></span>

##### <a name="requirements"></a><span data-ttu-id="a0efc-277">Requirements</span><span class="sxs-lookup"><span data-stu-id="a0efc-277">Requirements</span></span>

|<span data-ttu-id="a0efc-278">Требование</span><span class="sxs-lookup"><span data-stu-id="a0efc-278">Requirement</span></span>|<span data-ttu-id="a0efc-279">Значение</span><span class="sxs-lookup"><span data-stu-id="a0efc-279">Value</span></span>|
|---|---|
|[<span data-ttu-id="a0efc-280">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="a0efc-280">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a0efc-281">1.0</span><span class="sxs-lookup"><span data-stu-id="a0efc-281">1.0</span></span>|
|[<span data-ttu-id="a0efc-282">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="a0efc-282">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a0efc-283">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a0efc-283">ReadItem</span></span>|
|[<span data-ttu-id="a0efc-284">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="a0efc-284">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="a0efc-285">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="a0efc-285">Compose or Read</span></span>|

<br>

---
---

#### <a name="nullable-conversationid-string"></a><span data-ttu-id="a0efc-286">(nullable) conversationId: String</span><span class="sxs-lookup"><span data-stu-id="a0efc-286">(nullable) conversationId: String</span></span>

<span data-ttu-id="a0efc-287">Получает идентификатор разговора по электронной почте, содержащего конкретное сообщение.</span><span class="sxs-lookup"><span data-stu-id="a0efc-287">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="a0efc-p109">Вы можете получить целочисленное значение этого свойства, если ваше почтовое приложение активируется в формах просмотра или формах создания ответов. Если пользователь изменит тему ответа, после его отправки идентификатор беседы будет изменен, и полученное ранее значение будет недействительным.</span><span class="sxs-lookup"><span data-stu-id="a0efc-p109">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="a0efc-p110">Это свойство имеет значение NULL для нового элемента в форме создания. Свойство `conversationId` вернет значение, если пользователь задаст тему и сохранит элемент.</span><span class="sxs-lookup"><span data-stu-id="a0efc-p110">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="a0efc-292">Тип</span><span class="sxs-lookup"><span data-stu-id="a0efc-292">Type</span></span>

*   <span data-ttu-id="a0efc-293">String</span><span class="sxs-lookup"><span data-stu-id="a0efc-293">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="a0efc-294">Requirements</span><span class="sxs-lookup"><span data-stu-id="a0efc-294">Requirements</span></span>

|<span data-ttu-id="a0efc-295">Требование</span><span class="sxs-lookup"><span data-stu-id="a0efc-295">Requirement</span></span>|<span data-ttu-id="a0efc-296">Значение</span><span class="sxs-lookup"><span data-stu-id="a0efc-296">Value</span></span>|
|---|---|
|[<span data-ttu-id="a0efc-297">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="a0efc-297">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a0efc-298">1.0</span><span class="sxs-lookup"><span data-stu-id="a0efc-298">1.0</span></span>|
|[<span data-ttu-id="a0efc-299">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="a0efc-299">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a0efc-300">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a0efc-300">ReadItem</span></span>|
|[<span data-ttu-id="a0efc-301">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="a0efc-301">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="a0efc-302">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="a0efc-302">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a0efc-303">Пример</span><span class="sxs-lookup"><span data-stu-id="a0efc-303">Example</span></span>

```js
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

<br>

---
---

#### <a name="datetimecreated-date"></a><span data-ttu-id="a0efc-304">dateTimeCreated: Date</span><span class="sxs-lookup"><span data-stu-id="a0efc-304">dateTimeCreated: Date</span></span>

<span data-ttu-id="a0efc-p111">Получает дату и время создания элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="a0efc-p111">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="a0efc-307">Тип</span><span class="sxs-lookup"><span data-stu-id="a0efc-307">Type</span></span>

*   <span data-ttu-id="a0efc-308">Дата</span><span class="sxs-lookup"><span data-stu-id="a0efc-308">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="a0efc-309">Requirements</span><span class="sxs-lookup"><span data-stu-id="a0efc-309">Requirements</span></span>

|<span data-ttu-id="a0efc-310">Требование</span><span class="sxs-lookup"><span data-stu-id="a0efc-310">Requirement</span></span>|<span data-ttu-id="a0efc-311">Значение</span><span class="sxs-lookup"><span data-stu-id="a0efc-311">Value</span></span>|
|---|---|
|[<span data-ttu-id="a0efc-312">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="a0efc-312">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a0efc-313">1.0</span><span class="sxs-lookup"><span data-stu-id="a0efc-313">1.0</span></span>|
|[<span data-ttu-id="a0efc-314">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="a0efc-314">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a0efc-315">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a0efc-315">ReadItem</span></span>|
|[<span data-ttu-id="a0efc-316">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="a0efc-316">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="a0efc-317">Чтение</span><span class="sxs-lookup"><span data-stu-id="a0efc-317">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a0efc-318">Пример</span><span class="sxs-lookup"><span data-stu-id="a0efc-318">Example</span></span>

```js
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

<br>

---
---

#### <a name="datetimemodified-date"></a><span data-ttu-id="a0efc-319">dateTimeModified: Date</span><span class="sxs-lookup"><span data-stu-id="a0efc-319">dateTimeModified: Date</span></span>

<span data-ttu-id="a0efc-p112">Получает дату и время последнего изменения элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="a0efc-p112">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="a0efc-322">Этот элемент не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="a0efc-322">This member is not supported in Outlook on iOS or Android.</span></span>

##### <a name="type"></a><span data-ttu-id="a0efc-323">Тип</span><span class="sxs-lookup"><span data-stu-id="a0efc-323">Type</span></span>

*   <span data-ttu-id="a0efc-324">Дата</span><span class="sxs-lookup"><span data-stu-id="a0efc-324">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="a0efc-325">Requirements</span><span class="sxs-lookup"><span data-stu-id="a0efc-325">Requirements</span></span>

|<span data-ttu-id="a0efc-326">Требование</span><span class="sxs-lookup"><span data-stu-id="a0efc-326">Requirement</span></span>|<span data-ttu-id="a0efc-327">Значение</span><span class="sxs-lookup"><span data-stu-id="a0efc-327">Value</span></span>|
|---|---|
|[<span data-ttu-id="a0efc-328">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="a0efc-328">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a0efc-329">1.0</span><span class="sxs-lookup"><span data-stu-id="a0efc-329">1.0</span></span>|
|[<span data-ttu-id="a0efc-330">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="a0efc-330">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a0efc-331">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a0efc-331">ReadItem</span></span>|
|[<span data-ttu-id="a0efc-332">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="a0efc-332">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="a0efc-333">Чтение</span><span class="sxs-lookup"><span data-stu-id="a0efc-333">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a0efc-334">Пример</span><span class="sxs-lookup"><span data-stu-id="a0efc-334">Example</span></span>

```js
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

<br>

---
---

#### <a name="end-datetimejavascriptapioutlookofficetimeviewoutlook-js-17"></a><span data-ttu-id="a0efc-335">end: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="a0efc-335">end: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.7)</span></span>

<span data-ttu-id="a0efc-336">Получает или задает дату и время окончания встречи.</span><span class="sxs-lookup"><span data-stu-id="a0efc-336">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="a0efc-p113">Свойство `end` представлено в виде значения даты и времени в формате UTC. Преобразовать значение свойства end в местные значения даты и времени клиента можно с помощью метода [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime).</span><span class="sxs-lookup"><span data-stu-id="a0efc-p113">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="a0efc-339">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="a0efc-339">Read mode</span></span>

<span data-ttu-id="a0efc-340">Свойство `end` возвращает объект `Date`.</span><span class="sxs-lookup"><span data-stu-id="a0efc-340">The `end` property returns a `Date` object.</span></span>

```js
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="a0efc-341">Режим создания</span><span class="sxs-lookup"><span data-stu-id="a0efc-341">Compose mode</span></span>

<span data-ttu-id="a0efc-342">Свойство `end` возвращает объект `Time`.</span><span class="sxs-lookup"><span data-stu-id="a0efc-342">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="a0efc-343">Если вы задаете время окончания с помощью метода [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.7#setasync-datetime--options--callback-), необходимо использовать метод [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) для преобразования местного времени на клиенте в формат UTC для сервера.</span><span class="sxs-lookup"><span data-stu-id="a0efc-343">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.7#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="a0efc-344">В примере ниже показано, как с помощью метода [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.7#setasync-datetime--options--callback-) объекта `Time` задать время окончания встречи.</span><span class="sxs-lookup"><span data-stu-id="a0efc-344">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.7#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="a0efc-345">Тип</span><span class="sxs-lookup"><span data-stu-id="a0efc-345">Type</span></span>

*   <span data-ttu-id="a0efc-346">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="a0efc-346">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.7)</span></span>

##### <a name="requirements"></a><span data-ttu-id="a0efc-347">Требования</span><span class="sxs-lookup"><span data-stu-id="a0efc-347">Requirements</span></span>

|<span data-ttu-id="a0efc-348">Требование</span><span class="sxs-lookup"><span data-stu-id="a0efc-348">Requirement</span></span>|<span data-ttu-id="a0efc-349">Значение</span><span class="sxs-lookup"><span data-stu-id="a0efc-349">Value</span></span>|
|---|---|
|[<span data-ttu-id="a0efc-350">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="a0efc-350">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a0efc-351">1.0</span><span class="sxs-lookup"><span data-stu-id="a0efc-351">1.0</span></span>|
|[<span data-ttu-id="a0efc-352">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="a0efc-352">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a0efc-353">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a0efc-353">ReadItem</span></span>|
|[<span data-ttu-id="a0efc-354">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="a0efc-354">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="a0efc-355">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="a0efc-355">Compose or Read</span></span>|

<br>

---
---

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-17fromjavascriptapioutlookofficefromviewoutlook-js-17"></a><span data-ttu-id="a0efc-356">от: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)|[](/javascript/api/outlook/office.from?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="a0efc-356">from: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)|[From](/javascript/api/outlook/office.from?view=outlook-js-1.7)</span></span>

<span data-ttu-id="a0efc-357">Получает электронный адрес отправителя сообщения.</span><span class="sxs-lookup"><span data-stu-id="a0efc-357">Gets the email address of the sender of a message.</span></span>

<span data-ttu-id="a0efc-p114">Свойства `from` и [`sender`](#sender-emailaddressdetails) представляют одно лицо, если сообщение не отправлено представителем. В противном случае свойство `from` представляет лицо, делегировавшее полномочия, а свойство sender — представителя.</span><span class="sxs-lookup"><span data-stu-id="a0efc-p114">The `from` and [`sender`](#sender-emailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="a0efc-360">Свойству `recipientType`, принадлежащему объекту `EmailAddressDetails` в свойстве `from`, задано значение `undefined`.</span><span class="sxs-lookup"><span data-stu-id="a0efc-360">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="a0efc-361">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="a0efc-361">Read mode</span></span>

<span data-ttu-id="a0efc-362">`from` Свойство возвращает `EmailAddressDetails` объект.</span><span class="sxs-lookup"><span data-stu-id="a0efc-362">The `from` property returns an `EmailAddressDetails` object.</span></span>

```js
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

##### <a name="compose-mode"></a><span data-ttu-id="a0efc-363">Режим создания</span><span class="sxs-lookup"><span data-stu-id="a0efc-363">Compose mode</span></span>

<span data-ttu-id="a0efc-364">`from` Свойство возвращает `From` объект, который предоставляет метод для получения значения From.</span><span class="sxs-lookup"><span data-stu-id="a0efc-364">The `from` property returns a `From` object that provides a method to get the from value.</span></span>

```js
Office.context.mailbox.item.from.getAsync(callback);

function callback(asyncResult) {
  var from = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="a0efc-365">Тип</span><span class="sxs-lookup"><span data-stu-id="a0efc-365">Type</span></span>

*   <span data-ttu-id="a0efc-366">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7) | [из](/javascript/api/outlook/office.from?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="a0efc-366">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7) | [From](/javascript/api/outlook/office.from?view=outlook-js-1.7)</span></span>

##### <a name="requirements"></a><span data-ttu-id="a0efc-367">Requirements</span><span class="sxs-lookup"><span data-stu-id="a0efc-367">Requirements</span></span>

|<span data-ttu-id="a0efc-368">Требование</span><span class="sxs-lookup"><span data-stu-id="a0efc-368">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="a0efc-369">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="a0efc-369">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a0efc-370">1.0</span><span class="sxs-lookup"><span data-stu-id="a0efc-370">1.0</span></span>|<span data-ttu-id="a0efc-371">1.7</span><span class="sxs-lookup"><span data-stu-id="a0efc-371">1.7</span></span>|
|[<span data-ttu-id="a0efc-372">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="a0efc-372">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a0efc-373">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a0efc-373">ReadItem</span></span>|<span data-ttu-id="a0efc-374">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="a0efc-374">ReadWriteItem</span></span>|
|[<span data-ttu-id="a0efc-375">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="a0efc-375">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="a0efc-376">Чтение</span><span class="sxs-lookup"><span data-stu-id="a0efc-376">Read</span></span>|<span data-ttu-id="a0efc-377">Создание</span><span class="sxs-lookup"><span data-stu-id="a0efc-377">Compose</span></span>|

<br>

---
---

#### <a name="internetmessageid-string"></a><span data-ttu-id="a0efc-378">internetMessageId: String</span><span class="sxs-lookup"><span data-stu-id="a0efc-378">internetMessageId: String</span></span>

<span data-ttu-id="a0efc-p115">Получает идентификатор интернет-сообщения для электронного сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="a0efc-p115">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="a0efc-381">Тип</span><span class="sxs-lookup"><span data-stu-id="a0efc-381">Type</span></span>

*   <span data-ttu-id="a0efc-382">String</span><span class="sxs-lookup"><span data-stu-id="a0efc-382">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="a0efc-383">Requirements</span><span class="sxs-lookup"><span data-stu-id="a0efc-383">Requirements</span></span>

|<span data-ttu-id="a0efc-384">Требование</span><span class="sxs-lookup"><span data-stu-id="a0efc-384">Requirement</span></span>|<span data-ttu-id="a0efc-385">Значение</span><span class="sxs-lookup"><span data-stu-id="a0efc-385">Value</span></span>|
|---|---|
|[<span data-ttu-id="a0efc-386">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="a0efc-386">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a0efc-387">1.0</span><span class="sxs-lookup"><span data-stu-id="a0efc-387">1.0</span></span>|
|[<span data-ttu-id="a0efc-388">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="a0efc-388">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a0efc-389">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a0efc-389">ReadItem</span></span>|
|[<span data-ttu-id="a0efc-390">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="a0efc-390">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="a0efc-391">Чтение</span><span class="sxs-lookup"><span data-stu-id="a0efc-391">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a0efc-392">Пример</span><span class="sxs-lookup"><span data-stu-id="a0efc-392">Example</span></span>

```js
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

<br>

---
---

#### <a name="itemclass-string"></a><span data-ttu-id="a0efc-393">itemClass: String</span><span class="sxs-lookup"><span data-stu-id="a0efc-393">itemClass: String</span></span>

<span data-ttu-id="a0efc-p116">Получает класс элемента веб-служб Exchange для выбранного элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="a0efc-p116">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="a0efc-p117">Свойство `itemClass` указывает класс сообщения выбранного элемента. Ниже приводятся классы сообщения по умолчанию для элемента сообщения или встречи.</span><span class="sxs-lookup"><span data-stu-id="a0efc-p117">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

|<span data-ttu-id="a0efc-398">Тип</span><span class="sxs-lookup"><span data-stu-id="a0efc-398">Type</span></span>|<span data-ttu-id="a0efc-399">Описание</span><span class="sxs-lookup"><span data-stu-id="a0efc-399">Description</span></span>|<span data-ttu-id="a0efc-400">Класс элемента</span><span class="sxs-lookup"><span data-stu-id="a0efc-400">item class</span></span>|
|---|---|---|
|<span data-ttu-id="a0efc-401">Элементы встречи</span><span class="sxs-lookup"><span data-stu-id="a0efc-401">Appointment items</span></span>|<span data-ttu-id="a0efc-402">Это элементы календаря для класса элемента `IPM.Appointment` или `IPM.Appointment.Occurrence`.</span><span class="sxs-lookup"><span data-stu-id="a0efc-402">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span>|`IPM.Appointment`<br />`IPM.Appointment.Occurrence`|
|<span data-ttu-id="a0efc-403">Элементы сообщения</span><span class="sxs-lookup"><span data-stu-id="a0efc-403">Message items</span></span>|<span data-ttu-id="a0efc-404">Сюда входят электронные сообщения, для которых по умолчанию задан класс сообщения `IPM.Note`, а также приглашения на собрания, ответы на них и уведомления об их отмене, использующие `IPM.Schedule.Meeting` в качестве базового класса сообщения.</span><span class="sxs-lookup"><span data-stu-id="a0efc-404">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span>|`IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled`|

<span data-ttu-id="a0efc-405">Можно создавать настраиваемые классы сообщения, расширяющие классы сообщения по умолчанию, например настраиваемый класс сообщения о встрече `IPM.Appointment.Contoso`.</span><span class="sxs-lookup"><span data-stu-id="a0efc-405">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="a0efc-406">Тип</span><span class="sxs-lookup"><span data-stu-id="a0efc-406">Type</span></span>

*   <span data-ttu-id="a0efc-407">String</span><span class="sxs-lookup"><span data-stu-id="a0efc-407">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="a0efc-408">Requirements</span><span class="sxs-lookup"><span data-stu-id="a0efc-408">Requirements</span></span>

|<span data-ttu-id="a0efc-409">Требование</span><span class="sxs-lookup"><span data-stu-id="a0efc-409">Requirement</span></span>|<span data-ttu-id="a0efc-410">Значение</span><span class="sxs-lookup"><span data-stu-id="a0efc-410">Value</span></span>|
|---|---|
|[<span data-ttu-id="a0efc-411">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="a0efc-411">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a0efc-412">1.0</span><span class="sxs-lookup"><span data-stu-id="a0efc-412">1.0</span></span>|
|[<span data-ttu-id="a0efc-413">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="a0efc-413">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a0efc-414">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a0efc-414">ReadItem</span></span>|
|[<span data-ttu-id="a0efc-415">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="a0efc-415">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="a0efc-416">Чтение</span><span class="sxs-lookup"><span data-stu-id="a0efc-416">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a0efc-417">Пример</span><span class="sxs-lookup"><span data-stu-id="a0efc-417">Example</span></span>

```js
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

<br>

---
---

#### <a name="nullable-itemid-string"></a><span data-ttu-id="a0efc-418">(nullable) itemId: String</span><span class="sxs-lookup"><span data-stu-id="a0efc-418">(nullable) itemId: String</span></span>

<span data-ttu-id="a0efc-p118">Получает [идентификатор элемента веб-служб Exchange](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange) для текущего элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="a0efc-p118">Gets the [Exchange Web Services item identifier](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange) for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="a0efc-421">Идентификатор, возвращаемый свойством `itemId`, совпадает с [идентификатором элемента веб-служб Exchange](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange).</span><span class="sxs-lookup"><span data-stu-id="a0efc-421">The identifier returned by the `itemId` property is the same as the [Exchange Web Services item identifier](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange).</span></span> <span data-ttu-id="a0efc-422">Свойство `itemId` не совпадает с идентификатором записи Outlook, а также идентификатором, который используется REST API Outlook.</span><span class="sxs-lookup"><span data-stu-id="a0efc-422">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="a0efc-423">Прежде чем совершать вызовы REST API, используя это значение, его необходимо преобразовать с помощью [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="a0efc-423">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="a0efc-424">Дополнительные сведения см. в статье [Использование REST API Outlook из надстройки Outlook](/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="a0efc-424">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="a0efc-p120">Свойство `itemId` недоступно в режиме создания. Если требуется идентификатор элемента, с помощью метода [`saveAsync`](#saveasyncoptions-callback) можно сохранить элемент в хранилище. При этом в параметре [`AsyncResult.value`](/javascript/api/office/office.asyncresult) функции обратного вызова возвращается идентификатор элемента.</span><span class="sxs-lookup"><span data-stu-id="a0efc-p120">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="a0efc-427">Тип</span><span class="sxs-lookup"><span data-stu-id="a0efc-427">Type</span></span>

*   <span data-ttu-id="a0efc-428">String</span><span class="sxs-lookup"><span data-stu-id="a0efc-428">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="a0efc-429">Requirements</span><span class="sxs-lookup"><span data-stu-id="a0efc-429">Requirements</span></span>

|<span data-ttu-id="a0efc-430">Требование</span><span class="sxs-lookup"><span data-stu-id="a0efc-430">Requirement</span></span>|<span data-ttu-id="a0efc-431">Значение</span><span class="sxs-lookup"><span data-stu-id="a0efc-431">Value</span></span>|
|---|---|
|[<span data-ttu-id="a0efc-432">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="a0efc-432">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a0efc-433">1.0</span><span class="sxs-lookup"><span data-stu-id="a0efc-433">1.0</span></span>|
|[<span data-ttu-id="a0efc-434">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="a0efc-434">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a0efc-435">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a0efc-435">ReadItem</span></span>|
|[<span data-ttu-id="a0efc-436">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="a0efc-436">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="a0efc-437">Чтение</span><span class="sxs-lookup"><span data-stu-id="a0efc-437">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a0efc-438">Пример</span><span class="sxs-lookup"><span data-stu-id="a0efc-438">Example</span></span>

<span data-ttu-id="a0efc-p121">Указанный ниже код проверяет наличие идентификатора элемента. Если свойство `itemId` возвращает значение `null` или `undefined`, элемент будет сохранен в хранилище, а из асинхронного результата будет получен идентификатор элемента.</span><span class="sxs-lookup"><span data-stu-id="a0efc-p121">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

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

#### <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtypeviewoutlook-js-17"></a><span data-ttu-id="a0efc-441">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="a0efc-441">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.7)</span></span>

<span data-ttu-id="a0efc-442">Получает тип элемента, который представляет экземпляр.</span><span class="sxs-lookup"><span data-stu-id="a0efc-442">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="a0efc-443">Свойство `itemType` возвращает одно из значений перечисления `ItemType`, которое указывает, является ли экземпляр объекта `item` сообщением или встречей.</span><span class="sxs-lookup"><span data-stu-id="a0efc-443">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="a0efc-444">Тип</span><span class="sxs-lookup"><span data-stu-id="a0efc-444">Type</span></span>

*   [<span data-ttu-id="a0efc-445">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="a0efc-445">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.7)

##### <a name="requirements"></a><span data-ttu-id="a0efc-446">Requirements</span><span class="sxs-lookup"><span data-stu-id="a0efc-446">Requirements</span></span>

|<span data-ttu-id="a0efc-447">Требование</span><span class="sxs-lookup"><span data-stu-id="a0efc-447">Requirement</span></span>|<span data-ttu-id="a0efc-448">Значение</span><span class="sxs-lookup"><span data-stu-id="a0efc-448">Value</span></span>|
|---|---|
|[<span data-ttu-id="a0efc-449">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="a0efc-449">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a0efc-450">1.0</span><span class="sxs-lookup"><span data-stu-id="a0efc-450">1.0</span></span>|
|[<span data-ttu-id="a0efc-451">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="a0efc-451">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a0efc-452">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a0efc-452">ReadItem</span></span>|
|[<span data-ttu-id="a0efc-453">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="a0efc-453">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="a0efc-454">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="a0efc-454">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a0efc-455">Пример</span><span class="sxs-lookup"><span data-stu-id="a0efc-455">Example</span></span>

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

#### <a name="location-stringlocationjavascriptapioutlookofficelocationviewoutlook-js-17"></a><span data-ttu-id="a0efc-456">location: String|[Location](/javascript/api/outlook/office.location?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="a0efc-456">location: String|[Location](/javascript/api/outlook/office.location?view=outlook-js-1.7)</span></span>

<span data-ttu-id="a0efc-457">Получает или задает место встречи.</span><span class="sxs-lookup"><span data-stu-id="a0efc-457">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="a0efc-458">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="a0efc-458">Read mode</span></span>

<span data-ttu-id="a0efc-459">Свойство `location` возвращает строку, содержащую сведения о месте встречи.</span><span class="sxs-lookup"><span data-stu-id="a0efc-459">The `location` property returns a string that contains the location of the appointment.</span></span>

```js
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="a0efc-460">Режим создания</span><span class="sxs-lookup"><span data-stu-id="a0efc-460">Compose mode</span></span>

<span data-ttu-id="a0efc-461">Свойство `location` возвращает объект `Location`, предоставляющий методы, которые используются для получения и задания места встречи.</span><span class="sxs-lookup"><span data-stu-id="a0efc-461">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```js
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="a0efc-462">Тип</span><span class="sxs-lookup"><span data-stu-id="a0efc-462">Type</span></span>

*   <span data-ttu-id="a0efc-463">String | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="a0efc-463">String | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.7)</span></span>

##### <a name="requirements"></a><span data-ttu-id="a0efc-464">Requirements</span><span class="sxs-lookup"><span data-stu-id="a0efc-464">Requirements</span></span>

|<span data-ttu-id="a0efc-465">Требование</span><span class="sxs-lookup"><span data-stu-id="a0efc-465">Requirement</span></span>|<span data-ttu-id="a0efc-466">Значение</span><span class="sxs-lookup"><span data-stu-id="a0efc-466">Value</span></span>|
|---|---|
|[<span data-ttu-id="a0efc-467">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="a0efc-467">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a0efc-468">1.0</span><span class="sxs-lookup"><span data-stu-id="a0efc-468">1.0</span></span>|
|[<span data-ttu-id="a0efc-469">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="a0efc-469">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a0efc-470">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a0efc-470">ReadItem</span></span>|
|[<span data-ttu-id="a0efc-471">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="a0efc-471">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="a0efc-472">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="a0efc-472">Compose or Read</span></span>|

<br>

---
---

#### <a name="normalizedsubject-string"></a><span data-ttu-id="a0efc-473">normalizedSubject: String</span><span class="sxs-lookup"><span data-stu-id="a0efc-473">normalizedSubject: String</span></span>

<span data-ttu-id="a0efc-p122">Получает тему элемента со всеми удаленными префиксами (включая `RE:` и `FWD:`). Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="a0efc-p122">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="a0efc-p123">Свойство normalizedSubject получает тему элемента со стандартными префиксами (такими как `RE:` и `FW:`), добавляемыми почтовыми программами. Для получения темы элемента с неизмененными префиксами используйте свойство [`subject`](#subject-stringsubject).</span><span class="sxs-lookup"><span data-stu-id="a0efc-p123">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="a0efc-478">Тип</span><span class="sxs-lookup"><span data-stu-id="a0efc-478">Type</span></span>

*   <span data-ttu-id="a0efc-479">String</span><span class="sxs-lookup"><span data-stu-id="a0efc-479">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="a0efc-480">Requirements</span><span class="sxs-lookup"><span data-stu-id="a0efc-480">Requirements</span></span>

|<span data-ttu-id="a0efc-481">Требование</span><span class="sxs-lookup"><span data-stu-id="a0efc-481">Requirement</span></span>|<span data-ttu-id="a0efc-482">Значение</span><span class="sxs-lookup"><span data-stu-id="a0efc-482">Value</span></span>|
|---|---|
|[<span data-ttu-id="a0efc-483">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="a0efc-483">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a0efc-484">1.0</span><span class="sxs-lookup"><span data-stu-id="a0efc-484">1.0</span></span>|
|[<span data-ttu-id="a0efc-485">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="a0efc-485">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a0efc-486">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a0efc-486">ReadItem</span></span>|
|[<span data-ttu-id="a0efc-487">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="a0efc-487">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="a0efc-488">Чтение</span><span class="sxs-lookup"><span data-stu-id="a0efc-488">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a0efc-489">Пример</span><span class="sxs-lookup"><span data-stu-id="a0efc-489">Example</span></span>

```js
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

<br>

---
---

#### <a name="notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessagesviewoutlook-js-17"></a><span data-ttu-id="a0efc-490">notificationMessages: [NotificationMessages](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="a0efc-490">notificationMessages: [NotificationMessages](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.7)</span></span>

<span data-ttu-id="a0efc-491">Получает сообщения уведомления для элемента.</span><span class="sxs-lookup"><span data-stu-id="a0efc-491">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="a0efc-492">Тип</span><span class="sxs-lookup"><span data-stu-id="a0efc-492">Type</span></span>

*   [<span data-ttu-id="a0efc-493">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="a0efc-493">NotificationMessages</span></span>](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.7)

##### <a name="requirements"></a><span data-ttu-id="a0efc-494">Требования</span><span class="sxs-lookup"><span data-stu-id="a0efc-494">Requirements</span></span>

|<span data-ttu-id="a0efc-495">Требование</span><span class="sxs-lookup"><span data-stu-id="a0efc-495">Requirement</span></span>|<span data-ttu-id="a0efc-496">Значение</span><span class="sxs-lookup"><span data-stu-id="a0efc-496">Value</span></span>|
|---|---|
|[<span data-ttu-id="a0efc-497">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="a0efc-497">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a0efc-498">1.3</span><span class="sxs-lookup"><span data-stu-id="a0efc-498">1.3</span></span>|
|[<span data-ttu-id="a0efc-499">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="a0efc-499">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a0efc-500">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a0efc-500">ReadItem</span></span>|
|[<span data-ttu-id="a0efc-501">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="a0efc-501">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="a0efc-502">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="a0efc-502">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a0efc-503">Пример</span><span class="sxs-lookup"><span data-stu-id="a0efc-503">Example</span></span>

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

#### <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-17recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-17"></a><span data-ttu-id="a0efc-504">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="a0efc-504">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span></span>

<span data-ttu-id="a0efc-505">Предоставляет доступ к необязательным участникам события.</span><span class="sxs-lookup"><span data-stu-id="a0efc-505">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="a0efc-506">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="a0efc-506">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="a0efc-507">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="a0efc-507">Read mode</span></span>

<span data-ttu-id="a0efc-508">Свойство `optionalAttendees` возвращает массив, содержащий объект `EmailAddressDetails` для каждого необязательного участника собрания.</span><span class="sxs-lookup"><span data-stu-id="a0efc-508">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span> <span data-ttu-id="a0efc-509">Коллекция может включать не более 100 элементов по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="a0efc-509">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="a0efc-510">Однако на компьютерах с Windows и Mac OS может быть до 500 элементов.</span><span class="sxs-lookup"><span data-stu-id="a0efc-510">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="a0efc-511">Режим создания</span><span class="sxs-lookup"><span data-stu-id="a0efc-511">Compose mode</span></span>

<span data-ttu-id="a0efc-512">Свойство `optionalAttendees` возвращает объект `Recipients`, который предоставляет методы для получения или обновления необязательных участников собрания.</span><span class="sxs-lookup"><span data-stu-id="a0efc-512">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span> <span data-ttu-id="a0efc-513">Коллекция может включать не более 100 элементов по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="a0efc-513">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="a0efc-514">Однако для компьютеров с Windows и Mac OS действуют указанные ниже ограничения.</span><span class="sxs-lookup"><span data-stu-id="a0efc-514">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="a0efc-515">Максимальное количество элементов — 500.</span><span class="sxs-lookup"><span data-stu-id="a0efc-515">Get 500 members maximum.</span></span>
- <span data-ttu-id="a0efc-516">Установить ограничение количества элементов на вызов — не более 100, общего количества — не более 500.</span><span class="sxs-lookup"><span data-stu-id="a0efc-516">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="a0efc-517">Тип</span><span class="sxs-lookup"><span data-stu-id="a0efc-517">Type</span></span>

*   <span data-ttu-id="a0efc-518">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="a0efc-518">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span></span>

##### <a name="requirements"></a><span data-ttu-id="a0efc-519">Требования</span><span class="sxs-lookup"><span data-stu-id="a0efc-519">Requirements</span></span>

|<span data-ttu-id="a0efc-520">Требование</span><span class="sxs-lookup"><span data-stu-id="a0efc-520">Requirement</span></span>|<span data-ttu-id="a0efc-521">Значение</span><span class="sxs-lookup"><span data-stu-id="a0efc-521">Value</span></span>|
|---|---|
|[<span data-ttu-id="a0efc-522">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="a0efc-522">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a0efc-523">1.0</span><span class="sxs-lookup"><span data-stu-id="a0efc-523">1.0</span></span>|
|[<span data-ttu-id="a0efc-524">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="a0efc-524">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a0efc-525">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a0efc-525">ReadItem</span></span>|
|[<span data-ttu-id="a0efc-526">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="a0efc-526">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="a0efc-527">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="a0efc-527">Compose or Read</span></span>|

<br>

---
---

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-17organizerjavascriptapioutlookofficeorganizerviewoutlook-js-17"></a><span data-ttu-id="a0efc-528">Организатор: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)|[Организатор](/javascript/api/outlook/office.organizer?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="a0efc-528">organizer: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)|[Organizer](/javascript/api/outlook/office.organizer?view=outlook-js-1.7)</span></span>

<span data-ttu-id="a0efc-529">Получает адрес электронной почты организатора для указанного собрания.</span><span class="sxs-lookup"><span data-stu-id="a0efc-529">Gets the email address of the organizer for a specified meeting.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="a0efc-530">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="a0efc-530">Read mode</span></span>

<span data-ttu-id="a0efc-531">`organizer` Свойство возвращает объект [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7) , представляющий организатора собрания.</span><span class="sxs-lookup"><span data-stu-id="a0efc-531">The `organizer` property returns an [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7) object that represents the meeting organizer.</span></span>

```js
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

##### <a name="compose-mode"></a><span data-ttu-id="a0efc-532">Режим создания</span><span class="sxs-lookup"><span data-stu-id="a0efc-532">Compose mode</span></span>

<span data-ttu-id="a0efc-533">`organizer` Свойство возвращает объект [организатора](/javascript/api/outlook/office.organizer?view=outlook-js-1.7) , который предоставляет метод для получения значения организатора.</span><span class="sxs-lookup"><span data-stu-id="a0efc-533">The `organizer` property returns an [Organizer](/javascript/api/outlook/office.organizer?view=outlook-js-1.7) object that provides a method to get the organizer value.</span></span>

```js
Office.context.mailbox.item.organizer.getAsync(
  function(asyncResult) {
    console.log(JSON.stringify(asyncResult));
  }
);
```

##### <a name="type"></a><span data-ttu-id="a0efc-534">Тип</span><span class="sxs-lookup"><span data-stu-id="a0efc-534">Type</span></span>

*   <span data-ttu-id="a0efc-535">[](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7) | [Организатор](/javascript/api/outlook/office.organizer?view=outlook-js-1.7) EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="a0efc-535">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7) | [Organizer](/javascript/api/outlook/office.organizer?view=outlook-js-1.7)</span></span>

##### <a name="requirements"></a><span data-ttu-id="a0efc-536">Requirements</span><span class="sxs-lookup"><span data-stu-id="a0efc-536">Requirements</span></span>

|<span data-ttu-id="a0efc-537">Требование</span><span class="sxs-lookup"><span data-stu-id="a0efc-537">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="a0efc-538">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="a0efc-538">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a0efc-539">1.0</span><span class="sxs-lookup"><span data-stu-id="a0efc-539">1.0</span></span>|<span data-ttu-id="a0efc-540">1.7</span><span class="sxs-lookup"><span data-stu-id="a0efc-540">1.7</span></span>|
|[<span data-ttu-id="a0efc-541">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="a0efc-541">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a0efc-542">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a0efc-542">ReadItem</span></span>|<span data-ttu-id="a0efc-543">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="a0efc-543">ReadWriteItem</span></span>|
|[<span data-ttu-id="a0efc-544">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="a0efc-544">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="a0efc-545">Чтение</span><span class="sxs-lookup"><span data-stu-id="a0efc-545">Read</span></span>|<span data-ttu-id="a0efc-546">Создание</span><span class="sxs-lookup"><span data-stu-id="a0efc-546">Compose</span></span>|

<br>

---
---

#### <a name="nullable-recurrence-recurrencejavascriptapioutlookofficerecurrenceviewoutlook-js-17"></a><span data-ttu-id="a0efc-547">(Nullable) повторение: [повторение](/javascript/api/outlook/office.recurrence?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="a0efc-547">(nullable) recurrence: [Recurrence](/javascript/api/outlook/office.recurrence?view=outlook-js-1.7)</span></span>

<span data-ttu-id="a0efc-548">Получает или задает шаблон повторения встречи.</span><span class="sxs-lookup"><span data-stu-id="a0efc-548">Gets or sets the recurrence pattern of an appointment.</span></span> <span data-ttu-id="a0efc-549">Получает шаблон повторения приглашения на собрание.</span><span class="sxs-lookup"><span data-stu-id="a0efc-549">Gets the recurrence pattern of a meeting request.</span></span> <span data-ttu-id="a0efc-550">Режимы чтения и создания для элементов встречи.</span><span class="sxs-lookup"><span data-stu-id="a0efc-550">Read and compose modes for appointment items.</span></span> <span data-ttu-id="a0efc-551">Режим чтения для элементов приглашения на собрания.</span><span class="sxs-lookup"><span data-stu-id="a0efc-551">Read mode for meeting request items.</span></span>

<span data-ttu-id="a0efc-552">`recurrence` Свойство возвращает объект [повторения](/javascript/api/outlook/office.recurrence?view=outlook-js-1.7) для повторяющихся встреч или приглашений на собрания, если элемент представляет собой серию или экземпляр в ряду.</span><span class="sxs-lookup"><span data-stu-id="a0efc-552">The `recurrence` property returns a [recurrence](/javascript/api/outlook/office.recurrence?view=outlook-js-1.7) object for recurring appointments or meetings requests if an item is a series or an instance in a series.</span></span> <span data-ttu-id="a0efc-553">`null`возвращается для отдельных встреч и приглашений на собрание для отдельных встреч.</span><span class="sxs-lookup"><span data-stu-id="a0efc-553">`null` is returned for single appointments and meeting requests of single appointments.</span></span> <span data-ttu-id="a0efc-554">`undefined`возвращается для сообщений, которые не являются приглашениями на собрания.</span><span class="sxs-lookup"><span data-stu-id="a0efc-554">`undefined` is returned for messages that are not meeting requests.</span></span>

> <span data-ttu-id="a0efc-555">Note: приглашения на `itemClass` собрания имеют значение IPM. Schedule. Meeting. Request.</span><span class="sxs-lookup"><span data-stu-id="a0efc-555">Note: Meeting requests have an `itemClass` value of IPM.Schedule.Meeting.Request.</span></span>

> <span data-ttu-id="a0efc-556">Note: при наличии объекта `null`повторения это указывает на то, что объект является одной встречей или приглашением на собрание одной встречи, а не частью ряда.</span><span class="sxs-lookup"><span data-stu-id="a0efc-556">Note: If the recurrence object is `null`, this indicates that the object is a single appointment or a meeting request of a single appointment and NOT a part of a series.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="a0efc-557">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="a0efc-557">Read mode</span></span>

<span data-ttu-id="a0efc-558">`recurrence` Свойство возвращает объект [повторения](/javascript/api/outlook/office.recurrence?view=outlook-js-1.7) , представляющий повторение встречи.</span><span class="sxs-lookup"><span data-stu-id="a0efc-558">The `recurrence` property returns a [Recurrence](/javascript/api/outlook/office.recurrence?view=outlook-js-1.7) object that represents the appointment recurrence.</span></span> <span data-ttu-id="a0efc-559">Оно доступно для встреч и приглашений на собрания.</span><span class="sxs-lookup"><span data-stu-id="a0efc-559">This is available for appointments and meeting requests.</span></span>

```js
var recurrence = Office.context.mailbox.item.recurrence;
console.log("Recurrence: " + JSON.stringify(recurrence));
```

##### <a name="compose-mode"></a><span data-ttu-id="a0efc-560">Режим создания</span><span class="sxs-lookup"><span data-stu-id="a0efc-560">Compose mode</span></span>

<span data-ttu-id="a0efc-561">`recurrence` Свойство возвращает объект [повторения](/javascript/api/outlook/office.recurrence?view=outlook-js-1.7) , который предоставляет методы для управления повторением встречи.</span><span class="sxs-lookup"><span data-stu-id="a0efc-561">The `recurrence` property returns a [Recurrence](/javascript/api/outlook/office.recurrence?view=outlook-js-1.7) object that provides methods to manage the appointment recurrence.</span></span> <span data-ttu-id="a0efc-562">Оно доступно для встреч.</span><span class="sxs-lookup"><span data-stu-id="a0efc-562">This is available for appointments.</span></span>

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

##### <a name="type"></a><span data-ttu-id="a0efc-563">Тип</span><span class="sxs-lookup"><span data-stu-id="a0efc-563">Type</span></span>

* [<span data-ttu-id="a0efc-564">Повторения</span><span class="sxs-lookup"><span data-stu-id="a0efc-564">Recurrence</span></span>](/javascript/api/outlook/office.recurrence?view=outlook-js-1.7)

|<span data-ttu-id="a0efc-565">Requirement</span><span class="sxs-lookup"><span data-stu-id="a0efc-565">Requirement</span></span>|<span data-ttu-id="a0efc-566">Значение</span><span class="sxs-lookup"><span data-stu-id="a0efc-566">Value</span></span>|
|---|---|
|[<span data-ttu-id="a0efc-567">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="a0efc-567">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a0efc-568">1.7</span><span class="sxs-lookup"><span data-stu-id="a0efc-568">1.7</span></span>|
|[<span data-ttu-id="a0efc-569">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="a0efc-569">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a0efc-570">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a0efc-570">ReadItem</span></span>|
|[<span data-ttu-id="a0efc-571">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="a0efc-571">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="a0efc-572">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="a0efc-572">Compose or Read</span></span>|

<br>

---
---

#### <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-17recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-17"></a><span data-ttu-id="a0efc-573">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="a0efc-573">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span></span>

<span data-ttu-id="a0efc-574">Предоставляет доступ к обязательным участникам события.</span><span class="sxs-lookup"><span data-stu-id="a0efc-574">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="a0efc-575">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="a0efc-575">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="a0efc-576">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="a0efc-576">Read mode</span></span>

<span data-ttu-id="a0efc-577">Свойство `requiredAttendees` возвращает массив, содержащий объект `EmailAddressDetails` для каждого обязательного участника собрания.</span><span class="sxs-lookup"><span data-stu-id="a0efc-577">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span> <span data-ttu-id="a0efc-578">Коллекция может включать не более 100 элементов по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="a0efc-578">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="a0efc-579">Однако на компьютерах с Windows и Mac OS может быть до 500 элементов.</span><span class="sxs-lookup"><span data-stu-id="a0efc-579">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="a0efc-580">Режим создания</span><span class="sxs-lookup"><span data-stu-id="a0efc-580">Compose mode</span></span>

<span data-ttu-id="a0efc-581">Свойство `requiredAttendees` возвращает объект `Recipients`, предоставляющий методы, с помощью которых можно получить или обновить сведения об обязательных участниках собрания.</span><span class="sxs-lookup"><span data-stu-id="a0efc-581">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span> <span data-ttu-id="a0efc-582">Коллекция может включать не более 100 элементов по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="a0efc-582">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="a0efc-583">Однако для компьютеров с Windows и Mac OS действуют указанные ниже ограничения.</span><span class="sxs-lookup"><span data-stu-id="a0efc-583">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="a0efc-584">Максимальное количество элементов — 500.</span><span class="sxs-lookup"><span data-stu-id="a0efc-584">Get 500 members maximum.</span></span>
- <span data-ttu-id="a0efc-585">Установить ограничение количества элементов на вызов — не более 100, общего количества — не более 500.</span><span class="sxs-lookup"><span data-stu-id="a0efc-585">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="a0efc-586">Тип</span><span class="sxs-lookup"><span data-stu-id="a0efc-586">Type</span></span>

*   <span data-ttu-id="a0efc-587">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="a0efc-587">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span></span>

##### <a name="requirements"></a><span data-ttu-id="a0efc-588">Требования</span><span class="sxs-lookup"><span data-stu-id="a0efc-588">Requirements</span></span>

|<span data-ttu-id="a0efc-589">Требование</span><span class="sxs-lookup"><span data-stu-id="a0efc-589">Requirement</span></span>|<span data-ttu-id="a0efc-590">Значение</span><span class="sxs-lookup"><span data-stu-id="a0efc-590">Value</span></span>|
|---|---|
|[<span data-ttu-id="a0efc-591">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="a0efc-591">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a0efc-592">1.0</span><span class="sxs-lookup"><span data-stu-id="a0efc-592">1.0</span></span>|
|[<span data-ttu-id="a0efc-593">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="a0efc-593">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a0efc-594">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a0efc-594">ReadItem</span></span>|
|[<span data-ttu-id="a0efc-595">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="a0efc-595">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="a0efc-596">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="a0efc-596">Compose or Read</span></span>|

<br>

---
---

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-17"></a><span data-ttu-id="a0efc-597">sender: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="a0efc-597">sender: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)</span></span>

<span data-ttu-id="a0efc-p134">Получает электронный адрес отправителя электронного сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="a0efc-p134">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="a0efc-p135">Свойства [`from`](#from-emailaddressdetailsfrom) и `sender` представляют одно лицо, если сообщение не отправлено представителем. В противном случае свойство `from` представляет лицо, делегировавшее полномочия, а свойство sender — представителя.</span><span class="sxs-lookup"><span data-stu-id="a0efc-p135">The [`from`](#from-emailaddressdetailsfrom) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="a0efc-602">Свойству `recipientType`, принадлежащему объекту `EmailAddressDetails` в свойстве `sender`, задано значение `undefined`.</span><span class="sxs-lookup"><span data-stu-id="a0efc-602">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="a0efc-603">Тип</span><span class="sxs-lookup"><span data-stu-id="a0efc-603">Type</span></span>

*   [<span data-ttu-id="a0efc-604">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="a0efc-604">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)

##### <a name="requirements"></a><span data-ttu-id="a0efc-605">Requirements</span><span class="sxs-lookup"><span data-stu-id="a0efc-605">Requirements</span></span>

|<span data-ttu-id="a0efc-606">Требование</span><span class="sxs-lookup"><span data-stu-id="a0efc-606">Requirement</span></span>|<span data-ttu-id="a0efc-607">Значение</span><span class="sxs-lookup"><span data-stu-id="a0efc-607">Value</span></span>|
|---|---|
|[<span data-ttu-id="a0efc-608">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="a0efc-608">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a0efc-609">1.0</span><span class="sxs-lookup"><span data-stu-id="a0efc-609">1.0</span></span>|
|[<span data-ttu-id="a0efc-610">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="a0efc-610">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a0efc-611">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a0efc-611">ReadItem</span></span>|
|[<span data-ttu-id="a0efc-612">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="a0efc-612">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="a0efc-613">Чтение</span><span class="sxs-lookup"><span data-stu-id="a0efc-613">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a0efc-614">Пример</span><span class="sxs-lookup"><span data-stu-id="a0efc-614">Example</span></span>

```js
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

<br>

---
---

#### <a name="nullable-seriesid-string"></a><span data-ttu-id="a0efc-615">(Nullable) seriesId: строка</span><span class="sxs-lookup"><span data-stu-id="a0efc-615">(nullable) seriesId: String</span></span>

<span data-ttu-id="a0efc-616">Получает идентификатор ряда, к которому принадлежит экземпляр.</span><span class="sxs-lookup"><span data-stu-id="a0efc-616">Gets the id of the series that an instance belongs to.</span></span>

<span data-ttu-id="a0efc-617">В Outlook в Интернете и на настольных клиентах `seriesId` возвращается идентификатор веб-служб Exchange (EWS) родительского элемента (ряда), к которому принадлежит этот элемент.</span><span class="sxs-lookup"><span data-stu-id="a0efc-617">In Outlook on the web and desktop clients, the `seriesId` returns the Exchange Web Services (EWS) ID of the parent (series) item that this item belongs to.</span></span> <span data-ttu-id="a0efc-618">Однако в iOS и Android `seriesId` возвращается идентификатор REST родительского элемента.</span><span class="sxs-lookup"><span data-stu-id="a0efc-618">However, in iOS and Android, the `seriesId` returns the REST ID of the parent item.</span></span>

> [!NOTE]
> <span data-ttu-id="a0efc-619">Идентификатор, возвращаемый свойством `seriesId`, совпадает с идентификатором элемента веб-служб Exchange.</span><span class="sxs-lookup"><span data-stu-id="a0efc-619">The identifier returned by the `seriesId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="a0efc-620">`seriesId` Свойство не совпадает с идентификаторами Outlook, используемыми в REST API Outlook.</span><span class="sxs-lookup"><span data-stu-id="a0efc-620">The `seriesId` property is not identical to the Outlook IDs used by the Outlook REST API.</span></span> <span data-ttu-id="a0efc-621">Прежде чем совершать вызовы REST API, используя это значение, его необходимо преобразовать с помощью [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="a0efc-621">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="a0efc-622">Дополнительные сведения см. в статье [Использование REST API Outlook из надстройки Outlook](/outlook/add-ins/use-rest-api).</span><span class="sxs-lookup"><span data-stu-id="a0efc-622">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api).</span></span>

<span data-ttu-id="a0efc-623">`seriesId` Свойство возвращает `null` элементы, у которых нет родительских элементов, таких как одиночные встречи, элементы ряда или приглашения на собрание, `undefined` и возвращаемые для других элементов, не являющиеся приглашениями на собрания.</span><span class="sxs-lookup"><span data-stu-id="a0efc-623">The `seriesId` property returns `null` for items that do not have parent items such as single appointments, series items, or meeting requests and returns `undefined` for any other items that are not meeting requests.</span></span>

##### <a name="type"></a><span data-ttu-id="a0efc-624">Тип</span><span class="sxs-lookup"><span data-stu-id="a0efc-624">Type</span></span>

* <span data-ttu-id="a0efc-625">String</span><span class="sxs-lookup"><span data-stu-id="a0efc-625">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="a0efc-626">Requirements</span><span class="sxs-lookup"><span data-stu-id="a0efc-626">Requirements</span></span>

|<span data-ttu-id="a0efc-627">Требование</span><span class="sxs-lookup"><span data-stu-id="a0efc-627">Requirement</span></span>|<span data-ttu-id="a0efc-628">Значение</span><span class="sxs-lookup"><span data-stu-id="a0efc-628">Value</span></span>|
|---|---|
|[<span data-ttu-id="a0efc-629">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="a0efc-629">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a0efc-630">1.7</span><span class="sxs-lookup"><span data-stu-id="a0efc-630">1.7</span></span>|
|[<span data-ttu-id="a0efc-631">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="a0efc-631">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a0efc-632">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a0efc-632">ReadItem</span></span>|
|[<span data-ttu-id="a0efc-633">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="a0efc-633">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="a0efc-634">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="a0efc-634">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a0efc-635">Пример</span><span class="sxs-lookup"><span data-stu-id="a0efc-635">Example</span></span>

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

#### <a name="start-datetimejavascriptapioutlookofficetimeviewoutlook-js-17"></a><span data-ttu-id="a0efc-636">start: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="a0efc-636">start: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.7)</span></span>

<span data-ttu-id="a0efc-637">Получает или задает дату и время начала встречи.</span><span class="sxs-lookup"><span data-stu-id="a0efc-637">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="a0efc-p138">Свойство `start` представлено в виде значения даты и времени в формате UTC. Это значение можно преобразовать в местные значения даты и времени клиента с помощью метода [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime).</span><span class="sxs-lookup"><span data-stu-id="a0efc-p138">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="a0efc-640">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="a0efc-640">Read mode</span></span>

<span data-ttu-id="a0efc-641">Свойство `start` возвращает объект `Date`.</span><span class="sxs-lookup"><span data-stu-id="a0efc-641">The `start` property returns a `Date` object.</span></span>

```js
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="a0efc-642">Режим создания</span><span class="sxs-lookup"><span data-stu-id="a0efc-642">Compose mode</span></span>

<span data-ttu-id="a0efc-643">Свойство `start` возвращает объект `Time`.</span><span class="sxs-lookup"><span data-stu-id="a0efc-643">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="a0efc-644">Если вы задаете время начала с помощью метода [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.7#setasync-datetime--options--callback-), необходимо использовать метод [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) для преобразования местного времени на клиенте в формат UTC для сервера.</span><span class="sxs-lookup"><span data-stu-id="a0efc-644">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.7#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="a0efc-645">В примере ниже с помощью метода [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.7#setasync-datetime--options--callback-) объекта `Time` задается время начала встречи в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="a0efc-645">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.7#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="a0efc-646">Тип</span><span class="sxs-lookup"><span data-stu-id="a0efc-646">Type</span></span>

*   <span data-ttu-id="a0efc-647">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="a0efc-647">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.7)</span></span>

##### <a name="requirements"></a><span data-ttu-id="a0efc-648">Requirements</span><span class="sxs-lookup"><span data-stu-id="a0efc-648">Requirements</span></span>

|<span data-ttu-id="a0efc-649">Требование</span><span class="sxs-lookup"><span data-stu-id="a0efc-649">Requirement</span></span>|<span data-ttu-id="a0efc-650">Значение</span><span class="sxs-lookup"><span data-stu-id="a0efc-650">Value</span></span>|
|---|---|
|[<span data-ttu-id="a0efc-651">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="a0efc-651">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a0efc-652">1.0</span><span class="sxs-lookup"><span data-stu-id="a0efc-652">1.0</span></span>|
|[<span data-ttu-id="a0efc-653">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="a0efc-653">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a0efc-654">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a0efc-654">ReadItem</span></span>|
|[<span data-ttu-id="a0efc-655">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="a0efc-655">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="a0efc-656">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="a0efc-656">Compose or Read</span></span>|

<br>

---
---

#### <a name="subject-stringsubjectjavascriptapioutlookofficesubjectviewoutlook-js-17"></a><span data-ttu-id="a0efc-657">subject: String|[Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="a0efc-657">subject: String|[Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.7)</span></span>

<span data-ttu-id="a0efc-658">Получает или задает описание, которое отображается в поле темы элемента.</span><span class="sxs-lookup"><span data-stu-id="a0efc-658">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="a0efc-659">Свойство `subject` получает или задает всю тему элемента для отправки с почтового сервера.</span><span class="sxs-lookup"><span data-stu-id="a0efc-659">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="a0efc-660">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="a0efc-660">Read mode</span></span>

<span data-ttu-id="a0efc-p139">Свойство `subject` возвращает строку. С помощью свойства [`normalizedSubject`](#normalizedsubject-string) можно получить тему без начальных префиксов, таких как `RE:` и `FW:`.</span><span class="sxs-lookup"><span data-stu-id="a0efc-p139">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

<span data-ttu-id="a0efc-663">В примере кода JavaScript, приведенном ниже, показано, как получить доступ к свойству `subject` текущего элемента в Outlook.</span><span class="sxs-lookup"><span data-stu-id="a0efc-663">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

```js
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="a0efc-664">Режим создания</span><span class="sxs-lookup"><span data-stu-id="a0efc-664">Compose mode</span></span>

<span data-ttu-id="a0efc-665">Свойство `subject` возвращает объект `Subject`, который предоставляет методы для получения и задания темы.</span><span class="sxs-lookup"><span data-stu-id="a0efc-665">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```js
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="a0efc-666">Тип</span><span class="sxs-lookup"><span data-stu-id="a0efc-666">Type</span></span>

*   <span data-ttu-id="a0efc-667">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="a0efc-667">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.7)</span></span>

##### <a name="requirements"></a><span data-ttu-id="a0efc-668">Требования</span><span class="sxs-lookup"><span data-stu-id="a0efc-668">Requirements</span></span>

|<span data-ttu-id="a0efc-669">Требование</span><span class="sxs-lookup"><span data-stu-id="a0efc-669">Requirement</span></span>|<span data-ttu-id="a0efc-670">Значение</span><span class="sxs-lookup"><span data-stu-id="a0efc-670">Value</span></span>|
|---|---|
|[<span data-ttu-id="a0efc-671">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="a0efc-671">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a0efc-672">1.0</span><span class="sxs-lookup"><span data-stu-id="a0efc-672">1.0</span></span>|
|[<span data-ttu-id="a0efc-673">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="a0efc-673">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a0efc-674">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a0efc-674">ReadItem</span></span>|
|[<span data-ttu-id="a0efc-675">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="a0efc-675">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="a0efc-676">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="a0efc-676">Compose or Read</span></span>|

<br>

---
---

#### <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-17recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-17"></a><span data-ttu-id="a0efc-677">to: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="a0efc-677">to: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span></span>

<span data-ttu-id="a0efc-678">Предоставляет доступ к получателям, указанным в строке **Кому** сообщения.</span><span class="sxs-lookup"><span data-stu-id="a0efc-678">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="a0efc-679">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="a0efc-679">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="a0efc-680">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="a0efc-680">Read mode</span></span>

<span data-ttu-id="a0efc-681">Свойство `to` возвращает массив, содержащий объект `EmailAddressDetails` для каждого получателя в строке **Кому** сообщения.</span><span class="sxs-lookup"><span data-stu-id="a0efc-681">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message.</span></span> <span data-ttu-id="a0efc-682">Коллекция может включать не более 100 элементов по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="a0efc-682">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="a0efc-683">Однако на компьютерах с Windows и Mac OS может быть до 500 элементов.</span><span class="sxs-lookup"><span data-stu-id="a0efc-683">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="a0efc-684">Режим создания</span><span class="sxs-lookup"><span data-stu-id="a0efc-684">Compose mode</span></span>

<span data-ttu-id="a0efc-685">Свойство `to` возвращает объект `Recipients`, предоставляющий методы для получения или обновления получателей, которые указаны в строке **Кому** сообщения.</span><span class="sxs-lookup"><span data-stu-id="a0efc-685">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span> <span data-ttu-id="a0efc-686">Коллекция может включать не более 100 элементов по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="a0efc-686">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="a0efc-687">Однако для компьютеров с Windows и Mac OS действуют указанные ниже ограничения.</span><span class="sxs-lookup"><span data-stu-id="a0efc-687">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="a0efc-688">Максимальное количество элементов — 500.</span><span class="sxs-lookup"><span data-stu-id="a0efc-688">Get 500 members maximum.</span></span>
- <span data-ttu-id="a0efc-689">Установить ограничение количества элементов на вызов — не более 100, общего количества — не более 500.</span><span class="sxs-lookup"><span data-stu-id="a0efc-689">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="a0efc-690">Тип</span><span class="sxs-lookup"><span data-stu-id="a0efc-690">Type</span></span>

*   <span data-ttu-id="a0efc-691">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="a0efc-691">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span></span>

##### <a name="requirements"></a><span data-ttu-id="a0efc-692">Requirements</span><span class="sxs-lookup"><span data-stu-id="a0efc-692">Requirements</span></span>

|<span data-ttu-id="a0efc-693">Требование</span><span class="sxs-lookup"><span data-stu-id="a0efc-693">Requirement</span></span>|<span data-ttu-id="a0efc-694">Значение</span><span class="sxs-lookup"><span data-stu-id="a0efc-694">Value</span></span>|
|---|---|
|[<span data-ttu-id="a0efc-695">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="a0efc-695">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a0efc-696">1.0</span><span class="sxs-lookup"><span data-stu-id="a0efc-696">1.0</span></span>|
|[<span data-ttu-id="a0efc-697">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="a0efc-697">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a0efc-698">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a0efc-698">ReadItem</span></span>|
|[<span data-ttu-id="a0efc-699">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="a0efc-699">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="a0efc-700">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="a0efc-700">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="a0efc-701">Методы</span><span class="sxs-lookup"><span data-stu-id="a0efc-701">Methods</span></span>

#### <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="a0efc-702">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="a0efc-702">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="a0efc-703">Добавляет файл в сообщение или встречу в качестве вложения.</span><span class="sxs-lookup"><span data-stu-id="a0efc-703">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="a0efc-704">Метод `addFileAttachmentAsync` передает файл по указанному универсальному коду ресурса (URI) и вкладывает его в элемент в форме создания.</span><span class="sxs-lookup"><span data-stu-id="a0efc-704">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="a0efc-705">Идентификатор можно использовать с методом [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback), чтобы удалить вложение, добавленное во время текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="a0efc-705">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a0efc-706">Параметры</span><span class="sxs-lookup"><span data-stu-id="a0efc-706">Parameters</span></span>
|<span data-ttu-id="a0efc-707">Имя</span><span class="sxs-lookup"><span data-stu-id="a0efc-707">Name</span></span>|<span data-ttu-id="a0efc-708">Тип</span><span class="sxs-lookup"><span data-stu-id="a0efc-708">Type</span></span>|<span data-ttu-id="a0efc-709">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="a0efc-709">Attributes</span></span>|<span data-ttu-id="a0efc-710">Описание</span><span class="sxs-lookup"><span data-stu-id="a0efc-710">Description</span></span>|
|---|---|---|---|
|`uri`|<span data-ttu-id="a0efc-711">String</span><span class="sxs-lookup"><span data-stu-id="a0efc-711">String</span></span>||<span data-ttu-id="a0efc-p143">Универсальный код ресурса (URI), представляющий расположение файла, который нужно вложить в сообщение или встречу. Максимальная длина — 2048 символов.</span><span class="sxs-lookup"><span data-stu-id="a0efc-p143">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="a0efc-714">String</span><span class="sxs-lookup"><span data-stu-id="a0efc-714">String</span></span>||<span data-ttu-id="a0efc-p144">Имя вложения, которое отображается при передаче вложения. Максимальная длина — 255 символов.</span><span class="sxs-lookup"><span data-stu-id="a0efc-p144">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="a0efc-717">Объект</span><span class="sxs-lookup"><span data-stu-id="a0efc-717">Object</span></span>|<span data-ttu-id="a0efc-718">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="a0efc-718">&lt;optional&gt;</span></span>|<span data-ttu-id="a0efc-719">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="a0efc-719">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="a0efc-720">Объект</span><span class="sxs-lookup"><span data-stu-id="a0efc-720">Object</span></span>|<span data-ttu-id="a0efc-721">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="a0efc-721">&lt;optional&gt;</span></span>|<span data-ttu-id="a0efc-722">В методе обратного вызова разработчики могут указать любой объект, к которому необходимо получить доступ.</span><span class="sxs-lookup"><span data-stu-id="a0efc-722">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="a0efc-723">Boolean</span><span class="sxs-lookup"><span data-stu-id="a0efc-723">Boolean</span></span>|<span data-ttu-id="a0efc-724">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="a0efc-724">&lt;optional&gt;</span></span>|<span data-ttu-id="a0efc-725">Значение `true` указывает на то, что вложение будет встроено в текст сообщения и не должно отображаться в списке вложений.</span><span class="sxs-lookup"><span data-stu-id="a0efc-725">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="a0efc-726">function</span><span class="sxs-lookup"><span data-stu-id="a0efc-726">function</span></span>|<span data-ttu-id="a0efc-727">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="a0efc-727">&lt;optional&gt;</span></span>|<span data-ttu-id="a0efc-728">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="a0efc-728">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="a0efc-729">После успешного выполнения идентификатор вложения будет представлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="a0efc-729">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="a0efc-730">Если добавить вложение не удастся, объект `asyncResult` будет содержать объект `Error` с описанием ошибки.</span><span class="sxs-lookup"><span data-stu-id="a0efc-730">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="a0efc-731">Ошибки</span><span class="sxs-lookup"><span data-stu-id="a0efc-731">Errors</span></span>

|<span data-ttu-id="a0efc-732">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="a0efc-732">Error code</span></span>|<span data-ttu-id="a0efc-733">Описание</span><span class="sxs-lookup"><span data-stu-id="a0efc-733">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="a0efc-734">Вложение превышает максимальный размер.</span><span class="sxs-lookup"><span data-stu-id="a0efc-734">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="a0efc-735">Расширение вложения не поддерживается.</span><span class="sxs-lookup"><span data-stu-id="a0efc-735">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="a0efc-736">Сообщение или встреча содержат слишком много вложений.</span><span class="sxs-lookup"><span data-stu-id="a0efc-736">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a0efc-737">Requirements</span><span class="sxs-lookup"><span data-stu-id="a0efc-737">Requirements</span></span>

|<span data-ttu-id="a0efc-738">Требование</span><span class="sxs-lookup"><span data-stu-id="a0efc-738">Requirement</span></span>|<span data-ttu-id="a0efc-739">Значение</span><span class="sxs-lookup"><span data-stu-id="a0efc-739">Value</span></span>|
|---|---|
|[<span data-ttu-id="a0efc-740">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="a0efc-740">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a0efc-741">1.1</span><span class="sxs-lookup"><span data-stu-id="a0efc-741">1.1</span></span>|
|[<span data-ttu-id="a0efc-742">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="a0efc-742">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a0efc-743">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="a0efc-743">ReadWriteItem</span></span>|
|[<span data-ttu-id="a0efc-744">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="a0efc-744">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="a0efc-745">Создание</span><span class="sxs-lookup"><span data-stu-id="a0efc-745">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="a0efc-746">Примеры</span><span class="sxs-lookup"><span data-stu-id="a0efc-746">Examples</span></span>

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

<span data-ttu-id="a0efc-747">В приведенном ниже примере файл изображения добавляется в качестве встроенного вложения, а ссылка на вложение добавляется в текст сообщения.</span><span class="sxs-lookup"><span data-stu-id="a0efc-747">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

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

#### <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="a0efc-748">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="a0efc-748">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="a0efc-749">Добавляет обработчик для поддерживаемого события.</span><span class="sxs-lookup"><span data-stu-id="a0efc-749">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="a0efc-750">В настоящее время поддерживаются типы `Office.EventType.AppointmentTimeChanged`событий `Office.EventType.RecipientsChanged`, и`Office.EventType.RecurrenceChanged`</span><span class="sxs-lookup"><span data-stu-id="a0efc-750">Currently the supported event types are `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`</span></span>

##### <a name="parameters"></a><span data-ttu-id="a0efc-751">Параметры</span><span class="sxs-lookup"><span data-stu-id="a0efc-751">Parameters</span></span>

| <span data-ttu-id="a0efc-752">Имя</span><span class="sxs-lookup"><span data-stu-id="a0efc-752">Name</span></span> | <span data-ttu-id="a0efc-753">Тип</span><span class="sxs-lookup"><span data-stu-id="a0efc-753">Type</span></span> | <span data-ttu-id="a0efc-754">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="a0efc-754">Attributes</span></span> | <span data-ttu-id="a0efc-755">Описание</span><span class="sxs-lookup"><span data-stu-id="a0efc-755">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="a0efc-756">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="a0efc-756">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="a0efc-757">Событие, которое должно вызвать обработчик.</span><span class="sxs-lookup"><span data-stu-id="a0efc-757">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="a0efc-758">Function</span><span class="sxs-lookup"><span data-stu-id="a0efc-758">Function</span></span> || <span data-ttu-id="a0efc-p145">Функция для обработки события. Функция должна принимать один параметр, представляющий собой объектный литерал. Значение свойства `type` параметра совпадет со значением параметра `eventType`, переданного методу `addHandlerAsync`.</span><span class="sxs-lookup"><span data-stu-id="a0efc-p145">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="a0efc-762">Объект</span><span class="sxs-lookup"><span data-stu-id="a0efc-762">Object</span></span> | <span data-ttu-id="a0efc-763">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="a0efc-763">&lt;optional&gt;</span></span> | <span data-ttu-id="a0efc-764">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="a0efc-764">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="a0efc-765">Объект</span><span class="sxs-lookup"><span data-stu-id="a0efc-765">Object</span></span> | <span data-ttu-id="a0efc-766">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="a0efc-766">&lt;optional&gt;</span></span> | <span data-ttu-id="a0efc-767">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="a0efc-767">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="a0efc-768">функция</span><span class="sxs-lookup"><span data-stu-id="a0efc-768">function</span></span>| <span data-ttu-id="a0efc-769">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="a0efc-769">&lt;optional&gt;</span></span>|<span data-ttu-id="a0efc-770">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="a0efc-770">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a0efc-771">Требования</span><span class="sxs-lookup"><span data-stu-id="a0efc-771">Requirements</span></span>

|<span data-ttu-id="a0efc-772">Требование</span><span class="sxs-lookup"><span data-stu-id="a0efc-772">Requirement</span></span>| <span data-ttu-id="a0efc-773">Значение</span><span class="sxs-lookup"><span data-stu-id="a0efc-773">Value</span></span>|
|---|---|
|[<span data-ttu-id="a0efc-774">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="a0efc-774">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a0efc-775">1.7</span><span class="sxs-lookup"><span data-stu-id="a0efc-775">1.7</span></span> |
|[<span data-ttu-id="a0efc-776">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="a0efc-776">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a0efc-777">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a0efc-777">ReadItem</span></span> |
|[<span data-ttu-id="a0efc-778">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="a0efc-778">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a0efc-779">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="a0efc-779">Compose or Read</span></span> |

##### <a name="example"></a><span data-ttu-id="a0efc-780">Пример</span><span class="sxs-lookup"><span data-stu-id="a0efc-780">Example</span></span>

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

#### <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="a0efc-781">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="a0efc-781">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="a0efc-782">Добавляет к сообщению элемент Exchange, например сообщение, в виде вложения.</span><span class="sxs-lookup"><span data-stu-id="a0efc-782">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="a0efc-p146">С помощью метода `addItemAttachmentAsync` можно в элемент формы создания вложить элемент с указанным идентификатором Exchange. Если указать метод обратного вызова, то этот метод вызывается с помощью параметра `asyncResult`, который содержит идентификатор вложения или код, указывающий на ошибки, которые произошли при вложении элемента. При необходимости можно использовать параметр `options` для передачи сведений о состоянии методу обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="a0efc-p146">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="a0efc-786">Идентификатор можно использовать с методом [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback), чтобы удалить вложение, добавленное во время текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="a0efc-786">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="a0efc-787">Если ваша надстройка Office выполняется в Outlook в Интернете, метод `addItemAttachmentAsync` обеспечивает вложение элементов в элементы, отличные от редактируемого. Однако это действие не рекомендуется выполнять, так как оно не поддерживается.</span><span class="sxs-lookup"><span data-stu-id="a0efc-787">If your Office Add-in is running in Outlook on the web, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a0efc-788">Параметры</span><span class="sxs-lookup"><span data-stu-id="a0efc-788">Parameters</span></span>

|<span data-ttu-id="a0efc-789">Имя</span><span class="sxs-lookup"><span data-stu-id="a0efc-789">Name</span></span>|<span data-ttu-id="a0efc-790">Тип</span><span class="sxs-lookup"><span data-stu-id="a0efc-790">Type</span></span>|<span data-ttu-id="a0efc-791">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="a0efc-791">Attributes</span></span>|<span data-ttu-id="a0efc-792">Описание</span><span class="sxs-lookup"><span data-stu-id="a0efc-792">Description</span></span>|
|---|---|---|---|
|`itemId`|<span data-ttu-id="a0efc-793">String</span><span class="sxs-lookup"><span data-stu-id="a0efc-793">String</span></span>||<span data-ttu-id="a0efc-p147">Идентификатор Exchange для вкладываемого элемента. Максимальная длина — 100 символов.</span><span class="sxs-lookup"><span data-stu-id="a0efc-p147">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="a0efc-796">String</span><span class="sxs-lookup"><span data-stu-id="a0efc-796">String</span></span>||<span data-ttu-id="a0efc-797">Тема вкладываемого элемента.</span><span class="sxs-lookup"><span data-stu-id="a0efc-797">The subject of the item to be attached.</span></span> <span data-ttu-id="a0efc-798">Максимальная длина: 255 символов.</span><span class="sxs-lookup"><span data-stu-id="a0efc-798">The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="a0efc-799">Object</span><span class="sxs-lookup"><span data-stu-id="a0efc-799">Object</span></span>|<span data-ttu-id="a0efc-800">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="a0efc-800">&lt;optional&gt;</span></span>|<span data-ttu-id="a0efc-801">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="a0efc-801">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="a0efc-802">Объект</span><span class="sxs-lookup"><span data-stu-id="a0efc-802">Object</span></span>|<span data-ttu-id="a0efc-803">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="a0efc-803">&lt;optional&gt;</span></span>|<span data-ttu-id="a0efc-804">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="a0efc-804">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="a0efc-805">функция</span><span class="sxs-lookup"><span data-stu-id="a0efc-805">function</span></span>|<span data-ttu-id="a0efc-806">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="a0efc-806">&lt;optional&gt;</span></span>|<span data-ttu-id="a0efc-807">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="a0efc-807">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="a0efc-808">После успешного выполнения идентификатор вложения будет представлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="a0efc-808">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="a0efc-809">Если добавить вложение не удастся, объект `asyncResult` будет содержать объект `Error` с описанием ошибки.</span><span class="sxs-lookup"><span data-stu-id="a0efc-809">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="a0efc-810">Ошибки</span><span class="sxs-lookup"><span data-stu-id="a0efc-810">Errors</span></span>

|<span data-ttu-id="a0efc-811">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="a0efc-811">Error code</span></span>|<span data-ttu-id="a0efc-812">Описание</span><span class="sxs-lookup"><span data-stu-id="a0efc-812">Description</span></span>|
|------------|-------------|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="a0efc-813">Сообщение или встреча содержат слишком много вложений.</span><span class="sxs-lookup"><span data-stu-id="a0efc-813">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a0efc-814">Requirements</span><span class="sxs-lookup"><span data-stu-id="a0efc-814">Requirements</span></span>

|<span data-ttu-id="a0efc-815">Требование</span><span class="sxs-lookup"><span data-stu-id="a0efc-815">Requirement</span></span>|<span data-ttu-id="a0efc-816">Значение</span><span class="sxs-lookup"><span data-stu-id="a0efc-816">Value</span></span>|
|---|---|
|[<span data-ttu-id="a0efc-817">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="a0efc-817">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a0efc-818">1.1</span><span class="sxs-lookup"><span data-stu-id="a0efc-818">1.1</span></span>|
|[<span data-ttu-id="a0efc-819">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="a0efc-819">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a0efc-820">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="a0efc-820">ReadWriteItem</span></span>|
|[<span data-ttu-id="a0efc-821">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="a0efc-821">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="a0efc-822">Создание</span><span class="sxs-lookup"><span data-stu-id="a0efc-822">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="a0efc-823">Пример</span><span class="sxs-lookup"><span data-stu-id="a0efc-823">Example</span></span>

<span data-ttu-id="a0efc-824">В следующем примере существующий элемент Outlook добавляется в виде вложения с именем `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="a0efc-824">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

#### <a name="close"></a><span data-ttu-id="a0efc-825">close()</span><span class="sxs-lookup"><span data-stu-id="a0efc-825">close()</span></span>

<span data-ttu-id="a0efc-826">Закрывает текущий создаваемый элемент.</span><span class="sxs-lookup"><span data-stu-id="a0efc-826">Closes the current item that is being composed.</span></span>

<span data-ttu-id="a0efc-p149">Работа метода `close` зависит от текущего состояния создаваемого элемента. Если элемент содержит несохраненные изменения, клиент предложит пользователю сохранить или отклонить их либо отменить действие закрытия.</span><span class="sxs-lookup"><span data-stu-id="a0efc-p149">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="a0efc-829">Если в Outlook в Интернете элемент представляет собой встречу, ранее сохраненную с помощью метода `saveAsync`, пользователю предлагается сохранить, отклонить или отменить действие, даже если с момента последнего сохранения элемента не вносились какие-либо изменения.</span><span class="sxs-lookup"><span data-stu-id="a0efc-829">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="a0efc-830">Если в клиенте Outlook для настольных ПК сообщение представляет собой ответ в тексте, метод `close` не работает.</span><span class="sxs-lookup"><span data-stu-id="a0efc-830">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="a0efc-831">Requirements</span><span class="sxs-lookup"><span data-stu-id="a0efc-831">Requirements</span></span>

|<span data-ttu-id="a0efc-832">Требование</span><span class="sxs-lookup"><span data-stu-id="a0efc-832">Requirement</span></span>|<span data-ttu-id="a0efc-833">Значение</span><span class="sxs-lookup"><span data-stu-id="a0efc-833">Value</span></span>|
|---|---|
|[<span data-ttu-id="a0efc-834">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="a0efc-834">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a0efc-835">1.3</span><span class="sxs-lookup"><span data-stu-id="a0efc-835">1.3</span></span>|
|[<span data-ttu-id="a0efc-836">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="a0efc-836">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a0efc-837">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="a0efc-837">Restricted</span></span>|
|[<span data-ttu-id="a0efc-838">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="a0efc-838">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="a0efc-839">Создание</span><span class="sxs-lookup"><span data-stu-id="a0efc-839">Compose</span></span>|

<br>

---
---

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="a0efc-840">displayReplyAllForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="a0efc-840">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="a0efc-841">Отображает форму ответа, включающую отправителя и всех получателей выбранного сообщения или организатора и всех участников выбранной встречи.</span><span class="sxs-lookup"><span data-stu-id="a0efc-841">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="a0efc-842">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="a0efc-842">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="a0efc-843">В Outlook в Интернете форма ответа отображается в виде всплывающей формы в представлении с 3 либо 1 или 2 столбцами.</span><span class="sxs-lookup"><span data-stu-id="a0efc-843">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="a0efc-844">Если любой строковый параметр превышает указанные для него ограничения, `displayReplyAllForm` возвращает исключение.</span><span class="sxs-lookup"><span data-stu-id="a0efc-844">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="a0efc-p150">Если в параметре `formData.attachments` указаны вложения, Outlook в Интернете и классические клиенты пытаются скачать их и вложить в форму ответа. Если какие-либо вложения добавить не удается, в форме отображается сообщение об ошибке. Если сообщения об ошибках не предусмотрены, то они не отображаются.</span><span class="sxs-lookup"><span data-stu-id="a0efc-p150">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a0efc-848">Параметры</span><span class="sxs-lookup"><span data-stu-id="a0efc-848">Parameters</span></span>

|<span data-ttu-id="a0efc-849">Имя</span><span class="sxs-lookup"><span data-stu-id="a0efc-849">Name</span></span>|<span data-ttu-id="a0efc-850">Тип</span><span class="sxs-lookup"><span data-stu-id="a0efc-850">Type</span></span>|<span data-ttu-id="a0efc-851">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="a0efc-851">Attributes</span></span>|<span data-ttu-id="a0efc-852">Описание</span><span class="sxs-lookup"><span data-stu-id="a0efc-852">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="a0efc-853">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="a0efc-853">String &#124; Object</span></span>||<span data-ttu-id="a0efc-p151">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="a0efc-p151">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="a0efc-856">**ИЛИ**</span><span class="sxs-lookup"><span data-stu-id="a0efc-856">**OR**</span></span><br/><span data-ttu-id="a0efc-p152">Объект, который содержит текст или данные вложения и функцию обратного вызова. Ниже представлено определение этого объекта.</span><span class="sxs-lookup"><span data-stu-id="a0efc-p152">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="a0efc-859">String</span><span class="sxs-lookup"><span data-stu-id="a0efc-859">String</span></span>|<span data-ttu-id="a0efc-860">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="a0efc-860">&lt;optional&gt;</span></span>|<span data-ttu-id="a0efc-p153">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="a0efc-p153">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="a0efc-863">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="a0efc-863">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="a0efc-864">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="a0efc-864">&lt;optional&gt;</span></span>|<span data-ttu-id="a0efc-865">Массив объектов JSON, представляющих собой вложенные файлы или элементы.</span><span class="sxs-lookup"><span data-stu-id="a0efc-865">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="a0efc-866">String</span><span class="sxs-lookup"><span data-stu-id="a0efc-866">String</span></span>||<span data-ttu-id="a0efc-p154">Указывает тип вложения. Допустимые значения: `file` для вложенного файла и `item` для вложенного элемента.</span><span class="sxs-lookup"><span data-stu-id="a0efc-p154">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="a0efc-869">Строка</span><span class="sxs-lookup"><span data-stu-id="a0efc-869">String</span></span>||<span data-ttu-id="a0efc-870">Строка, содержащая имя вложения, длиной до 255 символов.</span><span class="sxs-lookup"><span data-stu-id="a0efc-870">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="a0efc-871">String</span><span class="sxs-lookup"><span data-stu-id="a0efc-871">String</span></span>||<span data-ttu-id="a0efc-p155">Используется, только если свойству `type` задано значение `file`. URI расположения файла.</span><span class="sxs-lookup"><span data-stu-id="a0efc-p155">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="a0efc-874">Логический</span><span class="sxs-lookup"><span data-stu-id="a0efc-874">Boolean</span></span>||<span data-ttu-id="a0efc-p156">Используется, только если свойству `type` задано значение `file`. Значение `true` указывает на то, что вложение будет встроено в текст сообщения и не должно отображаться в списке вложений.</span><span class="sxs-lookup"><span data-stu-id="a0efc-p156">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="a0efc-877">String</span><span class="sxs-lookup"><span data-stu-id="a0efc-877">String</span></span>||<span data-ttu-id="a0efc-p157">Используется, только если свойству `type` задано значение `item`. Идентификатор вложения EWS. Это строка длиной до 100 символов.</span><span class="sxs-lookup"><span data-stu-id="a0efc-p157">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="a0efc-881">function</span><span class="sxs-lookup"><span data-stu-id="a0efc-881">function</span></span>|<span data-ttu-id="a0efc-882">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="a0efc-882">&lt;optional&gt;</span></span>|<span data-ttu-id="a0efc-883">По завершении работы метода функция, переданная параметру `callback`, вызывается с помощью одного параметра `asyncResult`, представляющего собой объект [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="a0efc-883">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a0efc-884">Требования</span><span class="sxs-lookup"><span data-stu-id="a0efc-884">Requirements</span></span>

|<span data-ttu-id="a0efc-885">Требование</span><span class="sxs-lookup"><span data-stu-id="a0efc-885">Requirement</span></span>|<span data-ttu-id="a0efc-886">Значение</span><span class="sxs-lookup"><span data-stu-id="a0efc-886">Value</span></span>|
|---|---|
|[<span data-ttu-id="a0efc-887">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="a0efc-887">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a0efc-888">1.0</span><span class="sxs-lookup"><span data-stu-id="a0efc-888">1.0</span></span>|
|[<span data-ttu-id="a0efc-889">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="a0efc-889">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a0efc-890">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a0efc-890">ReadItem</span></span>|
|[<span data-ttu-id="a0efc-891">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="a0efc-891">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="a0efc-892">Чтение</span><span class="sxs-lookup"><span data-stu-id="a0efc-892">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="a0efc-893">Примеры</span><span class="sxs-lookup"><span data-stu-id="a0efc-893">Examples</span></span>

<span data-ttu-id="a0efc-894">Приведенный ниже код передает строку в функцию `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="a0efc-894">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="a0efc-895">Ответ с пустым текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="a0efc-895">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="a0efc-896">Ответ только с текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="a0efc-896">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="a0efc-897">Ответ с текстом сообщения и вложенным файлом.</span><span class="sxs-lookup"><span data-stu-id="a0efc-897">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="a0efc-898">Ответ с текстом сообщения и вложенным элементом.</span><span class="sxs-lookup"><span data-stu-id="a0efc-898">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="a0efc-899">Ответ с текстом сообщения, вложенным файлом, вложенным элементом и обратным вызовом.</span><span class="sxs-lookup"><span data-stu-id="a0efc-899">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="a0efc-900">displayReplyForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="a0efc-900">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="a0efc-901">Отображает форму ответа, включающую только отправителя выбранного сообщения или организатора выбранной встречи.</span><span class="sxs-lookup"><span data-stu-id="a0efc-901">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="a0efc-902">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="a0efc-902">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="a0efc-903">В Outlook в Интернете форма ответа отображается в виде всплывающей формы в представлении с 3 либо 1 или 2 столбцами.</span><span class="sxs-lookup"><span data-stu-id="a0efc-903">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="a0efc-904">Если любой строковый параметр превышает указанные для него ограничения, `displayReplyForm` возвращает исключение.</span><span class="sxs-lookup"><span data-stu-id="a0efc-904">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="a0efc-p158">Если в параметре `formData.attachments` указаны вложения, Outlook в Интернете и классические клиенты пытаются скачать их и вложить в форму ответа. Если какие-либо вложения добавить не удается, в форме отображается сообщение об ошибке. Если сообщения об ошибках не предусмотрены, то они не отображаются.</span><span class="sxs-lookup"><span data-stu-id="a0efc-p158">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a0efc-908">Параметры</span><span class="sxs-lookup"><span data-stu-id="a0efc-908">Parameters</span></span>

|<span data-ttu-id="a0efc-909">Имя</span><span class="sxs-lookup"><span data-stu-id="a0efc-909">Name</span></span>|<span data-ttu-id="a0efc-910">Тип</span><span class="sxs-lookup"><span data-stu-id="a0efc-910">Type</span></span>|<span data-ttu-id="a0efc-911">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="a0efc-911">Attributes</span></span>|<span data-ttu-id="a0efc-912">Описание</span><span class="sxs-lookup"><span data-stu-id="a0efc-912">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="a0efc-913">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="a0efc-913">String &#124; Object</span></span>||<span data-ttu-id="a0efc-p159">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="a0efc-p159">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="a0efc-916">**ИЛИ**</span><span class="sxs-lookup"><span data-stu-id="a0efc-916">**OR**</span></span><br/><span data-ttu-id="a0efc-p160">Объект, который содержит текст или данные вложения и функцию обратного вызова. Ниже представлено определение этого объекта.</span><span class="sxs-lookup"><span data-stu-id="a0efc-p160">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="a0efc-919">String</span><span class="sxs-lookup"><span data-stu-id="a0efc-919">String</span></span>|<span data-ttu-id="a0efc-920">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="a0efc-920">&lt;optional&gt;</span></span>|<span data-ttu-id="a0efc-p161">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="a0efc-p161">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="a0efc-923">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="a0efc-923">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="a0efc-924">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="a0efc-924">&lt;optional&gt;</span></span>|<span data-ttu-id="a0efc-925">Массив объектов JSON, представляющих собой вложенные файлы или элементы.</span><span class="sxs-lookup"><span data-stu-id="a0efc-925">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="a0efc-926">String</span><span class="sxs-lookup"><span data-stu-id="a0efc-926">String</span></span>||<span data-ttu-id="a0efc-p162">Указывает тип вложения. Допустимые значения: `file` для вложенного файла и `item` для вложенного элемента.</span><span class="sxs-lookup"><span data-stu-id="a0efc-p162">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="a0efc-929">Строка</span><span class="sxs-lookup"><span data-stu-id="a0efc-929">String</span></span>||<span data-ttu-id="a0efc-930">Строка, содержащая имя вложения, длиной до 255 символов.</span><span class="sxs-lookup"><span data-stu-id="a0efc-930">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="a0efc-931">Строка</span><span class="sxs-lookup"><span data-stu-id="a0efc-931">String</span></span>||<span data-ttu-id="a0efc-p163">Используется, только если свойству `type` задано значение `file`. URI расположения файла.</span><span class="sxs-lookup"><span data-stu-id="a0efc-p163">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="a0efc-934">Логический</span><span class="sxs-lookup"><span data-stu-id="a0efc-934">Boolean</span></span>||<span data-ttu-id="a0efc-p164">Используется, только если свойству `type` задано значение `file`. Значение `true` указывает на то, что вложение будет встроено в текст сообщения и не должно отображаться в списке вложений.</span><span class="sxs-lookup"><span data-stu-id="a0efc-p164">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="a0efc-937">String</span><span class="sxs-lookup"><span data-stu-id="a0efc-937">String</span></span>||<span data-ttu-id="a0efc-p165">Используется, только если свойству `type` задано значение `item`. Идентификатор вложения EWS. Это строка длиной до 100 символов.</span><span class="sxs-lookup"><span data-stu-id="a0efc-p165">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="a0efc-941">function</span><span class="sxs-lookup"><span data-stu-id="a0efc-941">function</span></span>|<span data-ttu-id="a0efc-942">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="a0efc-942">&lt;optional&gt;</span></span>|<span data-ttu-id="a0efc-943">По завершении работы метода функция, переданная параметру `callback`, вызывается с помощью одного параметра `asyncResult`, представляющего собой объект [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="a0efc-943">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a0efc-944">Требования</span><span class="sxs-lookup"><span data-stu-id="a0efc-944">Requirements</span></span>

|<span data-ttu-id="a0efc-945">Требование</span><span class="sxs-lookup"><span data-stu-id="a0efc-945">Requirement</span></span>|<span data-ttu-id="a0efc-946">Значение</span><span class="sxs-lookup"><span data-stu-id="a0efc-946">Value</span></span>|
|---|---|
|[<span data-ttu-id="a0efc-947">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="a0efc-947">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a0efc-948">1.0</span><span class="sxs-lookup"><span data-stu-id="a0efc-948">1.0</span></span>|
|[<span data-ttu-id="a0efc-949">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="a0efc-949">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a0efc-950">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a0efc-950">ReadItem</span></span>|
|[<span data-ttu-id="a0efc-951">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="a0efc-951">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="a0efc-952">Чтение</span><span class="sxs-lookup"><span data-stu-id="a0efc-952">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="a0efc-953">Примеры</span><span class="sxs-lookup"><span data-stu-id="a0efc-953">Examples</span></span>

<span data-ttu-id="a0efc-954">Приведенный ниже код передает строку в функцию `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="a0efc-954">The following code passes a string to the `displayReplyForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="a0efc-955">Ответ с пустым текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="a0efc-955">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="a0efc-956">Ответ только с текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="a0efc-956">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="a0efc-957">Ответ с текстом сообщения и вложенным файлом.</span><span class="sxs-lookup"><span data-stu-id="a0efc-957">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="a0efc-958">Ответ с текстом сообщения и вложенным элементом.</span><span class="sxs-lookup"><span data-stu-id="a0efc-958">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="a0efc-959">Ответ с текстом сообщения, вложенным файлом, вложенным элементом и обратным вызовом.</span><span class="sxs-lookup"><span data-stu-id="a0efc-959">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlookofficeentitiesviewoutlook-js-17"></a><span data-ttu-id="a0efc-960">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.7)}</span><span class="sxs-lookup"><span data-stu-id="a0efc-960">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.7)}</span></span>

<span data-ttu-id="a0efc-961">Получает сущности, обнаруженные в теле выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="a0efc-961">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="a0efc-962">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="a0efc-962">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="a0efc-963">Требования</span><span class="sxs-lookup"><span data-stu-id="a0efc-963">Requirements</span></span>

|<span data-ttu-id="a0efc-964">Требование</span><span class="sxs-lookup"><span data-stu-id="a0efc-964">Requirement</span></span>|<span data-ttu-id="a0efc-965">Значение</span><span class="sxs-lookup"><span data-stu-id="a0efc-965">Value</span></span>|
|---|---|
|[<span data-ttu-id="a0efc-966">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="a0efc-966">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a0efc-967">1.0</span><span class="sxs-lookup"><span data-stu-id="a0efc-967">1.0</span></span>|
|[<span data-ttu-id="a0efc-968">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="a0efc-968">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a0efc-969">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a0efc-969">ReadItem</span></span>|
|[<span data-ttu-id="a0efc-970">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="a0efc-970">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="a0efc-971">Чтение</span><span class="sxs-lookup"><span data-stu-id="a0efc-971">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="a0efc-972">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="a0efc-972">Returns:</span></span>

<span data-ttu-id="a0efc-973">Тип: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="a0efc-973">Type: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.7)</span></span>

##### <a name="example"></a><span data-ttu-id="a0efc-974">Пример</span><span class="sxs-lookup"><span data-stu-id="a0efc-974">Example</span></span>

<span data-ttu-id="a0efc-975">Ниже приведен пример получения доступа к сущностям контактов в теле текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="a0efc-975">The following example accesses the contacts entities in the current item's body.</span></span>

```js
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

<br>

---
---

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-17meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-17phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-17tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-17"></a><span data-ttu-id="a0efc-976">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.7)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.7)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.7)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.7))>}</span><span class="sxs-lookup"><span data-stu-id="a0efc-976">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.7)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.7)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.7)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.7))>}</span></span>

<span data-ttu-id="a0efc-977">Получает массив всех сущностей указанного типа, обнаруженных в теле выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="a0efc-977">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="a0efc-978">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="a0efc-978">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a0efc-979">Параметры</span><span class="sxs-lookup"><span data-stu-id="a0efc-979">Parameters</span></span>

|<span data-ttu-id="a0efc-980">Имя</span><span class="sxs-lookup"><span data-stu-id="a0efc-980">Name</span></span>|<span data-ttu-id="a0efc-981">Тип</span><span class="sxs-lookup"><span data-stu-id="a0efc-981">Type</span></span>|<span data-ttu-id="a0efc-982">Описание</span><span class="sxs-lookup"><span data-stu-id="a0efc-982">Description</span></span>|
|---|---|---|
|`entityType`|[<span data-ttu-id="a0efc-983">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="a0efc-983">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook/office.mailboxenums.entitytype?view=outlook-js-1.7)|<span data-ttu-id="a0efc-984">Одно из значений перечисления EntityType.</span><span class="sxs-lookup"><span data-stu-id="a0efc-984">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a0efc-985">Requirements</span><span class="sxs-lookup"><span data-stu-id="a0efc-985">Requirements</span></span>

|<span data-ttu-id="a0efc-986">Требование</span><span class="sxs-lookup"><span data-stu-id="a0efc-986">Requirement</span></span>|<span data-ttu-id="a0efc-987">Значение</span><span class="sxs-lookup"><span data-stu-id="a0efc-987">Value</span></span>|
|---|---|
|[<span data-ttu-id="a0efc-988">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="a0efc-988">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a0efc-989">1.0</span><span class="sxs-lookup"><span data-stu-id="a0efc-989">1.0</span></span>|
|[<span data-ttu-id="a0efc-990">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="a0efc-990">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a0efc-991">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="a0efc-991">Restricted</span></span>|
|[<span data-ttu-id="a0efc-992">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="a0efc-992">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="a0efc-993">Чтение</span><span class="sxs-lookup"><span data-stu-id="a0efc-993">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="a0efc-994">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="a0efc-994">Returns:</span></span>

<span data-ttu-id="a0efc-995">Если значение, переданное в `entityType`, не является допустимым членом перечисления `EntityType`, метод возвращает значение NULL.</span><span class="sxs-lookup"><span data-stu-id="a0efc-995">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="a0efc-996">Если в теле элемента отсутствуют сущности указанного типа, метод возвращает пустой массив.</span><span class="sxs-lookup"><span data-stu-id="a0efc-996">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="a0efc-997">В противном случае тип объектов в возвращаемом массиве зависит от типа сущности, запрошенной в параметре `entityType`.</span><span class="sxs-lookup"><span data-stu-id="a0efc-997">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="a0efc-998">Хотя минимальный уровень разрешений для использования этого метода — **Restricted**, для некоторых типов сущностей требуется доступ на уровне **ReadItem**, как указано в приведенной ниже таблице.</span><span class="sxs-lookup"><span data-stu-id="a0efc-998">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

|<span data-ttu-id="a0efc-999">Значение параметра `entityType`</span><span class="sxs-lookup"><span data-stu-id="a0efc-999">Value of `entityType`</span></span>|<span data-ttu-id="a0efc-1000">Тип объектов в возвращаемом массиве</span><span class="sxs-lookup"><span data-stu-id="a0efc-1000">Type of objects in returned array</span></span>|<span data-ttu-id="a0efc-1001">Необходимый уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="a0efc-1001">Required Permission Level</span></span>|
|---|---|---|
|`Address`|<span data-ttu-id="a0efc-1002">String</span><span class="sxs-lookup"><span data-stu-id="a0efc-1002">String</span></span>|<span data-ttu-id="a0efc-1003">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="a0efc-1003">**Restricted**</span></span>|
|`Contact`|<span data-ttu-id="a0efc-1004">Contact</span><span class="sxs-lookup"><span data-stu-id="a0efc-1004">Contact</span></span>|<span data-ttu-id="a0efc-1005">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="a0efc-1005">**ReadItem**</span></span>|
|`EmailAddress`|<span data-ttu-id="a0efc-1006">String</span><span class="sxs-lookup"><span data-stu-id="a0efc-1006">String</span></span>|<span data-ttu-id="a0efc-1007">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="a0efc-1007">**ReadItem**</span></span>|
|`MeetingSuggestion`|<span data-ttu-id="a0efc-1008">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="a0efc-1008">MeetingSuggestion</span></span>|<span data-ttu-id="a0efc-1009">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="a0efc-1009">**ReadItem**</span></span>|
|`PhoneNumber`|<span data-ttu-id="a0efc-1010">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="a0efc-1010">PhoneNumber</span></span>|<span data-ttu-id="a0efc-1011">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="a0efc-1011">**Restricted**</span></span>|
|`TaskSuggestion`|<span data-ttu-id="a0efc-1012">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="a0efc-1012">TaskSuggestion</span></span>|<span data-ttu-id="a0efc-1013">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="a0efc-1013">**ReadItem**</span></span>|
|`URL`|<span data-ttu-id="a0efc-1014">String</span><span class="sxs-lookup"><span data-stu-id="a0efc-1014">String</span></span>|<span data-ttu-id="a0efc-1015">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="a0efc-1015">**Restricted**</span></span>|

<span data-ttu-id="a0efc-1016">Тип: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.7)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.7)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.7)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.7))></span><span class="sxs-lookup"><span data-stu-id="a0efc-1016">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.7)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.7)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.7)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.7))></span></span>

##### <a name="example"></a><span data-ttu-id="a0efc-1017">Пример</span><span class="sxs-lookup"><span data-stu-id="a0efc-1017">Example</span></span>

<span data-ttu-id="a0efc-1018">В примере ниже показано, как получить доступ к массиву строк, которые представляют собой почтовые адреса в теле текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="a0efc-1018">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

<br>

---
---

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-17meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-17phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-17tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-17"></a><span data-ttu-id="a0efc-1019">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.7)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.7)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.7)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.7))>}</span><span class="sxs-lookup"><span data-stu-id="a0efc-1019">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.7)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.7)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.7)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.7))>}</span></span>

<span data-ttu-id="a0efc-1020">Возвращает известные сущности в выбранном элементе, которые проходят через именованный фильтр, определяемый в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="a0efc-1020">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="a0efc-1021">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="a0efc-1021">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="a0efc-1022">Метод `getFilteredEntitiesByName` возвращает сущности, соответствующие регулярному выражению, которое определяется в элементе правила [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) в XML-файле манифеста, с использованием указанного значения элемента `FilterName`.</span><span class="sxs-lookup"><span data-stu-id="a0efc-1022">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a0efc-1023">Параметры</span><span class="sxs-lookup"><span data-stu-id="a0efc-1023">Parameters</span></span>

|<span data-ttu-id="a0efc-1024">Имя</span><span class="sxs-lookup"><span data-stu-id="a0efc-1024">Name</span></span>|<span data-ttu-id="a0efc-1025">Тип</span><span class="sxs-lookup"><span data-stu-id="a0efc-1025">Type</span></span>|<span data-ttu-id="a0efc-1026">Описание</span><span class="sxs-lookup"><span data-stu-id="a0efc-1026">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="a0efc-1027">String</span><span class="sxs-lookup"><span data-stu-id="a0efc-1027">String</span></span>|<span data-ttu-id="a0efc-1028">Имя элемента правила `ItemHasKnownEntity`, определяющее соответствующий фильтр.</span><span class="sxs-lookup"><span data-stu-id="a0efc-1028">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a0efc-1029">Requirements</span><span class="sxs-lookup"><span data-stu-id="a0efc-1029">Requirements</span></span>

|<span data-ttu-id="a0efc-1030">Требование</span><span class="sxs-lookup"><span data-stu-id="a0efc-1030">Requirement</span></span>|<span data-ttu-id="a0efc-1031">Значение</span><span class="sxs-lookup"><span data-stu-id="a0efc-1031">Value</span></span>|
|---|---|
|[<span data-ttu-id="a0efc-1032">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="a0efc-1032">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a0efc-1033">1.0</span><span class="sxs-lookup"><span data-stu-id="a0efc-1033">1.0</span></span>|
|[<span data-ttu-id="a0efc-1034">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="a0efc-1034">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a0efc-1035">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a0efc-1035">ReadItem</span></span>|
|[<span data-ttu-id="a0efc-1036">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="a0efc-1036">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="a0efc-1037">Чтение</span><span class="sxs-lookup"><span data-stu-id="a0efc-1037">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="a0efc-1038">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="a0efc-1038">Returns:</span></span>

<span data-ttu-id="a0efc-p167">Если в манифесте нет элемента `ItemHasKnownEntity` со значением `FilterName`, соответствующим параметру `name`, метод возвращает `null`. Если параметр `name` соответствует элементу `ItemHasKnownEntity` в манифесте, но при этом в текущем элементе нет соответствующих сущностей, метод возвращает пустой массив.</span><span class="sxs-lookup"><span data-stu-id="a0efc-p167">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="a0efc-1041">Тип: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.7)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.7)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.7)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.7))></span><span class="sxs-lookup"><span data-stu-id="a0efc-1041">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.7)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.7)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.7)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.7))></span></span>

<br>

---
---

#### <a name="getregexmatches--object"></a><span data-ttu-id="a0efc-1042">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="a0efc-1042">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="a0efc-1043">Возвращает строковые значения в выбранном элементе, которые соответствуют регулярным выражениям, определенным в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="a0efc-1043">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="a0efc-1044">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="a0efc-1044">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="a0efc-p168">Метод `getRegExMatches` возвращает строки, соответствующие регулярному выражению, которое определяется в каждом элементе правила `ItemHasRegularExpressionMatch` или `ItemHasKnownEntity` в XML-файле манифеста. Для правила `ItemHasRegularExpressionMatch` соответствующую строку должно содержать свойство элемента, указанного этим правилом. Простой тип `PropertyName` определяет поддерживаемые свойства.</span><span class="sxs-lookup"><span data-stu-id="a0efc-p168">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="a0efc-1048">Например, рассмотрим манифест надстройки, который содержит указанный ниже элемент `Rule`.</span><span class="sxs-lookup"><span data-stu-id="a0efc-1048">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="a0efc-1049">Объект, возвращаемый методом `getRegExMatches`, будет содержать два свойства: `fruits` и `veggies`.</span><span class="sxs-lookup"><span data-stu-id="a0efc-1049">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="a0efc-p169">Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты. Лучше используйте метод [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.7#getasync-coerciontype--options--callback-) для этого.</span><span class="sxs-lookup"><span data-stu-id="a0efc-p169">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.7#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="a0efc-1053">Requirements</span><span class="sxs-lookup"><span data-stu-id="a0efc-1053">Requirements</span></span>

|<span data-ttu-id="a0efc-1054">Требование</span><span class="sxs-lookup"><span data-stu-id="a0efc-1054">Requirement</span></span>|<span data-ttu-id="a0efc-1055">Значение</span><span class="sxs-lookup"><span data-stu-id="a0efc-1055">Value</span></span>|
|---|---|
|[<span data-ttu-id="a0efc-1056">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="a0efc-1056">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a0efc-1057">1.0</span><span class="sxs-lookup"><span data-stu-id="a0efc-1057">1.0</span></span>|
|[<span data-ttu-id="a0efc-1058">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="a0efc-1058">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a0efc-1059">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a0efc-1059">ReadItem</span></span>|
|[<span data-ttu-id="a0efc-1060">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="a0efc-1060">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="a0efc-1061">Чтение</span><span class="sxs-lookup"><span data-stu-id="a0efc-1061">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="a0efc-1062">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="a0efc-1062">Returns:</span></span>

<span data-ttu-id="a0efc-p170">Объект, содержащий массив строк, которые соответствуют регулярным выражениям, определяемым в XML-файле манифеста. Имя каждого массива равно соответствующему значению атрибута `RegExName` подходящего правила `ItemHasRegularExpressionMatch` или атрибута `FilterName` соответствующего правила `ItemHasKnownEntity`.</span><span class="sxs-lookup"><span data-stu-id="a0efc-p170">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<span data-ttu-id="a0efc-1065">Тип: Object</span><span class="sxs-lookup"><span data-stu-id="a0efc-1065">Type: Object</span></span>

##### <a name="example"></a><span data-ttu-id="a0efc-1066">Пример</span><span class="sxs-lookup"><span data-stu-id="a0efc-1066">Example</span></span>

<span data-ttu-id="a0efc-1067">В приведенном ниже примере показано, как получить доступ к массиву совпадений с элементами `fruits` и `veggies` правил активации регулярных выражений, указанными в манифесте.</span><span class="sxs-lookup"><span data-stu-id="a0efc-1067">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```js
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

<br>

---
---

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="a0efc-1068">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="a0efc-1068">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="a0efc-1069">Возвращает строковые значения в выбранном элементе, которые соответствуют именованному регулярному выражению, определенному в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="a0efc-1069">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="a0efc-1070">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="a0efc-1070">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="a0efc-1071">Метод `getRegExMatchesByName` возвращает строки, соответствующие регулярному выражению, которое определяется в элементе правила `ItemHasRegularExpressionMatch` в XML-файле манифеста, с использованием указанного значения элемента `RegExName`.</span><span class="sxs-lookup"><span data-stu-id="a0efc-1071">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="a0efc-p171">Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты.</span><span class="sxs-lookup"><span data-stu-id="a0efc-p171">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a0efc-1074">Параметры</span><span class="sxs-lookup"><span data-stu-id="a0efc-1074">Parameters</span></span>

|<span data-ttu-id="a0efc-1075">Имя</span><span class="sxs-lookup"><span data-stu-id="a0efc-1075">Name</span></span>|<span data-ttu-id="a0efc-1076">Тип</span><span class="sxs-lookup"><span data-stu-id="a0efc-1076">Type</span></span>|<span data-ttu-id="a0efc-1077">Описание</span><span class="sxs-lookup"><span data-stu-id="a0efc-1077">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="a0efc-1078">String</span><span class="sxs-lookup"><span data-stu-id="a0efc-1078">String</span></span>|<span data-ttu-id="a0efc-1079">Имя элемента правила `ItemHasRegularExpressionMatch`, определяющее соответствующий фильтр.</span><span class="sxs-lookup"><span data-stu-id="a0efc-1079">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a0efc-1080">Requirements</span><span class="sxs-lookup"><span data-stu-id="a0efc-1080">Requirements</span></span>

|<span data-ttu-id="a0efc-1081">Требование</span><span class="sxs-lookup"><span data-stu-id="a0efc-1081">Requirement</span></span>|<span data-ttu-id="a0efc-1082">Значение</span><span class="sxs-lookup"><span data-stu-id="a0efc-1082">Value</span></span>|
|---|---|
|[<span data-ttu-id="a0efc-1083">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="a0efc-1083">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a0efc-1084">1.0</span><span class="sxs-lookup"><span data-stu-id="a0efc-1084">1.0</span></span>|
|[<span data-ttu-id="a0efc-1085">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="a0efc-1085">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a0efc-1086">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a0efc-1086">ReadItem</span></span>|
|[<span data-ttu-id="a0efc-1087">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="a0efc-1087">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="a0efc-1088">Чтение</span><span class="sxs-lookup"><span data-stu-id="a0efc-1088">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="a0efc-1089">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="a0efc-1089">Returns:</span></span>

<span data-ttu-id="a0efc-1090">Массив строк, соответствующих регулярному выражению, определяемому в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="a0efc-1090">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<span data-ttu-id="a0efc-1091">Тип: Array.< String ></span><span class="sxs-lookup"><span data-stu-id="a0efc-1091">Type: Array.< String ></span></span>

##### <a name="example"></a><span data-ttu-id="a0efc-1092">Пример</span><span class="sxs-lookup"><span data-stu-id="a0efc-1092">Example</span></span>

```js
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

<br>

---
---

#### <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="a0efc-1093">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="a0efc-1093">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="a0efc-1094">Асинхронно возвращает данные, выбранные в теме или тексте сообщения.</span><span class="sxs-lookup"><span data-stu-id="a0efc-1094">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="a0efc-p172">Если выделенный фрагмент отсутствует, но курсор находится в тексте или теме, метод возвращает пустую строку для выбранных данных. Если выбраны не текст и не тема, метод возвращает ошибку `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="a0efc-p172">If there is no selection but the cursor is in the body or subject, the method returns an empty string for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a0efc-1097">Параметры</span><span class="sxs-lookup"><span data-stu-id="a0efc-1097">Parameters</span></span>

|<span data-ttu-id="a0efc-1098">Имя</span><span class="sxs-lookup"><span data-stu-id="a0efc-1098">Name</span></span>|<span data-ttu-id="a0efc-1099">Тип</span><span class="sxs-lookup"><span data-stu-id="a0efc-1099">Type</span></span>|<span data-ttu-id="a0efc-1100">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="a0efc-1100">Attributes</span></span>|<span data-ttu-id="a0efc-1101">Описание</span><span class="sxs-lookup"><span data-stu-id="a0efc-1101">Description</span></span>|
|---|---|---|---|
|`coercionType`|[<span data-ttu-id="a0efc-1102">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="a0efc-1102">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="a0efc-p173">Запрашивает формат данных. Если задано значение Text, метод возвращает обычный текст как строку, удаляя все имеющиеся HTML-теги. Если задано значение HTML, метод возвращает выделенный текст (обычный текст или HTML).</span><span class="sxs-lookup"><span data-stu-id="a0efc-p173">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`|<span data-ttu-id="a0efc-1106">Object</span><span class="sxs-lookup"><span data-stu-id="a0efc-1106">Object</span></span>|<span data-ttu-id="a0efc-1107">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="a0efc-1107">&lt;optional&gt;</span></span>|<span data-ttu-id="a0efc-1108">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="a0efc-1108">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="a0efc-1109">Объект</span><span class="sxs-lookup"><span data-stu-id="a0efc-1109">Object</span></span>|<span data-ttu-id="a0efc-1110">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="a0efc-1110">&lt;optional&gt;</span></span>|<span data-ttu-id="a0efc-1111">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="a0efc-1111">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="a0efc-1112">функция</span><span class="sxs-lookup"><span data-stu-id="a0efc-1112">function</span></span>||<span data-ttu-id="a0efc-1113">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="a0efc-1113">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="a0efc-1114">Чтобы получить доступ к выбранным данным из метода обратного вызова, вызовите `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="a0efc-1114">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="a0efc-1115">Для доступа к исходному свойству, представляющему собой источник выбранных данных, вызовите параметр `asyncResult.value.sourceProperty`, который может иметь значение `body` или `subject`.</span><span class="sxs-lookup"><span data-stu-id="a0efc-1115">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a0efc-1116">Requirements</span><span class="sxs-lookup"><span data-stu-id="a0efc-1116">Requirements</span></span>

|<span data-ttu-id="a0efc-1117">Требование</span><span class="sxs-lookup"><span data-stu-id="a0efc-1117">Requirement</span></span>|<span data-ttu-id="a0efc-1118">Значение</span><span class="sxs-lookup"><span data-stu-id="a0efc-1118">Value</span></span>|
|---|---|
|[<span data-ttu-id="a0efc-1119">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="a0efc-1119">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a0efc-1120">1.2</span><span class="sxs-lookup"><span data-stu-id="a0efc-1120">1.2</span></span>|
|[<span data-ttu-id="a0efc-1121">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="a0efc-1121">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a0efc-1122">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a0efc-1122">ReadItem</span></span>|
|[<span data-ttu-id="a0efc-1123">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="a0efc-1123">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="a0efc-1124">Создание</span><span class="sxs-lookup"><span data-stu-id="a0efc-1124">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="a0efc-1125">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="a0efc-1125">Returns:</span></span>

<span data-ttu-id="a0efc-1126">Выбранные данные в виде строки с форматом, определенным в параметре `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="a0efc-1126">The selected data as a string with format determined by `coercionType`.</span></span>

<span data-ttu-id="a0efc-1127">Тип: строка</span><span class="sxs-lookup"><span data-stu-id="a0efc-1127">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="a0efc-1128">Пример</span><span class="sxs-lookup"><span data-stu-id="a0efc-1128">Example</span></span>

```js
// Get selected data.
Office.initialize = function () {
  Office.context.mailbox.item.getSelectedDataAsync(Office.CoercionType.Text, {}, getCallback);
};

function getCallback(asyncResult) {
  var text = asyncResult.value.data;
  var prop = asyncResult.value.sourceProperty;

  console.log("Selected text in " + prop + ": " + text);
}
```

<br>

---
---

#### <a name="getselectedentities--entitiesjavascriptapioutlookofficeentitiesviewoutlook-js-17"></a><span data-ttu-id="a0efc-1129">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.7)}</span><span class="sxs-lookup"><span data-stu-id="a0efc-1129">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.7)}</span></span>

<span data-ttu-id="a0efc-1130">Возвращает сущности, найденные в выделенном совпадении, выбранном пользователем.</span><span class="sxs-lookup"><span data-stu-id="a0efc-1130">Gets the entities found in a highlighted match a user has selected.</span></span> <span data-ttu-id="a0efc-1131">Выделенные совпадения применяются к [контекстным надстройкам](/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="a0efc-1131">Highlighted matches apply to [contextual add-ins](/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="a0efc-1132">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="a0efc-1132">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="a0efc-1133">Требования</span><span class="sxs-lookup"><span data-stu-id="a0efc-1133">Requirements</span></span>

|<span data-ttu-id="a0efc-1134">Требование</span><span class="sxs-lookup"><span data-stu-id="a0efc-1134">Requirement</span></span>|<span data-ttu-id="a0efc-1135">Значение</span><span class="sxs-lookup"><span data-stu-id="a0efc-1135">Value</span></span>|
|---|---|
|[<span data-ttu-id="a0efc-1136">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="a0efc-1136">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a0efc-1137">1.6</span><span class="sxs-lookup"><span data-stu-id="a0efc-1137">1.6</span></span>|
|[<span data-ttu-id="a0efc-1138">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="a0efc-1138">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a0efc-1139">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a0efc-1139">ReadItem</span></span>|
|[<span data-ttu-id="a0efc-1140">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="a0efc-1140">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="a0efc-1141">Чтение</span><span class="sxs-lookup"><span data-stu-id="a0efc-1141">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="a0efc-1142">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="a0efc-1142">Returns:</span></span>

<span data-ttu-id="a0efc-1143">Тип: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="a0efc-1143">Type: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.7)</span></span>

##### <a name="example"></a><span data-ttu-id="a0efc-1144">Пример</span><span class="sxs-lookup"><span data-stu-id="a0efc-1144">Example</span></span>

<span data-ttu-id="a0efc-1145">В приведенном ниже примере показано, как получить доступ к сущностям адресов в выделенном совпадении, выбранном пользователем.</span><span class="sxs-lookup"><span data-stu-id="a0efc-1145">The following example accesses the addresses entities in the highlighted match selected by the user.</span></span>

```js
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

<br>

---
---

#### <a name="getselectedregexmatches--object"></a><span data-ttu-id="a0efc-1146">getSelectedRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="a0efc-1146">getSelectedRegExMatches() → {Object}</span></span>

<span data-ttu-id="a0efc-p176">Возвращает строковые значения в выделенном совпадении, которые соответствуют регулярным выражениям, определенным в XML-файле манифеста. Выделенные совпадения применяются к [контекстным надстройкам](/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="a0efc-p176">Returns string values in a highlighted match that match the regular expressions defined in the manifest XML file. Highlighted matches apply to [contextual add-ins](/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="a0efc-1149">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="a0efc-1149">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="a0efc-p177">Метод `getSelectedRegExMatches` возвращает строки, соответствующие регулярному выражению, которое определяется в каждом элементе правила `ItemHasRegularExpressionMatch` или `ItemHasKnownEntity` в XML-файле манифеста. Для правила `ItemHasRegularExpressionMatch` соответствующую строку должно содержать свойство элемента, указанного этим правилом. Простой тип `PropertyName` определяет поддерживаемые свойства.</span><span class="sxs-lookup"><span data-stu-id="a0efc-p177">The `getSelectedRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="a0efc-1153">Например, рассмотрим манифест надстройки, который содержит указанный ниже элемент `Rule`.</span><span class="sxs-lookup"><span data-stu-id="a0efc-1153">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="a0efc-1154">Объект, возвращаемый методом `getRegExMatches`, будет содержать два свойства: `fruits` и `veggies`.</span><span class="sxs-lookup"><span data-stu-id="a0efc-1154">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="a0efc-p178">Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты. Лучше используйте метод [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.7#getasync-coerciontype--options--callback-) для этого.</span><span class="sxs-lookup"><span data-stu-id="a0efc-p178">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.7#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="a0efc-1158">Requirements</span><span class="sxs-lookup"><span data-stu-id="a0efc-1158">Requirements</span></span>

|<span data-ttu-id="a0efc-1159">Требование</span><span class="sxs-lookup"><span data-stu-id="a0efc-1159">Requirement</span></span>|<span data-ttu-id="a0efc-1160">Значение</span><span class="sxs-lookup"><span data-stu-id="a0efc-1160">Value</span></span>|
|---|---|
|[<span data-ttu-id="a0efc-1161">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="a0efc-1161">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a0efc-1162">1.6</span><span class="sxs-lookup"><span data-stu-id="a0efc-1162">1.6</span></span>|
|[<span data-ttu-id="a0efc-1163">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="a0efc-1163">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a0efc-1164">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a0efc-1164">ReadItem</span></span>|
|[<span data-ttu-id="a0efc-1165">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="a0efc-1165">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="a0efc-1166">Чтение</span><span class="sxs-lookup"><span data-stu-id="a0efc-1166">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="a0efc-1167">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="a0efc-1167">Returns:</span></span>

<span data-ttu-id="a0efc-p179">Объект, содержащий массив строк, которые соответствуют регулярным выражениям, определяемым в XML-файле манифеста. Имя каждого массива равно соответствующему значению атрибута `RegExName` подходящего правила `ItemHasRegularExpressionMatch` или атрибута `FilterName` соответствующего правила `ItemHasKnownEntity`.</span><span class="sxs-lookup"><span data-stu-id="a0efc-p179">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

##### <a name="example"></a><span data-ttu-id="a0efc-1170">Пример</span><span class="sxs-lookup"><span data-stu-id="a0efc-1170">Example</span></span>

<span data-ttu-id="a0efc-1171">В приведенном ниже примере показано, как получить доступ к массиву совпадений с элементами `fruits` и `veggies` правил активации регулярных выражений, указанными в манифесте.</span><span class="sxs-lookup"><span data-stu-id="a0efc-1171">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```js
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

<br>

---
---

#### <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="a0efc-1172">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="a0efc-1172">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="a0efc-1173">Асинхронно загружает настраиваемые свойства для надстройки для выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="a0efc-1173">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="a0efc-p180">Настраиваемые свойства сохраняются в виде пар "ключ-значение" для каждого приложения и каждого элемента. Этот метод возвращает объект `CustomProperties` при обратном вызове, который предоставляет методы для доступа к настраиваемым свойствам, характерным для текущего элемента и текущей надстройки. Настраиваемые свойства не шифруются для элемента, поэтому этот способ хранения не является безопасным.</span><span class="sxs-lookup"><span data-stu-id="a0efc-p180">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a0efc-1177">Параметры</span><span class="sxs-lookup"><span data-stu-id="a0efc-1177">Parameters</span></span>

|<span data-ttu-id="a0efc-1178">Имя</span><span class="sxs-lookup"><span data-stu-id="a0efc-1178">Name</span></span>|<span data-ttu-id="a0efc-1179">Тип</span><span class="sxs-lookup"><span data-stu-id="a0efc-1179">Type</span></span>|<span data-ttu-id="a0efc-1180">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="a0efc-1180">Attributes</span></span>|<span data-ttu-id="a0efc-1181">Описание</span><span class="sxs-lookup"><span data-stu-id="a0efc-1181">Description</span></span>|
|---|---|---|---|
|`callback`|<span data-ttu-id="a0efc-1182">function</span><span class="sxs-lookup"><span data-stu-id="a0efc-1182">function</span></span>||<span data-ttu-id="a0efc-1183">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="a0efc-1183">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="a0efc-1184">Настраиваемые свойства предоставляются в виде объекта [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.7) в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="a0efc-1184">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.7) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="a0efc-1185">Этот объект позволяет получить, задать и удалить настраиваемые свойства для элемента, а также сохранить изменения, внесенные в набор настраиваемых свойств, на сервере.</span><span class="sxs-lookup"><span data-stu-id="a0efc-1185">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`|<span data-ttu-id="a0efc-1186">Объект</span><span class="sxs-lookup"><span data-stu-id="a0efc-1186">Object</span></span>|<span data-ttu-id="a0efc-1187">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="a0efc-1187">&lt;optional&gt;</span></span>|<span data-ttu-id="a0efc-1188">Разработчики могут указать любой объект, к которому необходимо получить доступ, в функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="a0efc-1188">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="a0efc-1189">Доступ к этому объекту можно получить с помощью свойства `asyncResult.asyncContext` в функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="a0efc-1189">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a0efc-1190">Requirements</span><span class="sxs-lookup"><span data-stu-id="a0efc-1190">Requirements</span></span>

|<span data-ttu-id="a0efc-1191">Требование</span><span class="sxs-lookup"><span data-stu-id="a0efc-1191">Requirement</span></span>|<span data-ttu-id="a0efc-1192">Значение</span><span class="sxs-lookup"><span data-stu-id="a0efc-1192">Value</span></span>|
|---|---|
|[<span data-ttu-id="a0efc-1193">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="a0efc-1193">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a0efc-1194">1.0</span><span class="sxs-lookup"><span data-stu-id="a0efc-1194">1.0</span></span>|
|[<span data-ttu-id="a0efc-1195">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="a0efc-1195">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a0efc-1196">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a0efc-1196">ReadItem</span></span>|
|[<span data-ttu-id="a0efc-1197">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="a0efc-1197">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="a0efc-1198">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="a0efc-1198">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a0efc-1199">Пример</span><span class="sxs-lookup"><span data-stu-id="a0efc-1199">Example</span></span>

<span data-ttu-id="a0efc-p183">Приведенный ниже пример кода показывает, как асинхронно загружать настраиваемые свойства, характерные для текущего элемента, с помощью метода `loadCustomPropertiesAsync`. Этот пример также показывает, как сохранять эти свойства на сервере с помощью метода `CustomProperties.saveAsync`. После загрузки настраиваемых свойств в этом примере кода метод `CustomProperties.get` используется для считывания настраиваемого свойства `myProp`, метод `CustomProperties.set` — для записи настраиваемого свойства `otherProp`, а метод `saveAsync` — для сохранения настраиваемых свойств.</span><span class="sxs-lookup"><span data-stu-id="a0efc-p183">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

#### <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="a0efc-1203">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="a0efc-1203">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="a0efc-1204">Удаляет вложение из сообщения или встречи.</span><span class="sxs-lookup"><span data-stu-id="a0efc-1204">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="a0efc-1205">Метод `removeAttachmentAsync` удаляет из элемента вложение с указанным идентификатором.</span><span class="sxs-lookup"><span data-stu-id="a0efc-1205">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="a0efc-1206">Идентификатор вложения рекомендуется использовать для удаления вложения, только если оно добавлено тем же почтовым приложением в ходе текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="a0efc-1206">As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session.</span></span> <span data-ttu-id="a0efc-1207">В Outlook в Интернете и на мобильных устройствах идентификатор вложения действителен только в течение одного сеанса.</span><span class="sxs-lookup"><span data-stu-id="a0efc-1207">In Outlook on the web and mobile devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="a0efc-1208">Сеанс завершается, когда пользователь закрывает приложение или начинает создавать элемент во встроенной форме, а затем переходит из формы в отдельное окно.</span><span class="sxs-lookup"><span data-stu-id="a0efc-1208">A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a0efc-1209">Параметры</span><span class="sxs-lookup"><span data-stu-id="a0efc-1209">Parameters</span></span>

|<span data-ttu-id="a0efc-1210">Имя</span><span class="sxs-lookup"><span data-stu-id="a0efc-1210">Name</span></span>|<span data-ttu-id="a0efc-1211">Тип</span><span class="sxs-lookup"><span data-stu-id="a0efc-1211">Type</span></span>|<span data-ttu-id="a0efc-1212">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="a0efc-1212">Attributes</span></span>|<span data-ttu-id="a0efc-1213">Описание</span><span class="sxs-lookup"><span data-stu-id="a0efc-1213">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="a0efc-1214">String</span><span class="sxs-lookup"><span data-stu-id="a0efc-1214">String</span></span>||<span data-ttu-id="a0efc-1215">Идентификатор удаляемого вложения.</span><span class="sxs-lookup"><span data-stu-id="a0efc-1215">The identifier of the attachment to remove.</span></span>|
|`options`|<span data-ttu-id="a0efc-1216">Объект</span><span class="sxs-lookup"><span data-stu-id="a0efc-1216">Object</span></span>|<span data-ttu-id="a0efc-1217">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="a0efc-1217">&lt;optional&gt;</span></span>|<span data-ttu-id="a0efc-1218">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="a0efc-1218">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="a0efc-1219">Объект</span><span class="sxs-lookup"><span data-stu-id="a0efc-1219">Object</span></span>|<span data-ttu-id="a0efc-1220">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="a0efc-1220">&lt;optional&gt;</span></span>|<span data-ttu-id="a0efc-1221">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="a0efc-1221">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="a0efc-1222">функция</span><span class="sxs-lookup"><span data-stu-id="a0efc-1222">function</span></span>|<span data-ttu-id="a0efc-1223">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="a0efc-1223">&lt;optional&gt;</span></span>|<span data-ttu-id="a0efc-1224">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="a0efc-1224">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="a0efc-1225">Если удалить вложение не удается, свойство `asyncResult.error` содержит код ошибки с указанием ее причины.</span><span class="sxs-lookup"><span data-stu-id="a0efc-1225">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="a0efc-1226">Ошибки</span><span class="sxs-lookup"><span data-stu-id="a0efc-1226">Errors</span></span>

|<span data-ttu-id="a0efc-1227">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="a0efc-1227">Error code</span></span>|<span data-ttu-id="a0efc-1228">Описание</span><span class="sxs-lookup"><span data-stu-id="a0efc-1228">Description</span></span>|
|------------|-------------|
|`InvalidAttachmentId`|<span data-ttu-id="a0efc-1229">Идентификатор вложения не существует.</span><span class="sxs-lookup"><span data-stu-id="a0efc-1229">The attachment identifier does not exist.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a0efc-1230">Requirements</span><span class="sxs-lookup"><span data-stu-id="a0efc-1230">Requirements</span></span>

|<span data-ttu-id="a0efc-1231">Требование</span><span class="sxs-lookup"><span data-stu-id="a0efc-1231">Requirement</span></span>|<span data-ttu-id="a0efc-1232">Значение</span><span class="sxs-lookup"><span data-stu-id="a0efc-1232">Value</span></span>|
|---|---|
|[<span data-ttu-id="a0efc-1233">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="a0efc-1233">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a0efc-1234">1.1</span><span class="sxs-lookup"><span data-stu-id="a0efc-1234">1.1</span></span>|
|[<span data-ttu-id="a0efc-1235">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="a0efc-1235">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a0efc-1236">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="a0efc-1236">ReadWriteItem</span></span>|
|[<span data-ttu-id="a0efc-1237">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="a0efc-1237">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="a0efc-1238">Создание</span><span class="sxs-lookup"><span data-stu-id="a0efc-1238">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="a0efc-1239">Пример</span><span class="sxs-lookup"><span data-stu-id="a0efc-1239">Example</span></span>

<span data-ttu-id="a0efc-1240">Указанный ниже код удаляет вложение с идентификатором "0".</span><span class="sxs-lookup"><span data-stu-id="a0efc-1240">The following code removes an attachment with an identifier of '0'.</span></span>

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

#### <a name="removehandlerasynceventtype-options-callback"></a><span data-ttu-id="a0efc-1241">removeHandlerAsync(eventType, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="a0efc-1241">removeHandlerAsync(eventType, [options], [callback])</span></span>

<span data-ttu-id="a0efc-1242">Удаляет обработчиков для поддерживаемого типа события.</span><span class="sxs-lookup"><span data-stu-id="a0efc-1242">Removes the event handlers for a supported event type.</span></span>

<span data-ttu-id="a0efc-1243">В настоящее время поддерживаются типы `Office.EventType.AppointmentTimeChanged`событий `Office.EventType.RecipientsChanged`, и`Office.EventType.RecurrenceChanged`</span><span class="sxs-lookup"><span data-stu-id="a0efc-1243">Currently the supported event types are `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`</span></span>

##### <a name="parameters"></a><span data-ttu-id="a0efc-1244">Параметры</span><span class="sxs-lookup"><span data-stu-id="a0efc-1244">Parameters</span></span>

| <span data-ttu-id="a0efc-1245">Имя</span><span class="sxs-lookup"><span data-stu-id="a0efc-1245">Name</span></span> | <span data-ttu-id="a0efc-1246">Тип</span><span class="sxs-lookup"><span data-stu-id="a0efc-1246">Type</span></span> | <span data-ttu-id="a0efc-1247">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="a0efc-1247">Attributes</span></span> | <span data-ttu-id="a0efc-1248">Описание</span><span class="sxs-lookup"><span data-stu-id="a0efc-1248">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="a0efc-1249">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="a0efc-1249">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="a0efc-1250">Событие, которое должно вызвать обработчик.</span><span class="sxs-lookup"><span data-stu-id="a0efc-1250">The event that should invoke the handler.</span></span> |
| `options` | <span data-ttu-id="a0efc-1251">Объект</span><span class="sxs-lookup"><span data-stu-id="a0efc-1251">Object</span></span> | <span data-ttu-id="a0efc-1252">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="a0efc-1252">&lt;optional&gt;</span></span> | <span data-ttu-id="a0efc-1253">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="a0efc-1253">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="a0efc-1254">Объект</span><span class="sxs-lookup"><span data-stu-id="a0efc-1254">Object</span></span> | <span data-ttu-id="a0efc-1255">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="a0efc-1255">&lt;optional&gt;</span></span> | <span data-ttu-id="a0efc-1256">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="a0efc-1256">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="a0efc-1257">функция</span><span class="sxs-lookup"><span data-stu-id="a0efc-1257">function</span></span>| <span data-ttu-id="a0efc-1258">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="a0efc-1258">&lt;optional&gt;</span></span>|<span data-ttu-id="a0efc-1259">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="a0efc-1259">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a0efc-1260">Требования</span><span class="sxs-lookup"><span data-stu-id="a0efc-1260">Requirements</span></span>

|<span data-ttu-id="a0efc-1261">Требование</span><span class="sxs-lookup"><span data-stu-id="a0efc-1261">Requirement</span></span>| <span data-ttu-id="a0efc-1262">Значение</span><span class="sxs-lookup"><span data-stu-id="a0efc-1262">Value</span></span>|
|---|---|
|[<span data-ttu-id="a0efc-1263">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="a0efc-1263">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a0efc-1264">1.7</span><span class="sxs-lookup"><span data-stu-id="a0efc-1264">1.7</span></span> |
|[<span data-ttu-id="a0efc-1265">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="a0efc-1265">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a0efc-1266">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a0efc-1266">ReadItem</span></span> |
|[<span data-ttu-id="a0efc-1267">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="a0efc-1267">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a0efc-1268">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="a0efc-1268">Compose or Read</span></span> |

##### <a name="example"></a><span data-ttu-id="a0efc-1269">Пример</span><span class="sxs-lookup"><span data-stu-id="a0efc-1269">Example</span></span>

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

<br>

---
---

#### <a name="saveasyncoptions-callback"></a><span data-ttu-id="a0efc-1270">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="a0efc-1270">saveAsync([options], callback)</span></span>

<span data-ttu-id="a0efc-1271">Асинхронно сохраняет элемент.</span><span class="sxs-lookup"><span data-stu-id="a0efc-1271">Asynchronously saves an item.</span></span>

<span data-ttu-id="a0efc-1272">При вызове этот метод сохраняет текущее сообщение в виде черновика и возвращает идентификатор элемента с помощью метода обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="a0efc-1272">When invoked, this method saves the current message as a draft and returns the item id via the callback method.</span></span> <span data-ttu-id="a0efc-1273">В Outlook в Интернете или интерактивном режиме Outlook этот элемент сохраняется на сервере.</span><span class="sxs-lookup"><span data-stu-id="a0efc-1273">In Outlook on the web or Outlook in online mode, the item is saved to the server.</span></span> <span data-ttu-id="a0efc-1274">В Outlook в режиме кэширования этот элемент сохраняется в локальном кэше.</span><span class="sxs-lookup"><span data-stu-id="a0efc-1274">In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="a0efc-1275">Если в режиме создания надстройка вызывает для элемента метод `saveAsync`, чтобы получить параметр `itemId` для использования с EWS или REST API, необходимо помнить, что синхронизация элемента с сервером может занять много времени, если Outlook работает в режиме кэширования данных.</span><span class="sxs-lookup"><span data-stu-id="a0efc-1275">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="a0efc-1276">До окончания синхронизации элемента применение параметра `itemId` будет приводить к ошибке.</span><span class="sxs-lookup"><span data-stu-id="a0efc-1276">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="a0efc-p187">Если метод `saveAsync` вызывается для встречи в режиме создания, она сохраняется как обычная встреча в календаре пользователя, а не как черновик. При сохранении новой встречи приглашения не отправляются. При сохранении существующей встречи уведомления отправляются добавленным или удаленным участникам.</span><span class="sxs-lookup"><span data-stu-id="a0efc-p187">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="a0efc-1280">Следующие клиенты отличаются другим поведением `saveAsync` в отношении встреч в режиме создания:</span><span class="sxs-lookup"><span data-stu-id="a0efc-1280">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="a0efc-1281">Outlook для Mac не поддерживает сохранение собрания.</span><span class="sxs-lookup"><span data-stu-id="a0efc-1281">Outlook on Mac does not support saving a meeting.</span></span> <span data-ttu-id="a0efc-1282">Метод `saveAsync` не работает при вызове из собрания в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="a0efc-1282">The `saveAsync` method fails when called from a meeting in compose mode.</span></span> <span data-ttu-id="a0efc-1283">Временное решение представлено в статье [Не удается сохранить встречу как черновик в Outlook для Mac с помощью API JS для Office](https://support.microsoft.com/help/4505745).</span><span class="sxs-lookup"><span data-stu-id="a0efc-1283">See [Cannot save a meeting as a draft in Outlook for Mac by using Office JS API](https://support.microsoft.com/help/4505745) for a workaround.</span></span>
> - <span data-ttu-id="a0efc-1284">Outlook в Интернете всегда отправляет приглашение или обновление при вызове `saveAsync` для встречи в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="a0efc-1284">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a0efc-1285">Параметры</span><span class="sxs-lookup"><span data-stu-id="a0efc-1285">Parameters</span></span>

|<span data-ttu-id="a0efc-1286">Имя</span><span class="sxs-lookup"><span data-stu-id="a0efc-1286">Name</span></span>|<span data-ttu-id="a0efc-1287">Тип</span><span class="sxs-lookup"><span data-stu-id="a0efc-1287">Type</span></span>|<span data-ttu-id="a0efc-1288">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="a0efc-1288">Attributes</span></span>|<span data-ttu-id="a0efc-1289">Описание</span><span class="sxs-lookup"><span data-stu-id="a0efc-1289">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="a0efc-1290">Object</span><span class="sxs-lookup"><span data-stu-id="a0efc-1290">Object</span></span>|<span data-ttu-id="a0efc-1291">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="a0efc-1291">&lt;optional&gt;</span></span>|<span data-ttu-id="a0efc-1292">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="a0efc-1292">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="a0efc-1293">Объект</span><span class="sxs-lookup"><span data-stu-id="a0efc-1293">Object</span></span>|<span data-ttu-id="a0efc-1294">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="a0efc-1294">&lt;optional&gt;</span></span>|<span data-ttu-id="a0efc-1295">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="a0efc-1295">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="a0efc-1296">функция</span><span class="sxs-lookup"><span data-stu-id="a0efc-1296">function</span></span>||<span data-ttu-id="a0efc-1297">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="a0efc-1297">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="a0efc-1298">После успешного выполнения идентификатор элемента будет предоставлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="a0efc-1298">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a0efc-1299">Requirements</span><span class="sxs-lookup"><span data-stu-id="a0efc-1299">Requirements</span></span>

|<span data-ttu-id="a0efc-1300">Требование</span><span class="sxs-lookup"><span data-stu-id="a0efc-1300">Requirement</span></span>|<span data-ttu-id="a0efc-1301">Значение</span><span class="sxs-lookup"><span data-stu-id="a0efc-1301">Value</span></span>|
|---|---|
|[<span data-ttu-id="a0efc-1302">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="a0efc-1302">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a0efc-1303">1.3</span><span class="sxs-lookup"><span data-stu-id="a0efc-1303">1.3</span></span>|
|[<span data-ttu-id="a0efc-1304">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="a0efc-1304">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a0efc-1305">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="a0efc-1305">ReadWriteItem</span></span>|
|[<span data-ttu-id="a0efc-1306">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="a0efc-1306">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="a0efc-1307">Создание</span><span class="sxs-lookup"><span data-stu-id="a0efc-1307">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="a0efc-1308">Примеры</span><span class="sxs-lookup"><span data-stu-id="a0efc-1308">Examples</span></span>

```js
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result.
  });
```

<span data-ttu-id="a0efc-p189">Ниже приведен пример параметра `result`, переданного функции обратного вызова. Свойство `value` содержит идентификатор элемента.</span><span class="sxs-lookup"><span data-stu-id="a0efc-p189">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

<br>

---
---

#### <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="a0efc-1311">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="a0efc-1311">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="a0efc-1312">Асинхронно вставляет данные в текст или тему сообщения.</span><span class="sxs-lookup"><span data-stu-id="a0efc-1312">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="a0efc-p190">Метод `setSelectedDataAsync` вставляет указанную строку в местоположение курсора в теме или тексте элемента либо, если текст выделен в редакторе, он заменяет выделенный текст. Если курсор находится вне текста или темы элемента, возвращается ошибка. После вставки курсор помещается в конец вставленного содержимого.</span><span class="sxs-lookup"><span data-stu-id="a0efc-p190">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a0efc-1316">Параметры</span><span class="sxs-lookup"><span data-stu-id="a0efc-1316">Parameters</span></span>

|<span data-ttu-id="a0efc-1317">Имя</span><span class="sxs-lookup"><span data-stu-id="a0efc-1317">Name</span></span>|<span data-ttu-id="a0efc-1318">Тип</span><span class="sxs-lookup"><span data-stu-id="a0efc-1318">Type</span></span>|<span data-ttu-id="a0efc-1319">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="a0efc-1319">Attributes</span></span>|<span data-ttu-id="a0efc-1320">Описание</span><span class="sxs-lookup"><span data-stu-id="a0efc-1320">Description</span></span>|
|---|---|---|---|
|`data`|<span data-ttu-id="a0efc-1321">String</span><span class="sxs-lookup"><span data-stu-id="a0efc-1321">String</span></span>||<span data-ttu-id="a0efc-p191">Вставляемые данные. Объем данных не должен превышать 1 000 000 символов. Если передано больше 1 000 000 символов, возвращается исключение `ArgumentOutOfRange`.</span><span class="sxs-lookup"><span data-stu-id="a0efc-p191">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`|<span data-ttu-id="a0efc-1325">Object</span><span class="sxs-lookup"><span data-stu-id="a0efc-1325">Object</span></span>|<span data-ttu-id="a0efc-1326">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="a0efc-1326">&lt;optional&gt;</span></span>|<span data-ttu-id="a0efc-1327">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="a0efc-1327">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="a0efc-1328">Объект</span><span class="sxs-lookup"><span data-stu-id="a0efc-1328">Object</span></span>|<span data-ttu-id="a0efc-1329">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="a0efc-1329">&lt;optional&gt;</span></span>|<span data-ttu-id="a0efc-1330">В методе обратного вызова разработчики могут указать любой объект, к которому необходимо получить доступ.</span><span class="sxs-lookup"><span data-stu-id="a0efc-1330">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="a0efc-1331">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="a0efc-1331">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="a0efc-1332">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="a0efc-1332">&lt;optional&gt;</span></span>|<span data-ttu-id="a0efc-1333">Если задано значение `text`, текущий стиль применяется в Outlook в Интернете и классических клиентах.</span><span class="sxs-lookup"><span data-stu-id="a0efc-1333">If `text`, the current style is applied in Outlook on the web and desktop clients.</span></span> <span data-ttu-id="a0efc-1334">Если поле представляет собой редактор HTML, вставляются только текстовые данные, даже если они имеют формат HTML.</span><span class="sxs-lookup"><span data-stu-id="a0efc-1334">If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="a0efc-1335">Если задано значение `html` и поле (не тема) поддерживает HTML, в Outlook в Интернете применяется текущий стиль, а в классических клиентах Outlook — стиль по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="a0efc-1335">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook on the web and the default style is applied in Outlook desktop clients.</span></span> <span data-ttu-id="a0efc-1336">Если поле является текстовым, возвращается ошибка `InvalidDataFormat`.</span><span class="sxs-lookup"><span data-stu-id="a0efc-1336">If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="a0efc-1337">Если свойство `coercionType` не задано, результат зависит от поля: если поле имеет формат HTML, используется текст в формате HTML, а если поле текстовое, применяется обычный текст.</span><span class="sxs-lookup"><span data-stu-id="a0efc-1337">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`|<span data-ttu-id="a0efc-1338">функция</span><span class="sxs-lookup"><span data-stu-id="a0efc-1338">function</span></span>||<span data-ttu-id="a0efc-1339">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="a0efc-1339">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a0efc-1340">Требования</span><span class="sxs-lookup"><span data-stu-id="a0efc-1340">Requirements</span></span>

|<span data-ttu-id="a0efc-1341">Требование</span><span class="sxs-lookup"><span data-stu-id="a0efc-1341">Requirement</span></span>|<span data-ttu-id="a0efc-1342">Значение</span><span class="sxs-lookup"><span data-stu-id="a0efc-1342">Value</span></span>|
|---|---|
|[<span data-ttu-id="a0efc-1343">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="a0efc-1343">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="a0efc-1344">1.2</span><span class="sxs-lookup"><span data-stu-id="a0efc-1344">1.2</span></span>|
|[<span data-ttu-id="a0efc-1345">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="a0efc-1345">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="a0efc-1346">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="a0efc-1346">ReadWriteItem</span></span>|
|[<span data-ttu-id="a0efc-1347">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="a0efc-1347">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="a0efc-1348">Создание</span><span class="sxs-lookup"><span data-stu-id="a0efc-1348">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="a0efc-1349">Пример</span><span class="sxs-lookup"><span data-stu-id="a0efc-1349">Example</span></span>

```js
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
