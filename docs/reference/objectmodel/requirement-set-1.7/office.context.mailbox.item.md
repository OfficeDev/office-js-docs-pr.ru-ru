---
title: Office. Context. Mailbox. Item — набор требований 1,7
description: ''
ms.date: 11/06/2019
localization_priority: Normal
ms.openlocfilehash: 1c0948490c5c0b77252a8605b43f85dd529f2897
ms.sourcegitcommit: 08c0b9ff319c391922fa43d3c2e9783cf6b53b1b
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 11/08/2019
ms.locfileid: "38066216"
---
# <a name="item"></a><span data-ttu-id="4323f-102">item</span><span class="sxs-lookup"><span data-stu-id="4323f-102">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="4323f-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="4323f-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="4323f-p101">Пространство имен `item` используется для доступа к выбранному в данный момент сообщению, приглашению на собрание или описанию встречи. Вы можете определить тип пространства имен `item` с помощью свойства [itemType](#itemtype-officemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="4323f-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="4323f-106">Requirements</span><span class="sxs-lookup"><span data-stu-id="4323f-106">Requirements</span></span>

|<span data-ttu-id="4323f-107">Требование</span><span class="sxs-lookup"><span data-stu-id="4323f-107">Requirement</span></span>|<span data-ttu-id="4323f-108">Значение</span><span class="sxs-lookup"><span data-stu-id="4323f-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="4323f-109">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="4323f-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4323f-110">1.0</span><span class="sxs-lookup"><span data-stu-id="4323f-110">1.0</span></span>|
|[<span data-ttu-id="4323f-111">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4323f-111">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4323f-112">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="4323f-112">Restricted</span></span>|
|[<span data-ttu-id="4323f-113">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4323f-113">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="4323f-114">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="4323f-114">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="4323f-115">Элементы и методы</span><span class="sxs-lookup"><span data-stu-id="4323f-115">Members and methods</span></span>

| <span data-ttu-id="4323f-116">Элемент</span><span class="sxs-lookup"><span data-stu-id="4323f-116">Member</span></span> | <span data-ttu-id="4323f-117">Тип</span><span class="sxs-lookup"><span data-stu-id="4323f-117">Type</span></span> |
|--------|------|
| [<span data-ttu-id="4323f-118">attachments</span><span class="sxs-lookup"><span data-stu-id="4323f-118">attachments</span></span>](#attachments-arrayattachmentdetails) | <span data-ttu-id="4323f-119">Элемент</span><span class="sxs-lookup"><span data-stu-id="4323f-119">Member</span></span> |
| [<span data-ttu-id="4323f-120">bcc</span><span class="sxs-lookup"><span data-stu-id="4323f-120">bcc</span></span>](#bcc-recipients) | <span data-ttu-id="4323f-121">Элемент</span><span class="sxs-lookup"><span data-stu-id="4323f-121">Member</span></span> |
| [<span data-ttu-id="4323f-122">body</span><span class="sxs-lookup"><span data-stu-id="4323f-122">body</span></span>](#body-body) | <span data-ttu-id="4323f-123">Элемент</span><span class="sxs-lookup"><span data-stu-id="4323f-123">Member</span></span> |
| [<span data-ttu-id="4323f-124">cc</span><span class="sxs-lookup"><span data-stu-id="4323f-124">cc</span></span>](#cc-arrayemailaddressdetailsrecipients) | <span data-ttu-id="4323f-125">Элемент</span><span class="sxs-lookup"><span data-stu-id="4323f-125">Member</span></span> |
| [<span data-ttu-id="4323f-126">conversationId</span><span class="sxs-lookup"><span data-stu-id="4323f-126">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="4323f-127">Элемент</span><span class="sxs-lookup"><span data-stu-id="4323f-127">Member</span></span> |
| [<span data-ttu-id="4323f-128">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="4323f-128">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="4323f-129">Элемент</span><span class="sxs-lookup"><span data-stu-id="4323f-129">Member</span></span> |
| [<span data-ttu-id="4323f-130">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="4323f-130">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="4323f-131">Элемент</span><span class="sxs-lookup"><span data-stu-id="4323f-131">Member</span></span> |
| [<span data-ttu-id="4323f-132">end</span><span class="sxs-lookup"><span data-stu-id="4323f-132">end</span></span>](#end-datetime) | <span data-ttu-id="4323f-133">Элемент</span><span class="sxs-lookup"><span data-stu-id="4323f-133">Member</span></span> |
| [<span data-ttu-id="4323f-134">from</span><span class="sxs-lookup"><span data-stu-id="4323f-134">from</span></span>](#from-emailaddressdetailsfrom) | <span data-ttu-id="4323f-135">Элемент</span><span class="sxs-lookup"><span data-stu-id="4323f-135">Member</span></span> |
| [<span data-ttu-id="4323f-136">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="4323f-136">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="4323f-137">Элемент</span><span class="sxs-lookup"><span data-stu-id="4323f-137">Member</span></span> |
| [<span data-ttu-id="4323f-138">itemClass</span><span class="sxs-lookup"><span data-stu-id="4323f-138">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="4323f-139">Элемент</span><span class="sxs-lookup"><span data-stu-id="4323f-139">Member</span></span> |
| [<span data-ttu-id="4323f-140">itemId</span><span class="sxs-lookup"><span data-stu-id="4323f-140">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="4323f-141">Элемент</span><span class="sxs-lookup"><span data-stu-id="4323f-141">Member</span></span> |
| [<span data-ttu-id="4323f-142">itemType</span><span class="sxs-lookup"><span data-stu-id="4323f-142">itemType</span></span>](#itemtype-officemailboxenumsitemtype) | <span data-ttu-id="4323f-143">Элемент</span><span class="sxs-lookup"><span data-stu-id="4323f-143">Member</span></span> |
| [<span data-ttu-id="4323f-144">location</span><span class="sxs-lookup"><span data-stu-id="4323f-144">location</span></span>](#location-stringlocation) | <span data-ttu-id="4323f-145">Элемент</span><span class="sxs-lookup"><span data-stu-id="4323f-145">Member</span></span> |
| [<span data-ttu-id="4323f-146">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="4323f-146">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="4323f-147">Элемент</span><span class="sxs-lookup"><span data-stu-id="4323f-147">Member</span></span> |
| [<span data-ttu-id="4323f-148">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="4323f-148">notificationMessages</span></span>](#notificationmessages-notificationmessages) | <span data-ttu-id="4323f-149">Элемент</span><span class="sxs-lookup"><span data-stu-id="4323f-149">Member</span></span> |
| [<span data-ttu-id="4323f-150">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="4323f-150">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="4323f-151">Элемент</span><span class="sxs-lookup"><span data-stu-id="4323f-151">Member</span></span> |
| [<span data-ttu-id="4323f-152">organizer</span><span class="sxs-lookup"><span data-stu-id="4323f-152">organizer</span></span>](#organizer-emailaddressdetailsorganizer) | <span data-ttu-id="4323f-153">Элемент</span><span class="sxs-lookup"><span data-stu-id="4323f-153">Member</span></span> |
| [<span data-ttu-id="4323f-154">recurrence</span><span class="sxs-lookup"><span data-stu-id="4323f-154">recurrence</span></span>](#nullable-recurrence-recurrence) | <span data-ttu-id="4323f-155">Member</span><span class="sxs-lookup"><span data-stu-id="4323f-155">Member</span></span> |
| [<span data-ttu-id="4323f-156">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="4323f-156">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="4323f-157">Элемент</span><span class="sxs-lookup"><span data-stu-id="4323f-157">Member</span></span> |
| [<span data-ttu-id="4323f-158">sender</span><span class="sxs-lookup"><span data-stu-id="4323f-158">sender</span></span>](#sender-emailaddressdetails) | <span data-ttu-id="4323f-159">Элемент</span><span class="sxs-lookup"><span data-stu-id="4323f-159">Member</span></span> |
| [<span data-ttu-id="4323f-160">seriesId</span><span class="sxs-lookup"><span data-stu-id="4323f-160">seriesId</span></span>](#nullable-seriesid-string) | <span data-ttu-id="4323f-161">Элемент</span><span class="sxs-lookup"><span data-stu-id="4323f-161">Member</span></span> |
| [<span data-ttu-id="4323f-162">start</span><span class="sxs-lookup"><span data-stu-id="4323f-162">start</span></span>](#start-datetime) | <span data-ttu-id="4323f-163">Элемент</span><span class="sxs-lookup"><span data-stu-id="4323f-163">Member</span></span> |
| [<span data-ttu-id="4323f-164">subject</span><span class="sxs-lookup"><span data-stu-id="4323f-164">subject</span></span>](#subject-stringsubject) | <span data-ttu-id="4323f-165">Элемент</span><span class="sxs-lookup"><span data-stu-id="4323f-165">Member</span></span> |
| [<span data-ttu-id="4323f-166">to</span><span class="sxs-lookup"><span data-stu-id="4323f-166">to</span></span>](#to-arrayemailaddressdetailsrecipients) | <span data-ttu-id="4323f-167">Элемент</span><span class="sxs-lookup"><span data-stu-id="4323f-167">Member</span></span> |
| [<span data-ttu-id="4323f-168">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="4323f-168">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="4323f-169">Метод</span><span class="sxs-lookup"><span data-stu-id="4323f-169">Method</span></span> |
| [<span data-ttu-id="4323f-170">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="4323f-170">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="4323f-171">Метод</span><span class="sxs-lookup"><span data-stu-id="4323f-171">Method</span></span> |
| [<span data-ttu-id="4323f-172">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="4323f-172">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="4323f-173">Метод</span><span class="sxs-lookup"><span data-stu-id="4323f-173">Method</span></span> |
| [<span data-ttu-id="4323f-174">close</span><span class="sxs-lookup"><span data-stu-id="4323f-174">close</span></span>](#close) | <span data-ttu-id="4323f-175">Метод</span><span class="sxs-lookup"><span data-stu-id="4323f-175">Method</span></span> |
| [<span data-ttu-id="4323f-176">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="4323f-176">displayReplyAllForm</span></span>](#displayreplyallformformdata-callback) | <span data-ttu-id="4323f-177">Метод</span><span class="sxs-lookup"><span data-stu-id="4323f-177">Method</span></span> |
| [<span data-ttu-id="4323f-178">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="4323f-178">displayReplyForm</span></span>](#displayreplyformformdata-callback) | <span data-ttu-id="4323f-179">Метод</span><span class="sxs-lookup"><span data-stu-id="4323f-179">Method</span></span> |
| [<span data-ttu-id="4323f-180">getEntities</span><span class="sxs-lookup"><span data-stu-id="4323f-180">getEntities</span></span>](#getentities--entities) | <span data-ttu-id="4323f-181">Метод</span><span class="sxs-lookup"><span data-stu-id="4323f-181">Method</span></span> |
| [<span data-ttu-id="4323f-182">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="4323f-182">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="4323f-183">Метод</span><span class="sxs-lookup"><span data-stu-id="4323f-183">Method</span></span> |
| [<span data-ttu-id="4323f-184">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="4323f-184">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="4323f-185">Метод</span><span class="sxs-lookup"><span data-stu-id="4323f-185">Method</span></span> |
| [<span data-ttu-id="4323f-186">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="4323f-186">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="4323f-187">Метод</span><span class="sxs-lookup"><span data-stu-id="4323f-187">Method</span></span> |
| [<span data-ttu-id="4323f-188">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="4323f-188">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="4323f-189">Метод</span><span class="sxs-lookup"><span data-stu-id="4323f-189">Method</span></span> |
| [<span data-ttu-id="4323f-190">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="4323f-190">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="4323f-191">Метод</span><span class="sxs-lookup"><span data-stu-id="4323f-191">Method</span></span> |
| [<span data-ttu-id="4323f-192">жетселектедентитиес</span><span class="sxs-lookup"><span data-stu-id="4323f-192">getSelectedEntities</span></span>](#getselectedentities--entities) | <span data-ttu-id="4323f-193">Метод</span><span class="sxs-lookup"><span data-stu-id="4323f-193">Method</span></span> |
| [<span data-ttu-id="4323f-194">жетселектедрежексматчес</span><span class="sxs-lookup"><span data-stu-id="4323f-194">getSelectedRegExMatches</span></span>](#getselectedregexmatches--object) | <span data-ttu-id="4323f-195">Метод</span><span class="sxs-lookup"><span data-stu-id="4323f-195">Method</span></span> |
| [<span data-ttu-id="4323f-196">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="4323f-196">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="4323f-197">Метод</span><span class="sxs-lookup"><span data-stu-id="4323f-197">Method</span></span> |
| [<span data-ttu-id="4323f-198">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="4323f-198">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="4323f-199">Метод</span><span class="sxs-lookup"><span data-stu-id="4323f-199">Method</span></span> |
| [<span data-ttu-id="4323f-200">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="4323f-200">removeHandlerAsync</span></span>](#removehandlerasynceventtype-options-callback) | <span data-ttu-id="4323f-201">Метод</span><span class="sxs-lookup"><span data-stu-id="4323f-201">Method</span></span> |
| [<span data-ttu-id="4323f-202">saveAsync</span><span class="sxs-lookup"><span data-stu-id="4323f-202">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="4323f-203">Метод</span><span class="sxs-lookup"><span data-stu-id="4323f-203">Method</span></span> |
| [<span data-ttu-id="4323f-204">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="4323f-204">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="4323f-205">Метод</span><span class="sxs-lookup"><span data-stu-id="4323f-205">Method</span></span> |

### <a name="example"></a><span data-ttu-id="4323f-206">Пример</span><span class="sxs-lookup"><span data-stu-id="4323f-206">Example</span></span>

<span data-ttu-id="4323f-207">В примере кода JavaScript, приведенном ниже, показано, как получить доступ к свойству `subject` текущего элемента в Outlook.</span><span class="sxs-lookup"><span data-stu-id="4323f-207">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="4323f-208">Members</span><span class="sxs-lookup"><span data-stu-id="4323f-208">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetailsviewoutlook-js-17"></a><span data-ttu-id="4323f-209">attachments: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.7)></span><span class="sxs-lookup"><span data-stu-id="4323f-209">attachments: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.7)></span></span>

<span data-ttu-id="4323f-p102">Получает массив вложений для элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="4323f-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="4323f-212">Outlook блокирует определенные типы файлов из-за потенциальных проблем с безопасностью, поэтому они не возвращаются.</span><span class="sxs-lookup"><span data-stu-id="4323f-212">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="4323f-213">Дополнительные сведения см. в статье [Блокировка вложений в Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="4323f-213">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="4323f-214">Тип</span><span class="sxs-lookup"><span data-stu-id="4323f-214">Type</span></span>

*   <span data-ttu-id="4323f-215">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.7)></span><span class="sxs-lookup"><span data-stu-id="4323f-215">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.7)></span></span>

##### <a name="requirements"></a><span data-ttu-id="4323f-216">Requirements</span><span class="sxs-lookup"><span data-stu-id="4323f-216">Requirements</span></span>

|<span data-ttu-id="4323f-217">Требование</span><span class="sxs-lookup"><span data-stu-id="4323f-217">Requirement</span></span>|<span data-ttu-id="4323f-218">Значение</span><span class="sxs-lookup"><span data-stu-id="4323f-218">Value</span></span>|
|---|---|
|[<span data-ttu-id="4323f-219">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="4323f-219">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4323f-220">1.0</span><span class="sxs-lookup"><span data-stu-id="4323f-220">1.0</span></span>|
|[<span data-ttu-id="4323f-221">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4323f-221">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4323f-222">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4323f-222">ReadItem</span></span>|
|[<span data-ttu-id="4323f-223">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4323f-223">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="4323f-224">Чтение</span><span class="sxs-lookup"><span data-stu-id="4323f-224">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4323f-225">Пример</span><span class="sxs-lookup"><span data-stu-id="4323f-225">Example</span></span>

<span data-ttu-id="4323f-226">С помощью приведенного ниже кода можно создать HTML-строку с подробными сведениями обо всех вложениях для текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="4323f-226">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

#### <a name="bcc-recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-17"></a><span data-ttu-id="4323f-227">bcc: [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="4323f-227">bcc: [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span></span>

<span data-ttu-id="4323f-228">Получает объект, который предоставляет методы для получения или обновления получателей скрытой копии сообщения.</span><span class="sxs-lookup"><span data-stu-id="4323f-228">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="4323f-229">Только в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="4323f-229">Compose mode only.</span></span>

<span data-ttu-id="4323f-230">Коллекция может включать не более 100 элементов по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="4323f-230">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="4323f-231">Однако для компьютеров с Windows и Mac OS действуют указанные ниже ограничения.</span><span class="sxs-lookup"><span data-stu-id="4323f-231">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="4323f-232">Максимальное количество элементов — 500.</span><span class="sxs-lookup"><span data-stu-id="4323f-232">Get 500 members maximum.</span></span>
- <span data-ttu-id="4323f-233">Установить ограничение количества элементов на вызов — не более 100, общего количества — не более 500.</span><span class="sxs-lookup"><span data-stu-id="4323f-233">Set a maximum of 100 members per call, up to 500 members total.</span></span>

##### <a name="type"></a><span data-ttu-id="4323f-234">Тип</span><span class="sxs-lookup"><span data-stu-id="4323f-234">Type</span></span>

*   [<span data-ttu-id="4323f-235">Получатели</span><span class="sxs-lookup"><span data-stu-id="4323f-235">Recipients</span></span>](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)

##### <a name="requirements"></a><span data-ttu-id="4323f-236">Requirements</span><span class="sxs-lookup"><span data-stu-id="4323f-236">Requirements</span></span>

|<span data-ttu-id="4323f-237">Требование</span><span class="sxs-lookup"><span data-stu-id="4323f-237">Requirement</span></span>|<span data-ttu-id="4323f-238">Значение</span><span class="sxs-lookup"><span data-stu-id="4323f-238">Value</span></span>|
|---|---|
|[<span data-ttu-id="4323f-239">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="4323f-239">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4323f-240">1.1</span><span class="sxs-lookup"><span data-stu-id="4323f-240">1.1</span></span>|
|[<span data-ttu-id="4323f-241">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4323f-241">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4323f-242">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4323f-242">ReadItem</span></span>|
|[<span data-ttu-id="4323f-243">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4323f-243">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="4323f-244">Создание</span><span class="sxs-lookup"><span data-stu-id="4323f-244">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="4323f-245">Пример</span><span class="sxs-lookup"><span data-stu-id="4323f-245">Example</span></span>

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

#### <a name="body-bodyjavascriptapioutlookofficebodyviewoutlook-js-17"></a><span data-ttu-id="4323f-246">body: [Body](/javascript/api/outlook/office.body?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="4323f-246">body: [Body](/javascript/api/outlook/office.body?view=outlook-js-1.7)</span></span>

<span data-ttu-id="4323f-247">Получает объект, предоставляющий методы для работы с основным текстом элемента.</span><span class="sxs-lookup"><span data-stu-id="4323f-247">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="4323f-248">Тип</span><span class="sxs-lookup"><span data-stu-id="4323f-248">Type</span></span>

*   [<span data-ttu-id="4323f-249">Body</span><span class="sxs-lookup"><span data-stu-id="4323f-249">Body</span></span>](/javascript/api/outlook/office.body?view=outlook-js-1.7)

##### <a name="requirements"></a><span data-ttu-id="4323f-250">Требования</span><span class="sxs-lookup"><span data-stu-id="4323f-250">Requirements</span></span>

|<span data-ttu-id="4323f-251">Требование</span><span class="sxs-lookup"><span data-stu-id="4323f-251">Requirement</span></span>|<span data-ttu-id="4323f-252">Значение</span><span class="sxs-lookup"><span data-stu-id="4323f-252">Value</span></span>|
|---|---|
|[<span data-ttu-id="4323f-253">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="4323f-253">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4323f-254">1.1</span><span class="sxs-lookup"><span data-stu-id="4323f-254">1.1</span></span>|
|[<span data-ttu-id="4323f-255">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4323f-255">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4323f-256">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4323f-256">ReadItem</span></span>|
|[<span data-ttu-id="4323f-257">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4323f-257">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="4323f-258">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="4323f-258">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4323f-259">Пример</span><span class="sxs-lookup"><span data-stu-id="4323f-259">Example</span></span>

<span data-ttu-id="4323f-260">В этом примере возвращается текст сообщения в формате обычного текста.</span><span class="sxs-lookup"><span data-stu-id="4323f-260">This example gets the body of the message in plain text.</span></span>

```js
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="4323f-261">Ниже приведен пример итогового параметра, переданного функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="4323f-261">The following is an example of the result parameter passed to the callback function.</span></span>

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

#### <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-17recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-17"></a><span data-ttu-id="4323f-262">cc: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="4323f-262">cc: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span></span>

<span data-ttu-id="4323f-263">Предоставляет доступ к получателям копии сообщения.</span><span class="sxs-lookup"><span data-stu-id="4323f-263">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="4323f-264">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="4323f-264">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="4323f-265">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="4323f-265">Read mode</span></span>

<span data-ttu-id="4323f-266">Свойство `cc` возвращает массив, который содержит объект `EmailAddressDetails` для каждого получателя, указанного в строке **Копия** сообщения.</span><span class="sxs-lookup"><span data-stu-id="4323f-266">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message.</span></span> <span data-ttu-id="4323f-267">Коллекция может включать не более 100 элементов по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="4323f-267">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="4323f-268">Однако на компьютерах с Windows и Mac OS может быть до 500 элементов.</span><span class="sxs-lookup"><span data-stu-id="4323f-268">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="4323f-269">Режим создания</span><span class="sxs-lookup"><span data-stu-id="4323f-269">Compose mode</span></span>

<span data-ttu-id="4323f-270">Свойство `cc` возвращает объект `Recipients`, предоставляющий методы для получения или обновления получателей, которые указаны в строке **Копия** сообщения.</span><span class="sxs-lookup"><span data-stu-id="4323f-270">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span> <span data-ttu-id="4323f-271">Коллекция может включать не более 100 элементов по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="4323f-271">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="4323f-272">Однако для компьютеров с Windows и Mac OS действуют указанные ниже ограничения.</span><span class="sxs-lookup"><span data-stu-id="4323f-272">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="4323f-273">Максимальное количество элементов — 500.</span><span class="sxs-lookup"><span data-stu-id="4323f-273">Get 500 members maximum.</span></span>
- <span data-ttu-id="4323f-274">Установить ограничение количества элементов на вызов — не более 100, общего количества — не более 500.</span><span class="sxs-lookup"><span data-stu-id="4323f-274">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="4323f-275">Тип</span><span class="sxs-lookup"><span data-stu-id="4323f-275">Type</span></span>

*   <span data-ttu-id="4323f-276">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="4323f-276">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span></span>

##### <a name="requirements"></a><span data-ttu-id="4323f-277">Requirements</span><span class="sxs-lookup"><span data-stu-id="4323f-277">Requirements</span></span>

|<span data-ttu-id="4323f-278">Требование</span><span class="sxs-lookup"><span data-stu-id="4323f-278">Requirement</span></span>|<span data-ttu-id="4323f-279">Значение</span><span class="sxs-lookup"><span data-stu-id="4323f-279">Value</span></span>|
|---|---|
|[<span data-ttu-id="4323f-280">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="4323f-280">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4323f-281">1.0</span><span class="sxs-lookup"><span data-stu-id="4323f-281">1.0</span></span>|
|[<span data-ttu-id="4323f-282">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4323f-282">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4323f-283">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4323f-283">ReadItem</span></span>|
|[<span data-ttu-id="4323f-284">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4323f-284">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="4323f-285">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="4323f-285">Compose or Read</span></span>|

<br>

---
---

#### <a name="nullable-conversationid-string"></a><span data-ttu-id="4323f-286">(nullable) conversationId: String</span><span class="sxs-lookup"><span data-stu-id="4323f-286">(nullable) conversationId: String</span></span>

<span data-ttu-id="4323f-287">Получает идентификатор разговора по электронной почте, содержащего конкретное сообщение.</span><span class="sxs-lookup"><span data-stu-id="4323f-287">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="4323f-p109">Вы можете получить целочисленное значение этого свойства, если ваше почтовое приложение активируется в формах просмотра или формах создания ответов. Если пользователь изменит тему ответа, после его отправки идентификатор беседы будет изменен, и полученное ранее значение будет недействительным.</span><span class="sxs-lookup"><span data-stu-id="4323f-p109">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="4323f-p110">Это свойство имеет значение NULL для нового элемента в форме создания. Свойство `conversationId` вернет значение, если пользователь задаст тему и сохранит элемент.</span><span class="sxs-lookup"><span data-stu-id="4323f-p110">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="4323f-292">Тип</span><span class="sxs-lookup"><span data-stu-id="4323f-292">Type</span></span>

*   <span data-ttu-id="4323f-293">String</span><span class="sxs-lookup"><span data-stu-id="4323f-293">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="4323f-294">Requirements</span><span class="sxs-lookup"><span data-stu-id="4323f-294">Requirements</span></span>

|<span data-ttu-id="4323f-295">Требование</span><span class="sxs-lookup"><span data-stu-id="4323f-295">Requirement</span></span>|<span data-ttu-id="4323f-296">Значение</span><span class="sxs-lookup"><span data-stu-id="4323f-296">Value</span></span>|
|---|---|
|[<span data-ttu-id="4323f-297">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="4323f-297">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4323f-298">1.0</span><span class="sxs-lookup"><span data-stu-id="4323f-298">1.0</span></span>|
|[<span data-ttu-id="4323f-299">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4323f-299">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4323f-300">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4323f-300">ReadItem</span></span>|
|[<span data-ttu-id="4323f-301">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4323f-301">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="4323f-302">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="4323f-302">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4323f-303">Пример</span><span class="sxs-lookup"><span data-stu-id="4323f-303">Example</span></span>

```js
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

<br>

---
---

#### <a name="datetimecreated-date"></a><span data-ttu-id="4323f-304">dateTimeCreated: Date</span><span class="sxs-lookup"><span data-stu-id="4323f-304">dateTimeCreated: Date</span></span>

<span data-ttu-id="4323f-p111">Получает дату и время создания элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="4323f-p111">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="4323f-307">Тип</span><span class="sxs-lookup"><span data-stu-id="4323f-307">Type</span></span>

*   <span data-ttu-id="4323f-308">Дата</span><span class="sxs-lookup"><span data-stu-id="4323f-308">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="4323f-309">Requirements</span><span class="sxs-lookup"><span data-stu-id="4323f-309">Requirements</span></span>

|<span data-ttu-id="4323f-310">Требование</span><span class="sxs-lookup"><span data-stu-id="4323f-310">Requirement</span></span>|<span data-ttu-id="4323f-311">Значение</span><span class="sxs-lookup"><span data-stu-id="4323f-311">Value</span></span>|
|---|---|
|[<span data-ttu-id="4323f-312">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="4323f-312">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4323f-313">1.0</span><span class="sxs-lookup"><span data-stu-id="4323f-313">1.0</span></span>|
|[<span data-ttu-id="4323f-314">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4323f-314">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4323f-315">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4323f-315">ReadItem</span></span>|
|[<span data-ttu-id="4323f-316">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4323f-316">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="4323f-317">Чтение</span><span class="sxs-lookup"><span data-stu-id="4323f-317">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4323f-318">Пример</span><span class="sxs-lookup"><span data-stu-id="4323f-318">Example</span></span>

```js
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

<br>

---
---

#### <a name="datetimemodified-date"></a><span data-ttu-id="4323f-319">dateTimeModified: Date</span><span class="sxs-lookup"><span data-stu-id="4323f-319">dateTimeModified: Date</span></span>

<span data-ttu-id="4323f-p112">Получает дату и время последнего изменения элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="4323f-p112">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="4323f-322">Этот элемент не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="4323f-322">This member is not supported in Outlook on iOS or Android.</span></span>

##### <a name="type"></a><span data-ttu-id="4323f-323">Тип</span><span class="sxs-lookup"><span data-stu-id="4323f-323">Type</span></span>

*   <span data-ttu-id="4323f-324">Дата</span><span class="sxs-lookup"><span data-stu-id="4323f-324">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="4323f-325">Requirements</span><span class="sxs-lookup"><span data-stu-id="4323f-325">Requirements</span></span>

|<span data-ttu-id="4323f-326">Требование</span><span class="sxs-lookup"><span data-stu-id="4323f-326">Requirement</span></span>|<span data-ttu-id="4323f-327">Значение</span><span class="sxs-lookup"><span data-stu-id="4323f-327">Value</span></span>|
|---|---|
|[<span data-ttu-id="4323f-328">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="4323f-328">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4323f-329">1.0</span><span class="sxs-lookup"><span data-stu-id="4323f-329">1.0</span></span>|
|[<span data-ttu-id="4323f-330">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4323f-330">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4323f-331">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4323f-331">ReadItem</span></span>|
|[<span data-ttu-id="4323f-332">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4323f-332">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="4323f-333">Чтение</span><span class="sxs-lookup"><span data-stu-id="4323f-333">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4323f-334">Пример</span><span class="sxs-lookup"><span data-stu-id="4323f-334">Example</span></span>

```js
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

<br>

---
---

#### <a name="end-datetimejavascriptapioutlookofficetimeviewoutlook-js-17"></a><span data-ttu-id="4323f-335">end: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="4323f-335">end: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.7)</span></span>

<span data-ttu-id="4323f-336">Получает или задает дату и время окончания встречи.</span><span class="sxs-lookup"><span data-stu-id="4323f-336">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="4323f-p113">Свойство `end` представлено в виде значения даты и времени в формате UTC. Преобразовать значение свойства end в местные значения даты и времени клиента можно с помощью метода [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime).</span><span class="sxs-lookup"><span data-stu-id="4323f-p113">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="4323f-339">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="4323f-339">Read mode</span></span>

<span data-ttu-id="4323f-340">Свойство `end` возвращает объект `Date`.</span><span class="sxs-lookup"><span data-stu-id="4323f-340">The `end` property returns a `Date` object.</span></span>

```js
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="4323f-341">Режим создания</span><span class="sxs-lookup"><span data-stu-id="4323f-341">Compose mode</span></span>

<span data-ttu-id="4323f-342">Свойство `end` возвращает объект `Time`.</span><span class="sxs-lookup"><span data-stu-id="4323f-342">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="4323f-343">Если вы задаете время окончания с помощью метода [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.7#setasync-datetime--options--callback-), необходимо использовать метод [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) для преобразования местного времени на клиенте в формат UTC для сервера.</span><span class="sxs-lookup"><span data-stu-id="4323f-343">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.7#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="4323f-344">В примере ниже показано, как с помощью метода [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.7#setasync-datetime--options--callback-) объекта `Time` задать время окончания встречи.</span><span class="sxs-lookup"><span data-stu-id="4323f-344">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.7#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="4323f-345">Тип</span><span class="sxs-lookup"><span data-stu-id="4323f-345">Type</span></span>

*   <span data-ttu-id="4323f-346">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="4323f-346">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.7)</span></span>

##### <a name="requirements"></a><span data-ttu-id="4323f-347">Требования</span><span class="sxs-lookup"><span data-stu-id="4323f-347">Requirements</span></span>

|<span data-ttu-id="4323f-348">Требование</span><span class="sxs-lookup"><span data-stu-id="4323f-348">Requirement</span></span>|<span data-ttu-id="4323f-349">Значение</span><span class="sxs-lookup"><span data-stu-id="4323f-349">Value</span></span>|
|---|---|
|[<span data-ttu-id="4323f-350">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="4323f-350">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4323f-351">1.0</span><span class="sxs-lookup"><span data-stu-id="4323f-351">1.0</span></span>|
|[<span data-ttu-id="4323f-352">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4323f-352">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4323f-353">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4323f-353">ReadItem</span></span>|
|[<span data-ttu-id="4323f-354">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4323f-354">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="4323f-355">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="4323f-355">Compose or Read</span></span>|

<br>

---
---

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-17fromjavascriptapioutlookofficefromviewoutlook-js-17"></a><span data-ttu-id="4323f-356">от: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)|[](/javascript/api/outlook/office.from?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="4323f-356">from: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)|[From](/javascript/api/outlook/office.from?view=outlook-js-1.7)</span></span>

<span data-ttu-id="4323f-357">Получает электронный адрес отправителя сообщения.</span><span class="sxs-lookup"><span data-stu-id="4323f-357">Gets the email address of the sender of a message.</span></span>

<span data-ttu-id="4323f-p114">Свойства `from` и [`sender`](#sender-emailaddressdetails) представляют одно лицо, если сообщение не отправлено представителем. В противном случае свойство `from` представляет лицо, делегировавшее полномочия, а свойство sender — представителя.</span><span class="sxs-lookup"><span data-stu-id="4323f-p114">The `from` and [`sender`](#sender-emailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="4323f-360">Свойству `recipientType`, принадлежащему объекту `EmailAddressDetails` в свойстве `from`, задано значение `undefined`.</span><span class="sxs-lookup"><span data-stu-id="4323f-360">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="4323f-361">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="4323f-361">Read mode</span></span>

<span data-ttu-id="4323f-362">`from` Свойство возвращает `EmailAddressDetails` объект.</span><span class="sxs-lookup"><span data-stu-id="4323f-362">The `from` property returns an `EmailAddressDetails` object.</span></span>

```js
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

##### <a name="compose-mode"></a><span data-ttu-id="4323f-363">Режим создания</span><span class="sxs-lookup"><span data-stu-id="4323f-363">Compose mode</span></span>

<span data-ttu-id="4323f-364">`from` Свойство возвращает `From` объект, который предоставляет метод для получения значения From.</span><span class="sxs-lookup"><span data-stu-id="4323f-364">The `from` property returns a `From` object that provides a method to get the from value.</span></span>

```js
Office.context.mailbox.item.from.getAsync(callback);

function callback(asyncResult) {
  var from = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="4323f-365">Тип</span><span class="sxs-lookup"><span data-stu-id="4323f-365">Type</span></span>

*   <span data-ttu-id="4323f-366">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7) | [из](/javascript/api/outlook/office.from?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="4323f-366">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7) | [From](/javascript/api/outlook/office.from?view=outlook-js-1.7)</span></span>

##### <a name="requirements"></a><span data-ttu-id="4323f-367">Requirements</span><span class="sxs-lookup"><span data-stu-id="4323f-367">Requirements</span></span>

|<span data-ttu-id="4323f-368">Требование</span><span class="sxs-lookup"><span data-stu-id="4323f-368">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="4323f-369">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="4323f-369">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4323f-370">1.0</span><span class="sxs-lookup"><span data-stu-id="4323f-370">1.0</span></span>|<span data-ttu-id="4323f-371">1.7</span><span class="sxs-lookup"><span data-stu-id="4323f-371">1.7</span></span>|
|[<span data-ttu-id="4323f-372">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4323f-372">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4323f-373">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4323f-373">ReadItem</span></span>|<span data-ttu-id="4323f-374">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="4323f-374">ReadWriteItem</span></span>|
|[<span data-ttu-id="4323f-375">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4323f-375">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="4323f-376">Чтение</span><span class="sxs-lookup"><span data-stu-id="4323f-376">Read</span></span>|<span data-ttu-id="4323f-377">Создание</span><span class="sxs-lookup"><span data-stu-id="4323f-377">Compose</span></span>|

<br>

---
---

#### <a name="internetmessageid-string"></a><span data-ttu-id="4323f-378">internetMessageId: String</span><span class="sxs-lookup"><span data-stu-id="4323f-378">internetMessageId: String</span></span>

<span data-ttu-id="4323f-p115">Получает идентификатор интернет-сообщения для электронного сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="4323f-p115">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="4323f-381">Тип</span><span class="sxs-lookup"><span data-stu-id="4323f-381">Type</span></span>

*   <span data-ttu-id="4323f-382">String</span><span class="sxs-lookup"><span data-stu-id="4323f-382">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="4323f-383">Requirements</span><span class="sxs-lookup"><span data-stu-id="4323f-383">Requirements</span></span>

|<span data-ttu-id="4323f-384">Требование</span><span class="sxs-lookup"><span data-stu-id="4323f-384">Requirement</span></span>|<span data-ttu-id="4323f-385">Значение</span><span class="sxs-lookup"><span data-stu-id="4323f-385">Value</span></span>|
|---|---|
|[<span data-ttu-id="4323f-386">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="4323f-386">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4323f-387">1.0</span><span class="sxs-lookup"><span data-stu-id="4323f-387">1.0</span></span>|
|[<span data-ttu-id="4323f-388">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4323f-388">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4323f-389">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4323f-389">ReadItem</span></span>|
|[<span data-ttu-id="4323f-390">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4323f-390">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="4323f-391">Чтение</span><span class="sxs-lookup"><span data-stu-id="4323f-391">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4323f-392">Пример</span><span class="sxs-lookup"><span data-stu-id="4323f-392">Example</span></span>

```js
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

<br>

---
---

#### <a name="itemclass-string"></a><span data-ttu-id="4323f-393">itemClass: String</span><span class="sxs-lookup"><span data-stu-id="4323f-393">itemClass: String</span></span>

<span data-ttu-id="4323f-p116">Получает класс элемента веб-служб Exchange для выбранного элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="4323f-p116">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="4323f-p117">Свойство `itemClass` указывает класс сообщения выбранного элемента. Ниже приводятся классы сообщения по умолчанию для элемента сообщения или встречи.</span><span class="sxs-lookup"><span data-stu-id="4323f-p117">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

|<span data-ttu-id="4323f-398">Тип</span><span class="sxs-lookup"><span data-stu-id="4323f-398">Type</span></span>|<span data-ttu-id="4323f-399">Описание</span><span class="sxs-lookup"><span data-stu-id="4323f-399">Description</span></span>|<span data-ttu-id="4323f-400">Класс элемента</span><span class="sxs-lookup"><span data-stu-id="4323f-400">item class</span></span>|
|---|---|---|
|<span data-ttu-id="4323f-401">Элементы встречи</span><span class="sxs-lookup"><span data-stu-id="4323f-401">Appointment items</span></span>|<span data-ttu-id="4323f-402">Это элементы календаря для класса элемента `IPM.Appointment` или `IPM.Appointment.Occurrence`.</span><span class="sxs-lookup"><span data-stu-id="4323f-402">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span>|`IPM.Appointment`<br />`IPM.Appointment.Occurrence`|
|<span data-ttu-id="4323f-403">Элементы сообщения</span><span class="sxs-lookup"><span data-stu-id="4323f-403">Message items</span></span>|<span data-ttu-id="4323f-404">Сюда входят электронные сообщения, для которых по умолчанию задан класс сообщения `IPM.Note`, а также приглашения на собрания, ответы на них и уведомления об их отмене, использующие `IPM.Schedule.Meeting` в качестве базового класса сообщения.</span><span class="sxs-lookup"><span data-stu-id="4323f-404">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span>|`IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled`|

<span data-ttu-id="4323f-405">Можно создавать настраиваемые классы сообщения, расширяющие классы сообщения по умолчанию, например настраиваемый класс сообщения о встрече `IPM.Appointment.Contoso`.</span><span class="sxs-lookup"><span data-stu-id="4323f-405">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="4323f-406">Тип</span><span class="sxs-lookup"><span data-stu-id="4323f-406">Type</span></span>

*   <span data-ttu-id="4323f-407">String</span><span class="sxs-lookup"><span data-stu-id="4323f-407">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="4323f-408">Requirements</span><span class="sxs-lookup"><span data-stu-id="4323f-408">Requirements</span></span>

|<span data-ttu-id="4323f-409">Требование</span><span class="sxs-lookup"><span data-stu-id="4323f-409">Requirement</span></span>|<span data-ttu-id="4323f-410">Значение</span><span class="sxs-lookup"><span data-stu-id="4323f-410">Value</span></span>|
|---|---|
|[<span data-ttu-id="4323f-411">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="4323f-411">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4323f-412">1.0</span><span class="sxs-lookup"><span data-stu-id="4323f-412">1.0</span></span>|
|[<span data-ttu-id="4323f-413">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4323f-413">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4323f-414">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4323f-414">ReadItem</span></span>|
|[<span data-ttu-id="4323f-415">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4323f-415">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="4323f-416">Чтение</span><span class="sxs-lookup"><span data-stu-id="4323f-416">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4323f-417">Пример</span><span class="sxs-lookup"><span data-stu-id="4323f-417">Example</span></span>

```js
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

<br>

---
---

#### <a name="nullable-itemid-string"></a><span data-ttu-id="4323f-418">(nullable) itemId: String</span><span class="sxs-lookup"><span data-stu-id="4323f-418">(nullable) itemId: String</span></span>

<span data-ttu-id="4323f-p118">Получает [идентификатор элемента веб-служб Exchange](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange) для текущего элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="4323f-p118">Gets the [Exchange Web Services item identifier](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange) for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="4323f-421">Идентификатор, возвращаемый свойством `itemId`, совпадает с [идентификатором элемента веб-служб Exchange](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange).</span><span class="sxs-lookup"><span data-stu-id="4323f-421">The identifier returned by the `itemId` property is the same as the [Exchange Web Services item identifier](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange).</span></span> <span data-ttu-id="4323f-422">Свойство `itemId` не совпадает с идентификатором записи Outlook, а также идентификатором, который используется REST API Outlook.</span><span class="sxs-lookup"><span data-stu-id="4323f-422">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="4323f-423">Прежде чем совершать вызовы REST API, используя это значение, его необходимо преобразовать с помощью [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="4323f-423">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="4323f-424">Дополнительные сведения см. в статье [Использование REST API Outlook из надстройки Outlook](/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="4323f-424">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="4323f-p120">Свойство `itemId` недоступно в режиме создания. Если требуется идентификатор элемента, с помощью метода [`saveAsync`](#saveasyncoptions-callback) можно сохранить элемент в хранилище. При этом в параметре [`AsyncResult.value`](/javascript/api/office/office.asyncresult) функции обратного вызова возвращается идентификатор элемента.</span><span class="sxs-lookup"><span data-stu-id="4323f-p120">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="4323f-427">Тип</span><span class="sxs-lookup"><span data-stu-id="4323f-427">Type</span></span>

*   <span data-ttu-id="4323f-428">String</span><span class="sxs-lookup"><span data-stu-id="4323f-428">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="4323f-429">Requirements</span><span class="sxs-lookup"><span data-stu-id="4323f-429">Requirements</span></span>

|<span data-ttu-id="4323f-430">Требование</span><span class="sxs-lookup"><span data-stu-id="4323f-430">Requirement</span></span>|<span data-ttu-id="4323f-431">Значение</span><span class="sxs-lookup"><span data-stu-id="4323f-431">Value</span></span>|
|---|---|
|[<span data-ttu-id="4323f-432">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="4323f-432">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4323f-433">1.0</span><span class="sxs-lookup"><span data-stu-id="4323f-433">1.0</span></span>|
|[<span data-ttu-id="4323f-434">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4323f-434">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4323f-435">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4323f-435">ReadItem</span></span>|
|[<span data-ttu-id="4323f-436">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4323f-436">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="4323f-437">Чтение</span><span class="sxs-lookup"><span data-stu-id="4323f-437">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4323f-438">Пример</span><span class="sxs-lookup"><span data-stu-id="4323f-438">Example</span></span>

<span data-ttu-id="4323f-p121">Указанный ниже код проверяет наличие идентификатора элемента. Если свойство `itemId` возвращает значение `null` или `undefined`, элемент будет сохранен в хранилище, а из асинхронного результата будет получен идентификатор элемента.</span><span class="sxs-lookup"><span data-stu-id="4323f-p121">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

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

#### <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtypeviewoutlook-js-17"></a><span data-ttu-id="4323f-441">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="4323f-441">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.7)</span></span>

<span data-ttu-id="4323f-442">Получает тип элемента, который представляет экземпляр.</span><span class="sxs-lookup"><span data-stu-id="4323f-442">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="4323f-443">Свойство `itemType` возвращает одно из значений перечисления `ItemType`, которое указывает, является ли экземпляр объекта `item` сообщением или встречей.</span><span class="sxs-lookup"><span data-stu-id="4323f-443">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="4323f-444">Тип</span><span class="sxs-lookup"><span data-stu-id="4323f-444">Type</span></span>

*   [<span data-ttu-id="4323f-445">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="4323f-445">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.7)

##### <a name="requirements"></a><span data-ttu-id="4323f-446">Requirements</span><span class="sxs-lookup"><span data-stu-id="4323f-446">Requirements</span></span>

|<span data-ttu-id="4323f-447">Требование</span><span class="sxs-lookup"><span data-stu-id="4323f-447">Requirement</span></span>|<span data-ttu-id="4323f-448">Значение</span><span class="sxs-lookup"><span data-stu-id="4323f-448">Value</span></span>|
|---|---|
|[<span data-ttu-id="4323f-449">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="4323f-449">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4323f-450">1.0</span><span class="sxs-lookup"><span data-stu-id="4323f-450">1.0</span></span>|
|[<span data-ttu-id="4323f-451">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4323f-451">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4323f-452">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4323f-452">ReadItem</span></span>|
|[<span data-ttu-id="4323f-453">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4323f-453">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="4323f-454">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="4323f-454">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4323f-455">Пример</span><span class="sxs-lookup"><span data-stu-id="4323f-455">Example</span></span>

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

#### <a name="location-stringlocationjavascriptapioutlookofficelocationviewoutlook-js-17"></a><span data-ttu-id="4323f-456">location: String|[Location](/javascript/api/outlook/office.location?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="4323f-456">location: String|[Location](/javascript/api/outlook/office.location?view=outlook-js-1.7)</span></span>

<span data-ttu-id="4323f-457">Получает или задает место встречи.</span><span class="sxs-lookup"><span data-stu-id="4323f-457">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="4323f-458">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="4323f-458">Read mode</span></span>

<span data-ttu-id="4323f-459">Свойство `location` возвращает строку, содержащую сведения о месте встречи.</span><span class="sxs-lookup"><span data-stu-id="4323f-459">The `location` property returns a string that contains the location of the appointment.</span></span>

```js
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="4323f-460">Режим создания</span><span class="sxs-lookup"><span data-stu-id="4323f-460">Compose mode</span></span>

<span data-ttu-id="4323f-461">Свойство `location` возвращает объект `Location`, предоставляющий методы, которые используются для получения и задания места встречи.</span><span class="sxs-lookup"><span data-stu-id="4323f-461">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```js
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="4323f-462">Тип</span><span class="sxs-lookup"><span data-stu-id="4323f-462">Type</span></span>

*   <span data-ttu-id="4323f-463">String | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="4323f-463">String | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.7)</span></span>

##### <a name="requirements"></a><span data-ttu-id="4323f-464">Requirements</span><span class="sxs-lookup"><span data-stu-id="4323f-464">Requirements</span></span>

|<span data-ttu-id="4323f-465">Требование</span><span class="sxs-lookup"><span data-stu-id="4323f-465">Requirement</span></span>|<span data-ttu-id="4323f-466">Значение</span><span class="sxs-lookup"><span data-stu-id="4323f-466">Value</span></span>|
|---|---|
|[<span data-ttu-id="4323f-467">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="4323f-467">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4323f-468">1.0</span><span class="sxs-lookup"><span data-stu-id="4323f-468">1.0</span></span>|
|[<span data-ttu-id="4323f-469">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4323f-469">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4323f-470">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4323f-470">ReadItem</span></span>|
|[<span data-ttu-id="4323f-471">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4323f-471">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="4323f-472">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="4323f-472">Compose or Read</span></span>|

<br>

---
---

#### <a name="normalizedsubject-string"></a><span data-ttu-id="4323f-473">normalizedSubject: String</span><span class="sxs-lookup"><span data-stu-id="4323f-473">normalizedSubject: String</span></span>

<span data-ttu-id="4323f-p122">Получает тему элемента со всеми удаленными префиксами (включая `RE:` и `FWD:`). Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="4323f-p122">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="4323f-p123">Свойство normalizedSubject получает тему элемента со стандартными префиксами (такими как `RE:` и `FW:`), добавляемыми почтовыми программами. Для получения темы элемента с неизмененными префиксами используйте свойство [`subject`](#subject-stringsubject).</span><span class="sxs-lookup"><span data-stu-id="4323f-p123">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="4323f-478">Тип</span><span class="sxs-lookup"><span data-stu-id="4323f-478">Type</span></span>

*   <span data-ttu-id="4323f-479">String</span><span class="sxs-lookup"><span data-stu-id="4323f-479">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="4323f-480">Requirements</span><span class="sxs-lookup"><span data-stu-id="4323f-480">Requirements</span></span>

|<span data-ttu-id="4323f-481">Требование</span><span class="sxs-lookup"><span data-stu-id="4323f-481">Requirement</span></span>|<span data-ttu-id="4323f-482">Значение</span><span class="sxs-lookup"><span data-stu-id="4323f-482">Value</span></span>|
|---|---|
|[<span data-ttu-id="4323f-483">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="4323f-483">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4323f-484">1.0</span><span class="sxs-lookup"><span data-stu-id="4323f-484">1.0</span></span>|
|[<span data-ttu-id="4323f-485">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4323f-485">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4323f-486">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4323f-486">ReadItem</span></span>|
|[<span data-ttu-id="4323f-487">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4323f-487">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="4323f-488">Чтение</span><span class="sxs-lookup"><span data-stu-id="4323f-488">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4323f-489">Пример</span><span class="sxs-lookup"><span data-stu-id="4323f-489">Example</span></span>

```js
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

<br>

---
---

#### <a name="notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessagesviewoutlook-js-17"></a><span data-ttu-id="4323f-490">notificationMessages: [NotificationMessages](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="4323f-490">notificationMessages: [NotificationMessages](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.7)</span></span>

<span data-ttu-id="4323f-491">Получает сообщения уведомления для элемента.</span><span class="sxs-lookup"><span data-stu-id="4323f-491">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="4323f-492">Тип</span><span class="sxs-lookup"><span data-stu-id="4323f-492">Type</span></span>

*   [<span data-ttu-id="4323f-493">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="4323f-493">NotificationMessages</span></span>](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.7)

##### <a name="requirements"></a><span data-ttu-id="4323f-494">Требования</span><span class="sxs-lookup"><span data-stu-id="4323f-494">Requirements</span></span>

|<span data-ttu-id="4323f-495">Требование</span><span class="sxs-lookup"><span data-stu-id="4323f-495">Requirement</span></span>|<span data-ttu-id="4323f-496">Значение</span><span class="sxs-lookup"><span data-stu-id="4323f-496">Value</span></span>|
|---|---|
|[<span data-ttu-id="4323f-497">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="4323f-497">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4323f-498">1.3</span><span class="sxs-lookup"><span data-stu-id="4323f-498">1.3</span></span>|
|[<span data-ttu-id="4323f-499">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4323f-499">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4323f-500">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4323f-500">ReadItem</span></span>|
|[<span data-ttu-id="4323f-501">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4323f-501">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="4323f-502">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="4323f-502">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4323f-503">Пример</span><span class="sxs-lookup"><span data-stu-id="4323f-503">Example</span></span>

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

#### <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-17recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-17"></a><span data-ttu-id="4323f-504">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="4323f-504">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span></span>

<span data-ttu-id="4323f-505">Предоставляет доступ к необязательным участникам события.</span><span class="sxs-lookup"><span data-stu-id="4323f-505">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="4323f-506">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="4323f-506">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="4323f-507">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="4323f-507">Read mode</span></span>

<span data-ttu-id="4323f-508">Свойство `optionalAttendees` возвращает массив, содержащий объект `EmailAddressDetails` для каждого необязательного участника собрания.</span><span class="sxs-lookup"><span data-stu-id="4323f-508">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span> <span data-ttu-id="4323f-509">Коллекция может включать не более 100 элементов по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="4323f-509">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="4323f-510">Однако на компьютерах с Windows и Mac OS может быть до 500 элементов.</span><span class="sxs-lookup"><span data-stu-id="4323f-510">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="4323f-511">Режим создания</span><span class="sxs-lookup"><span data-stu-id="4323f-511">Compose mode</span></span>

<span data-ttu-id="4323f-512">Свойство `optionalAttendees` возвращает объект `Recipients`, который предоставляет методы для получения или обновления необязательных участников собрания.</span><span class="sxs-lookup"><span data-stu-id="4323f-512">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span> <span data-ttu-id="4323f-513">Коллекция может включать не более 100 элементов по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="4323f-513">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="4323f-514">Однако для компьютеров с Windows и Mac OS действуют указанные ниже ограничения.</span><span class="sxs-lookup"><span data-stu-id="4323f-514">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="4323f-515">Максимальное количество элементов — 500.</span><span class="sxs-lookup"><span data-stu-id="4323f-515">Get 500 members maximum.</span></span>
- <span data-ttu-id="4323f-516">Установить ограничение количества элементов на вызов — не более 100, общего количества — не более 500.</span><span class="sxs-lookup"><span data-stu-id="4323f-516">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="4323f-517">Тип</span><span class="sxs-lookup"><span data-stu-id="4323f-517">Type</span></span>

*   <span data-ttu-id="4323f-518">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="4323f-518">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span></span>

##### <a name="requirements"></a><span data-ttu-id="4323f-519">Требования</span><span class="sxs-lookup"><span data-stu-id="4323f-519">Requirements</span></span>

|<span data-ttu-id="4323f-520">Требование</span><span class="sxs-lookup"><span data-stu-id="4323f-520">Requirement</span></span>|<span data-ttu-id="4323f-521">Значение</span><span class="sxs-lookup"><span data-stu-id="4323f-521">Value</span></span>|
|---|---|
|[<span data-ttu-id="4323f-522">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="4323f-522">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4323f-523">1.0</span><span class="sxs-lookup"><span data-stu-id="4323f-523">1.0</span></span>|
|[<span data-ttu-id="4323f-524">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4323f-524">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4323f-525">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4323f-525">ReadItem</span></span>|
|[<span data-ttu-id="4323f-526">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4323f-526">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="4323f-527">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="4323f-527">Compose or Read</span></span>|

<br>

---
---

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-17organizerjavascriptapioutlookofficeorganizerviewoutlook-js-17"></a><span data-ttu-id="4323f-528">Организатор: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)|[Организатор](/javascript/api/outlook/office.organizer?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="4323f-528">organizer: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)|[Organizer](/javascript/api/outlook/office.organizer?view=outlook-js-1.7)</span></span>

<span data-ttu-id="4323f-529">Получает адрес электронной почты организатора для указанного собрания.</span><span class="sxs-lookup"><span data-stu-id="4323f-529">Gets the email address of the organizer for a specified meeting.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="4323f-530">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="4323f-530">Read mode</span></span>

<span data-ttu-id="4323f-531">`organizer` Свойство возвращает объект [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7) , представляющий организатора собрания.</span><span class="sxs-lookup"><span data-stu-id="4323f-531">The `organizer` property returns an [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7) object that represents the meeting organizer.</span></span>

```js
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

##### <a name="compose-mode"></a><span data-ttu-id="4323f-532">Режим создания</span><span class="sxs-lookup"><span data-stu-id="4323f-532">Compose mode</span></span>

<span data-ttu-id="4323f-533">`organizer` Свойство возвращает объект [организатора](/javascript/api/outlook/office.organizer?view=outlook-js-1.7) , который предоставляет метод для получения значения организатора.</span><span class="sxs-lookup"><span data-stu-id="4323f-533">The `organizer` property returns an [Organizer](/javascript/api/outlook/office.organizer?view=outlook-js-1.7) object that provides a method to get the organizer value.</span></span>

```js
Office.context.mailbox.item.organizer.getAsync(
  function(asyncResult) {
    console.log(JSON.stringify(asyncResult));
  }
);
```

##### <a name="type"></a><span data-ttu-id="4323f-534">Тип</span><span class="sxs-lookup"><span data-stu-id="4323f-534">Type</span></span>

*   <span data-ttu-id="4323f-535">[](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7) | [Организатор](/javascript/api/outlook/office.organizer?view=outlook-js-1.7) EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="4323f-535">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7) | [Organizer](/javascript/api/outlook/office.organizer?view=outlook-js-1.7)</span></span>

##### <a name="requirements"></a><span data-ttu-id="4323f-536">Requirements</span><span class="sxs-lookup"><span data-stu-id="4323f-536">Requirements</span></span>

|<span data-ttu-id="4323f-537">Требование</span><span class="sxs-lookup"><span data-stu-id="4323f-537">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="4323f-538">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="4323f-538">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4323f-539">1.0</span><span class="sxs-lookup"><span data-stu-id="4323f-539">1.0</span></span>|<span data-ttu-id="4323f-540">1.7</span><span class="sxs-lookup"><span data-stu-id="4323f-540">1.7</span></span>|
|[<span data-ttu-id="4323f-541">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4323f-541">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4323f-542">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4323f-542">ReadItem</span></span>|<span data-ttu-id="4323f-543">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="4323f-543">ReadWriteItem</span></span>|
|[<span data-ttu-id="4323f-544">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4323f-544">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="4323f-545">Чтение</span><span class="sxs-lookup"><span data-stu-id="4323f-545">Read</span></span>|<span data-ttu-id="4323f-546">Создание</span><span class="sxs-lookup"><span data-stu-id="4323f-546">Compose</span></span>|

<br>

---
---

#### <a name="nullable-recurrence-recurrencejavascriptapioutlookofficerecurrenceviewoutlook-js-17"></a><span data-ttu-id="4323f-547">(Nullable) повторение: [повторение](/javascript/api/outlook/office.recurrence?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="4323f-547">(nullable) recurrence: [Recurrence](/javascript/api/outlook/office.recurrence?view=outlook-js-1.7)</span></span>

<span data-ttu-id="4323f-548">Получает или задает шаблон повторения встречи.</span><span class="sxs-lookup"><span data-stu-id="4323f-548">Gets or sets the recurrence pattern of an appointment.</span></span> <span data-ttu-id="4323f-549">Получает шаблон повторения приглашения на собрание.</span><span class="sxs-lookup"><span data-stu-id="4323f-549">Gets the recurrence pattern of a meeting request.</span></span> <span data-ttu-id="4323f-550">Режимы чтения и создания для элементов встречи.</span><span class="sxs-lookup"><span data-stu-id="4323f-550">Read and compose modes for appointment items.</span></span> <span data-ttu-id="4323f-551">Режим чтения для элементов приглашения на собрания.</span><span class="sxs-lookup"><span data-stu-id="4323f-551">Read mode for meeting request items.</span></span>

<span data-ttu-id="4323f-552">`recurrence` Свойство возвращает объект [повторения](/javascript/api/outlook/office.recurrence?view=outlook-js-1.7) для повторяющихся встреч или приглашений на собрания, если элемент представляет собой серию или экземпляр в ряду.</span><span class="sxs-lookup"><span data-stu-id="4323f-552">The `recurrence` property returns a [recurrence](/javascript/api/outlook/office.recurrence?view=outlook-js-1.7) object for recurring appointments or meetings requests if an item is a series or an instance in a series.</span></span> <span data-ttu-id="4323f-553">`null`возвращается для отдельных встреч и приглашений на собрание для отдельных встреч.</span><span class="sxs-lookup"><span data-stu-id="4323f-553">`null` is returned for single appointments and meeting requests of single appointments.</span></span> <span data-ttu-id="4323f-554">`undefined`возвращается для сообщений, которые не являются приглашениями на собрания.</span><span class="sxs-lookup"><span data-stu-id="4323f-554">`undefined` is returned for messages that are not meeting requests.</span></span>

> <span data-ttu-id="4323f-555">Note: приглашения на `itemClass` собрания имеют значение IPM. Schedule. Meeting. Request.</span><span class="sxs-lookup"><span data-stu-id="4323f-555">Note: Meeting requests have an `itemClass` value of IPM.Schedule.Meeting.Request.</span></span>

> <span data-ttu-id="4323f-556">Note: при наличии объекта `null`повторения это указывает на то, что объект является одной встречей или приглашением на собрание одной встречи, а не частью ряда.</span><span class="sxs-lookup"><span data-stu-id="4323f-556">Note: If the recurrence object is `null`, this indicates that the object is a single appointment or a meeting request of a single appointment and NOT a part of a series.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="4323f-557">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="4323f-557">Read mode</span></span>

<span data-ttu-id="4323f-558">`recurrence` Свойство возвращает объект [повторения](/javascript/api/outlook/office.recurrence?view=outlook-js-1.7) , представляющий повторение встречи.</span><span class="sxs-lookup"><span data-stu-id="4323f-558">The `recurrence` property returns a [Recurrence](/javascript/api/outlook/office.recurrence?view=outlook-js-1.7) object that represents the appointment recurrence.</span></span> <span data-ttu-id="4323f-559">Оно доступно для встреч и приглашений на собрания.</span><span class="sxs-lookup"><span data-stu-id="4323f-559">This is available for appointments and meeting requests.</span></span>

```js
var recurrence = Office.context.mailbox.item.recurrence;
console.log("Recurrence: " + JSON.stringify(recurrence));
```

##### <a name="compose-mode"></a><span data-ttu-id="4323f-560">Режим создания</span><span class="sxs-lookup"><span data-stu-id="4323f-560">Compose mode</span></span>

<span data-ttu-id="4323f-561">`recurrence` Свойство возвращает объект [повторения](/javascript/api/outlook/office.recurrence?view=outlook-js-1.7) , который предоставляет методы для управления повторением встречи.</span><span class="sxs-lookup"><span data-stu-id="4323f-561">The `recurrence` property returns a [Recurrence](/javascript/api/outlook/office.recurrence?view=outlook-js-1.7) object that provides methods to manage the appointment recurrence.</span></span> <span data-ttu-id="4323f-562">Оно доступно для встреч.</span><span class="sxs-lookup"><span data-stu-id="4323f-562">This is available for appointments.</span></span>

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

##### <a name="type"></a><span data-ttu-id="4323f-563">Тип</span><span class="sxs-lookup"><span data-stu-id="4323f-563">Type</span></span>

* [<span data-ttu-id="4323f-564">Повторения</span><span class="sxs-lookup"><span data-stu-id="4323f-564">Recurrence</span></span>](/javascript/api/outlook/office.recurrence?view=outlook-js-1.7)

|<span data-ttu-id="4323f-565">Requirement</span><span class="sxs-lookup"><span data-stu-id="4323f-565">Requirement</span></span>|<span data-ttu-id="4323f-566">Значение</span><span class="sxs-lookup"><span data-stu-id="4323f-566">Value</span></span>|
|---|---|
|[<span data-ttu-id="4323f-567">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="4323f-567">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4323f-568">1.7</span><span class="sxs-lookup"><span data-stu-id="4323f-568">1.7</span></span>|
|[<span data-ttu-id="4323f-569">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4323f-569">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4323f-570">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4323f-570">ReadItem</span></span>|
|[<span data-ttu-id="4323f-571">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4323f-571">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="4323f-572">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="4323f-572">Compose or Read</span></span>|

<br>

---
---

#### <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-17recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-17"></a><span data-ttu-id="4323f-573">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="4323f-573">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span></span>

<span data-ttu-id="4323f-574">Предоставляет доступ к обязательным участникам события.</span><span class="sxs-lookup"><span data-stu-id="4323f-574">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="4323f-575">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="4323f-575">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="4323f-576">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="4323f-576">Read mode</span></span>

<span data-ttu-id="4323f-577">Свойство `requiredAttendees` возвращает массив, содержащий объект `EmailAddressDetails` для каждого обязательного участника собрания.</span><span class="sxs-lookup"><span data-stu-id="4323f-577">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span> <span data-ttu-id="4323f-578">Коллекция может включать не более 100 элементов по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="4323f-578">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="4323f-579">Однако на компьютерах с Windows и Mac OS может быть до 500 элементов.</span><span class="sxs-lookup"><span data-stu-id="4323f-579">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="4323f-580">Режим создания</span><span class="sxs-lookup"><span data-stu-id="4323f-580">Compose mode</span></span>

<span data-ttu-id="4323f-581">Свойство `requiredAttendees` возвращает объект `Recipients`, предоставляющий методы, с помощью которых можно получить или обновить сведения об обязательных участниках собрания.</span><span class="sxs-lookup"><span data-stu-id="4323f-581">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span> <span data-ttu-id="4323f-582">Коллекция может включать не более 100 элементов по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="4323f-582">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="4323f-583">Однако для компьютеров с Windows и Mac OS действуют указанные ниже ограничения.</span><span class="sxs-lookup"><span data-stu-id="4323f-583">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="4323f-584">Максимальное количество элементов — 500.</span><span class="sxs-lookup"><span data-stu-id="4323f-584">Get 500 members maximum.</span></span>
- <span data-ttu-id="4323f-585">Установить ограничение количества элементов на вызов — не более 100, общего количества — не более 500.</span><span class="sxs-lookup"><span data-stu-id="4323f-585">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="4323f-586">Тип</span><span class="sxs-lookup"><span data-stu-id="4323f-586">Type</span></span>

*   <span data-ttu-id="4323f-587">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="4323f-587">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span></span>

##### <a name="requirements"></a><span data-ttu-id="4323f-588">Требования</span><span class="sxs-lookup"><span data-stu-id="4323f-588">Requirements</span></span>

|<span data-ttu-id="4323f-589">Требование</span><span class="sxs-lookup"><span data-stu-id="4323f-589">Requirement</span></span>|<span data-ttu-id="4323f-590">Значение</span><span class="sxs-lookup"><span data-stu-id="4323f-590">Value</span></span>|
|---|---|
|[<span data-ttu-id="4323f-591">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="4323f-591">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4323f-592">1.0</span><span class="sxs-lookup"><span data-stu-id="4323f-592">1.0</span></span>|
|[<span data-ttu-id="4323f-593">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4323f-593">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4323f-594">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4323f-594">ReadItem</span></span>|
|[<span data-ttu-id="4323f-595">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4323f-595">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="4323f-596">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="4323f-596">Compose or Read</span></span>|

<br>

---
---

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-17"></a><span data-ttu-id="4323f-597">sender: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="4323f-597">sender: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)</span></span>

<span data-ttu-id="4323f-p134">Получает электронный адрес отправителя электронного сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="4323f-p134">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="4323f-p135">Свойства [`from`](#from-emailaddressdetailsfrom) и `sender` представляют одно лицо, если сообщение не отправлено представителем. В противном случае свойство `from` представляет лицо, делегировавшее полномочия, а свойство sender — представителя.</span><span class="sxs-lookup"><span data-stu-id="4323f-p135">The [`from`](#from-emailaddressdetailsfrom) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="4323f-602">Свойству `recipientType`, принадлежащему объекту `EmailAddressDetails` в свойстве `sender`, задано значение `undefined`.</span><span class="sxs-lookup"><span data-stu-id="4323f-602">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="4323f-603">Тип</span><span class="sxs-lookup"><span data-stu-id="4323f-603">Type</span></span>

*   [<span data-ttu-id="4323f-604">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="4323f-604">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)

##### <a name="requirements"></a><span data-ttu-id="4323f-605">Requirements</span><span class="sxs-lookup"><span data-stu-id="4323f-605">Requirements</span></span>

|<span data-ttu-id="4323f-606">Требование</span><span class="sxs-lookup"><span data-stu-id="4323f-606">Requirement</span></span>|<span data-ttu-id="4323f-607">Значение</span><span class="sxs-lookup"><span data-stu-id="4323f-607">Value</span></span>|
|---|---|
|[<span data-ttu-id="4323f-608">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="4323f-608">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4323f-609">1.0</span><span class="sxs-lookup"><span data-stu-id="4323f-609">1.0</span></span>|
|[<span data-ttu-id="4323f-610">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4323f-610">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4323f-611">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4323f-611">ReadItem</span></span>|
|[<span data-ttu-id="4323f-612">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4323f-612">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="4323f-613">Чтение</span><span class="sxs-lookup"><span data-stu-id="4323f-613">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4323f-614">Пример</span><span class="sxs-lookup"><span data-stu-id="4323f-614">Example</span></span>

```js
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

<br>

---
---

#### <a name="nullable-seriesid-string"></a><span data-ttu-id="4323f-615">(Nullable) seriesId: строка</span><span class="sxs-lookup"><span data-stu-id="4323f-615">(nullable) seriesId: String</span></span>

<span data-ttu-id="4323f-616">Получает идентификатор ряда, к которому принадлежит экземпляр.</span><span class="sxs-lookup"><span data-stu-id="4323f-616">Gets the id of the series that an instance belongs to.</span></span>

<span data-ttu-id="4323f-617">В Outlook в Интернете и на настольных клиентах `seriesId` возвращается идентификатор веб-служб Exchange (EWS) родительского элемента (ряда), к которому принадлежит этот элемент.</span><span class="sxs-lookup"><span data-stu-id="4323f-617">In Outlook on the web and desktop clients, the `seriesId` returns the Exchange Web Services (EWS) ID of the parent (series) item that this item belongs to.</span></span> <span data-ttu-id="4323f-618">Однако в iOS и Android `seriesId` возвращается идентификатор REST родительского элемента.</span><span class="sxs-lookup"><span data-stu-id="4323f-618">However, in iOS and Android, the `seriesId` returns the REST ID of the parent item.</span></span>

> [!NOTE]
> <span data-ttu-id="4323f-619">Идентификатор, возвращаемый свойством `seriesId`, совпадает с идентификатором элемента веб-служб Exchange.</span><span class="sxs-lookup"><span data-stu-id="4323f-619">The identifier returned by the `seriesId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="4323f-620">`seriesId` Свойство не совпадает с идентификаторами Outlook, используемыми в REST API Outlook.</span><span class="sxs-lookup"><span data-stu-id="4323f-620">The `seriesId` property is not identical to the Outlook IDs used by the Outlook REST API.</span></span> <span data-ttu-id="4323f-621">Прежде чем совершать вызовы REST API, используя это значение, его необходимо преобразовать с помощью [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="4323f-621">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="4323f-622">Дополнительные сведения см. в статье [Использование REST API Outlook из надстройки Outlook](/outlook/add-ins/use-rest-api).</span><span class="sxs-lookup"><span data-stu-id="4323f-622">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api).</span></span>

<span data-ttu-id="4323f-623">`seriesId` Свойство возвращает `null` элементы, у которых нет родительских элементов, таких как одиночные встречи, элементы ряда или приглашения на собрание, `undefined` и возвращаемые для других элементов, не являющиеся приглашениями на собрания.</span><span class="sxs-lookup"><span data-stu-id="4323f-623">The `seriesId` property returns `null` for items that do not have parent items such as single appointments, series items, or meeting requests and returns `undefined` for any other items that are not meeting requests.</span></span>

##### <a name="type"></a><span data-ttu-id="4323f-624">Тип</span><span class="sxs-lookup"><span data-stu-id="4323f-624">Type</span></span>

* <span data-ttu-id="4323f-625">String</span><span class="sxs-lookup"><span data-stu-id="4323f-625">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="4323f-626">Requirements</span><span class="sxs-lookup"><span data-stu-id="4323f-626">Requirements</span></span>

|<span data-ttu-id="4323f-627">Требование</span><span class="sxs-lookup"><span data-stu-id="4323f-627">Requirement</span></span>|<span data-ttu-id="4323f-628">Значение</span><span class="sxs-lookup"><span data-stu-id="4323f-628">Value</span></span>|
|---|---|
|[<span data-ttu-id="4323f-629">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="4323f-629">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4323f-630">1.7</span><span class="sxs-lookup"><span data-stu-id="4323f-630">1.7</span></span>|
|[<span data-ttu-id="4323f-631">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4323f-631">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4323f-632">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4323f-632">ReadItem</span></span>|
|[<span data-ttu-id="4323f-633">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4323f-633">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="4323f-634">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="4323f-634">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4323f-635">Пример</span><span class="sxs-lookup"><span data-stu-id="4323f-635">Example</span></span>

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

#### <a name="start-datetimejavascriptapioutlookofficetimeviewoutlook-js-17"></a><span data-ttu-id="4323f-636">start: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="4323f-636">start: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.7)</span></span>

<span data-ttu-id="4323f-637">Получает или задает дату и время начала встречи.</span><span class="sxs-lookup"><span data-stu-id="4323f-637">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="4323f-p138">Свойство `start` представлено в виде значения даты и времени в формате UTC. Это значение можно преобразовать в местные значения даты и времени клиента с помощью метода [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime).</span><span class="sxs-lookup"><span data-stu-id="4323f-p138">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="4323f-640">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="4323f-640">Read mode</span></span>

<span data-ttu-id="4323f-641">Свойство `start` возвращает объект `Date`.</span><span class="sxs-lookup"><span data-stu-id="4323f-641">The `start` property returns a `Date` object.</span></span>

```js
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="4323f-642">Режим создания</span><span class="sxs-lookup"><span data-stu-id="4323f-642">Compose mode</span></span>

<span data-ttu-id="4323f-643">Свойство `start` возвращает объект `Time`.</span><span class="sxs-lookup"><span data-stu-id="4323f-643">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="4323f-644">Если вы задаете время начала с помощью метода [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.7#setasync-datetime--options--callback-), необходимо использовать метод [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) для преобразования местного времени на клиенте в формат UTC для сервера.</span><span class="sxs-lookup"><span data-stu-id="4323f-644">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.7#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="4323f-645">В примере ниже с помощью метода [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.7#setasync-datetime--options--callback-) объекта `Time` задается время начала встречи в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="4323f-645">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.7#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="4323f-646">Тип</span><span class="sxs-lookup"><span data-stu-id="4323f-646">Type</span></span>

*   <span data-ttu-id="4323f-647">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="4323f-647">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.7)</span></span>

##### <a name="requirements"></a><span data-ttu-id="4323f-648">Requirements</span><span class="sxs-lookup"><span data-stu-id="4323f-648">Requirements</span></span>

|<span data-ttu-id="4323f-649">Требование</span><span class="sxs-lookup"><span data-stu-id="4323f-649">Requirement</span></span>|<span data-ttu-id="4323f-650">Значение</span><span class="sxs-lookup"><span data-stu-id="4323f-650">Value</span></span>|
|---|---|
|[<span data-ttu-id="4323f-651">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="4323f-651">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4323f-652">1.0</span><span class="sxs-lookup"><span data-stu-id="4323f-652">1.0</span></span>|
|[<span data-ttu-id="4323f-653">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4323f-653">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4323f-654">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4323f-654">ReadItem</span></span>|
|[<span data-ttu-id="4323f-655">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4323f-655">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="4323f-656">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="4323f-656">Compose or Read</span></span>|

<br>

---
---

#### <a name="subject-stringsubjectjavascriptapioutlookofficesubjectviewoutlook-js-17"></a><span data-ttu-id="4323f-657">subject: String|[Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="4323f-657">subject: String|[Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.7)</span></span>

<span data-ttu-id="4323f-658">Получает или задает описание, которое отображается в поле темы элемента.</span><span class="sxs-lookup"><span data-stu-id="4323f-658">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="4323f-659">Свойство `subject` получает или задает всю тему элемента для отправки с почтового сервера.</span><span class="sxs-lookup"><span data-stu-id="4323f-659">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="4323f-660">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="4323f-660">Read mode</span></span>

<span data-ttu-id="4323f-p139">Свойство `subject` возвращает строку. С помощью свойства [`normalizedSubject`](#normalizedsubject-string) можно получить тему без начальных префиксов, таких как `RE:` и `FW:`.</span><span class="sxs-lookup"><span data-stu-id="4323f-p139">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

<span data-ttu-id="4323f-663">В примере кода JavaScript, приведенном ниже, показано, как получить доступ к свойству `subject` текущего элемента в Outlook.</span><span class="sxs-lookup"><span data-stu-id="4323f-663">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

```js
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="4323f-664">Режим создания</span><span class="sxs-lookup"><span data-stu-id="4323f-664">Compose mode</span></span>

<span data-ttu-id="4323f-665">Свойство `subject` возвращает объект `Subject`, который предоставляет методы для получения и задания темы.</span><span class="sxs-lookup"><span data-stu-id="4323f-665">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```js
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="4323f-666">Тип</span><span class="sxs-lookup"><span data-stu-id="4323f-666">Type</span></span>

*   <span data-ttu-id="4323f-667">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="4323f-667">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.7)</span></span>

##### <a name="requirements"></a><span data-ttu-id="4323f-668">Требования</span><span class="sxs-lookup"><span data-stu-id="4323f-668">Requirements</span></span>

|<span data-ttu-id="4323f-669">Требование</span><span class="sxs-lookup"><span data-stu-id="4323f-669">Requirement</span></span>|<span data-ttu-id="4323f-670">Значение</span><span class="sxs-lookup"><span data-stu-id="4323f-670">Value</span></span>|
|---|---|
|[<span data-ttu-id="4323f-671">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="4323f-671">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4323f-672">1.0</span><span class="sxs-lookup"><span data-stu-id="4323f-672">1.0</span></span>|
|[<span data-ttu-id="4323f-673">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4323f-673">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4323f-674">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4323f-674">ReadItem</span></span>|
|[<span data-ttu-id="4323f-675">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4323f-675">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="4323f-676">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="4323f-676">Compose or Read</span></span>|

<br>

---
---

#### <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-17recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-17"></a><span data-ttu-id="4323f-677">to: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="4323f-677">to: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span></span>

<span data-ttu-id="4323f-678">Предоставляет доступ к получателям, указанным в строке **Кому** сообщения.</span><span class="sxs-lookup"><span data-stu-id="4323f-678">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="4323f-679">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="4323f-679">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="4323f-680">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="4323f-680">Read mode</span></span>

<span data-ttu-id="4323f-681">Свойство `to` возвращает массив, содержащий объект `EmailAddressDetails` для каждого получателя в строке **Кому** сообщения.</span><span class="sxs-lookup"><span data-stu-id="4323f-681">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message.</span></span> <span data-ttu-id="4323f-682">Коллекция может включать не более 100 элементов по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="4323f-682">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="4323f-683">Однако на компьютерах с Windows и Mac OS может быть до 500 элементов.</span><span class="sxs-lookup"><span data-stu-id="4323f-683">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="4323f-684">Режим создания</span><span class="sxs-lookup"><span data-stu-id="4323f-684">Compose mode</span></span>

<span data-ttu-id="4323f-685">Свойство `to` возвращает объект `Recipients`, предоставляющий методы для получения или обновления получателей, которые указаны в строке **Кому** сообщения.</span><span class="sxs-lookup"><span data-stu-id="4323f-685">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span> <span data-ttu-id="4323f-686">Коллекция может включать не более 100 элементов по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="4323f-686">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="4323f-687">Однако для компьютеров с Windows и Mac OS действуют указанные ниже ограничения.</span><span class="sxs-lookup"><span data-stu-id="4323f-687">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="4323f-688">Максимальное количество элементов — 500.</span><span class="sxs-lookup"><span data-stu-id="4323f-688">Get 500 members maximum.</span></span>
- <span data-ttu-id="4323f-689">Установить ограничение количества элементов на вызов — не более 100, общего количества — не более 500.</span><span class="sxs-lookup"><span data-stu-id="4323f-689">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="4323f-690">Тип</span><span class="sxs-lookup"><span data-stu-id="4323f-690">Type</span></span>

*   <span data-ttu-id="4323f-691">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="4323f-691">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span></span>

##### <a name="requirements"></a><span data-ttu-id="4323f-692">Requirements</span><span class="sxs-lookup"><span data-stu-id="4323f-692">Requirements</span></span>

|<span data-ttu-id="4323f-693">Требование</span><span class="sxs-lookup"><span data-stu-id="4323f-693">Requirement</span></span>|<span data-ttu-id="4323f-694">Значение</span><span class="sxs-lookup"><span data-stu-id="4323f-694">Value</span></span>|
|---|---|
|[<span data-ttu-id="4323f-695">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="4323f-695">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4323f-696">1.0</span><span class="sxs-lookup"><span data-stu-id="4323f-696">1.0</span></span>|
|[<span data-ttu-id="4323f-697">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4323f-697">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4323f-698">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4323f-698">ReadItem</span></span>|
|[<span data-ttu-id="4323f-699">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4323f-699">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="4323f-700">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="4323f-700">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="4323f-701">Методы</span><span class="sxs-lookup"><span data-stu-id="4323f-701">Methods</span></span>

#### <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="4323f-702">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="4323f-702">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="4323f-703">Добавляет файл в сообщение или встречу в качестве вложения.</span><span class="sxs-lookup"><span data-stu-id="4323f-703">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="4323f-704">Метод `addFileAttachmentAsync` передает файл по указанному универсальному коду ресурса (URI) и вкладывает его в элемент в форме создания.</span><span class="sxs-lookup"><span data-stu-id="4323f-704">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="4323f-705">Идентификатор можно использовать с методом [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback), чтобы удалить вложение, добавленное во время текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="4323f-705">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4323f-706">Параметры</span><span class="sxs-lookup"><span data-stu-id="4323f-706">Parameters</span></span>
|<span data-ttu-id="4323f-707">Имя</span><span class="sxs-lookup"><span data-stu-id="4323f-707">Name</span></span>|<span data-ttu-id="4323f-708">Тип</span><span class="sxs-lookup"><span data-stu-id="4323f-708">Type</span></span>|<span data-ttu-id="4323f-709">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="4323f-709">Attributes</span></span>|<span data-ttu-id="4323f-710">Описание</span><span class="sxs-lookup"><span data-stu-id="4323f-710">Description</span></span>|
|---|---|---|---|
|`uri`|<span data-ttu-id="4323f-711">String</span><span class="sxs-lookup"><span data-stu-id="4323f-711">String</span></span>||<span data-ttu-id="4323f-p143">Универсальный код ресурса (URI), представляющий расположение файла, который нужно вложить в сообщение или встречу. Максимальная длина — 2048 символов.</span><span class="sxs-lookup"><span data-stu-id="4323f-p143">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="4323f-714">String</span><span class="sxs-lookup"><span data-stu-id="4323f-714">String</span></span>||<span data-ttu-id="4323f-p144">Имя вложения, которое отображается при передаче вложения. Максимальная длина — 255 символов.</span><span class="sxs-lookup"><span data-stu-id="4323f-p144">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="4323f-717">Object</span><span class="sxs-lookup"><span data-stu-id="4323f-717">Object</span></span>|<span data-ttu-id="4323f-718">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="4323f-718">&lt;optional&gt;</span></span>|<span data-ttu-id="4323f-719">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="4323f-719">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="4323f-720">Object</span><span class="sxs-lookup"><span data-stu-id="4323f-720">Object</span></span>|<span data-ttu-id="4323f-721">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="4323f-721">&lt;optional&gt;</span></span>|<span data-ttu-id="4323f-722">В методе обратного вызова разработчики могут указать любой объект, к которому необходимо получить доступ.</span><span class="sxs-lookup"><span data-stu-id="4323f-722">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="4323f-723">Boolean</span><span class="sxs-lookup"><span data-stu-id="4323f-723">Boolean</span></span>|<span data-ttu-id="4323f-724">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="4323f-724">&lt;optional&gt;</span></span>|<span data-ttu-id="4323f-725">Значение `true` указывает на то, что вложение будет встроено в текст сообщения и не должно отображаться в списке вложений.</span><span class="sxs-lookup"><span data-stu-id="4323f-725">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="4323f-726">function</span><span class="sxs-lookup"><span data-stu-id="4323f-726">function</span></span>|<span data-ttu-id="4323f-727">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="4323f-727">&lt;optional&gt;</span></span>|<span data-ttu-id="4323f-728">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="4323f-728">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="4323f-729">После успешного выполнения идентификатор вложения будет представлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="4323f-729">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="4323f-730">Если добавить вложение не удастся, объект `asyncResult` будет содержать объект `Error` с описанием ошибки.</span><span class="sxs-lookup"><span data-stu-id="4323f-730">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="4323f-731">Ошибки</span><span class="sxs-lookup"><span data-stu-id="4323f-731">Errors</span></span>

|<span data-ttu-id="4323f-732">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="4323f-732">Error code</span></span>|<span data-ttu-id="4323f-733">Описание</span><span class="sxs-lookup"><span data-stu-id="4323f-733">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="4323f-734">Вложение превышает максимальный размер.</span><span class="sxs-lookup"><span data-stu-id="4323f-734">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="4323f-735">Расширение вложения не поддерживается.</span><span class="sxs-lookup"><span data-stu-id="4323f-735">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="4323f-736">Сообщение или встреча содержат слишком много вложений.</span><span class="sxs-lookup"><span data-stu-id="4323f-736">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4323f-737">Requirements</span><span class="sxs-lookup"><span data-stu-id="4323f-737">Requirements</span></span>

|<span data-ttu-id="4323f-738">Требование</span><span class="sxs-lookup"><span data-stu-id="4323f-738">Requirement</span></span>|<span data-ttu-id="4323f-739">Значение</span><span class="sxs-lookup"><span data-stu-id="4323f-739">Value</span></span>|
|---|---|
|[<span data-ttu-id="4323f-740">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="4323f-740">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4323f-741">1.1</span><span class="sxs-lookup"><span data-stu-id="4323f-741">1.1</span></span>|
|[<span data-ttu-id="4323f-742">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4323f-742">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4323f-743">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="4323f-743">ReadWriteItem</span></span>|
|[<span data-ttu-id="4323f-744">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4323f-744">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="4323f-745">Создание</span><span class="sxs-lookup"><span data-stu-id="4323f-745">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="4323f-746">Примеры</span><span class="sxs-lookup"><span data-stu-id="4323f-746">Examples</span></span>

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

<span data-ttu-id="4323f-747">В приведенном ниже примере файл изображения добавляется в качестве встроенного вложения, а ссылка на вложение добавляется в текст сообщения.</span><span class="sxs-lookup"><span data-stu-id="4323f-747">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

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

#### <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="4323f-748">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="4323f-748">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="4323f-749">Добавляет обработчик для поддерживаемого события.</span><span class="sxs-lookup"><span data-stu-id="4323f-749">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="4323f-750">В настоящее время поддерживаются типы `Office.EventType.AppointmentTimeChanged`событий `Office.EventType.RecipientsChanged`, и`Office.EventType.RecurrenceChanged`</span><span class="sxs-lookup"><span data-stu-id="4323f-750">Currently the supported event types are `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`</span></span>

##### <a name="parameters"></a><span data-ttu-id="4323f-751">Параметры</span><span class="sxs-lookup"><span data-stu-id="4323f-751">Parameters</span></span>

| <span data-ttu-id="4323f-752">Имя</span><span class="sxs-lookup"><span data-stu-id="4323f-752">Name</span></span> | <span data-ttu-id="4323f-753">Тип</span><span class="sxs-lookup"><span data-stu-id="4323f-753">Type</span></span> | <span data-ttu-id="4323f-754">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="4323f-754">Attributes</span></span> | <span data-ttu-id="4323f-755">Описание</span><span class="sxs-lookup"><span data-stu-id="4323f-755">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="4323f-756">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="4323f-756">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="4323f-757">Событие, которое должно вызвать обработчик.</span><span class="sxs-lookup"><span data-stu-id="4323f-757">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="4323f-758">Function</span><span class="sxs-lookup"><span data-stu-id="4323f-758">Function</span></span> || <span data-ttu-id="4323f-p145">Функция для обработки события. Функция должна принимать один параметр, представляющий собой объектный литерал. Значение свойства `type` параметра совпадет со значением параметра `eventType`, переданного методу `addHandlerAsync`.</span><span class="sxs-lookup"><span data-stu-id="4323f-p145">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="4323f-762">Объект</span><span class="sxs-lookup"><span data-stu-id="4323f-762">Object</span></span> | <span data-ttu-id="4323f-763">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="4323f-763">&lt;optional&gt;</span></span> | <span data-ttu-id="4323f-764">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="4323f-764">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="4323f-765">Object</span><span class="sxs-lookup"><span data-stu-id="4323f-765">Object</span></span> | <span data-ttu-id="4323f-766">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="4323f-766">&lt;optional&gt;</span></span> | <span data-ttu-id="4323f-767">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="4323f-767">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="4323f-768">функция</span><span class="sxs-lookup"><span data-stu-id="4323f-768">function</span></span>| <span data-ttu-id="4323f-769">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="4323f-769">&lt;optional&gt;</span></span>|<span data-ttu-id="4323f-770">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="4323f-770">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4323f-771">Требования</span><span class="sxs-lookup"><span data-stu-id="4323f-771">Requirements</span></span>

|<span data-ttu-id="4323f-772">Требование</span><span class="sxs-lookup"><span data-stu-id="4323f-772">Requirement</span></span>| <span data-ttu-id="4323f-773">Значение</span><span class="sxs-lookup"><span data-stu-id="4323f-773">Value</span></span>|
|---|---|
|[<span data-ttu-id="4323f-774">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="4323f-774">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4323f-775">1.7</span><span class="sxs-lookup"><span data-stu-id="4323f-775">1.7</span></span> |
|[<span data-ttu-id="4323f-776">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4323f-776">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4323f-777">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4323f-777">ReadItem</span></span> |
|[<span data-ttu-id="4323f-778">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4323f-778">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4323f-779">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="4323f-779">Compose or Read</span></span> |

##### <a name="example"></a><span data-ttu-id="4323f-780">Пример</span><span class="sxs-lookup"><span data-stu-id="4323f-780">Example</span></span>

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

#### <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="4323f-781">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="4323f-781">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="4323f-782">Добавляет к сообщению элемент Exchange, например сообщение, в виде вложения.</span><span class="sxs-lookup"><span data-stu-id="4323f-782">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="4323f-p146">С помощью метода `addItemAttachmentAsync` можно в элемент формы создания вложить элемент с указанным идентификатором Exchange. Если указать метод обратного вызова, то этот метод вызывается с помощью параметра `asyncResult`, который содержит идентификатор вложения или код, указывающий на ошибки, которые произошли при вложении элемента. При необходимости можно использовать параметр `options` для передачи сведений о состоянии методу обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="4323f-p146">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="4323f-786">Идентификатор можно использовать с методом [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback), чтобы удалить вложение, добавленное во время текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="4323f-786">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="4323f-787">Если ваша надстройка Office выполняется в Outlook в Интернете, метод `addItemAttachmentAsync` обеспечивает вложение элементов в элементы, отличные от редактируемого. Однако это действие не рекомендуется выполнять, так как оно не поддерживается.</span><span class="sxs-lookup"><span data-stu-id="4323f-787">If your Office Add-in is running in Outlook on the web, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4323f-788">Параметры</span><span class="sxs-lookup"><span data-stu-id="4323f-788">Parameters</span></span>

|<span data-ttu-id="4323f-789">Имя</span><span class="sxs-lookup"><span data-stu-id="4323f-789">Name</span></span>|<span data-ttu-id="4323f-790">Тип</span><span class="sxs-lookup"><span data-stu-id="4323f-790">Type</span></span>|<span data-ttu-id="4323f-791">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="4323f-791">Attributes</span></span>|<span data-ttu-id="4323f-792">Описание</span><span class="sxs-lookup"><span data-stu-id="4323f-792">Description</span></span>|
|---|---|---|---|
|`itemId`|<span data-ttu-id="4323f-793">String</span><span class="sxs-lookup"><span data-stu-id="4323f-793">String</span></span>||<span data-ttu-id="4323f-p147">Идентификатор Exchange для вкладываемого элемента. Максимальная длина — 100 символов.</span><span class="sxs-lookup"><span data-stu-id="4323f-p147">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="4323f-796">String</span><span class="sxs-lookup"><span data-stu-id="4323f-796">String</span></span>||<span data-ttu-id="4323f-797">Тема вкладываемого элемента.</span><span class="sxs-lookup"><span data-stu-id="4323f-797">The subject of the item to be attached.</span></span> <span data-ttu-id="4323f-798">Максимальная длина: 255 символов.</span><span class="sxs-lookup"><span data-stu-id="4323f-798">The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="4323f-799">Object</span><span class="sxs-lookup"><span data-stu-id="4323f-799">Object</span></span>|<span data-ttu-id="4323f-800">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="4323f-800">&lt;optional&gt;</span></span>|<span data-ttu-id="4323f-801">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="4323f-801">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="4323f-802">Object</span><span class="sxs-lookup"><span data-stu-id="4323f-802">Object</span></span>|<span data-ttu-id="4323f-803">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="4323f-803">&lt;optional&gt;</span></span>|<span data-ttu-id="4323f-804">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="4323f-804">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="4323f-805">функция</span><span class="sxs-lookup"><span data-stu-id="4323f-805">function</span></span>|<span data-ttu-id="4323f-806">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="4323f-806">&lt;optional&gt;</span></span>|<span data-ttu-id="4323f-807">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="4323f-807">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="4323f-808">После успешного выполнения идентификатор вложения будет представлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="4323f-808">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="4323f-809">Если добавить вложение не удастся, объект `asyncResult` будет содержать объект `Error` с описанием ошибки.</span><span class="sxs-lookup"><span data-stu-id="4323f-809">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="4323f-810">Ошибки</span><span class="sxs-lookup"><span data-stu-id="4323f-810">Errors</span></span>

|<span data-ttu-id="4323f-811">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="4323f-811">Error code</span></span>|<span data-ttu-id="4323f-812">Описание</span><span class="sxs-lookup"><span data-stu-id="4323f-812">Description</span></span>|
|------------|-------------|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="4323f-813">Сообщение или встреча содержат слишком много вложений.</span><span class="sxs-lookup"><span data-stu-id="4323f-813">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4323f-814">Requirements</span><span class="sxs-lookup"><span data-stu-id="4323f-814">Requirements</span></span>

|<span data-ttu-id="4323f-815">Требование</span><span class="sxs-lookup"><span data-stu-id="4323f-815">Requirement</span></span>|<span data-ttu-id="4323f-816">Значение</span><span class="sxs-lookup"><span data-stu-id="4323f-816">Value</span></span>|
|---|---|
|[<span data-ttu-id="4323f-817">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="4323f-817">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4323f-818">1.1</span><span class="sxs-lookup"><span data-stu-id="4323f-818">1.1</span></span>|
|[<span data-ttu-id="4323f-819">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4323f-819">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4323f-820">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="4323f-820">ReadWriteItem</span></span>|
|[<span data-ttu-id="4323f-821">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4323f-821">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="4323f-822">Создание</span><span class="sxs-lookup"><span data-stu-id="4323f-822">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="4323f-823">Пример</span><span class="sxs-lookup"><span data-stu-id="4323f-823">Example</span></span>

<span data-ttu-id="4323f-824">В следующем примере существующий элемент Outlook добавляется в виде вложения с именем `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="4323f-824">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

#### <a name="close"></a><span data-ttu-id="4323f-825">close()</span><span class="sxs-lookup"><span data-stu-id="4323f-825">close()</span></span>

<span data-ttu-id="4323f-826">Закрывает текущий создаваемый элемент.</span><span class="sxs-lookup"><span data-stu-id="4323f-826">Closes the current item that is being composed.</span></span>

<span data-ttu-id="4323f-p149">Работа метода `close` зависит от текущего состояния создаваемого элемента. Если элемент содержит несохраненные изменения, клиент предложит пользователю сохранить или отклонить их либо отменить действие закрытия.</span><span class="sxs-lookup"><span data-stu-id="4323f-p149">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="4323f-829">Если в Outlook в Интернете элемент представляет собой встречу, ранее сохраненную с помощью метода `saveAsync`, пользователю предлагается сохранить, отклонить или отменить действие, даже если с момента последнего сохранения элемента не вносились какие-либо изменения.</span><span class="sxs-lookup"><span data-stu-id="4323f-829">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="4323f-830">Если в клиенте Outlook для настольных ПК сообщение представляет собой ответ в тексте, метод `close` не работает.</span><span class="sxs-lookup"><span data-stu-id="4323f-830">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="4323f-831">Requirements</span><span class="sxs-lookup"><span data-stu-id="4323f-831">Requirements</span></span>

|<span data-ttu-id="4323f-832">Требование</span><span class="sxs-lookup"><span data-stu-id="4323f-832">Requirement</span></span>|<span data-ttu-id="4323f-833">Значение</span><span class="sxs-lookup"><span data-stu-id="4323f-833">Value</span></span>|
|---|---|
|[<span data-ttu-id="4323f-834">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="4323f-834">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4323f-835">1.3</span><span class="sxs-lookup"><span data-stu-id="4323f-835">1.3</span></span>|
|[<span data-ttu-id="4323f-836">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4323f-836">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4323f-837">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="4323f-837">Restricted</span></span>|
|[<span data-ttu-id="4323f-838">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4323f-838">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="4323f-839">Создание</span><span class="sxs-lookup"><span data-stu-id="4323f-839">Compose</span></span>|

<br>

---
---

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="4323f-840">displayReplyAllForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="4323f-840">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="4323f-841">Отображает форму ответа, включающую отправителя и всех получателей выбранного сообщения или организатора и всех участников выбранной встречи.</span><span class="sxs-lookup"><span data-stu-id="4323f-841">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="4323f-842">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="4323f-842">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="4323f-843">В Outlook в Интернете форма ответа отображается в виде всплывающей формы в представлении с 3 либо 1 или 2 столбцами.</span><span class="sxs-lookup"><span data-stu-id="4323f-843">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="4323f-844">Если любой строковый параметр превышает указанные для него ограничения, `displayReplyAllForm` возвращает исключение.</span><span class="sxs-lookup"><span data-stu-id="4323f-844">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="4323f-p150">Если в параметре `formData.attachments` указаны вложения, Outlook в Интернете и классические клиенты пытаются скачать их и вложить в форму ответа. Если какие-либо вложения добавить не удается, в форме отображается сообщение об ошибке. Если сообщения об ошибках не предусмотрены, то они не отображаются.</span><span class="sxs-lookup"><span data-stu-id="4323f-p150">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4323f-848">Параметры</span><span class="sxs-lookup"><span data-stu-id="4323f-848">Parameters</span></span>

|<span data-ttu-id="4323f-849">Имя</span><span class="sxs-lookup"><span data-stu-id="4323f-849">Name</span></span>|<span data-ttu-id="4323f-850">Тип</span><span class="sxs-lookup"><span data-stu-id="4323f-850">Type</span></span>|<span data-ttu-id="4323f-851">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="4323f-851">Attributes</span></span>|<span data-ttu-id="4323f-852">Описание</span><span class="sxs-lookup"><span data-stu-id="4323f-852">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="4323f-853">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="4323f-853">String &#124; Object</span></span>||<span data-ttu-id="4323f-p151">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="4323f-p151">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="4323f-856">**ИЛИ**</span><span class="sxs-lookup"><span data-stu-id="4323f-856">**OR**</span></span><br/><span data-ttu-id="4323f-p152">Объект, который содержит текст или данные вложения и функцию обратного вызова. Ниже представлено определение этого объекта.</span><span class="sxs-lookup"><span data-stu-id="4323f-p152">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="4323f-859">String</span><span class="sxs-lookup"><span data-stu-id="4323f-859">String</span></span>|<span data-ttu-id="4323f-860">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="4323f-860">&lt;optional&gt;</span></span>|<span data-ttu-id="4323f-p153">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="4323f-p153">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="4323f-863">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="4323f-863">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="4323f-864">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="4323f-864">&lt;optional&gt;</span></span>|<span data-ttu-id="4323f-865">Массив объектов JSON, представляющих собой вложенные файлы или элементы.</span><span class="sxs-lookup"><span data-stu-id="4323f-865">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="4323f-866">String</span><span class="sxs-lookup"><span data-stu-id="4323f-866">String</span></span>||<span data-ttu-id="4323f-p154">Указывает тип вложения. Допустимые значения: `file` для вложенного файла и `item` для вложенного элемента.</span><span class="sxs-lookup"><span data-stu-id="4323f-p154">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="4323f-869">Строка</span><span class="sxs-lookup"><span data-stu-id="4323f-869">String</span></span>||<span data-ttu-id="4323f-870">Строка, содержащая имя вложения, длиной до 255 символов.</span><span class="sxs-lookup"><span data-stu-id="4323f-870">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="4323f-871">String</span><span class="sxs-lookup"><span data-stu-id="4323f-871">String</span></span>||<span data-ttu-id="4323f-p155">Используется, только если свойству `type` задано значение `file`. URI расположения файла.</span><span class="sxs-lookup"><span data-stu-id="4323f-p155">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="4323f-874">Логический</span><span class="sxs-lookup"><span data-stu-id="4323f-874">Boolean</span></span>||<span data-ttu-id="4323f-p156">Используется, только если свойству `type` задано значение `file`. Значение `true` указывает на то, что вложение будет встроено в текст сообщения и не должно отображаться в списке вложений.</span><span class="sxs-lookup"><span data-stu-id="4323f-p156">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="4323f-877">String</span><span class="sxs-lookup"><span data-stu-id="4323f-877">String</span></span>||<span data-ttu-id="4323f-p157">Используется, только если свойству `type` задано значение `item`. Идентификатор вложения EWS. Это строка длиной до 100 символов.</span><span class="sxs-lookup"><span data-stu-id="4323f-p157">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="4323f-881">function</span><span class="sxs-lookup"><span data-stu-id="4323f-881">function</span></span>|<span data-ttu-id="4323f-882">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="4323f-882">&lt;optional&gt;</span></span>|<span data-ttu-id="4323f-883">По завершении работы метода функция, переданная параметру `callback`, вызывается с помощью одного параметра `asyncResult`, представляющего собой объект [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="4323f-883">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4323f-884">Требования</span><span class="sxs-lookup"><span data-stu-id="4323f-884">Requirements</span></span>

|<span data-ttu-id="4323f-885">Требование</span><span class="sxs-lookup"><span data-stu-id="4323f-885">Requirement</span></span>|<span data-ttu-id="4323f-886">Значение</span><span class="sxs-lookup"><span data-stu-id="4323f-886">Value</span></span>|
|---|---|
|[<span data-ttu-id="4323f-887">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="4323f-887">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4323f-888">1.0</span><span class="sxs-lookup"><span data-stu-id="4323f-888">1.0</span></span>|
|[<span data-ttu-id="4323f-889">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4323f-889">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4323f-890">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4323f-890">ReadItem</span></span>|
|[<span data-ttu-id="4323f-891">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4323f-891">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="4323f-892">Чтение</span><span class="sxs-lookup"><span data-stu-id="4323f-892">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="4323f-893">Примеры</span><span class="sxs-lookup"><span data-stu-id="4323f-893">Examples</span></span>

<span data-ttu-id="4323f-894">Приведенный ниже код передает строку в функцию `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="4323f-894">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="4323f-895">Ответ с пустым текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="4323f-895">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="4323f-896">Ответ только с текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="4323f-896">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="4323f-897">Ответ с текстом сообщения и вложенным файлом.</span><span class="sxs-lookup"><span data-stu-id="4323f-897">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="4323f-898">Ответ с текстом сообщения и вложенным элементом.</span><span class="sxs-lookup"><span data-stu-id="4323f-898">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="4323f-899">Ответ с текстом сообщения, вложенным файлом, вложенным элементом и обратным вызовом.</span><span class="sxs-lookup"><span data-stu-id="4323f-899">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="4323f-900">displayReplyForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="4323f-900">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="4323f-901">Отображает форму ответа, включающую только отправителя выбранного сообщения или организатора выбранной встречи.</span><span class="sxs-lookup"><span data-stu-id="4323f-901">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="4323f-902">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="4323f-902">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="4323f-903">В Outlook в Интернете форма ответа отображается в виде всплывающей формы в представлении с 3 либо 1 или 2 столбцами.</span><span class="sxs-lookup"><span data-stu-id="4323f-903">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="4323f-904">Если любой строковый параметр превышает указанные для него ограничения, `displayReplyForm` возвращает исключение.</span><span class="sxs-lookup"><span data-stu-id="4323f-904">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="4323f-p158">Если в параметре `formData.attachments` указаны вложения, Outlook в Интернете и классические клиенты пытаются скачать их и вложить в форму ответа. Если какие-либо вложения добавить не удается, в форме отображается сообщение об ошибке. Если сообщения об ошибках не предусмотрены, то они не отображаются.</span><span class="sxs-lookup"><span data-stu-id="4323f-p158">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4323f-908">Параметры</span><span class="sxs-lookup"><span data-stu-id="4323f-908">Parameters</span></span>

|<span data-ttu-id="4323f-909">Имя</span><span class="sxs-lookup"><span data-stu-id="4323f-909">Name</span></span>|<span data-ttu-id="4323f-910">Тип</span><span class="sxs-lookup"><span data-stu-id="4323f-910">Type</span></span>|<span data-ttu-id="4323f-911">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="4323f-911">Attributes</span></span>|<span data-ttu-id="4323f-912">Описание</span><span class="sxs-lookup"><span data-stu-id="4323f-912">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="4323f-913">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="4323f-913">String &#124; Object</span></span>||<span data-ttu-id="4323f-p159">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="4323f-p159">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="4323f-916">**ИЛИ**</span><span class="sxs-lookup"><span data-stu-id="4323f-916">**OR**</span></span><br/><span data-ttu-id="4323f-p160">Объект, который содержит текст или данные вложения и функцию обратного вызова. Ниже представлено определение этого объекта.</span><span class="sxs-lookup"><span data-stu-id="4323f-p160">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="4323f-919">String</span><span class="sxs-lookup"><span data-stu-id="4323f-919">String</span></span>|<span data-ttu-id="4323f-920">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="4323f-920">&lt;optional&gt;</span></span>|<span data-ttu-id="4323f-p161">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="4323f-p161">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="4323f-923">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="4323f-923">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="4323f-924">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="4323f-924">&lt;optional&gt;</span></span>|<span data-ttu-id="4323f-925">Массив объектов JSON, представляющих собой вложенные файлы или элементы.</span><span class="sxs-lookup"><span data-stu-id="4323f-925">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="4323f-926">String</span><span class="sxs-lookup"><span data-stu-id="4323f-926">String</span></span>||<span data-ttu-id="4323f-p162">Указывает тип вложения. Допустимые значения: `file` для вложенного файла и `item` для вложенного элемента.</span><span class="sxs-lookup"><span data-stu-id="4323f-p162">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="4323f-929">Строка</span><span class="sxs-lookup"><span data-stu-id="4323f-929">String</span></span>||<span data-ttu-id="4323f-930">Строка, содержащая имя вложения, длиной до 255 символов.</span><span class="sxs-lookup"><span data-stu-id="4323f-930">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="4323f-931">Строка</span><span class="sxs-lookup"><span data-stu-id="4323f-931">String</span></span>||<span data-ttu-id="4323f-p163">Используется, только если свойству `type` задано значение `file`. URI расположения файла.</span><span class="sxs-lookup"><span data-stu-id="4323f-p163">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="4323f-934">Логический</span><span class="sxs-lookup"><span data-stu-id="4323f-934">Boolean</span></span>||<span data-ttu-id="4323f-p164">Используется, только если свойству `type` задано значение `file`. Значение `true` указывает на то, что вложение будет встроено в текст сообщения и не должно отображаться в списке вложений.</span><span class="sxs-lookup"><span data-stu-id="4323f-p164">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="4323f-937">String</span><span class="sxs-lookup"><span data-stu-id="4323f-937">String</span></span>||<span data-ttu-id="4323f-p165">Используется, только если свойству `type` задано значение `item`. Идентификатор вложения EWS. Это строка длиной до 100 символов.</span><span class="sxs-lookup"><span data-stu-id="4323f-p165">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="4323f-941">function</span><span class="sxs-lookup"><span data-stu-id="4323f-941">function</span></span>|<span data-ttu-id="4323f-942">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="4323f-942">&lt;optional&gt;</span></span>|<span data-ttu-id="4323f-943">По завершении работы метода функция, переданная параметру `callback`, вызывается с помощью одного параметра `asyncResult`, представляющего собой объект [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="4323f-943">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4323f-944">Требования</span><span class="sxs-lookup"><span data-stu-id="4323f-944">Requirements</span></span>

|<span data-ttu-id="4323f-945">Требование</span><span class="sxs-lookup"><span data-stu-id="4323f-945">Requirement</span></span>|<span data-ttu-id="4323f-946">Значение</span><span class="sxs-lookup"><span data-stu-id="4323f-946">Value</span></span>|
|---|---|
|[<span data-ttu-id="4323f-947">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="4323f-947">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4323f-948">1.0</span><span class="sxs-lookup"><span data-stu-id="4323f-948">1.0</span></span>|
|[<span data-ttu-id="4323f-949">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4323f-949">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4323f-950">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4323f-950">ReadItem</span></span>|
|[<span data-ttu-id="4323f-951">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4323f-951">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="4323f-952">Чтение</span><span class="sxs-lookup"><span data-stu-id="4323f-952">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="4323f-953">Примеры</span><span class="sxs-lookup"><span data-stu-id="4323f-953">Examples</span></span>

<span data-ttu-id="4323f-954">Приведенный ниже код передает строку в функцию `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="4323f-954">The following code passes a string to the `displayReplyForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="4323f-955">Ответ с пустым текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="4323f-955">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="4323f-956">Ответ только с текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="4323f-956">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="4323f-957">Ответ с текстом сообщения и вложенным файлом.</span><span class="sxs-lookup"><span data-stu-id="4323f-957">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="4323f-958">Ответ с текстом сообщения и вложенным элементом.</span><span class="sxs-lookup"><span data-stu-id="4323f-958">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="4323f-959">Ответ с текстом сообщения, вложенным файлом, вложенным элементом и обратным вызовом.</span><span class="sxs-lookup"><span data-stu-id="4323f-959">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlookofficeentitiesviewoutlook-js-17"></a><span data-ttu-id="4323f-960">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.7)}</span><span class="sxs-lookup"><span data-stu-id="4323f-960">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.7)}</span></span>

<span data-ttu-id="4323f-961">Получает сущности, обнаруженные в теле выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="4323f-961">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="4323f-962">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="4323f-962">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="4323f-963">Требования</span><span class="sxs-lookup"><span data-stu-id="4323f-963">Requirements</span></span>

|<span data-ttu-id="4323f-964">Требование</span><span class="sxs-lookup"><span data-stu-id="4323f-964">Requirement</span></span>|<span data-ttu-id="4323f-965">Значение</span><span class="sxs-lookup"><span data-stu-id="4323f-965">Value</span></span>|
|---|---|
|[<span data-ttu-id="4323f-966">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="4323f-966">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4323f-967">1.0</span><span class="sxs-lookup"><span data-stu-id="4323f-967">1.0</span></span>|
|[<span data-ttu-id="4323f-968">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4323f-968">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4323f-969">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4323f-969">ReadItem</span></span>|
|[<span data-ttu-id="4323f-970">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4323f-970">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="4323f-971">Чтение</span><span class="sxs-lookup"><span data-stu-id="4323f-971">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="4323f-972">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="4323f-972">Returns:</span></span>

<span data-ttu-id="4323f-973">Тип: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="4323f-973">Type: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.7)</span></span>

##### <a name="example"></a><span data-ttu-id="4323f-974">Пример</span><span class="sxs-lookup"><span data-stu-id="4323f-974">Example</span></span>

<span data-ttu-id="4323f-975">Ниже приведен пример получения доступа к сущностям контактов в теле текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="4323f-975">The following example accesses the contacts entities in the current item's body.</span></span>

```js
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

<br>

---
---

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-17meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-17phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-17tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-17"></a><span data-ttu-id="4323f-976">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.7)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.7)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.7)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.7))>}</span><span class="sxs-lookup"><span data-stu-id="4323f-976">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.7)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.7)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.7)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.7))>}</span></span>

<span data-ttu-id="4323f-977">Получает массив всех сущностей указанного типа, обнаруженных в теле выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="4323f-977">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="4323f-978">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="4323f-978">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4323f-979">Параметры</span><span class="sxs-lookup"><span data-stu-id="4323f-979">Parameters</span></span>

|<span data-ttu-id="4323f-980">Имя</span><span class="sxs-lookup"><span data-stu-id="4323f-980">Name</span></span>|<span data-ttu-id="4323f-981">Тип</span><span class="sxs-lookup"><span data-stu-id="4323f-981">Type</span></span>|<span data-ttu-id="4323f-982">Описание</span><span class="sxs-lookup"><span data-stu-id="4323f-982">Description</span></span>|
|---|---|---|
|`entityType`|[<span data-ttu-id="4323f-983">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="4323f-983">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook/office.mailboxenums.entitytype?view=outlook-js-1.7)|<span data-ttu-id="4323f-984">Одно из значений перечисления EntityType.</span><span class="sxs-lookup"><span data-stu-id="4323f-984">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4323f-985">Requirements</span><span class="sxs-lookup"><span data-stu-id="4323f-985">Requirements</span></span>

|<span data-ttu-id="4323f-986">Требование</span><span class="sxs-lookup"><span data-stu-id="4323f-986">Requirement</span></span>|<span data-ttu-id="4323f-987">Значение</span><span class="sxs-lookup"><span data-stu-id="4323f-987">Value</span></span>|
|---|---|
|[<span data-ttu-id="4323f-988">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="4323f-988">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4323f-989">1.0</span><span class="sxs-lookup"><span data-stu-id="4323f-989">1.0</span></span>|
|[<span data-ttu-id="4323f-990">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4323f-990">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4323f-991">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="4323f-991">Restricted</span></span>|
|[<span data-ttu-id="4323f-992">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4323f-992">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="4323f-993">Чтение</span><span class="sxs-lookup"><span data-stu-id="4323f-993">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="4323f-994">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="4323f-994">Returns:</span></span>

<span data-ttu-id="4323f-995">Если значение, переданное в `entityType`, не является допустимым членом перечисления `EntityType`, метод возвращает значение NULL.</span><span class="sxs-lookup"><span data-stu-id="4323f-995">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="4323f-996">Если в теле элемента отсутствуют сущности указанного типа, метод возвращает пустой массив.</span><span class="sxs-lookup"><span data-stu-id="4323f-996">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="4323f-997">В противном случае тип объектов в возвращаемом массиве зависит от типа сущности, запрошенной в параметре `entityType`.</span><span class="sxs-lookup"><span data-stu-id="4323f-997">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="4323f-998">Хотя минимальный уровень разрешений для использования этого метода — **Restricted**, для некоторых типов сущностей требуется доступ на уровне **ReadItem**, как указано в приведенной ниже таблице.</span><span class="sxs-lookup"><span data-stu-id="4323f-998">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

|<span data-ttu-id="4323f-999">Значение параметра `entityType`</span><span class="sxs-lookup"><span data-stu-id="4323f-999">Value of `entityType`</span></span>|<span data-ttu-id="4323f-1000">Тип объектов в возвращаемом массиве</span><span class="sxs-lookup"><span data-stu-id="4323f-1000">Type of objects in returned array</span></span>|<span data-ttu-id="4323f-1001">Необходимый уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4323f-1001">Required Permission Level</span></span>|
|---|---|---|
|`Address`|<span data-ttu-id="4323f-1002">String</span><span class="sxs-lookup"><span data-stu-id="4323f-1002">String</span></span>|<span data-ttu-id="4323f-1003">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="4323f-1003">**Restricted**</span></span>|
|`Contact`|<span data-ttu-id="4323f-1004">Contact</span><span class="sxs-lookup"><span data-stu-id="4323f-1004">Contact</span></span>|<span data-ttu-id="4323f-1005">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="4323f-1005">**ReadItem**</span></span>|
|`EmailAddress`|<span data-ttu-id="4323f-1006">String</span><span class="sxs-lookup"><span data-stu-id="4323f-1006">String</span></span>|<span data-ttu-id="4323f-1007">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="4323f-1007">**ReadItem**</span></span>|
|`MeetingSuggestion`|<span data-ttu-id="4323f-1008">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="4323f-1008">MeetingSuggestion</span></span>|<span data-ttu-id="4323f-1009">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="4323f-1009">**ReadItem**</span></span>|
|`PhoneNumber`|<span data-ttu-id="4323f-1010">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="4323f-1010">PhoneNumber</span></span>|<span data-ttu-id="4323f-1011">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="4323f-1011">**Restricted**</span></span>|
|`TaskSuggestion`|<span data-ttu-id="4323f-1012">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="4323f-1012">TaskSuggestion</span></span>|<span data-ttu-id="4323f-1013">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="4323f-1013">**ReadItem**</span></span>|
|`URL`|<span data-ttu-id="4323f-1014">String</span><span class="sxs-lookup"><span data-stu-id="4323f-1014">String</span></span>|<span data-ttu-id="4323f-1015">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="4323f-1015">**Restricted**</span></span>|

<span data-ttu-id="4323f-1016">Тип: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.7)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.7)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.7)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.7))></span><span class="sxs-lookup"><span data-stu-id="4323f-1016">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.7)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.7)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.7)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.7))></span></span>

##### <a name="example"></a><span data-ttu-id="4323f-1017">Пример</span><span class="sxs-lookup"><span data-stu-id="4323f-1017">Example</span></span>

<span data-ttu-id="4323f-1018">В примере ниже показано, как получить доступ к массиву строк, которые представляют собой почтовые адреса в теле текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="4323f-1018">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-17meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-17phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-17tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-17"></a><span data-ttu-id="4323f-1019">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.7)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.7)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.7)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.7))>}</span><span class="sxs-lookup"><span data-stu-id="4323f-1019">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.7)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.7)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.7)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.7))>}</span></span>

<span data-ttu-id="4323f-1020">Возвращает известные сущности в выбранном элементе, которые проходят через именованный фильтр, определяемый в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="4323f-1020">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="4323f-1021">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="4323f-1021">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="4323f-1022">Метод `getFilteredEntitiesByName` возвращает сущности, соответствующие регулярному выражению, которое определяется в элементе правила [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) в XML-файле манифеста, с использованием указанного значения элемента `FilterName`.</span><span class="sxs-lookup"><span data-stu-id="4323f-1022">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4323f-1023">Параметры</span><span class="sxs-lookup"><span data-stu-id="4323f-1023">Parameters</span></span>

|<span data-ttu-id="4323f-1024">Имя</span><span class="sxs-lookup"><span data-stu-id="4323f-1024">Name</span></span>|<span data-ttu-id="4323f-1025">Тип</span><span class="sxs-lookup"><span data-stu-id="4323f-1025">Type</span></span>|<span data-ttu-id="4323f-1026">Описание</span><span class="sxs-lookup"><span data-stu-id="4323f-1026">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="4323f-1027">String</span><span class="sxs-lookup"><span data-stu-id="4323f-1027">String</span></span>|<span data-ttu-id="4323f-1028">Имя элемента правила `ItemHasKnownEntity`, определяющее соответствующий фильтр.</span><span class="sxs-lookup"><span data-stu-id="4323f-1028">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4323f-1029">Requirements</span><span class="sxs-lookup"><span data-stu-id="4323f-1029">Requirements</span></span>

|<span data-ttu-id="4323f-1030">Требование</span><span class="sxs-lookup"><span data-stu-id="4323f-1030">Requirement</span></span>|<span data-ttu-id="4323f-1031">Значение</span><span class="sxs-lookup"><span data-stu-id="4323f-1031">Value</span></span>|
|---|---|
|[<span data-ttu-id="4323f-1032">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="4323f-1032">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4323f-1033">1.0</span><span class="sxs-lookup"><span data-stu-id="4323f-1033">1.0</span></span>|
|[<span data-ttu-id="4323f-1034">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4323f-1034">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4323f-1035">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4323f-1035">ReadItem</span></span>|
|[<span data-ttu-id="4323f-1036">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4323f-1036">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="4323f-1037">Чтение</span><span class="sxs-lookup"><span data-stu-id="4323f-1037">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="4323f-1038">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="4323f-1038">Returns:</span></span>

<span data-ttu-id="4323f-p167">Если в манифесте нет элемента `ItemHasKnownEntity` со значением `FilterName`, соответствующим параметру `name`, метод возвращает `null`. Если параметр `name` соответствует элементу `ItemHasKnownEntity` в манифесте, но при этом в текущем элементе нет соответствующих сущностей, метод возвращает пустой массив.</span><span class="sxs-lookup"><span data-stu-id="4323f-p167">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="4323f-1041">Тип: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.7)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.7)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.7)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.7))></span><span class="sxs-lookup"><span data-stu-id="4323f-1041">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.7)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.7)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.7)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.7))></span></span>

<br>

---
---

#### <a name="getregexmatches--object"></a><span data-ttu-id="4323f-1042">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="4323f-1042">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="4323f-1043">Возвращает строковые значения в выбранном элементе, которые соответствуют регулярным выражениям, определенным в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="4323f-1043">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="4323f-1044">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="4323f-1044">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="4323f-p168">Метод `getRegExMatches` возвращает строки, соответствующие регулярному выражению, которое определяется в каждом элементе правила `ItemHasRegularExpressionMatch` или `ItemHasKnownEntity` в XML-файле манифеста. Для правила `ItemHasRegularExpressionMatch` соответствующую строку должно содержать свойство элемента, указанного этим правилом. Простой тип `PropertyName` определяет поддерживаемые свойства.</span><span class="sxs-lookup"><span data-stu-id="4323f-p168">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="4323f-1048">Например, рассмотрим манифест надстройки, который содержит указанный ниже элемент `Rule`.</span><span class="sxs-lookup"><span data-stu-id="4323f-1048">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="4323f-1049">Объект, возвращаемый методом `getRegExMatches`, будет содержать два свойства: `fruits` и `veggies`.</span><span class="sxs-lookup"><span data-stu-id="4323f-1049">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="4323f-p169">Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты. Лучше используйте метод [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.7#getasync-coerciontype--options--callback-) для этого.</span><span class="sxs-lookup"><span data-stu-id="4323f-p169">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.7#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="4323f-1053">Requirements</span><span class="sxs-lookup"><span data-stu-id="4323f-1053">Requirements</span></span>

|<span data-ttu-id="4323f-1054">Требование</span><span class="sxs-lookup"><span data-stu-id="4323f-1054">Requirement</span></span>|<span data-ttu-id="4323f-1055">Значение</span><span class="sxs-lookup"><span data-stu-id="4323f-1055">Value</span></span>|
|---|---|
|[<span data-ttu-id="4323f-1056">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="4323f-1056">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4323f-1057">1.0</span><span class="sxs-lookup"><span data-stu-id="4323f-1057">1.0</span></span>|
|[<span data-ttu-id="4323f-1058">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4323f-1058">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4323f-1059">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4323f-1059">ReadItem</span></span>|
|[<span data-ttu-id="4323f-1060">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4323f-1060">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="4323f-1061">Чтение</span><span class="sxs-lookup"><span data-stu-id="4323f-1061">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="4323f-1062">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="4323f-1062">Returns:</span></span>

<span data-ttu-id="4323f-p170">Объект, содержащий массив строк, которые соответствуют регулярным выражениям, определяемым в XML-файле манифеста. Имя каждого массива равно соответствующему значению атрибута `RegExName` подходящего правила `ItemHasRegularExpressionMatch` или атрибута `FilterName` соответствующего правила `ItemHasKnownEntity`.</span><span class="sxs-lookup"><span data-stu-id="4323f-p170">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<span data-ttu-id="4323f-1065">Тип: Object</span><span class="sxs-lookup"><span data-stu-id="4323f-1065">Type: Object</span></span>

##### <a name="example"></a><span data-ttu-id="4323f-1066">Пример</span><span class="sxs-lookup"><span data-stu-id="4323f-1066">Example</span></span>

<span data-ttu-id="4323f-1067">В приведенном ниже примере показано, как получить доступ к массиву совпадений с элементами `fruits` и `veggies` правил активации регулярных выражений, указанными в манифесте.</span><span class="sxs-lookup"><span data-stu-id="4323f-1067">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```js
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

<br>

---
---

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="4323f-1068">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="4323f-1068">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="4323f-1069">Возвращает строковые значения в выбранном элементе, которые соответствуют именованному регулярному выражению, определенному в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="4323f-1069">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="4323f-1070">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="4323f-1070">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="4323f-1071">Метод `getRegExMatchesByName` возвращает строки, соответствующие регулярному выражению, которое определяется в элементе правила `ItemHasRegularExpressionMatch` в XML-файле манифеста, с использованием указанного значения элемента `RegExName`.</span><span class="sxs-lookup"><span data-stu-id="4323f-1071">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="4323f-p171">Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты.</span><span class="sxs-lookup"><span data-stu-id="4323f-p171">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4323f-1074">Параметры</span><span class="sxs-lookup"><span data-stu-id="4323f-1074">Parameters</span></span>

|<span data-ttu-id="4323f-1075">Имя</span><span class="sxs-lookup"><span data-stu-id="4323f-1075">Name</span></span>|<span data-ttu-id="4323f-1076">Тип</span><span class="sxs-lookup"><span data-stu-id="4323f-1076">Type</span></span>|<span data-ttu-id="4323f-1077">Описание</span><span class="sxs-lookup"><span data-stu-id="4323f-1077">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="4323f-1078">String</span><span class="sxs-lookup"><span data-stu-id="4323f-1078">String</span></span>|<span data-ttu-id="4323f-1079">Имя элемента правила `ItemHasRegularExpressionMatch`, определяющее соответствующий фильтр.</span><span class="sxs-lookup"><span data-stu-id="4323f-1079">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4323f-1080">Requirements</span><span class="sxs-lookup"><span data-stu-id="4323f-1080">Requirements</span></span>

|<span data-ttu-id="4323f-1081">Требование</span><span class="sxs-lookup"><span data-stu-id="4323f-1081">Requirement</span></span>|<span data-ttu-id="4323f-1082">Значение</span><span class="sxs-lookup"><span data-stu-id="4323f-1082">Value</span></span>|
|---|---|
|[<span data-ttu-id="4323f-1083">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="4323f-1083">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4323f-1084">1.0</span><span class="sxs-lookup"><span data-stu-id="4323f-1084">1.0</span></span>|
|[<span data-ttu-id="4323f-1085">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4323f-1085">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4323f-1086">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4323f-1086">ReadItem</span></span>|
|[<span data-ttu-id="4323f-1087">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4323f-1087">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="4323f-1088">Чтение</span><span class="sxs-lookup"><span data-stu-id="4323f-1088">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="4323f-1089">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="4323f-1089">Returns:</span></span>

<span data-ttu-id="4323f-1090">Массив строк, соответствующих регулярному выражению, определяемому в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="4323f-1090">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<span data-ttu-id="4323f-1091">Тип: Array.< String ></span><span class="sxs-lookup"><span data-stu-id="4323f-1091">Type: Array.< String ></span></span>

##### <a name="example"></a><span data-ttu-id="4323f-1092">Пример</span><span class="sxs-lookup"><span data-stu-id="4323f-1092">Example</span></span>

```js
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

<br>

---
---

#### <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="4323f-1093">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="4323f-1093">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="4323f-1094">Асинхронно возвращает данные, выбранные в теме или тексте сообщения.</span><span class="sxs-lookup"><span data-stu-id="4323f-1094">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="4323f-1095">Если выделенный фрагмент отсутствует, но курсор находится в основном тексте или теме, метод возвращает пустую строку для выбранных данных.</span><span class="sxs-lookup"><span data-stu-id="4323f-1095">If there is no selection but the cursor is in the body or subject, the method returns an empty string for the selected data.</span></span> <span data-ttu-id="4323f-1096">Если выбраны не текст и не тема, метод возвращает ошибку `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="4323f-1096">If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

> [!NOTE]
> <span data-ttu-id="4323f-1097">В Outlook в Интернете метод возвращает строку null, если текст не выделен, но курсор находится в тексте.</span><span class="sxs-lookup"><span data-stu-id="4323f-1097">In Outlook on the web, the method returns the string "null" if no text is selected but the cursor is in the body.</span></span> <span data-ttu-id="4323f-1098">Чтобы проверить эту ситуацию, ознакомьтесь с приведенным далее в этом разделе.</span><span class="sxs-lookup"><span data-stu-id="4323f-1098">To check for this situation, see the example later in this section.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4323f-1099">Параметры</span><span class="sxs-lookup"><span data-stu-id="4323f-1099">Parameters</span></span>

|<span data-ttu-id="4323f-1100">Имя</span><span class="sxs-lookup"><span data-stu-id="4323f-1100">Name</span></span>|<span data-ttu-id="4323f-1101">Тип</span><span class="sxs-lookup"><span data-stu-id="4323f-1101">Type</span></span>|<span data-ttu-id="4323f-1102">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="4323f-1102">Attributes</span></span>|<span data-ttu-id="4323f-1103">Описание</span><span class="sxs-lookup"><span data-stu-id="4323f-1103">Description</span></span>|
|---|---|---|---|
|`coercionType`|[<span data-ttu-id="4323f-1104">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="4323f-1104">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="4323f-p174">Запрашивает формат данных. Если задано значение Text, метод возвращает обычный текст как строку, удаляя все имеющиеся HTML-теги. Если задано значение HTML, метод возвращает выделенный текст (обычный текст или HTML).</span><span class="sxs-lookup"><span data-stu-id="4323f-p174">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`|<span data-ttu-id="4323f-1108">Object</span><span class="sxs-lookup"><span data-stu-id="4323f-1108">Object</span></span>|<span data-ttu-id="4323f-1109">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="4323f-1109">&lt;optional&gt;</span></span>|<span data-ttu-id="4323f-1110">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="4323f-1110">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="4323f-1111">Object</span><span class="sxs-lookup"><span data-stu-id="4323f-1111">Object</span></span>|<span data-ttu-id="4323f-1112">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="4323f-1112">&lt;optional&gt;</span></span>|<span data-ttu-id="4323f-1113">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="4323f-1113">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="4323f-1114">функция</span><span class="sxs-lookup"><span data-stu-id="4323f-1114">function</span></span>||<span data-ttu-id="4323f-1115">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="4323f-1115">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="4323f-1116">Чтобы получить доступ к выбранным данным из метода обратного вызова, вызовите `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="4323f-1116">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="4323f-1117">Для доступа к исходному свойству, представляющему собой источник выбранных данных, вызовите параметр `asyncResult.value.sourceProperty`, который может иметь значение `body` или `subject`.</span><span class="sxs-lookup"><span data-stu-id="4323f-1117">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4323f-1118">Requirements</span><span class="sxs-lookup"><span data-stu-id="4323f-1118">Requirements</span></span>

|<span data-ttu-id="4323f-1119">Требование</span><span class="sxs-lookup"><span data-stu-id="4323f-1119">Requirement</span></span>|<span data-ttu-id="4323f-1120">Значение</span><span class="sxs-lookup"><span data-stu-id="4323f-1120">Value</span></span>|
|---|---|
|[<span data-ttu-id="4323f-1121">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="4323f-1121">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4323f-1122">1.2</span><span class="sxs-lookup"><span data-stu-id="4323f-1122">1.2</span></span>|
|[<span data-ttu-id="4323f-1123">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4323f-1123">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4323f-1124">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4323f-1124">ReadItem</span></span>|
|[<span data-ttu-id="4323f-1125">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4323f-1125">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="4323f-1126">Создание</span><span class="sxs-lookup"><span data-stu-id="4323f-1126">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="4323f-1127">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="4323f-1127">Returns:</span></span>

<span data-ttu-id="4323f-1128">Выбранные данные в виде строки с форматом, определенным в параметре `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="4323f-1128">The selected data as a string with format determined by `coercionType`.</span></span>

<span data-ttu-id="4323f-1129">Тип: строка</span><span class="sxs-lookup"><span data-stu-id="4323f-1129">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="4323f-1130">Пример</span><span class="sxs-lookup"><span data-stu-id="4323f-1130">Example</span></span>

```js
// Get selected data.
Office.initialize = function () {
  Office.context.mailbox.item.getSelectedDataAsync(Office.CoercionType.Text, {}, getCallback);
};

function getCallback(asyncResult) {
  var text = asyncResult.value.data;
  var prop = asyncResult.value.sourceProperty;

  // Handle where Outlook on the web erroneously returns "null" instead of empty string.
  if (Office.context.mailbox.diagnostics.hostName === 'OutlookWebApp'
      && asyncResult.value.endPosition === asyncResult.value.startPosition) {
    text = "";
  }

  console.log("Selected text in " + prop + ": " + text);
}
```

<br>

---
---

#### <a name="getselectedentities--entitiesjavascriptapioutlookofficeentitiesviewoutlook-js-17"></a><span data-ttu-id="4323f-1131">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.7)}</span><span class="sxs-lookup"><span data-stu-id="4323f-1131">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.7)}</span></span>

<span data-ttu-id="4323f-1132">Возвращает сущности, найденные в выделенном совпадении, выбранном пользователем.</span><span class="sxs-lookup"><span data-stu-id="4323f-1132">Gets the entities found in a highlighted match a user has selected.</span></span> <span data-ttu-id="4323f-1133">Выделенные совпадения применяются к [контекстным надстройкам](/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="4323f-1133">Highlighted matches apply to [contextual add-ins](/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="4323f-1134">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="4323f-1134">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="4323f-1135">Требования</span><span class="sxs-lookup"><span data-stu-id="4323f-1135">Requirements</span></span>

|<span data-ttu-id="4323f-1136">Требование</span><span class="sxs-lookup"><span data-stu-id="4323f-1136">Requirement</span></span>|<span data-ttu-id="4323f-1137">Значение</span><span class="sxs-lookup"><span data-stu-id="4323f-1137">Value</span></span>|
|---|---|
|[<span data-ttu-id="4323f-1138">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="4323f-1138">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4323f-1139">1.6</span><span class="sxs-lookup"><span data-stu-id="4323f-1139">1.6</span></span>|
|[<span data-ttu-id="4323f-1140">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4323f-1140">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4323f-1141">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4323f-1141">ReadItem</span></span>|
|[<span data-ttu-id="4323f-1142">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4323f-1142">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="4323f-1143">Чтение</span><span class="sxs-lookup"><span data-stu-id="4323f-1143">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="4323f-1144">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="4323f-1144">Returns:</span></span>

<span data-ttu-id="4323f-1145">Тип: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="4323f-1145">Type: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.7)</span></span>

##### <a name="example"></a><span data-ttu-id="4323f-1146">Пример</span><span class="sxs-lookup"><span data-stu-id="4323f-1146">Example</span></span>

<span data-ttu-id="4323f-1147">В приведенном ниже примере показано, как получить доступ к сущностям адресов в выделенном совпадении, выбранном пользователем.</span><span class="sxs-lookup"><span data-stu-id="4323f-1147">The following example accesses the addresses entities in the highlighted match selected by the user.</span></span>

```js
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

<br>

---
---

#### <a name="getselectedregexmatches--object"></a><span data-ttu-id="4323f-1148">getSelectedRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="4323f-1148">getSelectedRegExMatches() → {Object}</span></span>

<span data-ttu-id="4323f-p177">Возвращает строковые значения в выделенном совпадении, которые соответствуют регулярным выражениям, определенным в XML-файле манифеста. Выделенные совпадения применяются к [контекстным надстройкам](/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="4323f-p177">Returns string values in a highlighted match that match the regular expressions defined in the manifest XML file. Highlighted matches apply to [contextual add-ins](/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="4323f-1151">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="4323f-1151">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="4323f-p178">Метод `getSelectedRegExMatches` возвращает строки, соответствующие регулярному выражению, которое определяется в каждом элементе правила `ItemHasRegularExpressionMatch` или `ItemHasKnownEntity` в XML-файле манифеста. Для правила `ItemHasRegularExpressionMatch` соответствующую строку должно содержать свойство элемента, указанного этим правилом. Простой тип `PropertyName` определяет поддерживаемые свойства.</span><span class="sxs-lookup"><span data-stu-id="4323f-p178">The `getSelectedRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="4323f-1155">Например, рассмотрим манифест надстройки, который содержит указанный ниже элемент `Rule`.</span><span class="sxs-lookup"><span data-stu-id="4323f-1155">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="4323f-1156">Объект, возвращаемый методом `getRegExMatches`, будет содержать два свойства: `fruits` и `veggies`.</span><span class="sxs-lookup"><span data-stu-id="4323f-1156">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="4323f-p179">Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты. Лучше используйте метод [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.7#getasync-coerciontype--options--callback-) для этого.</span><span class="sxs-lookup"><span data-stu-id="4323f-p179">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.7#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="4323f-1160">Requirements</span><span class="sxs-lookup"><span data-stu-id="4323f-1160">Requirements</span></span>

|<span data-ttu-id="4323f-1161">Требование</span><span class="sxs-lookup"><span data-stu-id="4323f-1161">Requirement</span></span>|<span data-ttu-id="4323f-1162">Значение</span><span class="sxs-lookup"><span data-stu-id="4323f-1162">Value</span></span>|
|---|---|
|[<span data-ttu-id="4323f-1163">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="4323f-1163">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4323f-1164">1.6</span><span class="sxs-lookup"><span data-stu-id="4323f-1164">1.6</span></span>|
|[<span data-ttu-id="4323f-1165">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4323f-1165">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4323f-1166">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4323f-1166">ReadItem</span></span>|
|[<span data-ttu-id="4323f-1167">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4323f-1167">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="4323f-1168">Чтение</span><span class="sxs-lookup"><span data-stu-id="4323f-1168">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="4323f-1169">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="4323f-1169">Returns:</span></span>

<span data-ttu-id="4323f-p180">Объект, содержащий массив строк, которые соответствуют регулярным выражениям, определяемым в XML-файле манифеста. Имя каждого массива равно соответствующему значению атрибута `RegExName` подходящего правила `ItemHasRegularExpressionMatch` или атрибута `FilterName` соответствующего правила `ItemHasKnownEntity`.</span><span class="sxs-lookup"><span data-stu-id="4323f-p180">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

##### <a name="example"></a><span data-ttu-id="4323f-1172">Пример</span><span class="sxs-lookup"><span data-stu-id="4323f-1172">Example</span></span>

<span data-ttu-id="4323f-1173">В приведенном ниже примере показано, как получить доступ к массиву совпадений с элементами `fruits` и `veggies` правил активации регулярных выражений, указанными в манифесте.</span><span class="sxs-lookup"><span data-stu-id="4323f-1173">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```js
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

<br>

---
---

#### <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="4323f-1174">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="4323f-1174">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="4323f-1175">Асинхронно загружает настраиваемые свойства для надстройки для выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="4323f-1175">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="4323f-p181">Настраиваемые свойства сохраняются в виде пар "ключ-значение" для каждого приложения и каждого элемента. Этот метод возвращает объект `CustomProperties` при обратном вызове, который предоставляет методы для доступа к настраиваемым свойствам, характерным для текущего элемента и текущей надстройки. Настраиваемые свойства не шифруются для элемента, поэтому этот способ хранения не является безопасным.</span><span class="sxs-lookup"><span data-stu-id="4323f-p181">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4323f-1179">Параметры</span><span class="sxs-lookup"><span data-stu-id="4323f-1179">Parameters</span></span>

|<span data-ttu-id="4323f-1180">Имя</span><span class="sxs-lookup"><span data-stu-id="4323f-1180">Name</span></span>|<span data-ttu-id="4323f-1181">Тип</span><span class="sxs-lookup"><span data-stu-id="4323f-1181">Type</span></span>|<span data-ttu-id="4323f-1182">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="4323f-1182">Attributes</span></span>|<span data-ttu-id="4323f-1183">Описание</span><span class="sxs-lookup"><span data-stu-id="4323f-1183">Description</span></span>|
|---|---|---|---|
|`callback`|<span data-ttu-id="4323f-1184">function</span><span class="sxs-lookup"><span data-stu-id="4323f-1184">function</span></span>||<span data-ttu-id="4323f-1185">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="4323f-1185">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="4323f-1186">Настраиваемые свойства предоставляются в виде объекта [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.7) в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="4323f-1186">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.7) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="4323f-1187">Этот объект позволяет получить, задать и удалить настраиваемые свойства для элемента, а также сохранить изменения, внесенные в набор настраиваемых свойств, на сервере.</span><span class="sxs-lookup"><span data-stu-id="4323f-1187">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`|<span data-ttu-id="4323f-1188">Объект</span><span class="sxs-lookup"><span data-stu-id="4323f-1188">Object</span></span>|<span data-ttu-id="4323f-1189">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="4323f-1189">&lt;optional&gt;</span></span>|<span data-ttu-id="4323f-1190">Разработчики могут указать любой объект, к которому необходимо получить доступ, в функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="4323f-1190">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="4323f-1191">Доступ к этому объекту можно получить с помощью свойства `asyncResult.asyncContext` в функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="4323f-1191">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4323f-1192">Requirements</span><span class="sxs-lookup"><span data-stu-id="4323f-1192">Requirements</span></span>

|<span data-ttu-id="4323f-1193">Требование</span><span class="sxs-lookup"><span data-stu-id="4323f-1193">Requirement</span></span>|<span data-ttu-id="4323f-1194">Значение</span><span class="sxs-lookup"><span data-stu-id="4323f-1194">Value</span></span>|
|---|---|
|[<span data-ttu-id="4323f-1195">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="4323f-1195">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4323f-1196">1.0</span><span class="sxs-lookup"><span data-stu-id="4323f-1196">1.0</span></span>|
|[<span data-ttu-id="4323f-1197">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4323f-1197">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4323f-1198">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4323f-1198">ReadItem</span></span>|
|[<span data-ttu-id="4323f-1199">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4323f-1199">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="4323f-1200">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="4323f-1200">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4323f-1201">Пример</span><span class="sxs-lookup"><span data-stu-id="4323f-1201">Example</span></span>

<span data-ttu-id="4323f-p184">Приведенный ниже пример кода показывает, как асинхронно загружать настраиваемые свойства, характерные для текущего элемента, с помощью метода `loadCustomPropertiesAsync`. Этот пример также показывает, как сохранять эти свойства на сервере с помощью метода `CustomProperties.saveAsync`. После загрузки настраиваемых свойств в этом примере кода метод `CustomProperties.get` используется для считывания настраиваемого свойства `myProp`, метод `CustomProperties.set` — для записи настраиваемого свойства `otherProp`, а метод `saveAsync` — для сохранения настраиваемых свойств.</span><span class="sxs-lookup"><span data-stu-id="4323f-p184">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

#### <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="4323f-1205">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="4323f-1205">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="4323f-1206">Удаляет вложение из сообщения или встречи.</span><span class="sxs-lookup"><span data-stu-id="4323f-1206">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="4323f-1207">Метод `removeAttachmentAsync` удаляет из элемента вложение с указанным идентификатором.</span><span class="sxs-lookup"><span data-stu-id="4323f-1207">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="4323f-1208">Идентификатор вложения рекомендуется использовать для удаления вложения, только если оно добавлено тем же почтовым приложением в ходе текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="4323f-1208">As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session.</span></span> <span data-ttu-id="4323f-1209">В Outlook в Интернете и на мобильных устройствах идентификатор вложения действителен только в течение одного сеанса.</span><span class="sxs-lookup"><span data-stu-id="4323f-1209">In Outlook on the web and mobile devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="4323f-1210">Сеанс завершается, когда пользователь закрывает приложение или начинает создавать элемент во встроенной форме, а затем переходит из формы в отдельное окно.</span><span class="sxs-lookup"><span data-stu-id="4323f-1210">A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4323f-1211">Параметры</span><span class="sxs-lookup"><span data-stu-id="4323f-1211">Parameters</span></span>

|<span data-ttu-id="4323f-1212">Имя</span><span class="sxs-lookup"><span data-stu-id="4323f-1212">Name</span></span>|<span data-ttu-id="4323f-1213">Тип</span><span class="sxs-lookup"><span data-stu-id="4323f-1213">Type</span></span>|<span data-ttu-id="4323f-1214">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="4323f-1214">Attributes</span></span>|<span data-ttu-id="4323f-1215">Описание</span><span class="sxs-lookup"><span data-stu-id="4323f-1215">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="4323f-1216">String</span><span class="sxs-lookup"><span data-stu-id="4323f-1216">String</span></span>||<span data-ttu-id="4323f-1217">Идентификатор удаляемого вложения.</span><span class="sxs-lookup"><span data-stu-id="4323f-1217">The identifier of the attachment to remove.</span></span>|
|`options`|<span data-ttu-id="4323f-1218">Object</span><span class="sxs-lookup"><span data-stu-id="4323f-1218">Object</span></span>|<span data-ttu-id="4323f-1219">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="4323f-1219">&lt;optional&gt;</span></span>|<span data-ttu-id="4323f-1220">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="4323f-1220">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="4323f-1221">Object</span><span class="sxs-lookup"><span data-stu-id="4323f-1221">Object</span></span>|<span data-ttu-id="4323f-1222">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="4323f-1222">&lt;optional&gt;</span></span>|<span data-ttu-id="4323f-1223">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="4323f-1223">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="4323f-1224">функция</span><span class="sxs-lookup"><span data-stu-id="4323f-1224">function</span></span>|<span data-ttu-id="4323f-1225">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="4323f-1225">&lt;optional&gt;</span></span>|<span data-ttu-id="4323f-1226">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="4323f-1226">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="4323f-1227">Если удалить вложение не удается, свойство `asyncResult.error` содержит код ошибки с указанием ее причины.</span><span class="sxs-lookup"><span data-stu-id="4323f-1227">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="4323f-1228">Ошибки</span><span class="sxs-lookup"><span data-stu-id="4323f-1228">Errors</span></span>

|<span data-ttu-id="4323f-1229">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="4323f-1229">Error code</span></span>|<span data-ttu-id="4323f-1230">Описание</span><span class="sxs-lookup"><span data-stu-id="4323f-1230">Description</span></span>|
|------------|-------------|
|`InvalidAttachmentId`|<span data-ttu-id="4323f-1231">Идентификатор вложения не существует.</span><span class="sxs-lookup"><span data-stu-id="4323f-1231">The attachment identifier does not exist.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4323f-1232">Requirements</span><span class="sxs-lookup"><span data-stu-id="4323f-1232">Requirements</span></span>

|<span data-ttu-id="4323f-1233">Требование</span><span class="sxs-lookup"><span data-stu-id="4323f-1233">Requirement</span></span>|<span data-ttu-id="4323f-1234">Значение</span><span class="sxs-lookup"><span data-stu-id="4323f-1234">Value</span></span>|
|---|---|
|[<span data-ttu-id="4323f-1235">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="4323f-1235">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4323f-1236">1.1</span><span class="sxs-lookup"><span data-stu-id="4323f-1236">1.1</span></span>|
|[<span data-ttu-id="4323f-1237">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4323f-1237">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4323f-1238">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="4323f-1238">ReadWriteItem</span></span>|
|[<span data-ttu-id="4323f-1239">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4323f-1239">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="4323f-1240">Создание</span><span class="sxs-lookup"><span data-stu-id="4323f-1240">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="4323f-1241">Пример</span><span class="sxs-lookup"><span data-stu-id="4323f-1241">Example</span></span>

<span data-ttu-id="4323f-1242">Указанный ниже код удаляет вложение с идентификатором "0".</span><span class="sxs-lookup"><span data-stu-id="4323f-1242">The following code removes an attachment with an identifier of '0'.</span></span>

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

#### <a name="removehandlerasynceventtype-options-callback"></a><span data-ttu-id="4323f-1243">removeHandlerAsync(eventType, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="4323f-1243">removeHandlerAsync(eventType, [options], [callback])</span></span>

<span data-ttu-id="4323f-1244">Удаляет обработчиков для поддерживаемого типа события.</span><span class="sxs-lookup"><span data-stu-id="4323f-1244">Removes the event handlers for a supported event type.</span></span>

<span data-ttu-id="4323f-1245">В настоящее время поддерживаются типы `Office.EventType.AppointmentTimeChanged`событий `Office.EventType.RecipientsChanged`, и`Office.EventType.RecurrenceChanged`</span><span class="sxs-lookup"><span data-stu-id="4323f-1245">Currently the supported event types are `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`</span></span>

##### <a name="parameters"></a><span data-ttu-id="4323f-1246">Параметры</span><span class="sxs-lookup"><span data-stu-id="4323f-1246">Parameters</span></span>

| <span data-ttu-id="4323f-1247">Имя</span><span class="sxs-lookup"><span data-stu-id="4323f-1247">Name</span></span> | <span data-ttu-id="4323f-1248">Тип</span><span class="sxs-lookup"><span data-stu-id="4323f-1248">Type</span></span> | <span data-ttu-id="4323f-1249">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="4323f-1249">Attributes</span></span> | <span data-ttu-id="4323f-1250">Описание</span><span class="sxs-lookup"><span data-stu-id="4323f-1250">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="4323f-1251">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="4323f-1251">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="4323f-1252">Событие, которое должно вызвать обработчик.</span><span class="sxs-lookup"><span data-stu-id="4323f-1252">The event that should invoke the handler.</span></span> |
| `options` | <span data-ttu-id="4323f-1253">Object</span><span class="sxs-lookup"><span data-stu-id="4323f-1253">Object</span></span> | <span data-ttu-id="4323f-1254">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="4323f-1254">&lt;optional&gt;</span></span> | <span data-ttu-id="4323f-1255">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="4323f-1255">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="4323f-1256">Object</span><span class="sxs-lookup"><span data-stu-id="4323f-1256">Object</span></span> | <span data-ttu-id="4323f-1257">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="4323f-1257">&lt;optional&gt;</span></span> | <span data-ttu-id="4323f-1258">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="4323f-1258">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="4323f-1259">функция</span><span class="sxs-lookup"><span data-stu-id="4323f-1259">function</span></span>| <span data-ttu-id="4323f-1260">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="4323f-1260">&lt;optional&gt;</span></span>|<span data-ttu-id="4323f-1261">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="4323f-1261">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4323f-1262">Требования</span><span class="sxs-lookup"><span data-stu-id="4323f-1262">Requirements</span></span>

|<span data-ttu-id="4323f-1263">Требование</span><span class="sxs-lookup"><span data-stu-id="4323f-1263">Requirement</span></span>| <span data-ttu-id="4323f-1264">Значение</span><span class="sxs-lookup"><span data-stu-id="4323f-1264">Value</span></span>|
|---|---|
|[<span data-ttu-id="4323f-1265">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="4323f-1265">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4323f-1266">1.7</span><span class="sxs-lookup"><span data-stu-id="4323f-1266">1.7</span></span> |
|[<span data-ttu-id="4323f-1267">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4323f-1267">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4323f-1268">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4323f-1268">ReadItem</span></span> |
|[<span data-ttu-id="4323f-1269">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4323f-1269">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4323f-1270">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="4323f-1270">Compose or Read</span></span> |

##### <a name="example"></a><span data-ttu-id="4323f-1271">Пример</span><span class="sxs-lookup"><span data-stu-id="4323f-1271">Example</span></span>

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

#### <a name="saveasyncoptions-callback"></a><span data-ttu-id="4323f-1272">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="4323f-1272">saveAsync([options], callback)</span></span>

<span data-ttu-id="4323f-1273">Асинхронно сохраняет элемент.</span><span class="sxs-lookup"><span data-stu-id="4323f-1273">Asynchronously saves an item.</span></span>

<span data-ttu-id="4323f-1274">При вызове этот метод сохраняет текущее сообщение в виде черновика и возвращает идентификатор элемента с помощью метода обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="4323f-1274">When invoked, this method saves the current message as a draft and returns the item id via the callback method.</span></span> <span data-ttu-id="4323f-1275">В Outlook в Интернете или интерактивном режиме Outlook этот элемент сохраняется на сервере.</span><span class="sxs-lookup"><span data-stu-id="4323f-1275">In Outlook on the web or Outlook in online mode, the item is saved to the server.</span></span> <span data-ttu-id="4323f-1276">В Outlook в режиме кэширования этот элемент сохраняется в локальном кэше.</span><span class="sxs-lookup"><span data-stu-id="4323f-1276">In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="4323f-1277">Если в режиме создания надстройка вызывает для элемента метод `saveAsync`, чтобы получить параметр `itemId` для использования с EWS или REST API, необходимо помнить, что синхронизация элемента с сервером может занять много времени, если Outlook работает в режиме кэширования данных.</span><span class="sxs-lookup"><span data-stu-id="4323f-1277">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="4323f-1278">До окончания синхронизации элемента применение параметра `itemId` будет приводить к ошибке.</span><span class="sxs-lookup"><span data-stu-id="4323f-1278">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="4323f-p188">Если метод `saveAsync` вызывается для встречи в режиме создания, она сохраняется как обычная встреча в календаре пользователя, а не как черновик. При сохранении новой встречи приглашения не отправляются. При сохранении существующей встречи уведомления отправляются добавленным или удаленным участникам.</span><span class="sxs-lookup"><span data-stu-id="4323f-p188">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="4323f-1282">Следующие клиенты отличаются другим поведением `saveAsync` в отношении встреч в режиме создания:</span><span class="sxs-lookup"><span data-stu-id="4323f-1282">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="4323f-1283">Outlook для Mac не поддерживает сохранение собрания.</span><span class="sxs-lookup"><span data-stu-id="4323f-1283">Outlook on Mac does not support saving a meeting.</span></span> <span data-ttu-id="4323f-1284">Метод `saveAsync` не работает при вызове из собрания в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="4323f-1284">The `saveAsync` method fails when called from a meeting in compose mode.</span></span> <span data-ttu-id="4323f-1285">Временное решение представлено в статье [Не удается сохранить встречу как черновик в Outlook для Mac с помощью API JS для Office](https://support.microsoft.com/help/4505745).</span><span class="sxs-lookup"><span data-stu-id="4323f-1285">See [Cannot save a meeting as a draft in Outlook for Mac by using Office JS API](https://support.microsoft.com/help/4505745) for a workaround.</span></span>
> - <span data-ttu-id="4323f-1286">Outlook в Интернете всегда отправляет приглашение или обновление при вызове `saveAsync` для встречи в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="4323f-1286">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4323f-1287">Параметры</span><span class="sxs-lookup"><span data-stu-id="4323f-1287">Parameters</span></span>

|<span data-ttu-id="4323f-1288">Имя</span><span class="sxs-lookup"><span data-stu-id="4323f-1288">Name</span></span>|<span data-ttu-id="4323f-1289">Тип</span><span class="sxs-lookup"><span data-stu-id="4323f-1289">Type</span></span>|<span data-ttu-id="4323f-1290">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="4323f-1290">Attributes</span></span>|<span data-ttu-id="4323f-1291">Описание</span><span class="sxs-lookup"><span data-stu-id="4323f-1291">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="4323f-1292">Object</span><span class="sxs-lookup"><span data-stu-id="4323f-1292">Object</span></span>|<span data-ttu-id="4323f-1293">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="4323f-1293">&lt;optional&gt;</span></span>|<span data-ttu-id="4323f-1294">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="4323f-1294">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="4323f-1295">Object</span><span class="sxs-lookup"><span data-stu-id="4323f-1295">Object</span></span>|<span data-ttu-id="4323f-1296">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="4323f-1296">&lt;optional&gt;</span></span>|<span data-ttu-id="4323f-1297">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="4323f-1297">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="4323f-1298">функция</span><span class="sxs-lookup"><span data-stu-id="4323f-1298">function</span></span>||<span data-ttu-id="4323f-1299">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="4323f-1299">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="4323f-1300">После успешного выполнения идентификатор элемента будет предоставлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="4323f-1300">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4323f-1301">Requirements</span><span class="sxs-lookup"><span data-stu-id="4323f-1301">Requirements</span></span>

|<span data-ttu-id="4323f-1302">Требование</span><span class="sxs-lookup"><span data-stu-id="4323f-1302">Requirement</span></span>|<span data-ttu-id="4323f-1303">Значение</span><span class="sxs-lookup"><span data-stu-id="4323f-1303">Value</span></span>|
|---|---|
|[<span data-ttu-id="4323f-1304">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="4323f-1304">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4323f-1305">1.3</span><span class="sxs-lookup"><span data-stu-id="4323f-1305">1.3</span></span>|
|[<span data-ttu-id="4323f-1306">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4323f-1306">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4323f-1307">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="4323f-1307">ReadWriteItem</span></span>|
|[<span data-ttu-id="4323f-1308">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4323f-1308">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="4323f-1309">Создание</span><span class="sxs-lookup"><span data-stu-id="4323f-1309">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="4323f-1310">Примеры</span><span class="sxs-lookup"><span data-stu-id="4323f-1310">Examples</span></span>

```js
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result.
  });
```

<span data-ttu-id="4323f-p190">Ниже приведен пример параметра `result`, переданного функции обратного вызова. Свойство `value` содержит идентификатор элемента.</span><span class="sxs-lookup"><span data-stu-id="4323f-p190">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

<br>

---
---

#### <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="4323f-1313">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="4323f-1313">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="4323f-1314">Асинхронно вставляет данные в текст или тему сообщения.</span><span class="sxs-lookup"><span data-stu-id="4323f-1314">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="4323f-p191">Метод `setSelectedDataAsync` вставляет указанную строку в местоположение курсора в теме или тексте элемента либо, если текст выделен в редакторе, он заменяет выделенный текст. Если курсор находится вне текста или темы элемента, возвращается ошибка. После вставки курсор помещается в конец вставленного содержимого.</span><span class="sxs-lookup"><span data-stu-id="4323f-p191">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4323f-1318">Параметры</span><span class="sxs-lookup"><span data-stu-id="4323f-1318">Parameters</span></span>

|<span data-ttu-id="4323f-1319">Имя</span><span class="sxs-lookup"><span data-stu-id="4323f-1319">Name</span></span>|<span data-ttu-id="4323f-1320">Тип</span><span class="sxs-lookup"><span data-stu-id="4323f-1320">Type</span></span>|<span data-ttu-id="4323f-1321">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="4323f-1321">Attributes</span></span>|<span data-ttu-id="4323f-1322">Описание</span><span class="sxs-lookup"><span data-stu-id="4323f-1322">Description</span></span>|
|---|---|---|---|
|`data`|<span data-ttu-id="4323f-1323">String</span><span class="sxs-lookup"><span data-stu-id="4323f-1323">String</span></span>||<span data-ttu-id="4323f-p192">Вставляемые данные. Объем данных не должен превышать 1 000 000 символов. Если передано больше 1 000 000 символов, возвращается исключение `ArgumentOutOfRange`.</span><span class="sxs-lookup"><span data-stu-id="4323f-p192">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`|<span data-ttu-id="4323f-1327">Object</span><span class="sxs-lookup"><span data-stu-id="4323f-1327">Object</span></span>|<span data-ttu-id="4323f-1328">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="4323f-1328">&lt;optional&gt;</span></span>|<span data-ttu-id="4323f-1329">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="4323f-1329">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="4323f-1330">Object</span><span class="sxs-lookup"><span data-stu-id="4323f-1330">Object</span></span>|<span data-ttu-id="4323f-1331">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="4323f-1331">&lt;optional&gt;</span></span>|<span data-ttu-id="4323f-1332">В методе обратного вызова разработчики могут указать любой объект, к которому необходимо получить доступ.</span><span class="sxs-lookup"><span data-stu-id="4323f-1332">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="4323f-1333">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="4323f-1333">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="4323f-1334">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="4323f-1334">&lt;optional&gt;</span></span>|<span data-ttu-id="4323f-1335">Если задано значение `text`, текущий стиль применяется в Outlook в Интернете и классических клиентах.</span><span class="sxs-lookup"><span data-stu-id="4323f-1335">If `text`, the current style is applied in Outlook on the web and desktop clients.</span></span> <span data-ttu-id="4323f-1336">Если поле представляет собой редактор HTML, вставляются только текстовые данные, даже если они имеют формат HTML.</span><span class="sxs-lookup"><span data-stu-id="4323f-1336">If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="4323f-1337">Если задано значение `html` и поле (не тема) поддерживает HTML, в Outlook в Интернете применяется текущий стиль, а в классических клиентах Outlook — стиль по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="4323f-1337">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook on the web and the default style is applied in Outlook desktop clients.</span></span> <span data-ttu-id="4323f-1338">Если поле является текстовым, возвращается ошибка `InvalidDataFormat`.</span><span class="sxs-lookup"><span data-stu-id="4323f-1338">If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="4323f-1339">Если свойство `coercionType` не задано, результат зависит от поля: если поле имеет формат HTML, используется текст в формате HTML, а если поле текстовое, применяется обычный текст.</span><span class="sxs-lookup"><span data-stu-id="4323f-1339">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`|<span data-ttu-id="4323f-1340">функция</span><span class="sxs-lookup"><span data-stu-id="4323f-1340">function</span></span>||<span data-ttu-id="4323f-1341">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="4323f-1341">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4323f-1342">Требования</span><span class="sxs-lookup"><span data-stu-id="4323f-1342">Requirements</span></span>

|<span data-ttu-id="4323f-1343">Требование</span><span class="sxs-lookup"><span data-stu-id="4323f-1343">Requirement</span></span>|<span data-ttu-id="4323f-1344">Значение</span><span class="sxs-lookup"><span data-stu-id="4323f-1344">Value</span></span>|
|---|---|
|[<span data-ttu-id="4323f-1345">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="4323f-1345">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="4323f-1346">1.2</span><span class="sxs-lookup"><span data-stu-id="4323f-1346">1.2</span></span>|
|[<span data-ttu-id="4323f-1347">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4323f-1347">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="4323f-1348">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="4323f-1348">ReadWriteItem</span></span>|
|[<span data-ttu-id="4323f-1349">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4323f-1349">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="4323f-1350">Создание</span><span class="sxs-lookup"><span data-stu-id="4323f-1350">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="4323f-1351">Пример</span><span class="sxs-lookup"><span data-stu-id="4323f-1351">Example</span></span>

```js
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
