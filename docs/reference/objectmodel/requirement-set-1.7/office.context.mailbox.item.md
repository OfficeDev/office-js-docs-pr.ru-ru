---
title: Office. Context. Mailbox. Item — набор требований 1,7
description: ''
ms.date: 05/30/2019
localization_priority: Normal
ms.openlocfilehash: fd618b766a519c522f323e0a9d43105b3258c421
ms.sourcegitcommit: 567aa05d6ee6b3639f65c50188df2331b7685857
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/04/2019
ms.locfileid: "34706315"
---
# <a name="item"></a><span data-ttu-id="2085e-102">item</span><span class="sxs-lookup"><span data-stu-id="2085e-102">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="2085e-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="2085e-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="2085e-p101">Пространство имен `item` используется для доступа к выбранному в данный момент сообщению, приглашению на собрание или описанию встречи. Вы можете определить тип пространства имен `item` с помощью свойства [itemType](#itemtype-officemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="2085e-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="2085e-106">Requirements</span><span class="sxs-lookup"><span data-stu-id="2085e-106">Requirements</span></span>

|<span data-ttu-id="2085e-107">Требование</span><span class="sxs-lookup"><span data-stu-id="2085e-107">Requirement</span></span>|<span data-ttu-id="2085e-108">Значение</span><span class="sxs-lookup"><span data-stu-id="2085e-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="2085e-109">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="2085e-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2085e-110">1.0</span><span class="sxs-lookup"><span data-stu-id="2085e-110">1.0</span></span>|
|[<span data-ttu-id="2085e-111">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="2085e-111">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2085e-112">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="2085e-112">Restricted</span></span>|
|[<span data-ttu-id="2085e-113">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="2085e-113">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="2085e-114">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="2085e-114">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="2085e-115">Элементы и методы</span><span class="sxs-lookup"><span data-stu-id="2085e-115">Members and methods</span></span>

| <span data-ttu-id="2085e-116">Элемент</span><span class="sxs-lookup"><span data-stu-id="2085e-116">Member</span></span> | <span data-ttu-id="2085e-117">Тип</span><span class="sxs-lookup"><span data-stu-id="2085e-117">Type</span></span> |
|--------|------|
| [<span data-ttu-id="2085e-118">attachments</span><span class="sxs-lookup"><span data-stu-id="2085e-118">attachments</span></span>](#attachments-arrayattachmentdetails) | <span data-ttu-id="2085e-119">Элемент</span><span class="sxs-lookup"><span data-stu-id="2085e-119">Member</span></span> |
| [<span data-ttu-id="2085e-120">bcc</span><span class="sxs-lookup"><span data-stu-id="2085e-120">bcc</span></span>](#bcc-recipients) | <span data-ttu-id="2085e-121">Элемент</span><span class="sxs-lookup"><span data-stu-id="2085e-121">Member</span></span> |
| [<span data-ttu-id="2085e-122">body</span><span class="sxs-lookup"><span data-stu-id="2085e-122">body</span></span>](#body-body) | <span data-ttu-id="2085e-123">Элемент</span><span class="sxs-lookup"><span data-stu-id="2085e-123">Member</span></span> |
| [<span data-ttu-id="2085e-124">cc</span><span class="sxs-lookup"><span data-stu-id="2085e-124">cc</span></span>](#cc-arrayemailaddressdetailsrecipients) | <span data-ttu-id="2085e-125">Элемент</span><span class="sxs-lookup"><span data-stu-id="2085e-125">Member</span></span> |
| [<span data-ttu-id="2085e-126">conversationId</span><span class="sxs-lookup"><span data-stu-id="2085e-126">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="2085e-127">Элемент</span><span class="sxs-lookup"><span data-stu-id="2085e-127">Member</span></span> |
| [<span data-ttu-id="2085e-128">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="2085e-128">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="2085e-129">Элемент</span><span class="sxs-lookup"><span data-stu-id="2085e-129">Member</span></span> |
| [<span data-ttu-id="2085e-130">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="2085e-130">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="2085e-131">Элемент</span><span class="sxs-lookup"><span data-stu-id="2085e-131">Member</span></span> |
| [<span data-ttu-id="2085e-132">end</span><span class="sxs-lookup"><span data-stu-id="2085e-132">end</span></span>](#end-datetime) | <span data-ttu-id="2085e-133">Элемент</span><span class="sxs-lookup"><span data-stu-id="2085e-133">Member</span></span> |
| [<span data-ttu-id="2085e-134">from</span><span class="sxs-lookup"><span data-stu-id="2085e-134">from</span></span>](#from-emailaddressdetailsfrom) | <span data-ttu-id="2085e-135">Элемент</span><span class="sxs-lookup"><span data-stu-id="2085e-135">Member</span></span> |
| [<span data-ttu-id="2085e-136">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="2085e-136">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="2085e-137">Элемент</span><span class="sxs-lookup"><span data-stu-id="2085e-137">Member</span></span> |
| [<span data-ttu-id="2085e-138">itemClass</span><span class="sxs-lookup"><span data-stu-id="2085e-138">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="2085e-139">Элемент</span><span class="sxs-lookup"><span data-stu-id="2085e-139">Member</span></span> |
| [<span data-ttu-id="2085e-140">itemId</span><span class="sxs-lookup"><span data-stu-id="2085e-140">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="2085e-141">Элемент</span><span class="sxs-lookup"><span data-stu-id="2085e-141">Member</span></span> |
| [<span data-ttu-id="2085e-142">itemType</span><span class="sxs-lookup"><span data-stu-id="2085e-142">itemType</span></span>](#itemtype-officemailboxenumsitemtype) | <span data-ttu-id="2085e-143">Элемент</span><span class="sxs-lookup"><span data-stu-id="2085e-143">Member</span></span> |
| [<span data-ttu-id="2085e-144">location</span><span class="sxs-lookup"><span data-stu-id="2085e-144">location</span></span>](#location-stringlocation) | <span data-ttu-id="2085e-145">Элемент</span><span class="sxs-lookup"><span data-stu-id="2085e-145">Member</span></span> |
| [<span data-ttu-id="2085e-146">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="2085e-146">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="2085e-147">Элемент</span><span class="sxs-lookup"><span data-stu-id="2085e-147">Member</span></span> |
| [<span data-ttu-id="2085e-148">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="2085e-148">notificationMessages</span></span>](#notificationmessages-notificationmessages) | <span data-ttu-id="2085e-149">Элемент</span><span class="sxs-lookup"><span data-stu-id="2085e-149">Member</span></span> |
| [<span data-ttu-id="2085e-150">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="2085e-150">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="2085e-151">Элемент</span><span class="sxs-lookup"><span data-stu-id="2085e-151">Member</span></span> |
| [<span data-ttu-id="2085e-152">organizer</span><span class="sxs-lookup"><span data-stu-id="2085e-152">organizer</span></span>](#organizer-emailaddressdetailsorganizer) | <span data-ttu-id="2085e-153">Элемент</span><span class="sxs-lookup"><span data-stu-id="2085e-153">Member</span></span> |
| [<span data-ttu-id="2085e-154">recurrence</span><span class="sxs-lookup"><span data-stu-id="2085e-154">recurrence</span></span>](#nullable-recurrence-recurrence) | <span data-ttu-id="2085e-155">Member</span><span class="sxs-lookup"><span data-stu-id="2085e-155">Member</span></span> |
| [<span data-ttu-id="2085e-156">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="2085e-156">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="2085e-157">Элемент</span><span class="sxs-lookup"><span data-stu-id="2085e-157">Member</span></span> |
| [<span data-ttu-id="2085e-158">sender</span><span class="sxs-lookup"><span data-stu-id="2085e-158">sender</span></span>](#sender-emailaddressdetails) | <span data-ttu-id="2085e-159">Элемент</span><span class="sxs-lookup"><span data-stu-id="2085e-159">Member</span></span> |
| [<span data-ttu-id="2085e-160">seriesId</span><span class="sxs-lookup"><span data-stu-id="2085e-160">seriesId</span></span>](#nullable-seriesid-string) | <span data-ttu-id="2085e-161">Элемент</span><span class="sxs-lookup"><span data-stu-id="2085e-161">Member</span></span> |
| [<span data-ttu-id="2085e-162">start</span><span class="sxs-lookup"><span data-stu-id="2085e-162">start</span></span>](#start-datetime) | <span data-ttu-id="2085e-163">Элемент</span><span class="sxs-lookup"><span data-stu-id="2085e-163">Member</span></span> |
| [<span data-ttu-id="2085e-164">subject</span><span class="sxs-lookup"><span data-stu-id="2085e-164">subject</span></span>](#subject-stringsubject) | <span data-ttu-id="2085e-165">Элемент</span><span class="sxs-lookup"><span data-stu-id="2085e-165">Member</span></span> |
| [<span data-ttu-id="2085e-166">to</span><span class="sxs-lookup"><span data-stu-id="2085e-166">to</span></span>](#to-arrayemailaddressdetailsrecipients) | <span data-ttu-id="2085e-167">Элемент</span><span class="sxs-lookup"><span data-stu-id="2085e-167">Member</span></span> |
| [<span data-ttu-id="2085e-168">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="2085e-168">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="2085e-169">Метод</span><span class="sxs-lookup"><span data-stu-id="2085e-169">Method</span></span> |
| [<span data-ttu-id="2085e-170">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="2085e-170">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="2085e-171">Метод</span><span class="sxs-lookup"><span data-stu-id="2085e-171">Method</span></span> |
| [<span data-ttu-id="2085e-172">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="2085e-172">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="2085e-173">Метод</span><span class="sxs-lookup"><span data-stu-id="2085e-173">Method</span></span> |
| [<span data-ttu-id="2085e-174">close</span><span class="sxs-lookup"><span data-stu-id="2085e-174">close</span></span>](#close) | <span data-ttu-id="2085e-175">Метод</span><span class="sxs-lookup"><span data-stu-id="2085e-175">Method</span></span> |
| [<span data-ttu-id="2085e-176">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="2085e-176">displayReplyAllForm</span></span>](#displayreplyallformformdata-callback) | <span data-ttu-id="2085e-177">Метод</span><span class="sxs-lookup"><span data-stu-id="2085e-177">Method</span></span> |
| [<span data-ttu-id="2085e-178">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="2085e-178">displayReplyForm</span></span>](#displayreplyformformdata-callback) | <span data-ttu-id="2085e-179">Метод</span><span class="sxs-lookup"><span data-stu-id="2085e-179">Method</span></span> |
| [<span data-ttu-id="2085e-180">getEntities</span><span class="sxs-lookup"><span data-stu-id="2085e-180">getEntities</span></span>](#getentities--entities) | <span data-ttu-id="2085e-181">Метод</span><span class="sxs-lookup"><span data-stu-id="2085e-181">Method</span></span> |
| [<span data-ttu-id="2085e-182">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="2085e-182">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="2085e-183">Метод</span><span class="sxs-lookup"><span data-stu-id="2085e-183">Method</span></span> |
| [<span data-ttu-id="2085e-184">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="2085e-184">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="2085e-185">Метод</span><span class="sxs-lookup"><span data-stu-id="2085e-185">Method</span></span> |
| [<span data-ttu-id="2085e-186">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="2085e-186">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="2085e-187">Метод</span><span class="sxs-lookup"><span data-stu-id="2085e-187">Method</span></span> |
| [<span data-ttu-id="2085e-188">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="2085e-188">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="2085e-189">Метод</span><span class="sxs-lookup"><span data-stu-id="2085e-189">Method</span></span> |
| [<span data-ttu-id="2085e-190">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="2085e-190">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="2085e-191">Метод</span><span class="sxs-lookup"><span data-stu-id="2085e-191">Method</span></span> |
| [<span data-ttu-id="2085e-192">Жетселектедентитиес</span><span class="sxs-lookup"><span data-stu-id="2085e-192">getSelectedEntities</span></span>](#getselectedentities--entities) | <span data-ttu-id="2085e-193">Метод</span><span class="sxs-lookup"><span data-stu-id="2085e-193">Method</span></span> |
| [<span data-ttu-id="2085e-194">Жетселектедрежексматчес</span><span class="sxs-lookup"><span data-stu-id="2085e-194">getSelectedRegExMatches</span></span>](#getselectedregexmatches--object) | <span data-ttu-id="2085e-195">Метод</span><span class="sxs-lookup"><span data-stu-id="2085e-195">Method</span></span> |
| [<span data-ttu-id="2085e-196">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="2085e-196">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="2085e-197">Метод</span><span class="sxs-lookup"><span data-stu-id="2085e-197">Method</span></span> |
| [<span data-ttu-id="2085e-198">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="2085e-198">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="2085e-199">Метод</span><span class="sxs-lookup"><span data-stu-id="2085e-199">Method</span></span> |
| [<span data-ttu-id="2085e-200">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="2085e-200">removeHandlerAsync</span></span>](#removehandlerasynceventtype-options-callback) | <span data-ttu-id="2085e-201">Метод</span><span class="sxs-lookup"><span data-stu-id="2085e-201">Method</span></span> |
| [<span data-ttu-id="2085e-202">saveAsync</span><span class="sxs-lookup"><span data-stu-id="2085e-202">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="2085e-203">Метод</span><span class="sxs-lookup"><span data-stu-id="2085e-203">Method</span></span> |
| [<span data-ttu-id="2085e-204">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="2085e-204">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="2085e-205">Метод</span><span class="sxs-lookup"><span data-stu-id="2085e-205">Method</span></span> |

### <a name="example"></a><span data-ttu-id="2085e-206">Пример</span><span class="sxs-lookup"><span data-stu-id="2085e-206">Example</span></span>

<span data-ttu-id="2085e-207">В примере кода JavaScript, приведенном ниже, показано, как получить доступ к свойству `subject` текущего элемента в Outlook.</span><span class="sxs-lookup"><span data-stu-id="2085e-207">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="2085e-208">Элементы</span><span class="sxs-lookup"><span data-stu-id="2085e-208">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook17officeattachmentdetails"></a><span data-ttu-id="2085e-209">вложения: Array. <[AttachmentDetails](/javascript/api/outlook_1_7/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="2085e-209">attachments: Array.<[AttachmentDetails](/javascript/api/outlook_1_7/office.attachmentdetails)></span></span>

<span data-ttu-id="2085e-p102">Получает массив вложений для элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="2085e-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="2085e-212">Outlook блокирует определенные типы файлов из-за потенциальных проблем с безопасностью, поэтому они не возвращаются.</span><span class="sxs-lookup"><span data-stu-id="2085e-212">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="2085e-213">Дополнительные сведения см. в статье [Блокировка вложений в Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="2085e-213">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="2085e-214">Тип</span><span class="sxs-lookup"><span data-stu-id="2085e-214">Type</span></span>

*   <span data-ttu-id="2085e-215">Array.<[AttachmentDetails](/javascript/api/outlook_1_7/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="2085e-215">Array.<[AttachmentDetails](/javascript/api/outlook_1_7/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="2085e-216">Требования</span><span class="sxs-lookup"><span data-stu-id="2085e-216">Requirements</span></span>

|<span data-ttu-id="2085e-217">Требование</span><span class="sxs-lookup"><span data-stu-id="2085e-217">Requirement</span></span>|<span data-ttu-id="2085e-218">Значение</span><span class="sxs-lookup"><span data-stu-id="2085e-218">Value</span></span>|
|---|---|
|[<span data-ttu-id="2085e-219">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="2085e-219">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2085e-220">1.0</span><span class="sxs-lookup"><span data-stu-id="2085e-220">1.0</span></span>|
|[<span data-ttu-id="2085e-221">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="2085e-221">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2085e-222">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2085e-222">ReadItem</span></span>|
|[<span data-ttu-id="2085e-223">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="2085e-223">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="2085e-224">Чтение</span><span class="sxs-lookup"><span data-stu-id="2085e-224">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2085e-225">Пример</span><span class="sxs-lookup"><span data-stu-id="2085e-225">Example</span></span>

<span data-ttu-id="2085e-226">С помощью приведенного ниже кода можно создать HTML-строку с подробными сведениями обо всех вложениях для текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="2085e-226">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

#### <a name="bcc-recipientsjavascriptapioutlook17officerecipients"></a><span data-ttu-id="2085e-227">СК: [получатели](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="2085e-227">bcc: [Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

<span data-ttu-id="2085e-228">Получает объект, который предоставляет методы для получения или обновления получателей скрытой копии сообщения.</span><span class="sxs-lookup"><span data-stu-id="2085e-228">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="2085e-229">Только в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="2085e-229">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="2085e-230">Тип</span><span class="sxs-lookup"><span data-stu-id="2085e-230">Type</span></span>

*   [<span data-ttu-id="2085e-231">Получатели</span><span class="sxs-lookup"><span data-stu-id="2085e-231">Recipients</span></span>](/javascript/api/outlook_1_7/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="2085e-232">Требования</span><span class="sxs-lookup"><span data-stu-id="2085e-232">Requirements</span></span>

|<span data-ttu-id="2085e-233">Требование</span><span class="sxs-lookup"><span data-stu-id="2085e-233">Requirement</span></span>|<span data-ttu-id="2085e-234">Значение</span><span class="sxs-lookup"><span data-stu-id="2085e-234">Value</span></span>|
|---|---|
|[<span data-ttu-id="2085e-235">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="2085e-235">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2085e-236">1.1</span><span class="sxs-lookup"><span data-stu-id="2085e-236">1.1</span></span>|
|[<span data-ttu-id="2085e-237">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="2085e-237">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2085e-238">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2085e-238">ReadItem</span></span>|
|[<span data-ttu-id="2085e-239">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="2085e-239">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="2085e-240">Создание</span><span class="sxs-lookup"><span data-stu-id="2085e-240">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="2085e-241">Пример</span><span class="sxs-lookup"><span data-stu-id="2085e-241">Example</span></span>

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

#### <a name="body-bodyjavascriptapioutlook17officebody"></a><span data-ttu-id="2085e-242">основной текст: [Body](/javascript/api/outlook_1_7/office.body)</span><span class="sxs-lookup"><span data-stu-id="2085e-242">body: [Body](/javascript/api/outlook_1_7/office.body)</span></span>

<span data-ttu-id="2085e-243">Получает объект, предоставляющий методы для работы с основным текстом элемента.</span><span class="sxs-lookup"><span data-stu-id="2085e-243">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="2085e-244">Тип</span><span class="sxs-lookup"><span data-stu-id="2085e-244">Type</span></span>

*   [<span data-ttu-id="2085e-245">Body</span><span class="sxs-lookup"><span data-stu-id="2085e-245">Body</span></span>](/javascript/api/outlook_1_7/office.body)

##### <a name="requirements"></a><span data-ttu-id="2085e-246">Требования</span><span class="sxs-lookup"><span data-stu-id="2085e-246">Requirements</span></span>

|<span data-ttu-id="2085e-247">Требование</span><span class="sxs-lookup"><span data-stu-id="2085e-247">Requirement</span></span>|<span data-ttu-id="2085e-248">Значение</span><span class="sxs-lookup"><span data-stu-id="2085e-248">Value</span></span>|
|---|---|
|[<span data-ttu-id="2085e-249">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="2085e-249">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2085e-250">1.1</span><span class="sxs-lookup"><span data-stu-id="2085e-250">1.1</span></span>|
|[<span data-ttu-id="2085e-251">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="2085e-251">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2085e-252">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2085e-252">ReadItem</span></span>|
|[<span data-ttu-id="2085e-253">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="2085e-253">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="2085e-254">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="2085e-254">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2085e-255">Пример</span><span class="sxs-lookup"><span data-stu-id="2085e-255">Example</span></span>

<span data-ttu-id="2085e-256">В этом примере возвращается текст сообщения в формате обычного текста.</span><span class="sxs-lookup"><span data-stu-id="2085e-256">This example gets the body of the message in plain text.</span></span>

```javascript
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="2085e-257">Ниже приведен пример итогового параметра, переданного функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="2085e-257">The following is an example of the result parameter passed to the callback function.</span></span>

```json
{
  "value": "TEXT of whole body (including threads below)",
  "status": "succeeded",
  "asyncContext": "This is passed to the callback"
}
```

---
---

#### <a name="cc-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients"></a><span data-ttu-id="2085e-258">CC: Array. <[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[получатели](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="2085e-258">cc: Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

<span data-ttu-id="2085e-259">Предоставляет доступ к получателям копии сообщения.</span><span class="sxs-lookup"><span data-stu-id="2085e-259">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="2085e-260">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="2085e-260">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="2085e-261">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="2085e-261">Read mode</span></span>

<span data-ttu-id="2085e-p106">Свойство `cc` возвращает массив, который содержит объект `EmailAddressDetails` для каждого получателя, указанного в строке **Копия** сообщения. Коллекция может включать не более 100 элементов.</span><span class="sxs-lookup"><span data-stu-id="2085e-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="2085e-264">Режим создания</span><span class="sxs-lookup"><span data-stu-id="2085e-264">Compose mode</span></span>

<span data-ttu-id="2085e-265">Свойство `cc` возвращает объект `Recipients`, предоставляющий методы для получения или обновления получателей, которые указаны в строке **Копия** сообщения.</span><span class="sxs-lookup"><span data-stu-id="2085e-265">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

```javascript
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="2085e-266">Тип</span><span class="sxs-lookup"><span data-stu-id="2085e-266">Type</span></span>

*   <span data-ttu-id="2085e-267">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="2085e-267">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="2085e-268">Требования</span><span class="sxs-lookup"><span data-stu-id="2085e-268">Requirements</span></span>

|<span data-ttu-id="2085e-269">Требование</span><span class="sxs-lookup"><span data-stu-id="2085e-269">Requirement</span></span>|<span data-ttu-id="2085e-270">Значение</span><span class="sxs-lookup"><span data-stu-id="2085e-270">Value</span></span>|
|---|---|
|[<span data-ttu-id="2085e-271">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="2085e-271">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2085e-272">1.0</span><span class="sxs-lookup"><span data-stu-id="2085e-272">1.0</span></span>|
|[<span data-ttu-id="2085e-273">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="2085e-273">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2085e-274">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2085e-274">ReadItem</span></span>|
|[<span data-ttu-id="2085e-275">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="2085e-275">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="2085e-276">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="2085e-276">Compose or Read</span></span>|

---
---

#### <a name="nullable-conversationid-string"></a><span data-ttu-id="2085e-277">(Nullable) conversationId: строка</span><span class="sxs-lookup"><span data-stu-id="2085e-277">(nullable) conversationId: String</span></span>

<span data-ttu-id="2085e-278">Получает идентификатор разговора по электронной почте, содержащего конкретное сообщение.</span><span class="sxs-lookup"><span data-stu-id="2085e-278">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="2085e-p107">Вы можете получить целочисленное значение этого свойства, если ваше почтовое приложение активируется в формах просмотра или формах создания ответов. Если пользователь изменит тему ответа, после его отправки идентификатор беседы будет изменен, и полученное ранее значение будет недействительным.</span><span class="sxs-lookup"><span data-stu-id="2085e-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="2085e-p108">Это свойство имеет значение NULL для нового элемента в форме создания. Свойство `conversationId` вернет значение, если пользователь задаст тему и сохранит элемент.</span><span class="sxs-lookup"><span data-stu-id="2085e-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="2085e-283">Тип</span><span class="sxs-lookup"><span data-stu-id="2085e-283">Type</span></span>

*   <span data-ttu-id="2085e-284">String</span><span class="sxs-lookup"><span data-stu-id="2085e-284">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="2085e-285">Требования</span><span class="sxs-lookup"><span data-stu-id="2085e-285">Requirements</span></span>

|<span data-ttu-id="2085e-286">Требование</span><span class="sxs-lookup"><span data-stu-id="2085e-286">Requirement</span></span>|<span data-ttu-id="2085e-287">Значение</span><span class="sxs-lookup"><span data-stu-id="2085e-287">Value</span></span>|
|---|---|
|[<span data-ttu-id="2085e-288">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="2085e-288">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2085e-289">1.0</span><span class="sxs-lookup"><span data-stu-id="2085e-289">1.0</span></span>|
|[<span data-ttu-id="2085e-290">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="2085e-290">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2085e-291">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2085e-291">ReadItem</span></span>|
|[<span data-ttu-id="2085e-292">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="2085e-292">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="2085e-293">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="2085e-293">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2085e-294">Пример</span><span class="sxs-lookup"><span data-stu-id="2085e-294">Example</span></span>

```javascript
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

---
---

#### <a name="datetimecreated-date"></a><span data-ttu-id="2085e-295">dateTimeCreated: Дата</span><span class="sxs-lookup"><span data-stu-id="2085e-295">dateTimeCreated: Date</span></span>

<span data-ttu-id="2085e-p109">Получает дату и время создания элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="2085e-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="2085e-298">Тип</span><span class="sxs-lookup"><span data-stu-id="2085e-298">Type</span></span>

*   <span data-ttu-id="2085e-299">Дата</span><span class="sxs-lookup"><span data-stu-id="2085e-299">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="2085e-300">Требования</span><span class="sxs-lookup"><span data-stu-id="2085e-300">Requirements</span></span>

|<span data-ttu-id="2085e-301">Требование</span><span class="sxs-lookup"><span data-stu-id="2085e-301">Requirement</span></span>|<span data-ttu-id="2085e-302">Значение</span><span class="sxs-lookup"><span data-stu-id="2085e-302">Value</span></span>|
|---|---|
|[<span data-ttu-id="2085e-303">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="2085e-303">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2085e-304">1.0</span><span class="sxs-lookup"><span data-stu-id="2085e-304">1.0</span></span>|
|[<span data-ttu-id="2085e-305">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="2085e-305">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2085e-306">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2085e-306">ReadItem</span></span>|
|[<span data-ttu-id="2085e-307">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="2085e-307">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="2085e-308">Чтение</span><span class="sxs-lookup"><span data-stu-id="2085e-308">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2085e-309">Пример</span><span class="sxs-lookup"><span data-stu-id="2085e-309">Example</span></span>

```javascript
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

---
---

#### <a name="datetimemodified-date"></a><span data-ttu-id="2085e-310">dateTimeModified: Дата</span><span class="sxs-lookup"><span data-stu-id="2085e-310">dateTimeModified: Date</span></span>

<span data-ttu-id="2085e-p110">Получает дату и время последнего изменения элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="2085e-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="2085e-313">Этот элемент не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="2085e-313">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="2085e-314">Тип</span><span class="sxs-lookup"><span data-stu-id="2085e-314">Type</span></span>

*   <span data-ttu-id="2085e-315">Дата</span><span class="sxs-lookup"><span data-stu-id="2085e-315">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="2085e-316">Требования</span><span class="sxs-lookup"><span data-stu-id="2085e-316">Requirements</span></span>

|<span data-ttu-id="2085e-317">Требование</span><span class="sxs-lookup"><span data-stu-id="2085e-317">Requirement</span></span>|<span data-ttu-id="2085e-318">Значение</span><span class="sxs-lookup"><span data-stu-id="2085e-318">Value</span></span>|
|---|---|
|[<span data-ttu-id="2085e-319">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="2085e-319">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2085e-320">1.0</span><span class="sxs-lookup"><span data-stu-id="2085e-320">1.0</span></span>|
|[<span data-ttu-id="2085e-321">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="2085e-321">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2085e-322">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2085e-322">ReadItem</span></span>|
|[<span data-ttu-id="2085e-323">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="2085e-323">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="2085e-324">Чтение</span><span class="sxs-lookup"><span data-stu-id="2085e-324">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2085e-325">Пример</span><span class="sxs-lookup"><span data-stu-id="2085e-325">Example</span></span>

```javascript
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

---
---

#### <a name="end-datetimejavascriptapioutlook17officetime"></a><span data-ttu-id="2085e-326">конец: Дата | [Time (время](/javascript/api/outlook_1_7/office.time) )</span><span class="sxs-lookup"><span data-stu-id="2085e-326">end: Date|[Time](/javascript/api/outlook_1_7/office.time)</span></span>

<span data-ttu-id="2085e-327">Получает или задает дату и время окончания встречи.</span><span class="sxs-lookup"><span data-stu-id="2085e-327">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="2085e-p111">Свойство `end` представлено в виде значения даты и времени в формате UTC. Преобразовать значение свойства end в местные значения даты и времени клиента можно с помощью метода [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime).</span><span class="sxs-lookup"><span data-stu-id="2085e-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="2085e-330">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="2085e-330">Read mode</span></span>

<span data-ttu-id="2085e-331">Свойство `end` возвращает объект `Date`.</span><span class="sxs-lookup"><span data-stu-id="2085e-331">The `end` property returns a `Date` object.</span></span>

```javascript
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="2085e-332">Режим создания</span><span class="sxs-lookup"><span data-stu-id="2085e-332">Compose mode</span></span>

<span data-ttu-id="2085e-333">Свойство `end` возвращает объект `Time`.</span><span class="sxs-lookup"><span data-stu-id="2085e-333">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="2085e-334">Если вы задаете время окончания с помощью метода [`Time.setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-), необходимо использовать метод [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) для преобразования местного времени на клиенте в формат UTC для сервера.</span><span class="sxs-lookup"><span data-stu-id="2085e-334">When you use the [`Time.setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="2085e-335">В примере ниже показано, как с помощью метода [`setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) объекта `Time` задать время окончания встречи.</span><span class="sxs-lookup"><span data-stu-id="2085e-335">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="2085e-336">Тип</span><span class="sxs-lookup"><span data-stu-id="2085e-336">Type</span></span>

*   <span data-ttu-id="2085e-337">Date | [Time](/javascript/api/outlook_1_7/office.time)</span><span class="sxs-lookup"><span data-stu-id="2085e-337">Date | [Time](/javascript/api/outlook_1_7/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="2085e-338">Требования</span><span class="sxs-lookup"><span data-stu-id="2085e-338">Requirements</span></span>

|<span data-ttu-id="2085e-339">Требование</span><span class="sxs-lookup"><span data-stu-id="2085e-339">Requirement</span></span>|<span data-ttu-id="2085e-340">Значение</span><span class="sxs-lookup"><span data-stu-id="2085e-340">Value</span></span>|
|---|---|
|[<span data-ttu-id="2085e-341">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="2085e-341">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2085e-342">1.0</span><span class="sxs-lookup"><span data-stu-id="2085e-342">1.0</span></span>|
|[<span data-ttu-id="2085e-343">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="2085e-343">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2085e-344">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2085e-344">ReadItem</span></span>|
|[<span data-ttu-id="2085e-345">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="2085e-345">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="2085e-346">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="2085e-346">Compose or Read</span></span>|

---
---

#### <a name="from-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsfromjavascriptapioutlook17officefrom"></a><span data-ttu-id="2085e-347">от: [EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)|[](/javascript/api/outlook_1_7/office.from)</span><span class="sxs-lookup"><span data-stu-id="2085e-347">from: [EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)|[From](/javascript/api/outlook_1_7/office.from)</span></span>

<span data-ttu-id="2085e-348">Получает электронный адрес отправителя сообщения.</span><span class="sxs-lookup"><span data-stu-id="2085e-348">Gets the email address of the sender of a message.</span></span>

<span data-ttu-id="2085e-p112">Свойства `from` и [`sender`](#sender-emailaddressdetails) представляют одно лицо, если сообщение не отправлено представителем. В противном случае свойство `from` представляет лицо, делегировавшее полномочия, а свойство sender — представителя.</span><span class="sxs-lookup"><span data-stu-id="2085e-p112">The `from` and [`sender`](#sender-emailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="2085e-351">Свойству `recipientType`, принадлежащему объекту `EmailAddressDetails` в свойстве `from`, задано значение `undefined`.</span><span class="sxs-lookup"><span data-stu-id="2085e-351">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="2085e-352">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="2085e-352">Read mode</span></span>

<span data-ttu-id="2085e-353">`from` Свойство возвращает `EmailAddressDetails` объект.</span><span class="sxs-lookup"><span data-stu-id="2085e-353">The `from` property returns an `EmailAddressDetails` object.</span></span>

```javascript
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

##### <a name="compose-mode"></a><span data-ttu-id="2085e-354">Режим создания</span><span class="sxs-lookup"><span data-stu-id="2085e-354">Compose mode</span></span>

<span data-ttu-id="2085e-355">`from` Свойство возвращает `From` объект, который предоставляет метод для получения значения From.</span><span class="sxs-lookup"><span data-stu-id="2085e-355">The `from` property returns a `From` object that provides a method to get the from value.</span></span>

```javascript
Office.context.mailbox.item.from.getAsync(callback);

function callback(asyncResult) {
  var from = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="2085e-356">Тип</span><span class="sxs-lookup"><span data-stu-id="2085e-356">Type</span></span>

*   <span data-ttu-id="2085e-357">[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails) | [из](/javascript/api/outlook_1_7/office.from)</span><span class="sxs-lookup"><span data-stu-id="2085e-357">[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails) | [From](/javascript/api/outlook_1_7/office.from)</span></span>

##### <a name="requirements"></a><span data-ttu-id="2085e-358">Требования</span><span class="sxs-lookup"><span data-stu-id="2085e-358">Requirements</span></span>

|<span data-ttu-id="2085e-359">Требование</span><span class="sxs-lookup"><span data-stu-id="2085e-359">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="2085e-360">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="2085e-360">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2085e-361">1.0</span><span class="sxs-lookup"><span data-stu-id="2085e-361">1.0</span></span>|<span data-ttu-id="2085e-362">1.7</span><span class="sxs-lookup"><span data-stu-id="2085e-362">1.7</span></span>|
|[<span data-ttu-id="2085e-363">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="2085e-363">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2085e-364">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2085e-364">ReadItem</span></span>|<span data-ttu-id="2085e-365">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="2085e-365">ReadWriteItem</span></span>|
|[<span data-ttu-id="2085e-366">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="2085e-366">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="2085e-367">Чтение</span><span class="sxs-lookup"><span data-stu-id="2085e-367">Read</span></span>|<span data-ttu-id="2085e-368">Создание</span><span class="sxs-lookup"><span data-stu-id="2085e-368">Compose</span></span>|

---
---

#### <a name="internetmessageid-string"></a><span data-ttu-id="2085e-369">internetMessageId: строка</span><span class="sxs-lookup"><span data-stu-id="2085e-369">internetMessageId: String</span></span>

<span data-ttu-id="2085e-p113">Получает идентификатор интернет-сообщения для электронного сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="2085e-p113">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="2085e-372">Тип</span><span class="sxs-lookup"><span data-stu-id="2085e-372">Type</span></span>

*   <span data-ttu-id="2085e-373">String</span><span class="sxs-lookup"><span data-stu-id="2085e-373">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="2085e-374">Требования</span><span class="sxs-lookup"><span data-stu-id="2085e-374">Requirements</span></span>

|<span data-ttu-id="2085e-375">Требование</span><span class="sxs-lookup"><span data-stu-id="2085e-375">Requirement</span></span>|<span data-ttu-id="2085e-376">Значение</span><span class="sxs-lookup"><span data-stu-id="2085e-376">Value</span></span>|
|---|---|
|[<span data-ttu-id="2085e-377">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="2085e-377">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2085e-378">1.0</span><span class="sxs-lookup"><span data-stu-id="2085e-378">1.0</span></span>|
|[<span data-ttu-id="2085e-379">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="2085e-379">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2085e-380">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2085e-380">ReadItem</span></span>|
|[<span data-ttu-id="2085e-381">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="2085e-381">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="2085e-382">Чтение</span><span class="sxs-lookup"><span data-stu-id="2085e-382">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2085e-383">Пример</span><span class="sxs-lookup"><span data-stu-id="2085e-383">Example</span></span>

```javascript
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

---
---

#### <a name="itemclass-string"></a><span data-ttu-id="2085e-384">itemClass: строка</span><span class="sxs-lookup"><span data-stu-id="2085e-384">itemClass: String</span></span>

<span data-ttu-id="2085e-p114">Получает класс элемента веб-служб Exchange для выбранного элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="2085e-p114">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="2085e-p115">Свойство `itemClass` указывает класс сообщения выбранного элемента. Ниже приводятся классы сообщения по умолчанию для элемента сообщения или встречи.</span><span class="sxs-lookup"><span data-stu-id="2085e-p115">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

|<span data-ttu-id="2085e-389">Тип</span><span class="sxs-lookup"><span data-stu-id="2085e-389">Type</span></span>|<span data-ttu-id="2085e-390">Описание</span><span class="sxs-lookup"><span data-stu-id="2085e-390">Description</span></span>|<span data-ttu-id="2085e-391">Класс элемента</span><span class="sxs-lookup"><span data-stu-id="2085e-391">item class</span></span>|
|---|---|---|
|<span data-ttu-id="2085e-392">Элементы встречи</span><span class="sxs-lookup"><span data-stu-id="2085e-392">Appointment items</span></span>|<span data-ttu-id="2085e-393">Это элементы календаря для класса элемента `IPM.Appointment` или `IPM.Appointment.Occurrence`.</span><span class="sxs-lookup"><span data-stu-id="2085e-393">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span>|`IPM.Appointment`<br />`IPM.Appointment.Occurrence`|
|<span data-ttu-id="2085e-394">Элементы сообщения</span><span class="sxs-lookup"><span data-stu-id="2085e-394">Message items</span></span>|<span data-ttu-id="2085e-395">Сюда входят электронные сообщения, для которых по умолчанию задан класс сообщения `IPM.Note`, а также приглашения на собрания, ответы на них и уведомления об их отмене, использующие `IPM.Schedule.Meeting` в качестве базового класса сообщения.</span><span class="sxs-lookup"><span data-stu-id="2085e-395">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span>|`IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled`|

<span data-ttu-id="2085e-396">Можно создавать настраиваемые классы сообщения, расширяющие классы сообщения по умолчанию, например настраиваемый класс сообщения о встрече `IPM.Appointment.Contoso`.</span><span class="sxs-lookup"><span data-stu-id="2085e-396">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="2085e-397">Тип</span><span class="sxs-lookup"><span data-stu-id="2085e-397">Type</span></span>

*   <span data-ttu-id="2085e-398">String</span><span class="sxs-lookup"><span data-stu-id="2085e-398">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="2085e-399">Требования</span><span class="sxs-lookup"><span data-stu-id="2085e-399">Requirements</span></span>

|<span data-ttu-id="2085e-400">Требование</span><span class="sxs-lookup"><span data-stu-id="2085e-400">Requirement</span></span>|<span data-ttu-id="2085e-401">Значение</span><span class="sxs-lookup"><span data-stu-id="2085e-401">Value</span></span>|
|---|---|
|[<span data-ttu-id="2085e-402">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="2085e-402">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2085e-403">1.0</span><span class="sxs-lookup"><span data-stu-id="2085e-403">1.0</span></span>|
|[<span data-ttu-id="2085e-404">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="2085e-404">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2085e-405">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2085e-405">ReadItem</span></span>|
|[<span data-ttu-id="2085e-406">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="2085e-406">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="2085e-407">Чтение</span><span class="sxs-lookup"><span data-stu-id="2085e-407">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2085e-408">Пример</span><span class="sxs-lookup"><span data-stu-id="2085e-408">Example</span></span>

```javascript
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

---
---

#### <a name="nullable-itemid-string"></a><span data-ttu-id="2085e-409">(Nullable) itemId: строка</span><span class="sxs-lookup"><span data-stu-id="2085e-409">(nullable) itemId: String</span></span>

<span data-ttu-id="2085e-p116">Получает идентификатор элемента веб-служб Exchange для текущего элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="2085e-p116">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="2085e-412">Идентификатор, возвращаемый свойством `itemId`, совпадает с идентификатором элемента веб-служб Exchange.</span><span class="sxs-lookup"><span data-stu-id="2085e-412">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="2085e-413">Свойство `itemId` не совпадает с идентификатором записи Outlook, а также идентификатором, который используется REST API Outlook.</span><span class="sxs-lookup"><span data-stu-id="2085e-413">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="2085e-414">Прежде чем совершать вызовы REST API, используя это значение, его необходимо преобразовать с помощью [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="2085e-414">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="2085e-415">Дополнительные сведения см. в статье [Использование REST API Outlook из надстройки Outlook](/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="2085e-415">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="2085e-p118">Свойство `itemId` недоступно в режиме создания. Если требуется идентификатор элемента, с помощью метода [`saveAsync`](#saveasyncoptions-callback) можно сохранить элемент в хранилище. При этом в параметре [`AsyncResult.value`](/javascript/api/office/office.asyncresult) функции обратного вызова возвращается идентификатор элемента.</span><span class="sxs-lookup"><span data-stu-id="2085e-p118">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="2085e-418">Тип</span><span class="sxs-lookup"><span data-stu-id="2085e-418">Type</span></span>

*   <span data-ttu-id="2085e-419">String</span><span class="sxs-lookup"><span data-stu-id="2085e-419">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="2085e-420">Требования</span><span class="sxs-lookup"><span data-stu-id="2085e-420">Requirements</span></span>

|<span data-ttu-id="2085e-421">Требование</span><span class="sxs-lookup"><span data-stu-id="2085e-421">Requirement</span></span>|<span data-ttu-id="2085e-422">Значение</span><span class="sxs-lookup"><span data-stu-id="2085e-422">Value</span></span>|
|---|---|
|[<span data-ttu-id="2085e-423">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="2085e-423">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2085e-424">1.0</span><span class="sxs-lookup"><span data-stu-id="2085e-424">1.0</span></span>|
|[<span data-ttu-id="2085e-425">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="2085e-425">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2085e-426">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2085e-426">ReadItem</span></span>|
|[<span data-ttu-id="2085e-427">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="2085e-427">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="2085e-428">Чтение</span><span class="sxs-lookup"><span data-stu-id="2085e-428">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2085e-429">Пример</span><span class="sxs-lookup"><span data-stu-id="2085e-429">Example</span></span>

<span data-ttu-id="2085e-p119">Указанный ниже код проверяет наличие идентификатора элемента. Если свойство `itemId` возвращает значение `null` или `undefined`, элемент будет сохранен в хранилище, а из асинхронного результата будет получен идентификатор элемента.</span><span class="sxs-lookup"><span data-stu-id="2085e-p119">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

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

#### <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook17officemailboxenumsitemtype"></a><span data-ttu-id="2085e-432">itemType: [Office. MailboxEnums. ItemType](/javascript/api/outlook_1_7/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="2085e-432">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook_1_7/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="2085e-433">Получает тип элемента, который представляет экземпляр.</span><span class="sxs-lookup"><span data-stu-id="2085e-433">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="2085e-434">Свойство `itemType` возвращает одно из значений перечисления `ItemType`, которое указывает, является ли экземпляр объекта `item` сообщением или встречей.</span><span class="sxs-lookup"><span data-stu-id="2085e-434">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="2085e-435">Тип</span><span class="sxs-lookup"><span data-stu-id="2085e-435">Type</span></span>

*   [<span data-ttu-id="2085e-436">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="2085e-436">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook_1_7/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="2085e-437">Требования</span><span class="sxs-lookup"><span data-stu-id="2085e-437">Requirements</span></span>

|<span data-ttu-id="2085e-438">Требование</span><span class="sxs-lookup"><span data-stu-id="2085e-438">Requirement</span></span>|<span data-ttu-id="2085e-439">Значение</span><span class="sxs-lookup"><span data-stu-id="2085e-439">Value</span></span>|
|---|---|
|[<span data-ttu-id="2085e-440">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="2085e-440">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2085e-441">1.0</span><span class="sxs-lookup"><span data-stu-id="2085e-441">1.0</span></span>|
|[<span data-ttu-id="2085e-442">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="2085e-442">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2085e-443">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2085e-443">ReadItem</span></span>|
|[<span data-ttu-id="2085e-444">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="2085e-444">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="2085e-445">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="2085e-445">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2085e-446">Пример</span><span class="sxs-lookup"><span data-stu-id="2085e-446">Example</span></span>

```javascript
if (Office.context.mailbox.item.itemType === Office.MailboxEnums.ItemType.Message) {
  // Do something.
} else {
  // Do something else.
}
```

---
---

#### <a name="location-stringlocationjavascriptapioutlook17officelocation"></a><span data-ttu-id="2085e-447">Местоположение: строка | [Location (расположение](/javascript/api/outlook_1_7/office.location) )</span><span class="sxs-lookup"><span data-stu-id="2085e-447">location: String|[Location](/javascript/api/outlook_1_7/office.location)</span></span>

<span data-ttu-id="2085e-448">Получает или задает место встречи.</span><span class="sxs-lookup"><span data-stu-id="2085e-448">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="2085e-449">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="2085e-449">Read mode</span></span>

<span data-ttu-id="2085e-450">Свойство `location` возвращает строку, содержащую сведения о месте встречи.</span><span class="sxs-lookup"><span data-stu-id="2085e-450">The `location` property returns a string that contains the location of the appointment.</span></span>

```javascript
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="2085e-451">Режим создания</span><span class="sxs-lookup"><span data-stu-id="2085e-451">Compose mode</span></span>

<span data-ttu-id="2085e-452">Свойство `location` возвращает объект `Location`, предоставляющий методы, которые используются для получения и задания места встречи.</span><span class="sxs-lookup"><span data-stu-id="2085e-452">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```javascript
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="2085e-453">Тип</span><span class="sxs-lookup"><span data-stu-id="2085e-453">Type</span></span>

*   <span data-ttu-id="2085e-454">String | [Location](/javascript/api/outlook_1_7/office.location)</span><span class="sxs-lookup"><span data-stu-id="2085e-454">String | [Location](/javascript/api/outlook_1_7/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="2085e-455">Требования</span><span class="sxs-lookup"><span data-stu-id="2085e-455">Requirements</span></span>

|<span data-ttu-id="2085e-456">Требование</span><span class="sxs-lookup"><span data-stu-id="2085e-456">Requirement</span></span>|<span data-ttu-id="2085e-457">Значение</span><span class="sxs-lookup"><span data-stu-id="2085e-457">Value</span></span>|
|---|---|
|[<span data-ttu-id="2085e-458">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="2085e-458">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2085e-459">1.0</span><span class="sxs-lookup"><span data-stu-id="2085e-459">1.0</span></span>|
|[<span data-ttu-id="2085e-460">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="2085e-460">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2085e-461">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2085e-461">ReadItem</span></span>|
|[<span data-ttu-id="2085e-462">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="2085e-462">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="2085e-463">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="2085e-463">Compose or Read</span></span>|

---
---

#### <a name="normalizedsubject-string"></a><span data-ttu-id="2085e-464">normalizedSubject: строка</span><span class="sxs-lookup"><span data-stu-id="2085e-464">normalizedSubject: String</span></span>

<span data-ttu-id="2085e-p120">Получает тему элемента со всеми удаленными префиксами (включая `RE:` и `FWD:`). Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="2085e-p120">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="2085e-p121">Свойство normalizedSubject получает тему элемента со стандартными префиксами (такими как `RE:` и `FW:`), добавляемыми почтовыми программами. Для получения темы элемента с неизмененными префиксами используйте свойство [`subject`](#subject-stringsubject).</span><span class="sxs-lookup"><span data-stu-id="2085e-p121">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="2085e-469">Тип</span><span class="sxs-lookup"><span data-stu-id="2085e-469">Type</span></span>

*   <span data-ttu-id="2085e-470">String</span><span class="sxs-lookup"><span data-stu-id="2085e-470">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="2085e-471">Требования</span><span class="sxs-lookup"><span data-stu-id="2085e-471">Requirements</span></span>

|<span data-ttu-id="2085e-472">Требование</span><span class="sxs-lookup"><span data-stu-id="2085e-472">Requirement</span></span>|<span data-ttu-id="2085e-473">Значение</span><span class="sxs-lookup"><span data-stu-id="2085e-473">Value</span></span>|
|---|---|
|[<span data-ttu-id="2085e-474">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="2085e-474">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2085e-475">1.0</span><span class="sxs-lookup"><span data-stu-id="2085e-475">1.0</span></span>|
|[<span data-ttu-id="2085e-476">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="2085e-476">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2085e-477">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2085e-477">ReadItem</span></span>|
|[<span data-ttu-id="2085e-478">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="2085e-478">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="2085e-479">Чтение</span><span class="sxs-lookup"><span data-stu-id="2085e-479">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2085e-480">Пример</span><span class="sxs-lookup"><span data-stu-id="2085e-480">Example</span></span>

```javascript
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

---
---

#### <a name="notificationmessages-notificationmessagesjavascriptapioutlook17officenotificationmessages"></a><span data-ttu-id="2085e-481">notificationMessages: [notificationMessages](/javascript/api/outlook_1_7/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="2085e-481">notificationMessages: [NotificationMessages](/javascript/api/outlook_1_7/office.notificationmessages)</span></span>

<span data-ttu-id="2085e-482">Получает сообщения уведомления для элемента.</span><span class="sxs-lookup"><span data-stu-id="2085e-482">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="2085e-483">Тип</span><span class="sxs-lookup"><span data-stu-id="2085e-483">Type</span></span>

*   [<span data-ttu-id="2085e-484">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="2085e-484">NotificationMessages</span></span>](/javascript/api/outlook_1_7/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="2085e-485">Требования</span><span class="sxs-lookup"><span data-stu-id="2085e-485">Requirements</span></span>

|<span data-ttu-id="2085e-486">Требование</span><span class="sxs-lookup"><span data-stu-id="2085e-486">Requirement</span></span>|<span data-ttu-id="2085e-487">Значение</span><span class="sxs-lookup"><span data-stu-id="2085e-487">Value</span></span>|
|---|---|
|[<span data-ttu-id="2085e-488">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="2085e-488">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2085e-489">1.3</span><span class="sxs-lookup"><span data-stu-id="2085e-489">1.3</span></span>|
|[<span data-ttu-id="2085e-490">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="2085e-490">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2085e-491">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2085e-491">ReadItem</span></span>|
|[<span data-ttu-id="2085e-492">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="2085e-492">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="2085e-493">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="2085e-493">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2085e-494">Пример</span><span class="sxs-lookup"><span data-stu-id="2085e-494">Example</span></span>

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

#### <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients"></a><span data-ttu-id="2085e-495">optionalAttendees: Array. <[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[получатели](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="2085e-495">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

<span data-ttu-id="2085e-496">Предоставляет доступ к необязательным участникам события.</span><span class="sxs-lookup"><span data-stu-id="2085e-496">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="2085e-497">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="2085e-497">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="2085e-498">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="2085e-498">Read mode</span></span>

<span data-ttu-id="2085e-499">Свойство `optionalAttendees` возвращает массив, содержащий объект `EmailAddressDetails` для каждого необязательного участника собрания.</span><span class="sxs-lookup"><span data-stu-id="2085e-499">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

```javascript
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="2085e-500">Режим создания</span><span class="sxs-lookup"><span data-stu-id="2085e-500">Compose mode</span></span>

<span data-ttu-id="2085e-501">Свойство `optionalAttendees` возвращает объект `Recipients`, который предоставляет методы для получения или обновления необязательных участников собрания.</span><span class="sxs-lookup"><span data-stu-id="2085e-501">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="2085e-502">Тип</span><span class="sxs-lookup"><span data-stu-id="2085e-502">Type</span></span>

*   <span data-ttu-id="2085e-503">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="2085e-503">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="2085e-504">Требования</span><span class="sxs-lookup"><span data-stu-id="2085e-504">Requirements</span></span>

|<span data-ttu-id="2085e-505">Требование</span><span class="sxs-lookup"><span data-stu-id="2085e-505">Requirement</span></span>|<span data-ttu-id="2085e-506">Значение</span><span class="sxs-lookup"><span data-stu-id="2085e-506">Value</span></span>|
|---|---|
|[<span data-ttu-id="2085e-507">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="2085e-507">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2085e-508">1.0</span><span class="sxs-lookup"><span data-stu-id="2085e-508">1.0</span></span>|
|[<span data-ttu-id="2085e-509">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="2085e-509">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2085e-510">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2085e-510">ReadItem</span></span>|
|[<span data-ttu-id="2085e-511">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="2085e-511">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="2085e-512">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="2085e-512">Compose or Read</span></span>|

---
---

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsorganizerjavascriptapioutlook17officeorganizer"></a><span data-ttu-id="2085e-513">Организатор: [EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)|[Организатор](/javascript/api/outlook_1_7/office.organizer)</span><span class="sxs-lookup"><span data-stu-id="2085e-513">organizer: [EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)|[Organizer](/javascript/api/outlook_1_7/office.organizer)</span></span>

<span data-ttu-id="2085e-514">Получает адрес электронной почты организатора для указанного собрания.</span><span class="sxs-lookup"><span data-stu-id="2085e-514">Gets the email address of the organizer for a specified meeting.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="2085e-515">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="2085e-515">Read mode</span></span>

<span data-ttu-id="2085e-516">`organizer` Свойство возвращает объект [EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails) , представляющий организатора собрания.</span><span class="sxs-lookup"><span data-stu-id="2085e-516">The `organizer` property returns an [EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails) object that represents the meeting organizer.</span></span>

```javascript
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

##### <a name="compose-mode"></a><span data-ttu-id="2085e-517">Режим создания</span><span class="sxs-lookup"><span data-stu-id="2085e-517">Compose mode</span></span>

<span data-ttu-id="2085e-518">Свойство возвращает объект организатора, который предоставляет метод для получения значения организатора. [](/javascript/api/outlook_1_7/office.organizer) `organizer`</span><span class="sxs-lookup"><span data-stu-id="2085e-518">The `organizer` property returns an [Organizer](/javascript/api/outlook_1_7/office.organizer) object that provides a method to get the organizer value.</span></span>

```javascript
Office.context.mailbox.item.organizer.getAsync(
  function(asyncResult) {
    console.log(JSON.stringify(asyncResult));
  }
);
```

##### <a name="type"></a><span data-ttu-id="2085e-519">Тип</span><span class="sxs-lookup"><span data-stu-id="2085e-519">Type</span></span>

*   <span data-ttu-id="2085e-520">[](/javascript/api/outlook_1_7/office.emailaddressdetails) | [Организатор](/javascript/api/outlook_1_7/office.organizer) EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="2085e-520">[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails) | [Organizer](/javascript/api/outlook_1_7/office.organizer)</span></span>

##### <a name="requirements"></a><span data-ttu-id="2085e-521">Требования</span><span class="sxs-lookup"><span data-stu-id="2085e-521">Requirements</span></span>

|<span data-ttu-id="2085e-522">Требование</span><span class="sxs-lookup"><span data-stu-id="2085e-522">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="2085e-523">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="2085e-523">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2085e-524">1.0</span><span class="sxs-lookup"><span data-stu-id="2085e-524">1.0</span></span>|<span data-ttu-id="2085e-525">1.7</span><span class="sxs-lookup"><span data-stu-id="2085e-525">1.7</span></span>|
|[<span data-ttu-id="2085e-526">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="2085e-526">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2085e-527">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2085e-527">ReadItem</span></span>|<span data-ttu-id="2085e-528">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="2085e-528">ReadWriteItem</span></span>|
|[<span data-ttu-id="2085e-529">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="2085e-529">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="2085e-530">Чтение</span><span class="sxs-lookup"><span data-stu-id="2085e-530">Read</span></span>|<span data-ttu-id="2085e-531">Создание</span><span class="sxs-lookup"><span data-stu-id="2085e-531">Compose</span></span>|

---
---

#### <a name="nullable-recurrence-recurrencejavascriptapioutlook17officerecurrence"></a><span data-ttu-id="2085e-532">(Nullable) повторение [](/javascript/api/outlook_1_7/office.recurrence) : повторение</span><span class="sxs-lookup"><span data-stu-id="2085e-532">(nullable) recurrence: [Recurrence](/javascript/api/outlook_1_7/office.recurrence)</span></span>

<span data-ttu-id="2085e-533">Получает или задает шаблон повторения встречи.</span><span class="sxs-lookup"><span data-stu-id="2085e-533">Gets or sets the recurrence pattern of an appointment.</span></span> <span data-ttu-id="2085e-534">Получает шаблон повторения приглашения на собрание.</span><span class="sxs-lookup"><span data-stu-id="2085e-534">Gets the recurrence pattern of a meeting request.</span></span> <span data-ttu-id="2085e-535">Режимы чтения и создания для элементов встречи.</span><span class="sxs-lookup"><span data-stu-id="2085e-535">Read and compose modes for appointment items.</span></span> <span data-ttu-id="2085e-536">Режим чтения для элементов приглашения на собрания.</span><span class="sxs-lookup"><span data-stu-id="2085e-536">Read mode for meeting request items.</span></span>

<span data-ttu-id="2085e-537">`recurrence` Свойство возвращает объект [повторения](/javascript/api/outlook_1_7/office.recurrence) для повторяющихся встреч или приглашений на собрания, если элемент представляет собой серию или экземпляр в ряду.</span><span class="sxs-lookup"><span data-stu-id="2085e-537">The `recurrence` property returns a [recurrence](/javascript/api/outlook_1_7/office.recurrence) object for recurring appointments or meetings requests if an item is a series or an instance in a series.</span></span> <span data-ttu-id="2085e-538">`null`возвращается для отдельных встреч и приглашений на собрание для отдельных встреч.</span><span class="sxs-lookup"><span data-stu-id="2085e-538">`null` is returned for single appointments and meeting requests of single appointments.</span></span> <span data-ttu-id="2085e-539">`undefined`возвращается для сообщений, которые не являются приглашениями на собрания.</span><span class="sxs-lookup"><span data-stu-id="2085e-539">`undefined` is returned for messages that are not meeting requests.</span></span>

> <span data-ttu-id="2085e-540">Note: приглашения на `itemClass` собрания имеют значение IPM. Schedule. Meeting. Request.</span><span class="sxs-lookup"><span data-stu-id="2085e-540">Note: Meeting requests have an `itemClass` value of IPM.Schedule.Meeting.Request.</span></span>

> <span data-ttu-id="2085e-541">Note: при наличии объекта `null`повторения это указывает на то, что объект является одной встречей или приглашением на собрание одной встречи, а не частью ряда.</span><span class="sxs-lookup"><span data-stu-id="2085e-541">Note: If the recurrence object is `null`, this indicates that the object is a single appointment or a meeting request of a single appointment and NOT a part of a series.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="2085e-542">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="2085e-542">Read mode</span></span>

<span data-ttu-id="2085e-543">`recurrence` Свойство возвращает объект [повторения](/javascript/api/outlook_1_7/office.recurrence) , представляющий повторение встречи.</span><span class="sxs-lookup"><span data-stu-id="2085e-543">The `recurrence` property returns a [Recurrence](/javascript/api/outlook_1_7/office.recurrence) object that represents the appointment recurrence.</span></span> <span data-ttu-id="2085e-544">Оно доступно для встреч и приглашений на собрания.</span><span class="sxs-lookup"><span data-stu-id="2085e-544">This is available for appointments and meeting requests.</span></span>

```javascript
var recurrence = Office.context.mailbox.item.recurrence;
console.log("Recurrence: " + JSON.stringify(recurrence));
```

##### <a name="compose-mode"></a><span data-ttu-id="2085e-545">Режим создания</span><span class="sxs-lookup"><span data-stu-id="2085e-545">Compose mode</span></span>

<span data-ttu-id="2085e-546">`recurrence` Свойство возвращает объект [повторения](/javascript/api/outlook_1_7/office.recurrence) , который предоставляет методы для управления повторением встречи.</span><span class="sxs-lookup"><span data-stu-id="2085e-546">The `recurrence` property returns a [Recurrence](/javascript/api/outlook_1_7/office.recurrence) object that provides methods to manage the appointment recurrence.</span></span> <span data-ttu-id="2085e-547">Оно доступно для встреч.</span><span class="sxs-lookup"><span data-stu-id="2085e-547">This is available for appointments.</span></span>

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

##### <a name="type"></a><span data-ttu-id="2085e-548">Тип</span><span class="sxs-lookup"><span data-stu-id="2085e-548">Type</span></span>

* [<span data-ttu-id="2085e-549">Повторения</span><span class="sxs-lookup"><span data-stu-id="2085e-549">Recurrence</span></span>](/javascript/api/outlook_1_7/office.recurrence)

|<span data-ttu-id="2085e-550">Требование</span><span class="sxs-lookup"><span data-stu-id="2085e-550">Requirement</span></span>|<span data-ttu-id="2085e-551">Значение</span><span class="sxs-lookup"><span data-stu-id="2085e-551">Value</span></span>|
|---|---|
|[<span data-ttu-id="2085e-552">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="2085e-552">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2085e-553">1.7</span><span class="sxs-lookup"><span data-stu-id="2085e-553">1.7</span></span>|
|[<span data-ttu-id="2085e-554">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="2085e-554">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2085e-555">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2085e-555">ReadItem</span></span>|
|[<span data-ttu-id="2085e-556">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="2085e-556">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="2085e-557">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="2085e-557">Compose or Read</span></span>|

---
---

#### <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients"></a><span data-ttu-id="2085e-558">requiredAttendees: Array. <[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[получатели](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="2085e-558">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

<span data-ttu-id="2085e-559">Предоставляет доступ к обязательным участникам события.</span><span class="sxs-lookup"><span data-stu-id="2085e-559">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="2085e-560">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="2085e-560">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="2085e-561">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="2085e-561">Read mode</span></span>

<span data-ttu-id="2085e-562">Свойство `requiredAttendees` возвращает массив, содержащий объект `EmailAddressDetails` для каждого обязательного участника собрания.</span><span class="sxs-lookup"><span data-stu-id="2085e-562">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

```javascript
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="2085e-563">Режим создания</span><span class="sxs-lookup"><span data-stu-id="2085e-563">Compose mode</span></span>

<span data-ttu-id="2085e-564">Свойство `requiredAttendees` возвращает объект `Recipients`, предоставляющий методы, с помощью которых можно получить или обновить сведения об обязательных участниках собрания.</span><span class="sxs-lookup"><span data-stu-id="2085e-564">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="2085e-565">Тип</span><span class="sxs-lookup"><span data-stu-id="2085e-565">Type</span></span>

*   <span data-ttu-id="2085e-566">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="2085e-566">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="2085e-567">Требования</span><span class="sxs-lookup"><span data-stu-id="2085e-567">Requirements</span></span>

|<span data-ttu-id="2085e-568">Требование</span><span class="sxs-lookup"><span data-stu-id="2085e-568">Requirement</span></span>|<span data-ttu-id="2085e-569">Значение</span><span class="sxs-lookup"><span data-stu-id="2085e-569">Value</span></span>|
|---|---|
|[<span data-ttu-id="2085e-570">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="2085e-570">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2085e-571">1.0</span><span class="sxs-lookup"><span data-stu-id="2085e-571">1.0</span></span>|
|[<span data-ttu-id="2085e-572">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="2085e-572">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2085e-573">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2085e-573">ReadItem</span></span>|
|[<span data-ttu-id="2085e-574">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="2085e-574">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="2085e-575">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="2085e-575">Compose or Read</span></span>|

---
---

#### <a name="sender-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetails"></a><span data-ttu-id="2085e-576">Отправитель: [EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="2085e-576">sender: [EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)</span></span>

<span data-ttu-id="2085e-p128">Получает электронный адрес отправителя электронного сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="2085e-p128">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="2085e-p129">Свойства [`from`](#from-emailaddressdetailsfrom) и `sender` представляют одно лицо, если сообщение не отправлено представителем. В противном случае свойство `from` представляет лицо, делегировавшее полномочия, а свойство sender — представителя.</span><span class="sxs-lookup"><span data-stu-id="2085e-p129">The [`from`](#from-emailaddressdetailsfrom) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="2085e-581">Свойству `recipientType`, принадлежащему объекту `EmailAddressDetails` в свойстве `sender`, задано значение `undefined`.</span><span class="sxs-lookup"><span data-stu-id="2085e-581">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="2085e-582">Тип</span><span class="sxs-lookup"><span data-stu-id="2085e-582">Type</span></span>

*   [<span data-ttu-id="2085e-583">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="2085e-583">EmailAddressDetails</span></span>](/javascript/api/outlook_1_7/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="2085e-584">Требования</span><span class="sxs-lookup"><span data-stu-id="2085e-584">Requirements</span></span>

|<span data-ttu-id="2085e-585">Требование</span><span class="sxs-lookup"><span data-stu-id="2085e-585">Requirement</span></span>|<span data-ttu-id="2085e-586">Значение</span><span class="sxs-lookup"><span data-stu-id="2085e-586">Value</span></span>|
|---|---|
|[<span data-ttu-id="2085e-587">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="2085e-587">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2085e-588">1.0</span><span class="sxs-lookup"><span data-stu-id="2085e-588">1.0</span></span>|
|[<span data-ttu-id="2085e-589">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="2085e-589">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2085e-590">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2085e-590">ReadItem</span></span>|
|[<span data-ttu-id="2085e-591">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="2085e-591">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="2085e-592">Чтение</span><span class="sxs-lookup"><span data-stu-id="2085e-592">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2085e-593">Пример</span><span class="sxs-lookup"><span data-stu-id="2085e-593">Example</span></span>

```javascript
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

---
---

#### <a name="nullable-seriesid-string"></a><span data-ttu-id="2085e-594">(Nullable) seriesId: строка</span><span class="sxs-lookup"><span data-stu-id="2085e-594">(nullable) seriesId: String</span></span>

<span data-ttu-id="2085e-595">Получает идентификатор ряда, к которому принадлежит экземпляр.</span><span class="sxs-lookup"><span data-stu-id="2085e-595">Gets the id of the series that an instance belongs to.</span></span>

<span data-ttu-id="2085e-596">В OWA и Outlook `seriesId` возвращается идентификатор веб-служб Exchange (EWS) родительского элемента (ряда), к которому принадлежит этот элемент.</span><span class="sxs-lookup"><span data-stu-id="2085e-596">In OWA and Outlook, the `seriesId` returns the Exchange Web Services (EWS) ID of the parent (series) item that this item belongs to.</span></span> <span data-ttu-id="2085e-597">Однако в iOS и Android `seriesId` возвращается идентификатор REST родительского элемента.</span><span class="sxs-lookup"><span data-stu-id="2085e-597">However, in iOS and Android, the `seriesId` returns the REST ID of the parent item.</span></span>

> [!NOTE]
> <span data-ttu-id="2085e-598">Идентификатор, возвращаемый свойством `seriesId`, совпадает с идентификатором элемента веб-служб Exchange.</span><span class="sxs-lookup"><span data-stu-id="2085e-598">The identifier returned by the `seriesId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="2085e-599">`seriesId` Свойство не совпадает с идентификаторами Outlook, используемыми в REST API Outlook.</span><span class="sxs-lookup"><span data-stu-id="2085e-599">The `seriesId` property is not identical to the Outlook IDs used by the Outlook REST API.</span></span> <span data-ttu-id="2085e-600">Прежде чем совершать вызовы REST API, используя это значение, его необходимо преобразовать с помощью [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="2085e-600">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="2085e-601">Дополнительные сведения см. в статье [Использование REST API Outlook из надстройки Outlook](/outlook/add-ins/use-rest-api).</span><span class="sxs-lookup"><span data-stu-id="2085e-601">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api).</span></span>

<span data-ttu-id="2085e-602">`seriesId` Свойство возвращает `null` элементы, у которых нет родительских элементов, таких как одиночные встречи, элементы ряда или приглашения на собрание, `undefined` и возвращаемые для других элементов, не являющиеся приглашениями на собрания.</span><span class="sxs-lookup"><span data-stu-id="2085e-602">The `seriesId` property returns `null` for items that do not have parent items such as single appointments, series items, or meeting requests and returns `undefined` for any other items that are not meeting requests.</span></span>

##### <a name="type"></a><span data-ttu-id="2085e-603">Тип</span><span class="sxs-lookup"><span data-stu-id="2085e-603">Type</span></span>

* <span data-ttu-id="2085e-604">String</span><span class="sxs-lookup"><span data-stu-id="2085e-604">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="2085e-605">Требования</span><span class="sxs-lookup"><span data-stu-id="2085e-605">Requirements</span></span>

|<span data-ttu-id="2085e-606">Требование</span><span class="sxs-lookup"><span data-stu-id="2085e-606">Requirement</span></span>|<span data-ttu-id="2085e-607">Значение</span><span class="sxs-lookup"><span data-stu-id="2085e-607">Value</span></span>|
|---|---|
|[<span data-ttu-id="2085e-608">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="2085e-608">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2085e-609">1.7</span><span class="sxs-lookup"><span data-stu-id="2085e-609">1.7</span></span>|
|[<span data-ttu-id="2085e-610">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="2085e-610">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2085e-611">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2085e-611">ReadItem</span></span>|
|[<span data-ttu-id="2085e-612">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="2085e-612">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="2085e-613">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="2085e-613">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2085e-614">Пример</span><span class="sxs-lookup"><span data-stu-id="2085e-614">Example</span></span>

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

#### <a name="start-datetimejavascriptapioutlook17officetime"></a><span data-ttu-id="2085e-615">Начало: Дата | [Time (время](/javascript/api/outlook_1_7/office.time) )</span><span class="sxs-lookup"><span data-stu-id="2085e-615">start: Date|[Time](/javascript/api/outlook_1_7/office.time)</span></span>

<span data-ttu-id="2085e-616">Получает или задает дату и время начала встречи.</span><span class="sxs-lookup"><span data-stu-id="2085e-616">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="2085e-p132">Свойство `start` представлено в виде значения даты и времени в формате UTC. Это значение можно преобразовать в местные значения даты и времени клиента с помощью метода [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime).</span><span class="sxs-lookup"><span data-stu-id="2085e-p132">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="2085e-619">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="2085e-619">Read mode</span></span>

<span data-ttu-id="2085e-620">Свойство `start` возвращает объект `Date`.</span><span class="sxs-lookup"><span data-stu-id="2085e-620">The `start` property returns a `Date` object.</span></span>

```javascript
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="2085e-621">Режим создания</span><span class="sxs-lookup"><span data-stu-id="2085e-621">Compose mode</span></span>

<span data-ttu-id="2085e-622">Свойство `start` возвращает объект `Time`.</span><span class="sxs-lookup"><span data-stu-id="2085e-622">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="2085e-623">Если вы задаете время начала с помощью метода [`Time.setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-), необходимо использовать метод [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) для преобразования местного времени на клиенте в формат UTC для сервера.</span><span class="sxs-lookup"><span data-stu-id="2085e-623">When you use the [`Time.setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="2085e-624">В примере ниже с помощью метода [`setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) объекта `Time` задается время начала встречи в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="2085e-624">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="2085e-625">Тип</span><span class="sxs-lookup"><span data-stu-id="2085e-625">Type</span></span>

*   <span data-ttu-id="2085e-626">Date | [Time](/javascript/api/outlook_1_7/office.time)</span><span class="sxs-lookup"><span data-stu-id="2085e-626">Date | [Time](/javascript/api/outlook_1_7/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="2085e-627">Требования</span><span class="sxs-lookup"><span data-stu-id="2085e-627">Requirements</span></span>

|<span data-ttu-id="2085e-628">Требование</span><span class="sxs-lookup"><span data-stu-id="2085e-628">Requirement</span></span>|<span data-ttu-id="2085e-629">Значение</span><span class="sxs-lookup"><span data-stu-id="2085e-629">Value</span></span>|
|---|---|
|[<span data-ttu-id="2085e-630">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="2085e-630">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2085e-631">1.0</span><span class="sxs-lookup"><span data-stu-id="2085e-631">1.0</span></span>|
|[<span data-ttu-id="2085e-632">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="2085e-632">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2085e-633">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2085e-633">ReadItem</span></span>|
|[<span data-ttu-id="2085e-634">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="2085e-634">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="2085e-635">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="2085e-635">Compose or Read</span></span>|

---
---

#### <a name="subject-stringsubjectjavascriptapioutlook17officesubject"></a><span data-ttu-id="2085e-636">Тема: строка | [Subject (тема](/javascript/api/outlook_1_7/office.subject) )</span><span class="sxs-lookup"><span data-stu-id="2085e-636">subject: String|[Subject](/javascript/api/outlook_1_7/office.subject)</span></span>

<span data-ttu-id="2085e-637">Получает или задает описание, которое отображается в поле темы элемента.</span><span class="sxs-lookup"><span data-stu-id="2085e-637">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="2085e-638">Свойство `subject` получает или задает всю тему элемента для отправки с почтового сервера.</span><span class="sxs-lookup"><span data-stu-id="2085e-638">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="2085e-639">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="2085e-639">Read mode</span></span>

<span data-ttu-id="2085e-p133">Свойство `subject` возвращает строку. С помощью свойства [`normalizedSubject`](#normalizedsubject-string) можно получить тему без начальных префиксов, таких как `RE:` и `FW:`.</span><span class="sxs-lookup"><span data-stu-id="2085e-p133">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

<span data-ttu-id="2085e-642">В примере кода JavaScript, приведенном ниже, показано, как получить доступ к свойству `subject` текущего элемента в Outlook.</span><span class="sxs-lookup"><span data-stu-id="2085e-642">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

```javascript
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="2085e-643">Режим создания</span><span class="sxs-lookup"><span data-stu-id="2085e-643">Compose mode</span></span>

<span data-ttu-id="2085e-644">Свойство `subject` возвращает объект `Subject`, который предоставляет методы для получения и задания темы.</span><span class="sxs-lookup"><span data-stu-id="2085e-644">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```javascript
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="2085e-645">Тип</span><span class="sxs-lookup"><span data-stu-id="2085e-645">Type</span></span>

*   <span data-ttu-id="2085e-646">String | [Subject](/javascript/api/outlook_1_7/office.subject)</span><span class="sxs-lookup"><span data-stu-id="2085e-646">String | [Subject](/javascript/api/outlook_1_7/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="2085e-647">Требования</span><span class="sxs-lookup"><span data-stu-id="2085e-647">Requirements</span></span>

|<span data-ttu-id="2085e-648">Требование</span><span class="sxs-lookup"><span data-stu-id="2085e-648">Requirement</span></span>|<span data-ttu-id="2085e-649">Значение</span><span class="sxs-lookup"><span data-stu-id="2085e-649">Value</span></span>|
|---|---|
|[<span data-ttu-id="2085e-650">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="2085e-650">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2085e-651">1.0</span><span class="sxs-lookup"><span data-stu-id="2085e-651">1.0</span></span>|
|[<span data-ttu-id="2085e-652">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="2085e-652">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2085e-653">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2085e-653">ReadItem</span></span>|
|[<span data-ttu-id="2085e-654">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="2085e-654">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="2085e-655">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="2085e-655">Compose or Read</span></span>|

---
---

#### <a name="to-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients"></a><span data-ttu-id="2085e-656">Кому: Array. <[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[получатели](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="2085e-656">to: Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

<span data-ttu-id="2085e-657">Предоставляет доступ к получателям, указанным в строке **Кому** сообщения.</span><span class="sxs-lookup"><span data-stu-id="2085e-657">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="2085e-658">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="2085e-658">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="2085e-659">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="2085e-659">Read mode</span></span>

<span data-ttu-id="2085e-p135">Свойство `to` возвращает массив, содержащий объект `EmailAddressDetails` для каждого получателя в строке **Кому** сообщения. Коллекция может включать не более 100 элементов.</span><span class="sxs-lookup"><span data-stu-id="2085e-p135">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="2085e-662">Режим создания</span><span class="sxs-lookup"><span data-stu-id="2085e-662">Compose mode</span></span>

<span data-ttu-id="2085e-663">Свойство `to` возвращает объект `Recipients`, предоставляющий методы для получения или обновления получателей, которые указаны в строке **Кому** сообщения.</span><span class="sxs-lookup"><span data-stu-id="2085e-663">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

```javascript
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="2085e-664">Тип</span><span class="sxs-lookup"><span data-stu-id="2085e-664">Type</span></span>

*   <span data-ttu-id="2085e-665">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="2085e-665">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="2085e-666">Требования</span><span class="sxs-lookup"><span data-stu-id="2085e-666">Requirements</span></span>

|<span data-ttu-id="2085e-667">Требование</span><span class="sxs-lookup"><span data-stu-id="2085e-667">Requirement</span></span>|<span data-ttu-id="2085e-668">Значение</span><span class="sxs-lookup"><span data-stu-id="2085e-668">Value</span></span>|
|---|---|
|[<span data-ttu-id="2085e-669">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="2085e-669">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2085e-670">1.0</span><span class="sxs-lookup"><span data-stu-id="2085e-670">1.0</span></span>|
|[<span data-ttu-id="2085e-671">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="2085e-671">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2085e-672">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2085e-672">ReadItem</span></span>|
|[<span data-ttu-id="2085e-673">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="2085e-673">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="2085e-674">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="2085e-674">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="2085e-675">Методы</span><span class="sxs-lookup"><span data-stu-id="2085e-675">Methods</span></span>

#### <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="2085e-676">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="2085e-676">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="2085e-677">Добавляет файл в сообщение или встречу в качестве вложения.</span><span class="sxs-lookup"><span data-stu-id="2085e-677">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="2085e-678">Метод `addFileAttachmentAsync` передает файл по указанному универсальному коду ресурса (URI) и вкладывает его в элемент в форме создания.</span><span class="sxs-lookup"><span data-stu-id="2085e-678">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="2085e-679">Идентификатор можно использовать с методом [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback), чтобы удалить вложение, добавленное во время текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="2085e-679">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2085e-680">Параметры</span><span class="sxs-lookup"><span data-stu-id="2085e-680">Parameters</span></span>
|<span data-ttu-id="2085e-681">Имя</span><span class="sxs-lookup"><span data-stu-id="2085e-681">Name</span></span>|<span data-ttu-id="2085e-682">Тип</span><span class="sxs-lookup"><span data-stu-id="2085e-682">Type</span></span>|<span data-ttu-id="2085e-683">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="2085e-683">Attributes</span></span>|<span data-ttu-id="2085e-684">Описание</span><span class="sxs-lookup"><span data-stu-id="2085e-684">Description</span></span>|
|---|---|---|---|
|`uri`|<span data-ttu-id="2085e-685">String</span><span class="sxs-lookup"><span data-stu-id="2085e-685">String</span></span>||<span data-ttu-id="2085e-p136">Универсальный код ресурса (URI), представляющий расположение файла, который нужно вложить в сообщение или встречу. Максимальная длина — 2048 символов.</span><span class="sxs-lookup"><span data-stu-id="2085e-p136">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="2085e-688">String</span><span class="sxs-lookup"><span data-stu-id="2085e-688">String</span></span>||<span data-ttu-id="2085e-p137">Имя вложения, которое отображается при передаче вложения. Максимальная длина — 255 символов.</span><span class="sxs-lookup"><span data-stu-id="2085e-p137">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="2085e-691">Объект</span><span class="sxs-lookup"><span data-stu-id="2085e-691">Object</span></span>|<span data-ttu-id="2085e-692">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="2085e-692">&lt;optional&gt;</span></span>|<span data-ttu-id="2085e-693">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="2085e-693">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="2085e-694">Object</span><span class="sxs-lookup"><span data-stu-id="2085e-694">Object</span></span>|<span data-ttu-id="2085e-695">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="2085e-695">&lt;optional&gt;</span></span>|<span data-ttu-id="2085e-696">В методе обратного вызова разработчики могут указать любой объект, к которому необходимо получить доступ.</span><span class="sxs-lookup"><span data-stu-id="2085e-696">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="2085e-697">Boolean</span><span class="sxs-lookup"><span data-stu-id="2085e-697">Boolean</span></span>|<span data-ttu-id="2085e-698">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="2085e-698">&lt;optional&gt;</span></span>|<span data-ttu-id="2085e-699">Значение `true` указывает на то, что вложение будет встроено в текст сообщения и не должно отображаться в списке вложений.</span><span class="sxs-lookup"><span data-stu-id="2085e-699">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="2085e-700">function</span><span class="sxs-lookup"><span data-stu-id="2085e-700">function</span></span>|<span data-ttu-id="2085e-701">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="2085e-701">&lt;optional&gt;</span></span>|<span data-ttu-id="2085e-702">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="2085e-702">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="2085e-703">После успешного выполнения идентификатор вложения будет представлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="2085e-703">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="2085e-704">Если добавить вложение не удастся, объект `asyncResult` будет содержать объект `Error` с описанием ошибки.</span><span class="sxs-lookup"><span data-stu-id="2085e-704">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="2085e-705">Ошибки</span><span class="sxs-lookup"><span data-stu-id="2085e-705">Errors</span></span>

|<span data-ttu-id="2085e-706">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="2085e-706">Error code</span></span>|<span data-ttu-id="2085e-707">Описание</span><span class="sxs-lookup"><span data-stu-id="2085e-707">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="2085e-708">Вложение превышает максимальный размер.</span><span class="sxs-lookup"><span data-stu-id="2085e-708">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="2085e-709">Расширение вложения не поддерживается.</span><span class="sxs-lookup"><span data-stu-id="2085e-709">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="2085e-710">Сообщение или встреча содержат слишком много вложений.</span><span class="sxs-lookup"><span data-stu-id="2085e-710">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2085e-711">Требования</span><span class="sxs-lookup"><span data-stu-id="2085e-711">Requirements</span></span>

|<span data-ttu-id="2085e-712">Требование</span><span class="sxs-lookup"><span data-stu-id="2085e-712">Requirement</span></span>|<span data-ttu-id="2085e-713">Значение</span><span class="sxs-lookup"><span data-stu-id="2085e-713">Value</span></span>|
|---|---|
|[<span data-ttu-id="2085e-714">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="2085e-714">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2085e-715">1.1</span><span class="sxs-lookup"><span data-stu-id="2085e-715">1.1</span></span>|
|[<span data-ttu-id="2085e-716">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="2085e-716">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2085e-717">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="2085e-717">ReadWriteItem</span></span>|
|[<span data-ttu-id="2085e-718">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="2085e-718">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="2085e-719">Создание</span><span class="sxs-lookup"><span data-stu-id="2085e-719">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="2085e-720">Примеры</span><span class="sxs-lookup"><span data-stu-id="2085e-720">Examples</span></span>

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

<span data-ttu-id="2085e-721">В приведенном ниже примере файл изображения добавляется в качестве встроенного вложения, а ссылка на вложение добавляется в текст сообщения.</span><span class="sxs-lookup"><span data-stu-id="2085e-721">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

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

#### <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="2085e-722">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="2085e-722">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="2085e-723">Добавляет обработчик для поддерживаемого события.</span><span class="sxs-lookup"><span data-stu-id="2085e-723">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="2085e-724">В настоящее время поддерживаются типы `Office.EventType.AppointmentTimeChanged`событий `Office.EventType.RecipientsChanged`, и`Office.EventType.RecurrenceChanged`</span><span class="sxs-lookup"><span data-stu-id="2085e-724">Currently the supported event types are `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`</span></span>

##### <a name="parameters"></a><span data-ttu-id="2085e-725">Параметры</span><span class="sxs-lookup"><span data-stu-id="2085e-725">Parameters</span></span>

| <span data-ttu-id="2085e-726">Имя</span><span class="sxs-lookup"><span data-stu-id="2085e-726">Name</span></span> | <span data-ttu-id="2085e-727">Тип</span><span class="sxs-lookup"><span data-stu-id="2085e-727">Type</span></span> | <span data-ttu-id="2085e-728">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="2085e-728">Attributes</span></span> | <span data-ttu-id="2085e-729">Описание</span><span class="sxs-lookup"><span data-stu-id="2085e-729">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="2085e-730">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="2085e-730">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="2085e-731">Событие, которое должно вызвать обработчик.</span><span class="sxs-lookup"><span data-stu-id="2085e-731">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="2085e-732">Function</span><span class="sxs-lookup"><span data-stu-id="2085e-732">Function</span></span> || <span data-ttu-id="2085e-p138">Функция для обработки события. Функция должна принимать один параметр, представляющий собой объектный литерал. Значение свойства `type` параметра совпадет со значением параметра `eventType`, переданного методу `addHandlerAsync`.</span><span class="sxs-lookup"><span data-stu-id="2085e-p138">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="2085e-736">Объект</span><span class="sxs-lookup"><span data-stu-id="2085e-736">Object</span></span> | <span data-ttu-id="2085e-737">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="2085e-737">&lt;optional&gt;</span></span> | <span data-ttu-id="2085e-738">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="2085e-738">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="2085e-739">Объект</span><span class="sxs-lookup"><span data-stu-id="2085e-739">Object</span></span> | <span data-ttu-id="2085e-740">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="2085e-740">&lt;optional&gt;</span></span> | <span data-ttu-id="2085e-741">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="2085e-741">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="2085e-742">функция</span><span class="sxs-lookup"><span data-stu-id="2085e-742">function</span></span>| <span data-ttu-id="2085e-743">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="2085e-743">&lt;optional&gt;</span></span>|<span data-ttu-id="2085e-744">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="2085e-744">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2085e-745">Требования</span><span class="sxs-lookup"><span data-stu-id="2085e-745">Requirements</span></span>

|<span data-ttu-id="2085e-746">Требование</span><span class="sxs-lookup"><span data-stu-id="2085e-746">Requirement</span></span>| <span data-ttu-id="2085e-747">Значение</span><span class="sxs-lookup"><span data-stu-id="2085e-747">Value</span></span>|
|---|---|
|[<span data-ttu-id="2085e-748">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="2085e-748">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2085e-749">1.7</span><span class="sxs-lookup"><span data-stu-id="2085e-749">1.7</span></span> |
|[<span data-ttu-id="2085e-750">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="2085e-750">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2085e-751">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2085e-751">ReadItem</span></span> |
|[<span data-ttu-id="2085e-752">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="2085e-752">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2085e-753">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="2085e-753">Compose or Read</span></span> |

##### <a name="example"></a><span data-ttu-id="2085e-754">Пример</span><span class="sxs-lookup"><span data-stu-id="2085e-754">Example</span></span>

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

#### <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="2085e-755">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="2085e-755">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="2085e-756">Добавляет к сообщению элемент Exchange, например сообщение, в виде вложения.</span><span class="sxs-lookup"><span data-stu-id="2085e-756">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="2085e-p139">С помощью метода `addItemAttachmentAsync` можно в элемент формы создания вложить элемент с указанным идентификатором Exchange. Если указать метод обратного вызова, то этот метод вызывается с помощью параметра `asyncResult`, который содержит идентификатор вложения или код, указывающий на ошибки, которые произошли при вложении элемента. При необходимости можно использовать параметр `options` для передачи сведений о состоянии методу обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="2085e-p139">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="2085e-760">Идентификатор можно использовать с методом [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback), чтобы удалить вложение, добавленное во время текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="2085e-760">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="2085e-761">Если ваша надстройка Office выполняется в Outlook Web App, метод `addItemAttachmentAsync` обеспечивает вложение элементов в элементы, отличные от редактируемого. Однако это действие не рекомендуем выполнять, так как оно не поддерживается.</span><span class="sxs-lookup"><span data-stu-id="2085e-761">If your Office Add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2085e-762">Параметры</span><span class="sxs-lookup"><span data-stu-id="2085e-762">Parameters</span></span>

|<span data-ttu-id="2085e-763">Имя</span><span class="sxs-lookup"><span data-stu-id="2085e-763">Name</span></span>|<span data-ttu-id="2085e-764">Тип</span><span class="sxs-lookup"><span data-stu-id="2085e-764">Type</span></span>|<span data-ttu-id="2085e-765">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="2085e-765">Attributes</span></span>|<span data-ttu-id="2085e-766">Описание</span><span class="sxs-lookup"><span data-stu-id="2085e-766">Description</span></span>|
|---|---|---|---|
|`itemId`|<span data-ttu-id="2085e-767">String</span><span class="sxs-lookup"><span data-stu-id="2085e-767">String</span></span>||<span data-ttu-id="2085e-p140">Идентификатор Exchange для вкладываемого элемента. Максимальная длина — 100 символов.</span><span class="sxs-lookup"><span data-stu-id="2085e-p140">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="2085e-770">String</span><span class="sxs-lookup"><span data-stu-id="2085e-770">String</span></span>||<span data-ttu-id="2085e-771">Тема вкладываемого элемента.</span><span class="sxs-lookup"><span data-stu-id="2085e-771">The subject of the item to be attached.</span></span> <span data-ttu-id="2085e-772">Максимальная длина: 255 символов.</span><span class="sxs-lookup"><span data-stu-id="2085e-772">The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="2085e-773">Object</span><span class="sxs-lookup"><span data-stu-id="2085e-773">Object</span></span>|<span data-ttu-id="2085e-774">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="2085e-774">&lt;optional&gt;</span></span>|<span data-ttu-id="2085e-775">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="2085e-775">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="2085e-776">Объект</span><span class="sxs-lookup"><span data-stu-id="2085e-776">Object</span></span>|<span data-ttu-id="2085e-777">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="2085e-777">&lt;optional&gt;</span></span>|<span data-ttu-id="2085e-778">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="2085e-778">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="2085e-779">функция</span><span class="sxs-lookup"><span data-stu-id="2085e-779">function</span></span>|<span data-ttu-id="2085e-780">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="2085e-780">&lt;optional&gt;</span></span>|<span data-ttu-id="2085e-781">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="2085e-781">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="2085e-782">После успешного выполнения идентификатор вложения будет представлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="2085e-782">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="2085e-783">Если добавить вложение не удастся, объект `asyncResult` будет содержать объект `Error` с описанием ошибки.</span><span class="sxs-lookup"><span data-stu-id="2085e-783">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="2085e-784">Ошибки</span><span class="sxs-lookup"><span data-stu-id="2085e-784">Errors</span></span>

|<span data-ttu-id="2085e-785">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="2085e-785">Error code</span></span>|<span data-ttu-id="2085e-786">Описание</span><span class="sxs-lookup"><span data-stu-id="2085e-786">Description</span></span>|
|------------|-------------|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="2085e-787">Сообщение или встреча содержат слишком много вложений.</span><span class="sxs-lookup"><span data-stu-id="2085e-787">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2085e-788">Требования</span><span class="sxs-lookup"><span data-stu-id="2085e-788">Requirements</span></span>

|<span data-ttu-id="2085e-789">Требование</span><span class="sxs-lookup"><span data-stu-id="2085e-789">Requirement</span></span>|<span data-ttu-id="2085e-790">Значение</span><span class="sxs-lookup"><span data-stu-id="2085e-790">Value</span></span>|
|---|---|
|[<span data-ttu-id="2085e-791">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="2085e-791">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2085e-792">1.1</span><span class="sxs-lookup"><span data-stu-id="2085e-792">1.1</span></span>|
|[<span data-ttu-id="2085e-793">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="2085e-793">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2085e-794">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="2085e-794">ReadWriteItem</span></span>|
|[<span data-ttu-id="2085e-795">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="2085e-795">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="2085e-796">Создание</span><span class="sxs-lookup"><span data-stu-id="2085e-796">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="2085e-797">Пример</span><span class="sxs-lookup"><span data-stu-id="2085e-797">Example</span></span>

<span data-ttu-id="2085e-798">В следующем примере существующий элемент Outlook добавляется в виде вложения с именем `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="2085e-798">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

#### <a name="close"></a><span data-ttu-id="2085e-799">close()</span><span class="sxs-lookup"><span data-stu-id="2085e-799">close()</span></span>

<span data-ttu-id="2085e-800">Закрывает текущий создаваемый элемент.</span><span class="sxs-lookup"><span data-stu-id="2085e-800">Closes the current item that is being composed.</span></span>

<span data-ttu-id="2085e-p142">Работа метода `close` зависит от текущего состояния создаваемого элемента. Если элемент содержит несохраненные изменения, клиент предложит пользователю сохранить или отклонить их либо отменить действие закрытия.</span><span class="sxs-lookup"><span data-stu-id="2085e-p142">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="2085e-803">Если в Outlook в Интернете элемент представляет собой встречу, ранее сохраненную с помощью метода `saveAsync`, пользователю предлагается сохранить, отклонить или отменить действие, даже если с момента последнего сохранения элемента не вносились какие-либо изменения.</span><span class="sxs-lookup"><span data-stu-id="2085e-803">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="2085e-804">Если в клиенте Outlook для настольных ПК сообщение представляет собой ответ в тексте, метод `close` не работает.</span><span class="sxs-lookup"><span data-stu-id="2085e-804">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="2085e-805">Требования</span><span class="sxs-lookup"><span data-stu-id="2085e-805">Requirements</span></span>

|<span data-ttu-id="2085e-806">Требование</span><span class="sxs-lookup"><span data-stu-id="2085e-806">Requirement</span></span>|<span data-ttu-id="2085e-807">Значение</span><span class="sxs-lookup"><span data-stu-id="2085e-807">Value</span></span>|
|---|---|
|[<span data-ttu-id="2085e-808">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="2085e-808">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2085e-809">1.3</span><span class="sxs-lookup"><span data-stu-id="2085e-809">1.3</span></span>|
|[<span data-ttu-id="2085e-810">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="2085e-810">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2085e-811">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="2085e-811">Restricted</span></span>|
|[<span data-ttu-id="2085e-812">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="2085e-812">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="2085e-813">Создание</span><span class="sxs-lookup"><span data-stu-id="2085e-813">Compose</span></span>|

---
---

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="2085e-814">displayReplyAllForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="2085e-814">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="2085e-815">Отображает форму ответа, включающую отправителя и всех получателей выбранного сообщения или организатора и всех участников выбранной встречи.</span><span class="sxs-lookup"><span data-stu-id="2085e-815">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="2085e-816">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="2085e-816">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="2085e-817">В Outlook Web App форма ответа отображается в виде всплывающей формы в представлении с 3 либо 1 или 2 колонками.</span><span class="sxs-lookup"><span data-stu-id="2085e-817">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="2085e-818">Если любой строковый параметр превышает указанные для него ограничения, `displayReplyAllForm` возвращает исключение.</span><span class="sxs-lookup"><span data-stu-id="2085e-818">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="2085e-p143">Если в параметре `formData.attachments` указаны вложения, Outlook и Outlook Web App пытаются скачать их и вложить в форму ответа. Если какие-либо вложения добавить не удается, в форме отображается сообщение об ошибке. Если сообщения об ошибках не предусмотрены, то они не отображаются.</span><span class="sxs-lookup"><span data-stu-id="2085e-p143">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2085e-822">Параметры</span><span class="sxs-lookup"><span data-stu-id="2085e-822">Parameters</span></span>

|<span data-ttu-id="2085e-823">Имя</span><span class="sxs-lookup"><span data-stu-id="2085e-823">Name</span></span>|<span data-ttu-id="2085e-824">Тип</span><span class="sxs-lookup"><span data-stu-id="2085e-824">Type</span></span>|<span data-ttu-id="2085e-825">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="2085e-825">Attributes</span></span>|<span data-ttu-id="2085e-826">Описание</span><span class="sxs-lookup"><span data-stu-id="2085e-826">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="2085e-827">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="2085e-827">String &#124; Object</span></span>||<span data-ttu-id="2085e-p144">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="2085e-p144">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="2085e-830">**ИЛИ**</span><span class="sxs-lookup"><span data-stu-id="2085e-830">**OR**</span></span><br/><span data-ttu-id="2085e-p145">Объект, который содержит текст или данные вложения и функцию обратного вызова. Ниже представлено определение этого объекта.</span><span class="sxs-lookup"><span data-stu-id="2085e-p145">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="2085e-833">String</span><span class="sxs-lookup"><span data-stu-id="2085e-833">String</span></span>|<span data-ttu-id="2085e-834">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="2085e-834">&lt;optional&gt;</span></span>|<span data-ttu-id="2085e-p146">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="2085e-p146">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="2085e-837">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="2085e-837">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="2085e-838">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="2085e-838">&lt;optional&gt;</span></span>|<span data-ttu-id="2085e-839">Массив объектов JSON, представляющих собой вложенные файлы или элементы.</span><span class="sxs-lookup"><span data-stu-id="2085e-839">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="2085e-840">String</span><span class="sxs-lookup"><span data-stu-id="2085e-840">String</span></span>||<span data-ttu-id="2085e-p147">Указывает тип вложения. Допустимые значения: `file` для вложенного файла и `item` для вложенного элемента.</span><span class="sxs-lookup"><span data-stu-id="2085e-p147">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="2085e-843">Строка</span><span class="sxs-lookup"><span data-stu-id="2085e-843">String</span></span>||<span data-ttu-id="2085e-844">Строка, содержащая имя вложения, длиной до 255 символов.</span><span class="sxs-lookup"><span data-stu-id="2085e-844">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="2085e-845">Строка</span><span class="sxs-lookup"><span data-stu-id="2085e-845">String</span></span>||<span data-ttu-id="2085e-p148">Используется, только если свойству `type` задано значение `file`. URI расположения файла.</span><span class="sxs-lookup"><span data-stu-id="2085e-p148">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="2085e-848">Логический</span><span class="sxs-lookup"><span data-stu-id="2085e-848">Boolean</span></span>||<span data-ttu-id="2085e-p149">Используется, только если свойству `type` задано значение `file`. Значение `true` указывает на то, что вложение будет встроено в текст сообщения и не должно отображаться в списке вложений.</span><span class="sxs-lookup"><span data-stu-id="2085e-p149">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="2085e-851">String</span><span class="sxs-lookup"><span data-stu-id="2085e-851">String</span></span>||<span data-ttu-id="2085e-p150">Используется, только если свойству `type` задано значение `item`. Идентификатор вложения EWS. Это строка длиной до 100 символов.</span><span class="sxs-lookup"><span data-stu-id="2085e-p150">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="2085e-855">function</span><span class="sxs-lookup"><span data-stu-id="2085e-855">function</span></span>|<span data-ttu-id="2085e-856">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="2085e-856">&lt;optional&gt;</span></span>|<span data-ttu-id="2085e-857">По завершении работы метода функция, переданная параметру `callback`, вызывается с помощью одного параметра `asyncResult`, представляющего собой объект [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="2085e-857">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2085e-858">Требования</span><span class="sxs-lookup"><span data-stu-id="2085e-858">Requirements</span></span>

|<span data-ttu-id="2085e-859">Требование</span><span class="sxs-lookup"><span data-stu-id="2085e-859">Requirement</span></span>|<span data-ttu-id="2085e-860">Значение</span><span class="sxs-lookup"><span data-stu-id="2085e-860">Value</span></span>|
|---|---|
|[<span data-ttu-id="2085e-861">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="2085e-861">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2085e-862">1.0</span><span class="sxs-lookup"><span data-stu-id="2085e-862">1.0</span></span>|
|[<span data-ttu-id="2085e-863">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="2085e-863">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2085e-864">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2085e-864">ReadItem</span></span>|
|[<span data-ttu-id="2085e-865">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="2085e-865">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="2085e-866">Чтение</span><span class="sxs-lookup"><span data-stu-id="2085e-866">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="2085e-867">Примеры</span><span class="sxs-lookup"><span data-stu-id="2085e-867">Examples</span></span>

<span data-ttu-id="2085e-868">Приведенный ниже код передает строку в функцию `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="2085e-868">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="2085e-869">Ответ с пустым текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="2085e-869">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="2085e-870">Ответ только с текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="2085e-870">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="2085e-871">Ответ с текстом сообщения и вложенным файлом.</span><span class="sxs-lookup"><span data-stu-id="2085e-871">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="2085e-872">Ответ с текстом сообщения и вложенным элементом.</span><span class="sxs-lookup"><span data-stu-id="2085e-872">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="2085e-873">Ответ с текстом сообщения, вложенным файлом, вложенным элементом и обратным вызовом.</span><span class="sxs-lookup"><span data-stu-id="2085e-873">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="2085e-874">displayReplyForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="2085e-874">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="2085e-875">Отображает форму ответа, включающую только отправителя выбранного сообщения или организатора выбранной встречи.</span><span class="sxs-lookup"><span data-stu-id="2085e-875">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="2085e-876">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="2085e-876">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="2085e-877">В Outlook Web App форма ответа отображается в виде всплывающей формы в представлении с 3 либо 1 или 2 колонками.</span><span class="sxs-lookup"><span data-stu-id="2085e-877">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="2085e-878">Если любой строковый параметр превышает указанные для него ограничения, `displayReplyForm` возвращает исключение.</span><span class="sxs-lookup"><span data-stu-id="2085e-878">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="2085e-p151">Если в параметре `formData.attachments` указаны вложения, Outlook и Outlook Web App пытаются скачать их и вложить в форму ответа. Если какие-либо вложения добавить не удается, в форме отображается сообщение об ошибке. Если сообщения об ошибках не предусмотрены, то они не отображаются.</span><span class="sxs-lookup"><span data-stu-id="2085e-p151">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2085e-882">Параметры</span><span class="sxs-lookup"><span data-stu-id="2085e-882">Parameters</span></span>

|<span data-ttu-id="2085e-883">Имя</span><span class="sxs-lookup"><span data-stu-id="2085e-883">Name</span></span>|<span data-ttu-id="2085e-884">Тип</span><span class="sxs-lookup"><span data-stu-id="2085e-884">Type</span></span>|<span data-ttu-id="2085e-885">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="2085e-885">Attributes</span></span>|<span data-ttu-id="2085e-886">Описание</span><span class="sxs-lookup"><span data-stu-id="2085e-886">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="2085e-887">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="2085e-887">String &#124; Object</span></span>||<span data-ttu-id="2085e-p152">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="2085e-p152">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="2085e-890">**ИЛИ**</span><span class="sxs-lookup"><span data-stu-id="2085e-890">**OR**</span></span><br/><span data-ttu-id="2085e-p153">Объект, который содержит текст или данные вложения и функцию обратного вызова. Ниже представлено определение этого объекта.</span><span class="sxs-lookup"><span data-stu-id="2085e-p153">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="2085e-893">String</span><span class="sxs-lookup"><span data-stu-id="2085e-893">String</span></span>|<span data-ttu-id="2085e-894">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="2085e-894">&lt;optional&gt;</span></span>|<span data-ttu-id="2085e-p154">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="2085e-p154">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="2085e-897">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="2085e-897">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="2085e-898">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="2085e-898">&lt;optional&gt;</span></span>|<span data-ttu-id="2085e-899">Массив объектов JSON, представляющих собой вложенные файлы или элементы.</span><span class="sxs-lookup"><span data-stu-id="2085e-899">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="2085e-900">String</span><span class="sxs-lookup"><span data-stu-id="2085e-900">String</span></span>||<span data-ttu-id="2085e-p155">Указывает тип вложения. Допустимые значения: `file` для вложенного файла и `item` для вложенного элемента.</span><span class="sxs-lookup"><span data-stu-id="2085e-p155">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="2085e-903">Строка</span><span class="sxs-lookup"><span data-stu-id="2085e-903">String</span></span>||<span data-ttu-id="2085e-904">Строка, содержащая имя вложения, длиной до 255 символов.</span><span class="sxs-lookup"><span data-stu-id="2085e-904">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="2085e-905">Строка</span><span class="sxs-lookup"><span data-stu-id="2085e-905">String</span></span>||<span data-ttu-id="2085e-p156">Используется, только если свойству `type` задано значение `file`. URI расположения файла.</span><span class="sxs-lookup"><span data-stu-id="2085e-p156">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="2085e-908">Логический</span><span class="sxs-lookup"><span data-stu-id="2085e-908">Boolean</span></span>||<span data-ttu-id="2085e-p157">Используется, только если свойству `type` задано значение `file`. Значение `true` указывает на то, что вложение будет встроено в текст сообщения и не должно отображаться в списке вложений.</span><span class="sxs-lookup"><span data-stu-id="2085e-p157">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="2085e-911">String</span><span class="sxs-lookup"><span data-stu-id="2085e-911">String</span></span>||<span data-ttu-id="2085e-p158">Используется, только если свойству `type` задано значение `item`. Идентификатор вложения EWS. Это строка длиной до 100 символов.</span><span class="sxs-lookup"><span data-stu-id="2085e-p158">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="2085e-915">function</span><span class="sxs-lookup"><span data-stu-id="2085e-915">function</span></span>|<span data-ttu-id="2085e-916">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="2085e-916">&lt;optional&gt;</span></span>|<span data-ttu-id="2085e-917">По завершении работы метода функция, переданная параметру `callback`, вызывается с помощью одного параметра `asyncResult`, представляющего собой объект [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="2085e-917">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2085e-918">Требования</span><span class="sxs-lookup"><span data-stu-id="2085e-918">Requirements</span></span>

|<span data-ttu-id="2085e-919">Требование</span><span class="sxs-lookup"><span data-stu-id="2085e-919">Requirement</span></span>|<span data-ttu-id="2085e-920">Значение</span><span class="sxs-lookup"><span data-stu-id="2085e-920">Value</span></span>|
|---|---|
|[<span data-ttu-id="2085e-921">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="2085e-921">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2085e-922">1.0</span><span class="sxs-lookup"><span data-stu-id="2085e-922">1.0</span></span>|
|[<span data-ttu-id="2085e-923">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="2085e-923">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2085e-924">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2085e-924">ReadItem</span></span>|
|[<span data-ttu-id="2085e-925">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="2085e-925">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="2085e-926">Чтение</span><span class="sxs-lookup"><span data-stu-id="2085e-926">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="2085e-927">Примеры</span><span class="sxs-lookup"><span data-stu-id="2085e-927">Examples</span></span>

<span data-ttu-id="2085e-928">Приведенный ниже код передает строку в функцию `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="2085e-928">The following code passes a string to the `displayReplyForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="2085e-929">Ответ с пустым текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="2085e-929">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="2085e-930">Ответ только с текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="2085e-930">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="2085e-931">Ответ с текстом сообщения и вложенным файлом.</span><span class="sxs-lookup"><span data-stu-id="2085e-931">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="2085e-932">Ответ с текстом сообщения и вложенным элементом.</span><span class="sxs-lookup"><span data-stu-id="2085e-932">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="2085e-933">Ответ с текстом сообщения, вложенным файлом, вложенным элементом и обратным вызовом.</span><span class="sxs-lookup"><span data-stu-id="2085e-933">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlook17officeentities"></a><span data-ttu-id="2085e-934">getEntities() → {[Entities](/javascript/api/outlook_1_7/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="2085e-934">getEntities() → {[Entities](/javascript/api/outlook_1_7/office.entities)}</span></span>

<span data-ttu-id="2085e-935">Получает сущности, обнаруженные в теле выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="2085e-935">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="2085e-936">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="2085e-936">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="2085e-937">Требования</span><span class="sxs-lookup"><span data-stu-id="2085e-937">Requirements</span></span>

|<span data-ttu-id="2085e-938">Требование</span><span class="sxs-lookup"><span data-stu-id="2085e-938">Requirement</span></span>|<span data-ttu-id="2085e-939">Значение</span><span class="sxs-lookup"><span data-stu-id="2085e-939">Value</span></span>|
|---|---|
|[<span data-ttu-id="2085e-940">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="2085e-940">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2085e-941">1.0</span><span class="sxs-lookup"><span data-stu-id="2085e-941">1.0</span></span>|
|[<span data-ttu-id="2085e-942">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="2085e-942">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2085e-943">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2085e-943">ReadItem</span></span>|
|[<span data-ttu-id="2085e-944">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="2085e-944">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="2085e-945">Чтение</span><span class="sxs-lookup"><span data-stu-id="2085e-945">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="2085e-946">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="2085e-946">Returns:</span></span>

<span data-ttu-id="2085e-947">Тип: [Entities](/javascript/api/outlook_1_7/office.entities)</span><span class="sxs-lookup"><span data-stu-id="2085e-947">Type: [Entities](/javascript/api/outlook_1_7/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="2085e-948">Пример</span><span class="sxs-lookup"><span data-stu-id="2085e-948">Example</span></span>

<span data-ttu-id="2085e-949">Ниже приведен пример получения доступа к сущностям контактов в теле текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="2085e-949">The following example accesses the contacts entities in the current item's body.</span></span>

```javascript
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

---
---

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook17officecontactmeetingsuggestionjavascriptapioutlook17officemeetingsuggestionphonenumberjavascriptapioutlook17officephonenumbertasksuggestionjavascriptapioutlook17officetasksuggestion"></a><span data-ttu-id="2085e-950">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="2085e-950">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))>}</span></span>

<span data-ttu-id="2085e-951">Получает массив всех сущностей указанного типа, обнаруженных в теле выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="2085e-951">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="2085e-952">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="2085e-952">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2085e-953">Параметры</span><span class="sxs-lookup"><span data-stu-id="2085e-953">Parameters</span></span>

|<span data-ttu-id="2085e-954">Имя</span><span class="sxs-lookup"><span data-stu-id="2085e-954">Name</span></span>|<span data-ttu-id="2085e-955">Тип</span><span class="sxs-lookup"><span data-stu-id="2085e-955">Type</span></span>|<span data-ttu-id="2085e-956">Описание</span><span class="sxs-lookup"><span data-stu-id="2085e-956">Description</span></span>|
|---|---|---|
|`entityType`|[<span data-ttu-id="2085e-957">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="2085e-957">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook_1_7/office.mailboxenums.entitytype)|<span data-ttu-id="2085e-958">Одно из значений перечисления EntityType.</span><span class="sxs-lookup"><span data-stu-id="2085e-958">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2085e-959">Требования</span><span class="sxs-lookup"><span data-stu-id="2085e-959">Requirements</span></span>

|<span data-ttu-id="2085e-960">Требование</span><span class="sxs-lookup"><span data-stu-id="2085e-960">Requirement</span></span>|<span data-ttu-id="2085e-961">Значение</span><span class="sxs-lookup"><span data-stu-id="2085e-961">Value</span></span>|
|---|---|
|[<span data-ttu-id="2085e-962">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="2085e-962">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2085e-963">1.0</span><span class="sxs-lookup"><span data-stu-id="2085e-963">1.0</span></span>|
|[<span data-ttu-id="2085e-964">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="2085e-964">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2085e-965">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="2085e-965">Restricted</span></span>|
|[<span data-ttu-id="2085e-966">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="2085e-966">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="2085e-967">Чтение</span><span class="sxs-lookup"><span data-stu-id="2085e-967">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="2085e-968">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="2085e-968">Returns:</span></span>

<span data-ttu-id="2085e-969">Если значение, переданное в `entityType`, не является допустимым членом перечисления `EntityType`, метод возвращает значение NULL.</span><span class="sxs-lookup"><span data-stu-id="2085e-969">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="2085e-970">Если в теле элемента отсутствуют сущности указанного типа, метод возвращает пустой массив.</span><span class="sxs-lookup"><span data-stu-id="2085e-970">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="2085e-971">В противном случае тип объектов в возвращаемом массиве зависит от типа сущности, запрошенной в параметре `entityType`.</span><span class="sxs-lookup"><span data-stu-id="2085e-971">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="2085e-972">Хотя минимальный уровень разрешений для использования этого метода — **Restricted**, для некоторых типов сущностей требуется доступ на уровне **ReadItem**, как указано в приведенной ниже таблице.</span><span class="sxs-lookup"><span data-stu-id="2085e-972">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

|<span data-ttu-id="2085e-973">Значение параметра `entityType`</span><span class="sxs-lookup"><span data-stu-id="2085e-973">Value of `entityType`</span></span>|<span data-ttu-id="2085e-974">Тип объектов в возвращаемом массиве</span><span class="sxs-lookup"><span data-stu-id="2085e-974">Type of objects in returned array</span></span>|<span data-ttu-id="2085e-975">Необходимый уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="2085e-975">Required Permission Level</span></span>|
|---|---|---|
|`Address`|<span data-ttu-id="2085e-976">String</span><span class="sxs-lookup"><span data-stu-id="2085e-976">String</span></span>|<span data-ttu-id="2085e-977">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="2085e-977">**Restricted**</span></span>|
|`Contact`|<span data-ttu-id="2085e-978">Contact</span><span class="sxs-lookup"><span data-stu-id="2085e-978">Contact</span></span>|<span data-ttu-id="2085e-979">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="2085e-979">**ReadItem**</span></span>|
|`EmailAddress`|<span data-ttu-id="2085e-980">String</span><span class="sxs-lookup"><span data-stu-id="2085e-980">String</span></span>|<span data-ttu-id="2085e-981">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="2085e-981">**ReadItem**</span></span>|
|`MeetingSuggestion`|<span data-ttu-id="2085e-982">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="2085e-982">MeetingSuggestion</span></span>|<span data-ttu-id="2085e-983">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="2085e-983">**ReadItem**</span></span>|
|`PhoneNumber`|<span data-ttu-id="2085e-984">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="2085e-984">PhoneNumber</span></span>|<span data-ttu-id="2085e-985">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="2085e-985">**Restricted**</span></span>|
|`TaskSuggestion`|<span data-ttu-id="2085e-986">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="2085e-986">TaskSuggestion</span></span>|<span data-ttu-id="2085e-987">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="2085e-987">**ReadItem**</span></span>|
|`URL`|<span data-ttu-id="2085e-988">String</span><span class="sxs-lookup"><span data-stu-id="2085e-988">String</span></span>|<span data-ttu-id="2085e-989">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="2085e-989">**Restricted**</span></span>|

<span data-ttu-id="2085e-990">Тип: Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="2085e-990">Type: Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="2085e-991">Пример</span><span class="sxs-lookup"><span data-stu-id="2085e-991">Example</span></span>

<span data-ttu-id="2085e-992">В примере ниже показано, как получить доступ к массиву строк, которые представляют собой почтовые адреса в теле текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="2085e-992">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

---
---

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook17officecontactmeetingsuggestionjavascriptapioutlook17officemeetingsuggestionphonenumberjavascriptapioutlook17officephonenumbertasksuggestionjavascriptapioutlook17officetasksuggestion"></a><span data-ttu-id="2085e-993">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="2085e-993">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))>}</span></span>

<span data-ttu-id="2085e-994">Возвращает известные сущности в выбранном элементе, которые проходят через именованный фильтр, определяемый в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="2085e-994">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="2085e-995">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="2085e-995">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="2085e-996">Метод `getFilteredEntitiesByName` возвращает сущности, соответствующие регулярному выражению, которое определяется в элементе правила [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) в XML-файле манифеста, с использованием указанного значения элемента `FilterName`.</span><span class="sxs-lookup"><span data-stu-id="2085e-996">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2085e-997">Параметры</span><span class="sxs-lookup"><span data-stu-id="2085e-997">Parameters</span></span>

|<span data-ttu-id="2085e-998">Имя</span><span class="sxs-lookup"><span data-stu-id="2085e-998">Name</span></span>|<span data-ttu-id="2085e-999">Тип</span><span class="sxs-lookup"><span data-stu-id="2085e-999">Type</span></span>|<span data-ttu-id="2085e-1000">Описание</span><span class="sxs-lookup"><span data-stu-id="2085e-1000">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="2085e-1001">String</span><span class="sxs-lookup"><span data-stu-id="2085e-1001">String</span></span>|<span data-ttu-id="2085e-1002">Имя элемента правила `ItemHasKnownEntity`, определяющее соответствующий фильтр.</span><span class="sxs-lookup"><span data-stu-id="2085e-1002">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2085e-1003">Требования</span><span class="sxs-lookup"><span data-stu-id="2085e-1003">Requirements</span></span>

|<span data-ttu-id="2085e-1004">Требование</span><span class="sxs-lookup"><span data-stu-id="2085e-1004">Requirement</span></span>|<span data-ttu-id="2085e-1005">Значение</span><span class="sxs-lookup"><span data-stu-id="2085e-1005">Value</span></span>|
|---|---|
|[<span data-ttu-id="2085e-1006">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="2085e-1006">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2085e-1007">1.0</span><span class="sxs-lookup"><span data-stu-id="2085e-1007">1.0</span></span>|
|[<span data-ttu-id="2085e-1008">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="2085e-1008">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2085e-1009">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2085e-1009">ReadItem</span></span>|
|[<span data-ttu-id="2085e-1010">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="2085e-1010">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="2085e-1011">Чтение</span><span class="sxs-lookup"><span data-stu-id="2085e-1011">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="2085e-1012">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="2085e-1012">Returns:</span></span>

<span data-ttu-id="2085e-p160">Если в манифесте нет элемента `ItemHasKnownEntity` со значением `FilterName`, соответствующим параметру `name`, метод возвращает `null`. Если параметр `name` соответствует элементу `ItemHasKnownEntity` в манифесте, но при этом в текущем элементе нет соответствующих сущностей, метод возвращает пустой массив.</span><span class="sxs-lookup"><span data-stu-id="2085e-p160">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="2085e-1015">Тип: Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="2085e-1015">Type: Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))></span></span>

---
---

#### <a name="getregexmatches--object"></a><span data-ttu-id="2085e-1016">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="2085e-1016">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="2085e-1017">Возвращает строковые значения в выбранном элементе, которые соответствуют регулярным выражениям, определенным в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="2085e-1017">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="2085e-1018">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="2085e-1018">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="2085e-p161">Метод `getRegExMatches` возвращает строки, соответствующие регулярному выражению, которое определяется в каждом элементе правила `ItemHasRegularExpressionMatch` или `ItemHasKnownEntity` в XML-файле манифеста. Для правила `ItemHasRegularExpressionMatch` соответствующую строку должно содержать свойство элемента, указанного этим правилом. Простой тип `PropertyName` определяет поддерживаемые свойства.</span><span class="sxs-lookup"><span data-stu-id="2085e-p161">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="2085e-1022">Например, рассмотрим манифест надстройки, который содержит указанный ниже элемент `Rule`.</span><span class="sxs-lookup"><span data-stu-id="2085e-1022">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="2085e-1023">Объект, возвращаемый методом `getRegExMatches`, будет содержать два свойства: `fruits` и `veggies`.</span><span class="sxs-lookup"><span data-stu-id="2085e-1023">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="2085e-p162">Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты. Лучше используйте метод [`Body.getAsync`](/javascript/api/outlook_1_7/office.body#getasync-coerciontype--options--callback-) для этого.</span><span class="sxs-lookup"><span data-stu-id="2085e-p162">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook_1_7/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="2085e-1027">Requirements</span><span class="sxs-lookup"><span data-stu-id="2085e-1027">Requirements</span></span>

|<span data-ttu-id="2085e-1028">Требование</span><span class="sxs-lookup"><span data-stu-id="2085e-1028">Requirement</span></span>|<span data-ttu-id="2085e-1029">Значение</span><span class="sxs-lookup"><span data-stu-id="2085e-1029">Value</span></span>|
|---|---|
|[<span data-ttu-id="2085e-1030">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="2085e-1030">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2085e-1031">1.0</span><span class="sxs-lookup"><span data-stu-id="2085e-1031">1.0</span></span>|
|[<span data-ttu-id="2085e-1032">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="2085e-1032">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2085e-1033">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2085e-1033">ReadItem</span></span>|
|[<span data-ttu-id="2085e-1034">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="2085e-1034">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="2085e-1035">Чтение</span><span class="sxs-lookup"><span data-stu-id="2085e-1035">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="2085e-1036">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="2085e-1036">Returns:</span></span>

<span data-ttu-id="2085e-p163">Объект, содержащий массив строк, которые соответствуют регулярным выражениям, определяемым в XML-файле манифеста. Имя каждого массива равно соответствующему значению атрибута `RegExName` подходящего правила `ItemHasRegularExpressionMatch` или атрибута `FilterName` соответствующего правила `ItemHasKnownEntity`.</span><span class="sxs-lookup"><span data-stu-id="2085e-p163">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="2085e-1039">Тип:</span><span class="sxs-lookup"><span data-stu-id="2085e-1039">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="2085e-1040">Object</span><span class="sxs-lookup"><span data-stu-id="2085e-1040">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="2085e-1041">Пример</span><span class="sxs-lookup"><span data-stu-id="2085e-1041">Example</span></span>

<span data-ttu-id="2085e-1042">В приведенном ниже примере показано, как получить доступ к массиву совпадений с элементами `fruits` и `veggies` правил активации регулярных выражений, указанными в манифесте.</span><span class="sxs-lookup"><span data-stu-id="2085e-1042">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```javascript
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

---
---

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="2085e-1043">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="2085e-1043">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="2085e-1044">Возвращает строковые значения в выбранном элементе, которые соответствуют именованному регулярному выражению, определенному в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="2085e-1044">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="2085e-1045">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="2085e-1045">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="2085e-1046">Метод `getRegExMatchesByName` возвращает строки, соответствующие регулярному выражению, которое определяется в элементе правила `ItemHasRegularExpressionMatch` в XML-файле манифеста, с использованием указанного значения элемента `RegExName`.</span><span class="sxs-lookup"><span data-stu-id="2085e-1046">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="2085e-p164">Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты.</span><span class="sxs-lookup"><span data-stu-id="2085e-p164">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2085e-1049">Параметры</span><span class="sxs-lookup"><span data-stu-id="2085e-1049">Parameters</span></span>

|<span data-ttu-id="2085e-1050">Имя</span><span class="sxs-lookup"><span data-stu-id="2085e-1050">Name</span></span>|<span data-ttu-id="2085e-1051">Тип</span><span class="sxs-lookup"><span data-stu-id="2085e-1051">Type</span></span>|<span data-ttu-id="2085e-1052">Описание</span><span class="sxs-lookup"><span data-stu-id="2085e-1052">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="2085e-1053">String</span><span class="sxs-lookup"><span data-stu-id="2085e-1053">String</span></span>|<span data-ttu-id="2085e-1054">Имя элемента правила `ItemHasRegularExpressionMatch`, определяющее соответствующий фильтр.</span><span class="sxs-lookup"><span data-stu-id="2085e-1054">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2085e-1055">Требования</span><span class="sxs-lookup"><span data-stu-id="2085e-1055">Requirements</span></span>

|<span data-ttu-id="2085e-1056">Требование</span><span class="sxs-lookup"><span data-stu-id="2085e-1056">Requirement</span></span>|<span data-ttu-id="2085e-1057">Значение</span><span class="sxs-lookup"><span data-stu-id="2085e-1057">Value</span></span>|
|---|---|
|[<span data-ttu-id="2085e-1058">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="2085e-1058">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2085e-1059">1.0</span><span class="sxs-lookup"><span data-stu-id="2085e-1059">1.0</span></span>|
|[<span data-ttu-id="2085e-1060">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="2085e-1060">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2085e-1061">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2085e-1061">ReadItem</span></span>|
|[<span data-ttu-id="2085e-1062">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="2085e-1062">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="2085e-1063">Чтение</span><span class="sxs-lookup"><span data-stu-id="2085e-1063">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="2085e-1064">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="2085e-1064">Returns:</span></span>

<span data-ttu-id="2085e-1065">Массив строк, соответствующих регулярному выражению, определяемому в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="2085e-1065">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="2085e-1066">

<dt>Тип</dt>

</span><span class="sxs-lookup"><span data-stu-id="2085e-1066">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="2085e-1067">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="2085e-1067">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="2085e-1068">Пример</span><span class="sxs-lookup"><span data-stu-id="2085e-1068">Example</span></span>

```javascript
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

---
---

#### <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="2085e-1069">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="2085e-1069">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="2085e-1070">Асинхронно возвращает данные, выбранные в теме или тексте сообщения.</span><span class="sxs-lookup"><span data-stu-id="2085e-1070">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="2085e-p165">Если выделенный фрагмент отсутствует, но курсор находится в тексте или теме, метод возвращает значение NULL для выбранных данных. Если выбраны не текст и не тема, метод возвращает ошибку `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="2085e-p165">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2085e-1073">Параметры</span><span class="sxs-lookup"><span data-stu-id="2085e-1073">Parameters</span></span>

|<span data-ttu-id="2085e-1074">Имя</span><span class="sxs-lookup"><span data-stu-id="2085e-1074">Name</span></span>|<span data-ttu-id="2085e-1075">Тип</span><span class="sxs-lookup"><span data-stu-id="2085e-1075">Type</span></span>|<span data-ttu-id="2085e-1076">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="2085e-1076">Attributes</span></span>|<span data-ttu-id="2085e-1077">Описание</span><span class="sxs-lookup"><span data-stu-id="2085e-1077">Description</span></span>|
|---|---|---|---|
|`coercionType`|[<span data-ttu-id="2085e-1078">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="2085e-1078">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="2085e-p166">Запрашивает формат данных. Если задано значение Text, метод возвращает обычный текст как строку, удаляя все имеющиеся HTML-теги. Если задано значение HTML, метод возвращает выделенный текст (обычный текст или HTML).</span><span class="sxs-lookup"><span data-stu-id="2085e-p166">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`|<span data-ttu-id="2085e-1082">Объект</span><span class="sxs-lookup"><span data-stu-id="2085e-1082">Object</span></span>|<span data-ttu-id="2085e-1083">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="2085e-1083">&lt;optional&gt;</span></span>|<span data-ttu-id="2085e-1084">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="2085e-1084">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="2085e-1085">Объект</span><span class="sxs-lookup"><span data-stu-id="2085e-1085">Object</span></span>|<span data-ttu-id="2085e-1086">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="2085e-1086">&lt;optional&gt;</span></span>|<span data-ttu-id="2085e-1087">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="2085e-1087">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="2085e-1088">функция</span><span class="sxs-lookup"><span data-stu-id="2085e-1088">function</span></span>||<span data-ttu-id="2085e-1089">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="2085e-1089">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="2085e-1090">Чтобы получить доступ к выбранным данным из метода обратного вызова, вызовите `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="2085e-1090">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="2085e-1091">Для доступа к исходному свойству, представляющему собой источник выбранных данных, вызовите параметр `asyncResult.value.sourceProperty`, который может иметь значение `body` или `subject`.</span><span class="sxs-lookup"><span data-stu-id="2085e-1091">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2085e-1092">Требования</span><span class="sxs-lookup"><span data-stu-id="2085e-1092">Requirements</span></span>

|<span data-ttu-id="2085e-1093">Требование</span><span class="sxs-lookup"><span data-stu-id="2085e-1093">Requirement</span></span>|<span data-ttu-id="2085e-1094">Значение</span><span class="sxs-lookup"><span data-stu-id="2085e-1094">Value</span></span>|
|---|---|
|[<span data-ttu-id="2085e-1095">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="2085e-1095">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2085e-1096">1.2</span><span class="sxs-lookup"><span data-stu-id="2085e-1096">1.2</span></span>|
|[<span data-ttu-id="2085e-1097">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="2085e-1097">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2085e-1098">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="2085e-1098">ReadWriteItem</span></span>|
|[<span data-ttu-id="2085e-1099">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="2085e-1099">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="2085e-1100">Создание</span><span class="sxs-lookup"><span data-stu-id="2085e-1100">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="2085e-1101">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="2085e-1101">Returns:</span></span>

<span data-ttu-id="2085e-1102">Выбранные данные в виде строки с форматом, определенным в параметре `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="2085e-1102">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="2085e-1103">

<dt>Тип</dt>

</span><span class="sxs-lookup"><span data-stu-id="2085e-1103">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="2085e-1104">String</span><span class="sxs-lookup"><span data-stu-id="2085e-1104">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="2085e-1105">Пример</span><span class="sxs-lookup"><span data-stu-id="2085e-1105">Example</span></span>

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

#### <a name="getselectedentities--entitiesjavascriptapioutlook17officeentities"></a><span data-ttu-id="2085e-1106">getSelectedEntities() → {[Entities](/javascript/api/outlook_1_7/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="2085e-1106">getSelectedEntities() → {[Entities](/javascript/api/outlook_1_7/office.entities)}</span></span>

<span data-ttu-id="2085e-1107">Возвращает сущности, найденные в выделенном совпадении, выбранном пользователем.</span><span class="sxs-lookup"><span data-stu-id="2085e-1107">Gets the entities found in a highlighted match a user has selected.</span></span> <span data-ttu-id="2085e-1108">Выделенные совпадения применяются к [контекстным надстройкам](/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="2085e-1108">Highlighted matches apply to [contextual add-ins](/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="2085e-1109">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="2085e-1109">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="2085e-1110">Требования</span><span class="sxs-lookup"><span data-stu-id="2085e-1110">Requirements</span></span>

|<span data-ttu-id="2085e-1111">Требование</span><span class="sxs-lookup"><span data-stu-id="2085e-1111">Requirement</span></span>|<span data-ttu-id="2085e-1112">Значение</span><span class="sxs-lookup"><span data-stu-id="2085e-1112">Value</span></span>|
|---|---|
|[<span data-ttu-id="2085e-1113">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="2085e-1113">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2085e-1114">1.6</span><span class="sxs-lookup"><span data-stu-id="2085e-1114">1.6</span></span>|
|[<span data-ttu-id="2085e-1115">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="2085e-1115">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2085e-1116">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2085e-1116">ReadItem</span></span>|
|[<span data-ttu-id="2085e-1117">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="2085e-1117">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="2085e-1118">Чтение</span><span class="sxs-lookup"><span data-stu-id="2085e-1118">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="2085e-1119">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="2085e-1119">Returns:</span></span>

<span data-ttu-id="2085e-1120">Тип: [Entities](/javascript/api/outlook_1_7/office.entities)</span><span class="sxs-lookup"><span data-stu-id="2085e-1120">Type: [Entities](/javascript/api/outlook_1_7/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="2085e-1121">Пример</span><span class="sxs-lookup"><span data-stu-id="2085e-1121">Example</span></span>

<span data-ttu-id="2085e-1122">В приведенном ниже примере показано, как получить доступ к сущностям адресов в выделенном совпадении, выбранном пользователем.</span><span class="sxs-lookup"><span data-stu-id="2085e-1122">The following example accesses the addresses entities in the highlighted match selected by the user.</span></span>

```javascript
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

---
---

#### <a name="getselectedregexmatches--object"></a><span data-ttu-id="2085e-1123">getSelectedRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="2085e-1123">getSelectedRegExMatches() → {Object}</span></span>

<span data-ttu-id="2085e-p169">Возвращает строковые значения в выделенном совпадении, которые соответствуют регулярным выражениям, определенным в XML-файле манифеста. Выделенные совпадения применяются к [контекстным надстройкам](/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="2085e-p169">Returns string values in a highlighted match that match the regular expressions defined in the manifest XML file. Highlighted matches apply to [contextual add-ins](/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="2085e-1126">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="2085e-1126">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="2085e-p170">Метод `getSelectedRegExMatches` возвращает строки, соответствующие регулярному выражению, которое определяется в каждом элементе правила `ItemHasRegularExpressionMatch` или `ItemHasKnownEntity` в XML-файле манифеста. Для правила `ItemHasRegularExpressionMatch` соответствующую строку должно содержать свойство элемента, указанного этим правилом. Простой тип `PropertyName` определяет поддерживаемые свойства.</span><span class="sxs-lookup"><span data-stu-id="2085e-p170">The `getSelectedRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="2085e-1130">Например, рассмотрим манифест надстройки, который содержит указанный ниже элемент `Rule`.</span><span class="sxs-lookup"><span data-stu-id="2085e-1130">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="2085e-1131">Объект, возвращаемый методом `getRegExMatches`, будет содержать два свойства: `fruits` и `veggies`.</span><span class="sxs-lookup"><span data-stu-id="2085e-1131">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="2085e-p171">Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты. Лучше используйте метод [`Body.getAsync`](/javascript/api/outlook_1_7/office.body#getasync-coerciontype--options--callback-) для этого.</span><span class="sxs-lookup"><span data-stu-id="2085e-p171">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook_1_7/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="2085e-1135">Requirements</span><span class="sxs-lookup"><span data-stu-id="2085e-1135">Requirements</span></span>

|<span data-ttu-id="2085e-1136">Требование</span><span class="sxs-lookup"><span data-stu-id="2085e-1136">Requirement</span></span>|<span data-ttu-id="2085e-1137">Значение</span><span class="sxs-lookup"><span data-stu-id="2085e-1137">Value</span></span>|
|---|---|
|[<span data-ttu-id="2085e-1138">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="2085e-1138">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2085e-1139">1.6</span><span class="sxs-lookup"><span data-stu-id="2085e-1139">1.6</span></span>|
|[<span data-ttu-id="2085e-1140">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="2085e-1140">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2085e-1141">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2085e-1141">ReadItem</span></span>|
|[<span data-ttu-id="2085e-1142">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="2085e-1142">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="2085e-1143">Чтение</span><span class="sxs-lookup"><span data-stu-id="2085e-1143">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="2085e-1144">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="2085e-1144">Returns:</span></span>

<span data-ttu-id="2085e-p172">Объект, содержащий массив строк, которые соответствуют регулярным выражениям, определяемым в XML-файле манифеста. Имя каждого массива равно соответствующему значению атрибута `RegExName` подходящего правила `ItemHasRegularExpressionMatch` или атрибута `FilterName` соответствующего правила `ItemHasKnownEntity`.</span><span class="sxs-lookup"><span data-stu-id="2085e-p172">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

##### <a name="example"></a><span data-ttu-id="2085e-1147">Пример</span><span class="sxs-lookup"><span data-stu-id="2085e-1147">Example</span></span>

<span data-ttu-id="2085e-1148">В приведенном ниже примере показано, как получить доступ к массиву совпадений с элементами `fruits` и `veggies` правил активации регулярных выражений, указанными в манифесте.</span><span class="sxs-lookup"><span data-stu-id="2085e-1148">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```javascript
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

---
---

#### <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="2085e-1149">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="2085e-1149">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="2085e-1150">Асинхронно загружает настраиваемые свойства для надстройки для выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="2085e-1150">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="2085e-p173">Настраиваемые свойства сохраняются в виде пар "ключ-значение" для каждого приложения и каждого элемента. Этот метод возвращает объект `CustomProperties` при обратном вызове, который предоставляет методы для доступа к настраиваемым свойствам, характерным для текущего элемента и текущей надстройки. Настраиваемые свойства не шифруются для элемента, поэтому этот способ хранения не является безопасным.</span><span class="sxs-lookup"><span data-stu-id="2085e-p173">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2085e-1154">Параметры</span><span class="sxs-lookup"><span data-stu-id="2085e-1154">Parameters</span></span>

|<span data-ttu-id="2085e-1155">Имя</span><span class="sxs-lookup"><span data-stu-id="2085e-1155">Name</span></span>|<span data-ttu-id="2085e-1156">Тип</span><span class="sxs-lookup"><span data-stu-id="2085e-1156">Type</span></span>|<span data-ttu-id="2085e-1157">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="2085e-1157">Attributes</span></span>|<span data-ttu-id="2085e-1158">Описание</span><span class="sxs-lookup"><span data-stu-id="2085e-1158">Description</span></span>|
|---|---|---|---|
|`callback`|<span data-ttu-id="2085e-1159">function</span><span class="sxs-lookup"><span data-stu-id="2085e-1159">function</span></span>||<span data-ttu-id="2085e-1160">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="2085e-1160">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="2085e-1161">Настраиваемые свойства предоставляются в виде объекта [`CustomProperties`](/javascript/api/outlook_1_7/office.customproperties) в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="2085e-1161">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook_1_7/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="2085e-1162">Этот объект позволяет получить, задать и удалить настраиваемые свойства для элемента, а также сохранить изменения, внесенные в набор настраиваемых свойств, на сервере.</span><span class="sxs-lookup"><span data-stu-id="2085e-1162">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`|<span data-ttu-id="2085e-1163">Объект</span><span class="sxs-lookup"><span data-stu-id="2085e-1163">Object</span></span>|<span data-ttu-id="2085e-1164">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="2085e-1164">&lt;optional&gt;</span></span>|<span data-ttu-id="2085e-1165">Разработчики могут указать любой объект, к которому необходимо получить доступ, в функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="2085e-1165">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="2085e-1166">Доступ к этому объекту можно получить с помощью свойства `asyncResult.asyncContext` в функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="2085e-1166">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2085e-1167">Требования</span><span class="sxs-lookup"><span data-stu-id="2085e-1167">Requirements</span></span>

|<span data-ttu-id="2085e-1168">Требование</span><span class="sxs-lookup"><span data-stu-id="2085e-1168">Requirement</span></span>|<span data-ttu-id="2085e-1169">Значение</span><span class="sxs-lookup"><span data-stu-id="2085e-1169">Value</span></span>|
|---|---|
|[<span data-ttu-id="2085e-1170">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="2085e-1170">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2085e-1171">1.0</span><span class="sxs-lookup"><span data-stu-id="2085e-1171">1.0</span></span>|
|[<span data-ttu-id="2085e-1172">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="2085e-1172">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2085e-1173">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2085e-1173">ReadItem</span></span>|
|[<span data-ttu-id="2085e-1174">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="2085e-1174">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="2085e-1175">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="2085e-1175">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2085e-1176">Пример</span><span class="sxs-lookup"><span data-stu-id="2085e-1176">Example</span></span>

<span data-ttu-id="2085e-p176">Приведенный ниже пример кода показывает, как асинхронно загружать настраиваемые свойства, характерные для текущего элемента, с помощью метода `loadCustomPropertiesAsync`. Этот пример также показывает, как сохранять эти свойства на сервере с помощью метода `CustomProperties.saveAsync`. После загрузки настраиваемых свойств в этом примере кода метод `CustomProperties.get` используется для считывания настраиваемого свойства `myProp`, метод `CustomProperties.set` — для записи настраиваемого свойства `otherProp`, а метод `saveAsync` — для сохранения настраиваемых свойств.</span><span class="sxs-lookup"><span data-stu-id="2085e-p176">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

#### <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="2085e-1180">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="2085e-1180">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="2085e-1181">Удаляет вложение из сообщения или встречи.</span><span class="sxs-lookup"><span data-stu-id="2085e-1181">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="2085e-p177">Метод `removeAttachmentAsync` удаляет из элемента вложение с указанным идентификатором. Идентификатор вложения рекомендуется использовать для удаления вложения, только если оно добавлено тем же почтовым приложением в ходе текущего сеанса. В Outlook Web App и Outlook Web App для устройств идентификатор вложения действителен только в рамках одного сеанса. Сеанс завершается, когда пользователь закрывает приложение или начинает создавать элемент во встроенной форме, а затем переходит из формы в отдельное окно.</span><span class="sxs-lookup"><span data-stu-id="2085e-p177">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2085e-1186">Параметры</span><span class="sxs-lookup"><span data-stu-id="2085e-1186">Parameters</span></span>

|<span data-ttu-id="2085e-1187">Имя</span><span class="sxs-lookup"><span data-stu-id="2085e-1187">Name</span></span>|<span data-ttu-id="2085e-1188">Тип</span><span class="sxs-lookup"><span data-stu-id="2085e-1188">Type</span></span>|<span data-ttu-id="2085e-1189">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="2085e-1189">Attributes</span></span>|<span data-ttu-id="2085e-1190">Описание</span><span class="sxs-lookup"><span data-stu-id="2085e-1190">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="2085e-1191">String</span><span class="sxs-lookup"><span data-stu-id="2085e-1191">String</span></span>||<span data-ttu-id="2085e-1192">Идентификатор удаляемого вложения.</span><span class="sxs-lookup"><span data-stu-id="2085e-1192">The identifier of the attachment to remove.</span></span>|
|`options`|<span data-ttu-id="2085e-1193">Object</span><span class="sxs-lookup"><span data-stu-id="2085e-1193">Object</span></span>|<span data-ttu-id="2085e-1194">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="2085e-1194">&lt;optional&gt;</span></span>|<span data-ttu-id="2085e-1195">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="2085e-1195">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="2085e-1196">Объект</span><span class="sxs-lookup"><span data-stu-id="2085e-1196">Object</span></span>|<span data-ttu-id="2085e-1197">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="2085e-1197">&lt;optional&gt;</span></span>|<span data-ttu-id="2085e-1198">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="2085e-1198">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="2085e-1199">функция</span><span class="sxs-lookup"><span data-stu-id="2085e-1199">function</span></span>|<span data-ttu-id="2085e-1200">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="2085e-1200">&lt;optional&gt;</span></span>|<span data-ttu-id="2085e-1201">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="2085e-1201">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="2085e-1202">Если удалить вложение не удается, свойство `asyncResult.error` содержит код ошибки с указанием ее причины.</span><span class="sxs-lookup"><span data-stu-id="2085e-1202">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="2085e-1203">Ошибки</span><span class="sxs-lookup"><span data-stu-id="2085e-1203">Errors</span></span>

|<span data-ttu-id="2085e-1204">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="2085e-1204">Error code</span></span>|<span data-ttu-id="2085e-1205">Описание</span><span class="sxs-lookup"><span data-stu-id="2085e-1205">Description</span></span>|
|------------|-------------|
|`InvalidAttachmentId`|<span data-ttu-id="2085e-1206">Идентификатор вложения не существует.</span><span class="sxs-lookup"><span data-stu-id="2085e-1206">The attachment identifier does not exist.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2085e-1207">Требования</span><span class="sxs-lookup"><span data-stu-id="2085e-1207">Requirements</span></span>

|<span data-ttu-id="2085e-1208">Требование</span><span class="sxs-lookup"><span data-stu-id="2085e-1208">Requirement</span></span>|<span data-ttu-id="2085e-1209">Значение</span><span class="sxs-lookup"><span data-stu-id="2085e-1209">Value</span></span>|
|---|---|
|[<span data-ttu-id="2085e-1210">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="2085e-1210">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2085e-1211">1.1</span><span class="sxs-lookup"><span data-stu-id="2085e-1211">1.1</span></span>|
|[<span data-ttu-id="2085e-1212">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="2085e-1212">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2085e-1213">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="2085e-1213">ReadWriteItem</span></span>|
|[<span data-ttu-id="2085e-1214">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="2085e-1214">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="2085e-1215">Создание</span><span class="sxs-lookup"><span data-stu-id="2085e-1215">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="2085e-1216">Пример</span><span class="sxs-lookup"><span data-stu-id="2085e-1216">Example</span></span>

<span data-ttu-id="2085e-1217">Указанный ниже код удаляет вложение с идентификатором "0".</span><span class="sxs-lookup"><span data-stu-id="2085e-1217">The following code removes an attachment with an identifier of '0'.</span></span>

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

#### <a name="removehandlerasynceventtype-options-callback"></a><span data-ttu-id="2085e-1218">removeHandlerAsync(eventType, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="2085e-1218">removeHandlerAsync(eventType, [options], [callback])</span></span>

<span data-ttu-id="2085e-1219">Удаляет обработчиков для поддерживаемого типа события.</span><span class="sxs-lookup"><span data-stu-id="2085e-1219">Removes the event handlers for a supported event type.</span></span>

<span data-ttu-id="2085e-1220">В настоящее время поддерживаются типы `Office.EventType.AppointmentTimeChanged`событий `Office.EventType.RecipientsChanged`, и`Office.EventType.RecurrenceChanged`</span><span class="sxs-lookup"><span data-stu-id="2085e-1220">Currently the supported event types are `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`</span></span>

##### <a name="parameters"></a><span data-ttu-id="2085e-1221">Параметры</span><span class="sxs-lookup"><span data-stu-id="2085e-1221">Parameters</span></span>

| <span data-ttu-id="2085e-1222">Имя</span><span class="sxs-lookup"><span data-stu-id="2085e-1222">Name</span></span> | <span data-ttu-id="2085e-1223">Тип</span><span class="sxs-lookup"><span data-stu-id="2085e-1223">Type</span></span> | <span data-ttu-id="2085e-1224">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="2085e-1224">Attributes</span></span> | <span data-ttu-id="2085e-1225">Описание</span><span class="sxs-lookup"><span data-stu-id="2085e-1225">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="2085e-1226">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="2085e-1226">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="2085e-1227">Событие, которое должно вызвать обработчик.</span><span class="sxs-lookup"><span data-stu-id="2085e-1227">The event that should invoke the handler.</span></span> |
| `options` | <span data-ttu-id="2085e-1228">Объект</span><span class="sxs-lookup"><span data-stu-id="2085e-1228">Object</span></span> | <span data-ttu-id="2085e-1229">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="2085e-1229">&lt;optional&gt;</span></span> | <span data-ttu-id="2085e-1230">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="2085e-1230">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="2085e-1231">Объект</span><span class="sxs-lookup"><span data-stu-id="2085e-1231">Object</span></span> | <span data-ttu-id="2085e-1232">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="2085e-1232">&lt;optional&gt;</span></span> | <span data-ttu-id="2085e-1233">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="2085e-1233">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="2085e-1234">функция</span><span class="sxs-lookup"><span data-stu-id="2085e-1234">function</span></span>| <span data-ttu-id="2085e-1235">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="2085e-1235">&lt;optional&gt;</span></span>|<span data-ttu-id="2085e-1236">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="2085e-1236">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2085e-1237">Требования</span><span class="sxs-lookup"><span data-stu-id="2085e-1237">Requirements</span></span>

|<span data-ttu-id="2085e-1238">Требование</span><span class="sxs-lookup"><span data-stu-id="2085e-1238">Requirement</span></span>| <span data-ttu-id="2085e-1239">Значение</span><span class="sxs-lookup"><span data-stu-id="2085e-1239">Value</span></span>|
|---|---|
|[<span data-ttu-id="2085e-1240">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="2085e-1240">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2085e-1241">1.7</span><span class="sxs-lookup"><span data-stu-id="2085e-1241">1.7</span></span> |
|[<span data-ttu-id="2085e-1242">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="2085e-1242">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2085e-1243">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2085e-1243">ReadItem</span></span> |
|[<span data-ttu-id="2085e-1244">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="2085e-1244">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2085e-1245">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="2085e-1245">Compose or Read</span></span> |

##### <a name="example"></a><span data-ttu-id="2085e-1246">Пример</span><span class="sxs-lookup"><span data-stu-id="2085e-1246">Example</span></span>

```javascript
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

---
---

#### <a name="saveasyncoptions-callback"></a><span data-ttu-id="2085e-1247">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="2085e-1247">saveAsync([options], callback)</span></span>

<span data-ttu-id="2085e-1248">Асинхронно сохраняет элемент.</span><span class="sxs-lookup"><span data-stu-id="2085e-1248">Asynchronously saves an item.</span></span>

<span data-ttu-id="2085e-p178">При вызове этот метод сохраняет текущее сообщение в виде черновика и возвращает идентификатор элемента с помощью метода обратного вызова. В Outlook Web App или интерактивном режиме Outlook этот элемент сохраняется на сервере. В Outlook в режиме кэширования этот элемент сохраняется в локальном кэше.</span><span class="sxs-lookup"><span data-stu-id="2085e-p178">When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="2085e-1252">Если в режиме создания надстройка вызывает для элемента метод `saveAsync`, чтобы получить параметр `itemId` для использования с EWS или REST API, необходимо помнить, что синхронизация элемента с сервером может занять много времени, если Outlook работает в режиме кэширования данных.</span><span class="sxs-lookup"><span data-stu-id="2085e-1252">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="2085e-1253">До окончания синхронизации элемента применение параметра `itemId` будет приводить к ошибке.</span><span class="sxs-lookup"><span data-stu-id="2085e-1253">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="2085e-p180">Если метод `saveAsync` вызывается для встречи в режиме создания, она сохраняется как обычная встреча в календаре пользователя, а не как черновик. При сохранении новой встречи приглашения не отправляются. При сохранении существующей встречи уведомления отправляются добавленным или удаленным участникам.</span><span class="sxs-lookup"><span data-stu-id="2085e-p180">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="2085e-1257">Следующие клиенты отличаются другим поведением `saveAsync` в отношении встреч в режиме создания:</span><span class="sxs-lookup"><span data-stu-id="2085e-1257">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="2085e-1258">Outlook для Mac не поддерживает сохранение собраний.</span><span class="sxs-lookup"><span data-stu-id="2085e-1258">Outlook for Mac does not support saving a meeting.</span></span> <span data-ttu-id="2085e-1259">`saveAsync` Метод завершается с ошибкой при вызове из собрания в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="2085e-1259">The `saveAsync` method fails when called from a meeting in compose mode.</span></span> <span data-ttu-id="2085e-1260">Просмотреть [не удается сохранить собрание в виде черновика в Outlook для Mac с помощью API Office JS](https://support.microsoft.com/help/4505745) для обхода.</span><span class="sxs-lookup"><span data-stu-id="2085e-1260">See [Cannot save a meeting as a draft in Outlook for Mac by using Office JS API](https://support.microsoft.com/help/4505745) for a workaround.</span></span>
> - <span data-ttu-id="2085e-1261">Outlook в Интернете всегда отправляет приглашение или обновление при вызове `saveAsync` для встречи в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="2085e-1261">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2085e-1262">Параметры</span><span class="sxs-lookup"><span data-stu-id="2085e-1262">Parameters</span></span>

|<span data-ttu-id="2085e-1263">Имя</span><span class="sxs-lookup"><span data-stu-id="2085e-1263">Name</span></span>|<span data-ttu-id="2085e-1264">Тип</span><span class="sxs-lookup"><span data-stu-id="2085e-1264">Type</span></span>|<span data-ttu-id="2085e-1265">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="2085e-1265">Attributes</span></span>|<span data-ttu-id="2085e-1266">Описание</span><span class="sxs-lookup"><span data-stu-id="2085e-1266">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="2085e-1267">Объект</span><span class="sxs-lookup"><span data-stu-id="2085e-1267">Object</span></span>|<span data-ttu-id="2085e-1268">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="2085e-1268">&lt;optional&gt;</span></span>|<span data-ttu-id="2085e-1269">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="2085e-1269">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="2085e-1270">Объект</span><span class="sxs-lookup"><span data-stu-id="2085e-1270">Object</span></span>|<span data-ttu-id="2085e-1271">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="2085e-1271">&lt;optional&gt;</span></span>|<span data-ttu-id="2085e-1272">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="2085e-1272">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="2085e-1273">функция</span><span class="sxs-lookup"><span data-stu-id="2085e-1273">function</span></span>||<span data-ttu-id="2085e-1274">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="2085e-1274">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="2085e-1275">После успешного выполнения идентификатор элемента будет предоставлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="2085e-1275">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2085e-1276">Требования</span><span class="sxs-lookup"><span data-stu-id="2085e-1276">Requirements</span></span>

|<span data-ttu-id="2085e-1277">Требование</span><span class="sxs-lookup"><span data-stu-id="2085e-1277">Requirement</span></span>|<span data-ttu-id="2085e-1278">Значение</span><span class="sxs-lookup"><span data-stu-id="2085e-1278">Value</span></span>|
|---|---|
|[<span data-ttu-id="2085e-1279">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="2085e-1279">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2085e-1280">1.3</span><span class="sxs-lookup"><span data-stu-id="2085e-1280">1.3</span></span>|
|[<span data-ttu-id="2085e-1281">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="2085e-1281">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2085e-1282">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="2085e-1282">ReadWriteItem</span></span>|
|[<span data-ttu-id="2085e-1283">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="2085e-1283">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="2085e-1284">Создание</span><span class="sxs-lookup"><span data-stu-id="2085e-1284">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="2085e-1285">Примеры</span><span class="sxs-lookup"><span data-stu-id="2085e-1285">Examples</span></span>

```javascript
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result.
  });
```

<span data-ttu-id="2085e-p182">Ниже приведен пример параметра `result`, переданного функции обратного вызова. Свойство `value` содержит идентификатор элемента.</span><span class="sxs-lookup"><span data-stu-id="2085e-p182">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

---
---

#### <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="2085e-1288">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="2085e-1288">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="2085e-1289">Асинхронно вставляет данные в текст или тему сообщения.</span><span class="sxs-lookup"><span data-stu-id="2085e-1289">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="2085e-p183">Метод `setSelectedDataAsync` вставляет указанную строку в местоположение курсора в теме или тексте элемента либо, если текст выделен в редакторе, он заменяет выделенный текст. Если курсор находится вне текста или темы элемента, возвращается ошибка. После вставки курсор помещается в конец вставленного содержимого.</span><span class="sxs-lookup"><span data-stu-id="2085e-p183">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2085e-1293">Параметры</span><span class="sxs-lookup"><span data-stu-id="2085e-1293">Parameters</span></span>

|<span data-ttu-id="2085e-1294">Имя</span><span class="sxs-lookup"><span data-stu-id="2085e-1294">Name</span></span>|<span data-ttu-id="2085e-1295">Тип</span><span class="sxs-lookup"><span data-stu-id="2085e-1295">Type</span></span>|<span data-ttu-id="2085e-1296">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="2085e-1296">Attributes</span></span>|<span data-ttu-id="2085e-1297">Описание</span><span class="sxs-lookup"><span data-stu-id="2085e-1297">Description</span></span>|
|---|---|---|---|
|`data`|<span data-ttu-id="2085e-1298">String</span><span class="sxs-lookup"><span data-stu-id="2085e-1298">String</span></span>||<span data-ttu-id="2085e-p184">Вставляемые данные. Объем данных не должен превышать 1 000 000 символов. Если передано больше 1 000 000 символов, возвращается исключение `ArgumentOutOfRange`.</span><span class="sxs-lookup"><span data-stu-id="2085e-p184">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`|<span data-ttu-id="2085e-1302">Object</span><span class="sxs-lookup"><span data-stu-id="2085e-1302">Object</span></span>|<span data-ttu-id="2085e-1303">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="2085e-1303">&lt;optional&gt;</span></span>|<span data-ttu-id="2085e-1304">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="2085e-1304">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="2085e-1305">Объект</span><span class="sxs-lookup"><span data-stu-id="2085e-1305">Object</span></span>|<span data-ttu-id="2085e-1306">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="2085e-1306">&lt;optional&gt;</span></span>|<span data-ttu-id="2085e-1307">В методе обратного вызова разработчики могут указать любой объект, к которому необходимо получить доступ.</span><span class="sxs-lookup"><span data-stu-id="2085e-1307">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="2085e-1308">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="2085e-1308">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="2085e-1309">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="2085e-1309">&lt;optional&gt;</span></span>|<span data-ttu-id="2085e-p185">Если задано значение `text`, текущий стиль применяется в Outlook Web App и Outlook. Если поле представляет собой редактор HTML, вставляются только текстовые данные, даже если они имеют формат HTML.</span><span class="sxs-lookup"><span data-stu-id="2085e-p185">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="2085e-p186">Если задано значение `html` и поле (не тема) поддерживает HTML, в Outlook Web App применяется текущий стиль, а в Outlook — стиль по умолчанию. Если поле является текстовым, возвращается ошибка `InvalidDataFormat`.</span><span class="sxs-lookup"><span data-stu-id="2085e-p186">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="2085e-1314">Если свойство `coercionType` не задано, результат зависит от поля: если поле имеет формат HTML, используется текст в формате HTML, а если поле текстовое, применяется обычный текст.</span><span class="sxs-lookup"><span data-stu-id="2085e-1314">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`|<span data-ttu-id="2085e-1315">функция</span><span class="sxs-lookup"><span data-stu-id="2085e-1315">function</span></span>||<span data-ttu-id="2085e-1316">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="2085e-1316">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2085e-1317">Требования</span><span class="sxs-lookup"><span data-stu-id="2085e-1317">Requirements</span></span>

|<span data-ttu-id="2085e-1318">Требование</span><span class="sxs-lookup"><span data-stu-id="2085e-1318">Requirement</span></span>|<span data-ttu-id="2085e-1319">Значение</span><span class="sxs-lookup"><span data-stu-id="2085e-1319">Value</span></span>|
|---|---|
|[<span data-ttu-id="2085e-1320">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="2085e-1320">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="2085e-1321">1.2</span><span class="sxs-lookup"><span data-stu-id="2085e-1321">1.2</span></span>|
|[<span data-ttu-id="2085e-1322">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="2085e-1322">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="2085e-1323">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="2085e-1323">ReadWriteItem</span></span>|
|[<span data-ttu-id="2085e-1324">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="2085e-1324">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="2085e-1325">Создание</span><span class="sxs-lookup"><span data-stu-id="2085e-1325">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="2085e-1326">Пример</span><span class="sxs-lookup"><span data-stu-id="2085e-1326">Example</span></span>

```javascript
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
