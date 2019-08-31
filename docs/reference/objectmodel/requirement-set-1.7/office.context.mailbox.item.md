---
title: Office. Context. Mailbox. Item — набор требований 1,7
description: ''
ms.date: 08/08/2019
localization_priority: Normal
ms.openlocfilehash: ec990b79ef1f1daaa0bdedc77688a927c1a87e23
ms.sourcegitcommit: 1fb99b1b4e63868a0e81a928c69a34c42bf7e209
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/30/2019
ms.locfileid: "36695975"
---
# <a name="item"></a><span data-ttu-id="df815-102">item</span><span class="sxs-lookup"><span data-stu-id="df815-102">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="df815-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="df815-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="df815-p101">Пространство имен `item` используется для доступа к выбранному в данный момент сообщению, приглашению на собрание или описанию встречи. Вы можете определить тип пространства имен `item` с помощью свойства [itemType](#itemtype-officemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="df815-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="df815-106">Requirements</span><span class="sxs-lookup"><span data-stu-id="df815-106">Requirements</span></span>

|<span data-ttu-id="df815-107">Требование</span><span class="sxs-lookup"><span data-stu-id="df815-107">Requirement</span></span>|<span data-ttu-id="df815-108">Значение</span><span class="sxs-lookup"><span data-stu-id="df815-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="df815-109">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="df815-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="df815-110">1.0</span><span class="sxs-lookup"><span data-stu-id="df815-110">1.0</span></span>|
|[<span data-ttu-id="df815-111">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="df815-111">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="df815-112">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="df815-112">Restricted</span></span>|
|[<span data-ttu-id="df815-113">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="df815-113">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="df815-114">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="df815-114">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="df815-115">Элементы и методы</span><span class="sxs-lookup"><span data-stu-id="df815-115">Members and methods</span></span>

| <span data-ttu-id="df815-116">Элемент</span><span class="sxs-lookup"><span data-stu-id="df815-116">Member</span></span> | <span data-ttu-id="df815-117">Тип</span><span class="sxs-lookup"><span data-stu-id="df815-117">Type</span></span> |
|--------|------|
| [<span data-ttu-id="df815-118">attachments</span><span class="sxs-lookup"><span data-stu-id="df815-118">attachments</span></span>](#attachments-arrayattachmentdetails) | <span data-ttu-id="df815-119">Элемент</span><span class="sxs-lookup"><span data-stu-id="df815-119">Member</span></span> |
| [<span data-ttu-id="df815-120">bcc</span><span class="sxs-lookup"><span data-stu-id="df815-120">bcc</span></span>](#bcc-recipients) | <span data-ttu-id="df815-121">Элемент</span><span class="sxs-lookup"><span data-stu-id="df815-121">Member</span></span> |
| [<span data-ttu-id="df815-122">body</span><span class="sxs-lookup"><span data-stu-id="df815-122">body</span></span>](#body-body) | <span data-ttu-id="df815-123">Элемент</span><span class="sxs-lookup"><span data-stu-id="df815-123">Member</span></span> |
| [<span data-ttu-id="df815-124">cc</span><span class="sxs-lookup"><span data-stu-id="df815-124">cc</span></span>](#cc-arrayemailaddressdetailsrecipients) | <span data-ttu-id="df815-125">Элемент</span><span class="sxs-lookup"><span data-stu-id="df815-125">Member</span></span> |
| [<span data-ttu-id="df815-126">conversationId</span><span class="sxs-lookup"><span data-stu-id="df815-126">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="df815-127">Элемент</span><span class="sxs-lookup"><span data-stu-id="df815-127">Member</span></span> |
| [<span data-ttu-id="df815-128">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="df815-128">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="df815-129">Элемент</span><span class="sxs-lookup"><span data-stu-id="df815-129">Member</span></span> |
| [<span data-ttu-id="df815-130">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="df815-130">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="df815-131">Элемент</span><span class="sxs-lookup"><span data-stu-id="df815-131">Member</span></span> |
| [<span data-ttu-id="df815-132">end</span><span class="sxs-lookup"><span data-stu-id="df815-132">end</span></span>](#end-datetime) | <span data-ttu-id="df815-133">Элемент</span><span class="sxs-lookup"><span data-stu-id="df815-133">Member</span></span> |
| [<span data-ttu-id="df815-134">from</span><span class="sxs-lookup"><span data-stu-id="df815-134">from</span></span>](#from-emailaddressdetailsfrom) | <span data-ttu-id="df815-135">Элемент</span><span class="sxs-lookup"><span data-stu-id="df815-135">Member</span></span> |
| [<span data-ttu-id="df815-136">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="df815-136">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="df815-137">Элемент</span><span class="sxs-lookup"><span data-stu-id="df815-137">Member</span></span> |
| [<span data-ttu-id="df815-138">itemClass</span><span class="sxs-lookup"><span data-stu-id="df815-138">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="df815-139">Элемент</span><span class="sxs-lookup"><span data-stu-id="df815-139">Member</span></span> |
| [<span data-ttu-id="df815-140">itemId</span><span class="sxs-lookup"><span data-stu-id="df815-140">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="df815-141">Элемент</span><span class="sxs-lookup"><span data-stu-id="df815-141">Member</span></span> |
| [<span data-ttu-id="df815-142">itemType</span><span class="sxs-lookup"><span data-stu-id="df815-142">itemType</span></span>](#itemtype-officemailboxenumsitemtype) | <span data-ttu-id="df815-143">Элемент</span><span class="sxs-lookup"><span data-stu-id="df815-143">Member</span></span> |
| [<span data-ttu-id="df815-144">location</span><span class="sxs-lookup"><span data-stu-id="df815-144">location</span></span>](#location-stringlocation) | <span data-ttu-id="df815-145">Элемент</span><span class="sxs-lookup"><span data-stu-id="df815-145">Member</span></span> |
| [<span data-ttu-id="df815-146">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="df815-146">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="df815-147">Элемент</span><span class="sxs-lookup"><span data-stu-id="df815-147">Member</span></span> |
| [<span data-ttu-id="df815-148">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="df815-148">notificationMessages</span></span>](#notificationmessages-notificationmessages) | <span data-ttu-id="df815-149">Элемент</span><span class="sxs-lookup"><span data-stu-id="df815-149">Member</span></span> |
| [<span data-ttu-id="df815-150">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="df815-150">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="df815-151">Элемент</span><span class="sxs-lookup"><span data-stu-id="df815-151">Member</span></span> |
| [<span data-ttu-id="df815-152">organizer</span><span class="sxs-lookup"><span data-stu-id="df815-152">organizer</span></span>](#organizer-emailaddressdetailsorganizer) | <span data-ttu-id="df815-153">Элемент</span><span class="sxs-lookup"><span data-stu-id="df815-153">Member</span></span> |
| [<span data-ttu-id="df815-154">recurrence</span><span class="sxs-lookup"><span data-stu-id="df815-154">recurrence</span></span>](#nullable-recurrence-recurrence) | <span data-ttu-id="df815-155">Member</span><span class="sxs-lookup"><span data-stu-id="df815-155">Member</span></span> |
| [<span data-ttu-id="df815-156">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="df815-156">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="df815-157">Элемент</span><span class="sxs-lookup"><span data-stu-id="df815-157">Member</span></span> |
| [<span data-ttu-id="df815-158">sender</span><span class="sxs-lookup"><span data-stu-id="df815-158">sender</span></span>](#sender-emailaddressdetails) | <span data-ttu-id="df815-159">Элемент</span><span class="sxs-lookup"><span data-stu-id="df815-159">Member</span></span> |
| [<span data-ttu-id="df815-160">seriesId</span><span class="sxs-lookup"><span data-stu-id="df815-160">seriesId</span></span>](#nullable-seriesid-string) | <span data-ttu-id="df815-161">Элемент</span><span class="sxs-lookup"><span data-stu-id="df815-161">Member</span></span> |
| [<span data-ttu-id="df815-162">start</span><span class="sxs-lookup"><span data-stu-id="df815-162">start</span></span>](#start-datetime) | <span data-ttu-id="df815-163">Элемент</span><span class="sxs-lookup"><span data-stu-id="df815-163">Member</span></span> |
| [<span data-ttu-id="df815-164">subject</span><span class="sxs-lookup"><span data-stu-id="df815-164">subject</span></span>](#subject-stringsubject) | <span data-ttu-id="df815-165">Элемент</span><span class="sxs-lookup"><span data-stu-id="df815-165">Member</span></span> |
| [<span data-ttu-id="df815-166">to</span><span class="sxs-lookup"><span data-stu-id="df815-166">to</span></span>](#to-arrayemailaddressdetailsrecipients) | <span data-ttu-id="df815-167">Элемент</span><span class="sxs-lookup"><span data-stu-id="df815-167">Member</span></span> |
| [<span data-ttu-id="df815-168">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="df815-168">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="df815-169">Метод</span><span class="sxs-lookup"><span data-stu-id="df815-169">Method</span></span> |
| [<span data-ttu-id="df815-170">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="df815-170">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="df815-171">Метод</span><span class="sxs-lookup"><span data-stu-id="df815-171">Method</span></span> |
| [<span data-ttu-id="df815-172">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="df815-172">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="df815-173">Метод</span><span class="sxs-lookup"><span data-stu-id="df815-173">Method</span></span> |
| [<span data-ttu-id="df815-174">close</span><span class="sxs-lookup"><span data-stu-id="df815-174">close</span></span>](#close) | <span data-ttu-id="df815-175">Метод</span><span class="sxs-lookup"><span data-stu-id="df815-175">Method</span></span> |
| [<span data-ttu-id="df815-176">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="df815-176">displayReplyAllForm</span></span>](#displayreplyallformformdata-callback) | <span data-ttu-id="df815-177">Метод</span><span class="sxs-lookup"><span data-stu-id="df815-177">Method</span></span> |
| [<span data-ttu-id="df815-178">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="df815-178">displayReplyForm</span></span>](#displayreplyformformdata-callback) | <span data-ttu-id="df815-179">Метод</span><span class="sxs-lookup"><span data-stu-id="df815-179">Method</span></span> |
| [<span data-ttu-id="df815-180">getEntities</span><span class="sxs-lookup"><span data-stu-id="df815-180">getEntities</span></span>](#getentities--entities) | <span data-ttu-id="df815-181">Метод</span><span class="sxs-lookup"><span data-stu-id="df815-181">Method</span></span> |
| [<span data-ttu-id="df815-182">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="df815-182">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="df815-183">Метод</span><span class="sxs-lookup"><span data-stu-id="df815-183">Method</span></span> |
| [<span data-ttu-id="df815-184">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="df815-184">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="df815-185">Метод</span><span class="sxs-lookup"><span data-stu-id="df815-185">Method</span></span> |
| [<span data-ttu-id="df815-186">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="df815-186">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="df815-187">Метод</span><span class="sxs-lookup"><span data-stu-id="df815-187">Method</span></span> |
| [<span data-ttu-id="df815-188">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="df815-188">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="df815-189">Метод</span><span class="sxs-lookup"><span data-stu-id="df815-189">Method</span></span> |
| [<span data-ttu-id="df815-190">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="df815-190">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="df815-191">Метод</span><span class="sxs-lookup"><span data-stu-id="df815-191">Method</span></span> |
| [<span data-ttu-id="df815-192">жетселектедентитиес</span><span class="sxs-lookup"><span data-stu-id="df815-192">getSelectedEntities</span></span>](#getselectedentities--entities) | <span data-ttu-id="df815-193">Метод</span><span class="sxs-lookup"><span data-stu-id="df815-193">Method</span></span> |
| [<span data-ttu-id="df815-194">жетселектедрежексматчес</span><span class="sxs-lookup"><span data-stu-id="df815-194">getSelectedRegExMatches</span></span>](#getselectedregexmatches--object) | <span data-ttu-id="df815-195">Метод</span><span class="sxs-lookup"><span data-stu-id="df815-195">Method</span></span> |
| [<span data-ttu-id="df815-196">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="df815-196">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="df815-197">Метод</span><span class="sxs-lookup"><span data-stu-id="df815-197">Method</span></span> |
| [<span data-ttu-id="df815-198">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="df815-198">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="df815-199">Метод</span><span class="sxs-lookup"><span data-stu-id="df815-199">Method</span></span> |
| [<span data-ttu-id="df815-200">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="df815-200">removeHandlerAsync</span></span>](#removehandlerasynceventtype-options-callback) | <span data-ttu-id="df815-201">Метод</span><span class="sxs-lookup"><span data-stu-id="df815-201">Method</span></span> |
| [<span data-ttu-id="df815-202">saveAsync</span><span class="sxs-lookup"><span data-stu-id="df815-202">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="df815-203">Метод</span><span class="sxs-lookup"><span data-stu-id="df815-203">Method</span></span> |
| [<span data-ttu-id="df815-204">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="df815-204">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="df815-205">Метод</span><span class="sxs-lookup"><span data-stu-id="df815-205">Method</span></span> |

### <a name="example"></a><span data-ttu-id="df815-206">Пример</span><span class="sxs-lookup"><span data-stu-id="df815-206">Example</span></span>

<span data-ttu-id="df815-207">В примере кода JavaScript, приведенном ниже, показано, как получить доступ к свойству `subject` текущего элемента в Outlook.</span><span class="sxs-lookup"><span data-stu-id="df815-207">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="df815-208">Элементы</span><span class="sxs-lookup"><span data-stu-id="df815-208">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetailsviewoutlook-js-17"></a><span data-ttu-id="df815-209">вложения: Array. <[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.7)></span><span class="sxs-lookup"><span data-stu-id="df815-209">attachments: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.7)></span></span>

<span data-ttu-id="df815-p102">Получает массив вложений для элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="df815-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="df815-212">Outlook блокирует определенные типы файлов из-за потенциальных проблем с безопасностью, поэтому они не возвращаются.</span><span class="sxs-lookup"><span data-stu-id="df815-212">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="df815-213">Дополнительные сведения см. в статье [Блокировка вложений в Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="df815-213">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="df815-214">Тип</span><span class="sxs-lookup"><span data-stu-id="df815-214">Type</span></span>

*   <span data-ttu-id="df815-215">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.7)></span><span class="sxs-lookup"><span data-stu-id="df815-215">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.7)></span></span>

##### <a name="requirements"></a><span data-ttu-id="df815-216">Требования</span><span class="sxs-lookup"><span data-stu-id="df815-216">Requirements</span></span>

|<span data-ttu-id="df815-217">Требование</span><span class="sxs-lookup"><span data-stu-id="df815-217">Requirement</span></span>|<span data-ttu-id="df815-218">Значение</span><span class="sxs-lookup"><span data-stu-id="df815-218">Value</span></span>|
|---|---|
|[<span data-ttu-id="df815-219">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="df815-219">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="df815-220">1.0</span><span class="sxs-lookup"><span data-stu-id="df815-220">1.0</span></span>|
|[<span data-ttu-id="df815-221">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="df815-221">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="df815-222">ReadItem</span><span class="sxs-lookup"><span data-stu-id="df815-222">ReadItem</span></span>|
|[<span data-ttu-id="df815-223">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="df815-223">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="df815-224">Чтение</span><span class="sxs-lookup"><span data-stu-id="df815-224">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="df815-225">Пример</span><span class="sxs-lookup"><span data-stu-id="df815-225">Example</span></span>

<span data-ttu-id="df815-226">С помощью приведенного ниже кода можно создать HTML-строку с подробными сведениями обо всех вложениях для текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="df815-226">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

#### <a name="bcc-recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-17"></a><span data-ttu-id="df815-227">СК: [получатели](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="df815-227">bcc: [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span></span>

<span data-ttu-id="df815-228">Получает объект, который предоставляет методы для получения или обновления получателей скрытой копии сообщения.</span><span class="sxs-lookup"><span data-stu-id="df815-228">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="df815-229">Только в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="df815-229">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="df815-230">Тип</span><span class="sxs-lookup"><span data-stu-id="df815-230">Type</span></span>

*   [<span data-ttu-id="df815-231">Получатели</span><span class="sxs-lookup"><span data-stu-id="df815-231">Recipients</span></span>](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)

##### <a name="requirements"></a><span data-ttu-id="df815-232">Требования</span><span class="sxs-lookup"><span data-stu-id="df815-232">Requirements</span></span>

|<span data-ttu-id="df815-233">Требование</span><span class="sxs-lookup"><span data-stu-id="df815-233">Requirement</span></span>|<span data-ttu-id="df815-234">Значение</span><span class="sxs-lookup"><span data-stu-id="df815-234">Value</span></span>|
|---|---|
|[<span data-ttu-id="df815-235">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="df815-235">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="df815-236">1.1</span><span class="sxs-lookup"><span data-stu-id="df815-236">1.1</span></span>|
|[<span data-ttu-id="df815-237">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="df815-237">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="df815-238">ReadItem</span><span class="sxs-lookup"><span data-stu-id="df815-238">ReadItem</span></span>|
|[<span data-ttu-id="df815-239">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="df815-239">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="df815-240">Создание</span><span class="sxs-lookup"><span data-stu-id="df815-240">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="df815-241">Пример</span><span class="sxs-lookup"><span data-stu-id="df815-241">Example</span></span>

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

#### <a name="body-bodyjavascriptapioutlookofficebodyviewoutlook-js-17"></a><span data-ttu-id="df815-242">основной текст: [Body](/javascript/api/outlook/office.body?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="df815-242">body: [Body](/javascript/api/outlook/office.body?view=outlook-js-1.7)</span></span>

<span data-ttu-id="df815-243">Получает объект, предоставляющий методы для работы с основным текстом элемента.</span><span class="sxs-lookup"><span data-stu-id="df815-243">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="df815-244">Тип</span><span class="sxs-lookup"><span data-stu-id="df815-244">Type</span></span>

*   [<span data-ttu-id="df815-245">Body</span><span class="sxs-lookup"><span data-stu-id="df815-245">Body</span></span>](/javascript/api/outlook/office.body?view=outlook-js-1.7)

##### <a name="requirements"></a><span data-ttu-id="df815-246">Требования</span><span class="sxs-lookup"><span data-stu-id="df815-246">Requirements</span></span>

|<span data-ttu-id="df815-247">Требование</span><span class="sxs-lookup"><span data-stu-id="df815-247">Requirement</span></span>|<span data-ttu-id="df815-248">Значение</span><span class="sxs-lookup"><span data-stu-id="df815-248">Value</span></span>|
|---|---|
|[<span data-ttu-id="df815-249">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="df815-249">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="df815-250">1.1</span><span class="sxs-lookup"><span data-stu-id="df815-250">1.1</span></span>|
|[<span data-ttu-id="df815-251">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="df815-251">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="df815-252">ReadItem</span><span class="sxs-lookup"><span data-stu-id="df815-252">ReadItem</span></span>|
|[<span data-ttu-id="df815-253">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="df815-253">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="df815-254">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="df815-254">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="df815-255">Пример</span><span class="sxs-lookup"><span data-stu-id="df815-255">Example</span></span>

<span data-ttu-id="df815-256">В этом примере возвращается текст сообщения в формате обычного текста.</span><span class="sxs-lookup"><span data-stu-id="df815-256">This example gets the body of the message in plain text.</span></span>

```js
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="df815-257">Ниже приведен пример итогового параметра, переданного функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="df815-257">The following is an example of the result parameter passed to the callback function.</span></span>

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

#### <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-17recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-17"></a><span data-ttu-id="df815-258">CC: Array. <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)>|[получатели](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="df815-258">cc: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span></span>

<span data-ttu-id="df815-259">Предоставляет доступ к получателям копии сообщения.</span><span class="sxs-lookup"><span data-stu-id="df815-259">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="df815-260">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="df815-260">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="df815-261">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="df815-261">Read mode</span></span>

<span data-ttu-id="df815-p106">Свойство `cc` возвращает массив, который содержит объект `EmailAddressDetails` для каждого получателя, указанного в строке **Копия** сообщения. Коллекция может включать не более 100 элементов.</span><span class="sxs-lookup"><span data-stu-id="df815-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="df815-264">Режим создания</span><span class="sxs-lookup"><span data-stu-id="df815-264">Compose mode</span></span>

<span data-ttu-id="df815-265">Свойство `cc` возвращает объект `Recipients`, предоставляющий методы для получения или обновления получателей, которые указаны в строке **Копия** сообщения.</span><span class="sxs-lookup"><span data-stu-id="df815-265">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

```js
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="df815-266">Тип</span><span class="sxs-lookup"><span data-stu-id="df815-266">Type</span></span>

*   <span data-ttu-id="df815-267">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="df815-267">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span></span>

##### <a name="requirements"></a><span data-ttu-id="df815-268">Требования</span><span class="sxs-lookup"><span data-stu-id="df815-268">Requirements</span></span>

|<span data-ttu-id="df815-269">Требование</span><span class="sxs-lookup"><span data-stu-id="df815-269">Requirement</span></span>|<span data-ttu-id="df815-270">Значение</span><span class="sxs-lookup"><span data-stu-id="df815-270">Value</span></span>|
|---|---|
|[<span data-ttu-id="df815-271">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="df815-271">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="df815-272">1.0</span><span class="sxs-lookup"><span data-stu-id="df815-272">1.0</span></span>|
|[<span data-ttu-id="df815-273">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="df815-273">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="df815-274">ReadItem</span><span class="sxs-lookup"><span data-stu-id="df815-274">ReadItem</span></span>|
|[<span data-ttu-id="df815-275">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="df815-275">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="df815-276">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="df815-276">Compose or Read</span></span>|

<br>

---
---

#### <a name="nullable-conversationid-string"></a><span data-ttu-id="df815-277">(Nullable) conversationId: строка</span><span class="sxs-lookup"><span data-stu-id="df815-277">(nullable) conversationId: String</span></span>

<span data-ttu-id="df815-278">Получает идентификатор разговора по электронной почте, содержащего конкретное сообщение.</span><span class="sxs-lookup"><span data-stu-id="df815-278">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="df815-p107">Вы можете получить целочисленное значение этого свойства, если ваше почтовое приложение активируется в формах просмотра или формах создания ответов. Если пользователь изменит тему ответа, после его отправки идентификатор беседы будет изменен, и полученное ранее значение будет недействительным.</span><span class="sxs-lookup"><span data-stu-id="df815-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="df815-p108">Это свойство имеет значение NULL для нового элемента в форме создания. Свойство `conversationId` вернет значение, если пользователь задаст тему и сохранит элемент.</span><span class="sxs-lookup"><span data-stu-id="df815-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="df815-283">Тип</span><span class="sxs-lookup"><span data-stu-id="df815-283">Type</span></span>

*   <span data-ttu-id="df815-284">String</span><span class="sxs-lookup"><span data-stu-id="df815-284">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="df815-285">Требования</span><span class="sxs-lookup"><span data-stu-id="df815-285">Requirements</span></span>

|<span data-ttu-id="df815-286">Требование</span><span class="sxs-lookup"><span data-stu-id="df815-286">Requirement</span></span>|<span data-ttu-id="df815-287">Значение</span><span class="sxs-lookup"><span data-stu-id="df815-287">Value</span></span>|
|---|---|
|[<span data-ttu-id="df815-288">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="df815-288">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="df815-289">1.0</span><span class="sxs-lookup"><span data-stu-id="df815-289">1.0</span></span>|
|[<span data-ttu-id="df815-290">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="df815-290">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="df815-291">ReadItem</span><span class="sxs-lookup"><span data-stu-id="df815-291">ReadItem</span></span>|
|[<span data-ttu-id="df815-292">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="df815-292">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="df815-293">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="df815-293">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="df815-294">Пример</span><span class="sxs-lookup"><span data-stu-id="df815-294">Example</span></span>

```js
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

<br>

---
---

#### <a name="datetimecreated-date"></a><span data-ttu-id="df815-295">dateTimeCreated: Дата</span><span class="sxs-lookup"><span data-stu-id="df815-295">dateTimeCreated: Date</span></span>

<span data-ttu-id="df815-p109">Получает дату и время создания элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="df815-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="df815-298">Тип</span><span class="sxs-lookup"><span data-stu-id="df815-298">Type</span></span>

*   <span data-ttu-id="df815-299">Дата</span><span class="sxs-lookup"><span data-stu-id="df815-299">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="df815-300">Требования</span><span class="sxs-lookup"><span data-stu-id="df815-300">Requirements</span></span>

|<span data-ttu-id="df815-301">Требование</span><span class="sxs-lookup"><span data-stu-id="df815-301">Requirement</span></span>|<span data-ttu-id="df815-302">Значение</span><span class="sxs-lookup"><span data-stu-id="df815-302">Value</span></span>|
|---|---|
|[<span data-ttu-id="df815-303">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="df815-303">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="df815-304">1.0</span><span class="sxs-lookup"><span data-stu-id="df815-304">1.0</span></span>|
|[<span data-ttu-id="df815-305">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="df815-305">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="df815-306">ReadItem</span><span class="sxs-lookup"><span data-stu-id="df815-306">ReadItem</span></span>|
|[<span data-ttu-id="df815-307">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="df815-307">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="df815-308">Чтение</span><span class="sxs-lookup"><span data-stu-id="df815-308">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="df815-309">Пример</span><span class="sxs-lookup"><span data-stu-id="df815-309">Example</span></span>

```js
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

<br>

---
---

#### <a name="datetimemodified-date"></a><span data-ttu-id="df815-310">dateTimeModified: Дата</span><span class="sxs-lookup"><span data-stu-id="df815-310">dateTimeModified: Date</span></span>

<span data-ttu-id="df815-311">Получает дату и время последнего изменения элемента.</span><span class="sxs-lookup"><span data-stu-id="df815-311">Gets the date and time that an item was last modified.</span></span> <span data-ttu-id="df815-312">Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="df815-312">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="df815-313">Этот элемент не поддерживается в Outlook на iOS или Android.</span><span class="sxs-lookup"><span data-stu-id="df815-313">This member is not supported in Outlook on iOS or Android.</span></span>

##### <a name="type"></a><span data-ttu-id="df815-314">Тип</span><span class="sxs-lookup"><span data-stu-id="df815-314">Type</span></span>

*   <span data-ttu-id="df815-315">Дата</span><span class="sxs-lookup"><span data-stu-id="df815-315">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="df815-316">Требования</span><span class="sxs-lookup"><span data-stu-id="df815-316">Requirements</span></span>

|<span data-ttu-id="df815-317">Требование</span><span class="sxs-lookup"><span data-stu-id="df815-317">Requirement</span></span>|<span data-ttu-id="df815-318">Значение</span><span class="sxs-lookup"><span data-stu-id="df815-318">Value</span></span>|
|---|---|
|[<span data-ttu-id="df815-319">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="df815-319">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="df815-320">1.0</span><span class="sxs-lookup"><span data-stu-id="df815-320">1.0</span></span>|
|[<span data-ttu-id="df815-321">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="df815-321">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="df815-322">ReadItem</span><span class="sxs-lookup"><span data-stu-id="df815-322">ReadItem</span></span>|
|[<span data-ttu-id="df815-323">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="df815-323">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="df815-324">Чтение</span><span class="sxs-lookup"><span data-stu-id="df815-324">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="df815-325">Пример</span><span class="sxs-lookup"><span data-stu-id="df815-325">Example</span></span>

```js
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

<br>

---
---

#### <a name="end-datetimejavascriptapioutlookofficetimeviewoutlook-js-17"></a><span data-ttu-id="df815-326">конец: Дата | [Time (время](/javascript/api/outlook/office.time?view=outlook-js-1.7) )</span><span class="sxs-lookup"><span data-stu-id="df815-326">end: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.7)</span></span>

<span data-ttu-id="df815-327">Получает или задает дату и время окончания встречи.</span><span class="sxs-lookup"><span data-stu-id="df815-327">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="df815-p111">Свойство `end` представлено в виде значения даты и времени в формате UTC. Преобразовать значение свойства end в местные значения даты и времени клиента можно с помощью метода [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime).</span><span class="sxs-lookup"><span data-stu-id="df815-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="df815-330">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="df815-330">Read mode</span></span>

<span data-ttu-id="df815-331">Свойство `end` возвращает объект `Date`.</span><span class="sxs-lookup"><span data-stu-id="df815-331">The `end` property returns a `Date` object.</span></span>

```js
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="df815-332">Режим создания</span><span class="sxs-lookup"><span data-stu-id="df815-332">Compose mode</span></span>

<span data-ttu-id="df815-333">Свойство `end` возвращает объект `Time`.</span><span class="sxs-lookup"><span data-stu-id="df815-333">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="df815-334">Если вы задаете время окончания с помощью метода [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.7#setasync-datetime--options--callback-), необходимо использовать метод [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) для преобразования местного времени на клиенте в формат UTC для сервера.</span><span class="sxs-lookup"><span data-stu-id="df815-334">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.7#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="df815-335">В примере ниже показано, как с помощью метода [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.7#setasync-datetime--options--callback-) объекта `Time` задать время окончания встречи.</span><span class="sxs-lookup"><span data-stu-id="df815-335">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.7#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="df815-336">Тип</span><span class="sxs-lookup"><span data-stu-id="df815-336">Type</span></span>

*   <span data-ttu-id="df815-337">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="df815-337">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.7)</span></span>

##### <a name="requirements"></a><span data-ttu-id="df815-338">Требования</span><span class="sxs-lookup"><span data-stu-id="df815-338">Requirements</span></span>

|<span data-ttu-id="df815-339">Требование</span><span class="sxs-lookup"><span data-stu-id="df815-339">Requirement</span></span>|<span data-ttu-id="df815-340">Значение</span><span class="sxs-lookup"><span data-stu-id="df815-340">Value</span></span>|
|---|---|
|[<span data-ttu-id="df815-341">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="df815-341">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="df815-342">1.0</span><span class="sxs-lookup"><span data-stu-id="df815-342">1.0</span></span>|
|[<span data-ttu-id="df815-343">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="df815-343">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="df815-344">ReadItem</span><span class="sxs-lookup"><span data-stu-id="df815-344">ReadItem</span></span>|
|[<span data-ttu-id="df815-345">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="df815-345">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="df815-346">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="df815-346">Compose or Read</span></span>|

<br>

---
---

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-17fromjavascriptapioutlookofficefromviewoutlook-js-17"></a><span data-ttu-id="df815-347">от: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)|[](/javascript/api/outlook/office.from?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="df815-347">from: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)|[From](/javascript/api/outlook/office.from?view=outlook-js-1.7)</span></span>

<span data-ttu-id="df815-348">Получает электронный адрес отправителя сообщения.</span><span class="sxs-lookup"><span data-stu-id="df815-348">Gets the email address of the sender of a message.</span></span>

<span data-ttu-id="df815-p112">Свойства `from` и [`sender`](#sender-emailaddressdetails) представляют одно лицо, если сообщение не отправлено представителем. В противном случае свойство `from` представляет лицо, делегировавшее полномочия, а свойство sender — представителя.</span><span class="sxs-lookup"><span data-stu-id="df815-p112">The `from` and [`sender`](#sender-emailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="df815-351">Свойству `recipientType`, принадлежащему объекту `EmailAddressDetails` в свойстве `from`, задано значение `undefined`.</span><span class="sxs-lookup"><span data-stu-id="df815-351">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="df815-352">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="df815-352">Read mode</span></span>

<span data-ttu-id="df815-353">`from` Свойство возвращает `EmailAddressDetails` объект.</span><span class="sxs-lookup"><span data-stu-id="df815-353">The `from` property returns an `EmailAddressDetails` object.</span></span>

```js
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

##### <a name="compose-mode"></a><span data-ttu-id="df815-354">Режим создания</span><span class="sxs-lookup"><span data-stu-id="df815-354">Compose mode</span></span>

<span data-ttu-id="df815-355">`from` Свойство возвращает `From` объект, который предоставляет метод для получения значения From.</span><span class="sxs-lookup"><span data-stu-id="df815-355">The `from` property returns a `From` object that provides a method to get the from value.</span></span>

```js
Office.context.mailbox.item.from.getAsync(callback);

function callback(asyncResult) {
  var from = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="df815-356">Тип</span><span class="sxs-lookup"><span data-stu-id="df815-356">Type</span></span>

*   <span data-ttu-id="df815-357">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7) | [из](/javascript/api/outlook/office.from?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="df815-357">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7) | [From](/javascript/api/outlook/office.from?view=outlook-js-1.7)</span></span>

##### <a name="requirements"></a><span data-ttu-id="df815-358">Требования</span><span class="sxs-lookup"><span data-stu-id="df815-358">Requirements</span></span>

|<span data-ttu-id="df815-359">Требование</span><span class="sxs-lookup"><span data-stu-id="df815-359">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="df815-360">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="df815-360">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="df815-361">1.0</span><span class="sxs-lookup"><span data-stu-id="df815-361">1.0</span></span>|<span data-ttu-id="df815-362">1.7</span><span class="sxs-lookup"><span data-stu-id="df815-362">1.7</span></span>|
|[<span data-ttu-id="df815-363">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="df815-363">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="df815-364">ReadItem</span><span class="sxs-lookup"><span data-stu-id="df815-364">ReadItem</span></span>|<span data-ttu-id="df815-365">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="df815-365">ReadWriteItem</span></span>|
|[<span data-ttu-id="df815-366">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="df815-366">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="df815-367">Чтение</span><span class="sxs-lookup"><span data-stu-id="df815-367">Read</span></span>|<span data-ttu-id="df815-368">Создание</span><span class="sxs-lookup"><span data-stu-id="df815-368">Compose</span></span>|

<br>

---
---

#### <a name="internetmessageid-string"></a><span data-ttu-id="df815-369">internetMessageId: строка</span><span class="sxs-lookup"><span data-stu-id="df815-369">internetMessageId: String</span></span>

<span data-ttu-id="df815-p113">Получает идентификатор интернет-сообщения для электронного сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="df815-p113">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="df815-372">Тип</span><span class="sxs-lookup"><span data-stu-id="df815-372">Type</span></span>

*   <span data-ttu-id="df815-373">String</span><span class="sxs-lookup"><span data-stu-id="df815-373">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="df815-374">Требования</span><span class="sxs-lookup"><span data-stu-id="df815-374">Requirements</span></span>

|<span data-ttu-id="df815-375">Требование</span><span class="sxs-lookup"><span data-stu-id="df815-375">Requirement</span></span>|<span data-ttu-id="df815-376">Значение</span><span class="sxs-lookup"><span data-stu-id="df815-376">Value</span></span>|
|---|---|
|[<span data-ttu-id="df815-377">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="df815-377">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="df815-378">1.0</span><span class="sxs-lookup"><span data-stu-id="df815-378">1.0</span></span>|
|[<span data-ttu-id="df815-379">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="df815-379">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="df815-380">ReadItem</span><span class="sxs-lookup"><span data-stu-id="df815-380">ReadItem</span></span>|
|[<span data-ttu-id="df815-381">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="df815-381">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="df815-382">Чтение</span><span class="sxs-lookup"><span data-stu-id="df815-382">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="df815-383">Пример</span><span class="sxs-lookup"><span data-stu-id="df815-383">Example</span></span>

```js
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

<br>

---
---

#### <a name="itemclass-string"></a><span data-ttu-id="df815-384">itemClass: строка</span><span class="sxs-lookup"><span data-stu-id="df815-384">itemClass: String</span></span>

<span data-ttu-id="df815-p114">Получает класс элемента веб-служб Exchange для выбранного элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="df815-p114">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="df815-p115">Свойство `itemClass` указывает класс сообщения выбранного элемента. Ниже приводятся классы сообщения по умолчанию для элемента сообщения или встречи.</span><span class="sxs-lookup"><span data-stu-id="df815-p115">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

|<span data-ttu-id="df815-389">Тип</span><span class="sxs-lookup"><span data-stu-id="df815-389">Type</span></span>|<span data-ttu-id="df815-390">Описание</span><span class="sxs-lookup"><span data-stu-id="df815-390">Description</span></span>|<span data-ttu-id="df815-391">Класс элемента</span><span class="sxs-lookup"><span data-stu-id="df815-391">item class</span></span>|
|---|---|---|
|<span data-ttu-id="df815-392">Элементы встречи</span><span class="sxs-lookup"><span data-stu-id="df815-392">Appointment items</span></span>|<span data-ttu-id="df815-393">Это элементы календаря для класса элемента `IPM.Appointment` или `IPM.Appointment.Occurrence`.</span><span class="sxs-lookup"><span data-stu-id="df815-393">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span>|`IPM.Appointment`<br />`IPM.Appointment.Occurrence`|
|<span data-ttu-id="df815-394">Элементы сообщения</span><span class="sxs-lookup"><span data-stu-id="df815-394">Message items</span></span>|<span data-ttu-id="df815-395">Сюда входят электронные сообщения, для которых по умолчанию задан класс сообщения `IPM.Note`, а также приглашения на собрания, ответы на них и уведомления об их отмене, использующие `IPM.Schedule.Meeting` в качестве базового класса сообщения.</span><span class="sxs-lookup"><span data-stu-id="df815-395">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span>|`IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled`|

<span data-ttu-id="df815-396">Можно создавать настраиваемые классы сообщения, расширяющие классы сообщения по умолчанию, например настраиваемый класс сообщения о встрече `IPM.Appointment.Contoso`.</span><span class="sxs-lookup"><span data-stu-id="df815-396">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="df815-397">Тип</span><span class="sxs-lookup"><span data-stu-id="df815-397">Type</span></span>

*   <span data-ttu-id="df815-398">String</span><span class="sxs-lookup"><span data-stu-id="df815-398">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="df815-399">Требования</span><span class="sxs-lookup"><span data-stu-id="df815-399">Requirements</span></span>

|<span data-ttu-id="df815-400">Требование</span><span class="sxs-lookup"><span data-stu-id="df815-400">Requirement</span></span>|<span data-ttu-id="df815-401">Значение</span><span class="sxs-lookup"><span data-stu-id="df815-401">Value</span></span>|
|---|---|
|[<span data-ttu-id="df815-402">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="df815-402">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="df815-403">1.0</span><span class="sxs-lookup"><span data-stu-id="df815-403">1.0</span></span>|
|[<span data-ttu-id="df815-404">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="df815-404">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="df815-405">ReadItem</span><span class="sxs-lookup"><span data-stu-id="df815-405">ReadItem</span></span>|
|[<span data-ttu-id="df815-406">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="df815-406">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="df815-407">Чтение</span><span class="sxs-lookup"><span data-stu-id="df815-407">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="df815-408">Пример</span><span class="sxs-lookup"><span data-stu-id="df815-408">Example</span></span>

```js
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

<br>

---
---

#### <a name="nullable-itemid-string"></a><span data-ttu-id="df815-409">(Nullable) itemId: строка</span><span class="sxs-lookup"><span data-stu-id="df815-409">(nullable) itemId: String</span></span>

<span data-ttu-id="df815-p116">Получает идентификатор элемента веб-служб Exchange для текущего элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="df815-p116">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="df815-412">Идентификатор, возвращаемый свойством `itemId`, совпадает с идентификатором элемента веб-служб Exchange.</span><span class="sxs-lookup"><span data-stu-id="df815-412">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="df815-413">Свойство `itemId` не совпадает с идентификатором записи Outlook, а также идентификатором, который используется REST API Outlook.</span><span class="sxs-lookup"><span data-stu-id="df815-413">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="df815-414">Прежде чем совершать вызовы REST API, используя это значение, его необходимо преобразовать с помощью [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="df815-414">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="df815-415">Дополнительные сведения см. в статье [Использование REST API Outlook из надстройки Outlook](/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="df815-415">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="df815-p118">Свойство `itemId` недоступно в режиме создания. Если требуется идентификатор элемента, с помощью метода [`saveAsync`](#saveasyncoptions-callback) можно сохранить элемент в хранилище. При этом в параметре [`AsyncResult.value`](/javascript/api/office/office.asyncresult) функции обратного вызова возвращается идентификатор элемента.</span><span class="sxs-lookup"><span data-stu-id="df815-p118">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="df815-418">Тип</span><span class="sxs-lookup"><span data-stu-id="df815-418">Type</span></span>

*   <span data-ttu-id="df815-419">String</span><span class="sxs-lookup"><span data-stu-id="df815-419">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="df815-420">Требования</span><span class="sxs-lookup"><span data-stu-id="df815-420">Requirements</span></span>

|<span data-ttu-id="df815-421">Требование</span><span class="sxs-lookup"><span data-stu-id="df815-421">Requirement</span></span>|<span data-ttu-id="df815-422">Значение</span><span class="sxs-lookup"><span data-stu-id="df815-422">Value</span></span>|
|---|---|
|[<span data-ttu-id="df815-423">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="df815-423">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="df815-424">1.0</span><span class="sxs-lookup"><span data-stu-id="df815-424">1.0</span></span>|
|[<span data-ttu-id="df815-425">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="df815-425">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="df815-426">ReadItem</span><span class="sxs-lookup"><span data-stu-id="df815-426">ReadItem</span></span>|
|[<span data-ttu-id="df815-427">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="df815-427">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="df815-428">Чтение</span><span class="sxs-lookup"><span data-stu-id="df815-428">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="df815-429">Пример</span><span class="sxs-lookup"><span data-stu-id="df815-429">Example</span></span>

<span data-ttu-id="df815-p119">Указанный ниже код проверяет наличие идентификатора элемента. Если свойство `itemId` возвращает значение `null` или `undefined`, элемент будет сохранен в хранилище, а из асинхронного результата будет получен идентификатор элемента.</span><span class="sxs-lookup"><span data-stu-id="df815-p119">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

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

#### <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtypeviewoutlook-js-17"></a><span data-ttu-id="df815-432">itemType: [Office. MailboxEnums. ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="df815-432">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.7)</span></span>

<span data-ttu-id="df815-433">Получает тип элемента, который представляет экземпляр.</span><span class="sxs-lookup"><span data-stu-id="df815-433">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="df815-434">Свойство `itemType` возвращает одно из значений перечисления `ItemType`, которое указывает, является ли экземпляр объекта `item` сообщением или встречей.</span><span class="sxs-lookup"><span data-stu-id="df815-434">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="df815-435">Тип</span><span class="sxs-lookup"><span data-stu-id="df815-435">Type</span></span>

*   [<span data-ttu-id="df815-436">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="df815-436">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.7)

##### <a name="requirements"></a><span data-ttu-id="df815-437">Требования</span><span class="sxs-lookup"><span data-stu-id="df815-437">Requirements</span></span>

|<span data-ttu-id="df815-438">Требование</span><span class="sxs-lookup"><span data-stu-id="df815-438">Requirement</span></span>|<span data-ttu-id="df815-439">Значение</span><span class="sxs-lookup"><span data-stu-id="df815-439">Value</span></span>|
|---|---|
|[<span data-ttu-id="df815-440">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="df815-440">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="df815-441">1.0</span><span class="sxs-lookup"><span data-stu-id="df815-441">1.0</span></span>|
|[<span data-ttu-id="df815-442">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="df815-442">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="df815-443">ReadItem</span><span class="sxs-lookup"><span data-stu-id="df815-443">ReadItem</span></span>|
|[<span data-ttu-id="df815-444">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="df815-444">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="df815-445">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="df815-445">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="df815-446">Пример</span><span class="sxs-lookup"><span data-stu-id="df815-446">Example</span></span>

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

#### <a name="location-stringlocationjavascriptapioutlookofficelocationviewoutlook-js-17"></a><span data-ttu-id="df815-447">Местоположение: строка | [Location (расположение](/javascript/api/outlook/office.location?view=outlook-js-1.7) )</span><span class="sxs-lookup"><span data-stu-id="df815-447">location: String|[Location](/javascript/api/outlook/office.location?view=outlook-js-1.7)</span></span>

<span data-ttu-id="df815-448">Получает или задает место встречи.</span><span class="sxs-lookup"><span data-stu-id="df815-448">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="df815-449">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="df815-449">Read mode</span></span>

<span data-ttu-id="df815-450">Свойство `location` возвращает строку, содержащую сведения о месте встречи.</span><span class="sxs-lookup"><span data-stu-id="df815-450">The `location` property returns a string that contains the location of the appointment.</span></span>

```js
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="df815-451">Режим создания</span><span class="sxs-lookup"><span data-stu-id="df815-451">Compose mode</span></span>

<span data-ttu-id="df815-452">Свойство `location` возвращает объект `Location`, предоставляющий методы, которые используются для получения и задания места встречи.</span><span class="sxs-lookup"><span data-stu-id="df815-452">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```js
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="df815-453">Тип</span><span class="sxs-lookup"><span data-stu-id="df815-453">Type</span></span>

*   <span data-ttu-id="df815-454">String | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="df815-454">String | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.7)</span></span>

##### <a name="requirements"></a><span data-ttu-id="df815-455">Требования</span><span class="sxs-lookup"><span data-stu-id="df815-455">Requirements</span></span>

|<span data-ttu-id="df815-456">Требование</span><span class="sxs-lookup"><span data-stu-id="df815-456">Requirement</span></span>|<span data-ttu-id="df815-457">Значение</span><span class="sxs-lookup"><span data-stu-id="df815-457">Value</span></span>|
|---|---|
|[<span data-ttu-id="df815-458">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="df815-458">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="df815-459">1.0</span><span class="sxs-lookup"><span data-stu-id="df815-459">1.0</span></span>|
|[<span data-ttu-id="df815-460">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="df815-460">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="df815-461">ReadItem</span><span class="sxs-lookup"><span data-stu-id="df815-461">ReadItem</span></span>|
|[<span data-ttu-id="df815-462">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="df815-462">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="df815-463">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="df815-463">Compose or Read</span></span>|

<br>

---
---

#### <a name="normalizedsubject-string"></a><span data-ttu-id="df815-464">normalizedSubject: строка</span><span class="sxs-lookup"><span data-stu-id="df815-464">normalizedSubject: String</span></span>

<span data-ttu-id="df815-p120">Получает тему элемента со всеми удаленными префиксами (включая `RE:` и `FWD:`). Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="df815-p120">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="df815-p121">Свойство normalizedSubject получает тему элемента со стандартными префиксами (такими как `RE:` и `FW:`), добавляемыми почтовыми программами. Для получения темы элемента с неизмененными префиксами используйте свойство [`subject`](#subject-stringsubject).</span><span class="sxs-lookup"><span data-stu-id="df815-p121">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="df815-469">Тип</span><span class="sxs-lookup"><span data-stu-id="df815-469">Type</span></span>

*   <span data-ttu-id="df815-470">String</span><span class="sxs-lookup"><span data-stu-id="df815-470">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="df815-471">Требования</span><span class="sxs-lookup"><span data-stu-id="df815-471">Requirements</span></span>

|<span data-ttu-id="df815-472">Требование</span><span class="sxs-lookup"><span data-stu-id="df815-472">Requirement</span></span>|<span data-ttu-id="df815-473">Значение</span><span class="sxs-lookup"><span data-stu-id="df815-473">Value</span></span>|
|---|---|
|[<span data-ttu-id="df815-474">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="df815-474">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="df815-475">1.0</span><span class="sxs-lookup"><span data-stu-id="df815-475">1.0</span></span>|
|[<span data-ttu-id="df815-476">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="df815-476">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="df815-477">ReadItem</span><span class="sxs-lookup"><span data-stu-id="df815-477">ReadItem</span></span>|
|[<span data-ttu-id="df815-478">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="df815-478">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="df815-479">Чтение</span><span class="sxs-lookup"><span data-stu-id="df815-479">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="df815-480">Пример</span><span class="sxs-lookup"><span data-stu-id="df815-480">Example</span></span>

```js
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

<br>

---
---

#### <a name="notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessagesviewoutlook-js-17"></a><span data-ttu-id="df815-481">notificationMessages: [notificationMessages](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="df815-481">notificationMessages: [NotificationMessages](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.7)</span></span>

<span data-ttu-id="df815-482">Получает сообщения уведомления для элемента.</span><span class="sxs-lookup"><span data-stu-id="df815-482">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="df815-483">Тип</span><span class="sxs-lookup"><span data-stu-id="df815-483">Type</span></span>

*   [<span data-ttu-id="df815-484">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="df815-484">NotificationMessages</span></span>](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.7)

##### <a name="requirements"></a><span data-ttu-id="df815-485">Требования</span><span class="sxs-lookup"><span data-stu-id="df815-485">Requirements</span></span>

|<span data-ttu-id="df815-486">Требование</span><span class="sxs-lookup"><span data-stu-id="df815-486">Requirement</span></span>|<span data-ttu-id="df815-487">Значение</span><span class="sxs-lookup"><span data-stu-id="df815-487">Value</span></span>|
|---|---|
|[<span data-ttu-id="df815-488">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="df815-488">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="df815-489">1.3</span><span class="sxs-lookup"><span data-stu-id="df815-489">1.3</span></span>|
|[<span data-ttu-id="df815-490">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="df815-490">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="df815-491">ReadItem</span><span class="sxs-lookup"><span data-stu-id="df815-491">ReadItem</span></span>|
|[<span data-ttu-id="df815-492">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="df815-492">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="df815-493">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="df815-493">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="df815-494">Пример</span><span class="sxs-lookup"><span data-stu-id="df815-494">Example</span></span>

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

#### <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-17recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-17"></a><span data-ttu-id="df815-495">optionalAttendees: Array. <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)>|[получатели](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="df815-495">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span></span>

<span data-ttu-id="df815-496">Предоставляет доступ к необязательным участникам события.</span><span class="sxs-lookup"><span data-stu-id="df815-496">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="df815-497">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="df815-497">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="df815-498">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="df815-498">Read mode</span></span>

<span data-ttu-id="df815-499">Свойство `optionalAttendees` возвращает массив, содержащий объект `EmailAddressDetails` для каждого необязательного участника собрания.</span><span class="sxs-lookup"><span data-stu-id="df815-499">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

```js
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="df815-500">Режим создания</span><span class="sxs-lookup"><span data-stu-id="df815-500">Compose mode</span></span>

<span data-ttu-id="df815-501">Свойство `optionalAttendees` возвращает объект `Recipients`, который предоставляет методы для получения или обновления необязательных участников собрания.</span><span class="sxs-lookup"><span data-stu-id="df815-501">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

```js
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="df815-502">Тип</span><span class="sxs-lookup"><span data-stu-id="df815-502">Type</span></span>

*   <span data-ttu-id="df815-503">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="df815-503">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span></span>

##### <a name="requirements"></a><span data-ttu-id="df815-504">Требования</span><span class="sxs-lookup"><span data-stu-id="df815-504">Requirements</span></span>

|<span data-ttu-id="df815-505">Требование</span><span class="sxs-lookup"><span data-stu-id="df815-505">Requirement</span></span>|<span data-ttu-id="df815-506">Значение</span><span class="sxs-lookup"><span data-stu-id="df815-506">Value</span></span>|
|---|---|
|[<span data-ttu-id="df815-507">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="df815-507">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="df815-508">1.0</span><span class="sxs-lookup"><span data-stu-id="df815-508">1.0</span></span>|
|[<span data-ttu-id="df815-509">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="df815-509">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="df815-510">ReadItem</span><span class="sxs-lookup"><span data-stu-id="df815-510">ReadItem</span></span>|
|[<span data-ttu-id="df815-511">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="df815-511">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="df815-512">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="df815-512">Compose or Read</span></span>|

<br>

---
---

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-17organizerjavascriptapioutlookofficeorganizerviewoutlook-js-17"></a><span data-ttu-id="df815-513">Организатор: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)|[Организатор](/javascript/api/outlook/office.organizer?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="df815-513">organizer: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)|[Organizer](/javascript/api/outlook/office.organizer?view=outlook-js-1.7)</span></span>

<span data-ttu-id="df815-514">Получает адрес электронной почты организатора для указанного собрания.</span><span class="sxs-lookup"><span data-stu-id="df815-514">Gets the email address of the organizer for a specified meeting.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="df815-515">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="df815-515">Read mode</span></span>

<span data-ttu-id="df815-516">`organizer` Свойство возвращает объект [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7) , представляющий организатора собрания.</span><span class="sxs-lookup"><span data-stu-id="df815-516">The `organizer` property returns an [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7) object that represents the meeting organizer.</span></span>

```js
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

##### <a name="compose-mode"></a><span data-ttu-id="df815-517">Режим создания</span><span class="sxs-lookup"><span data-stu-id="df815-517">Compose mode</span></span>

<span data-ttu-id="df815-518">Свойство возвращает объект организатора, который предоставляет метод для получения значения организатора. [](/javascript/api/outlook/office.organizer?view=outlook-js-1.7) `organizer`</span><span class="sxs-lookup"><span data-stu-id="df815-518">The `organizer` property returns an [Organizer](/javascript/api/outlook/office.organizer?view=outlook-js-1.7) object that provides a method to get the organizer value.</span></span>

```js
Office.context.mailbox.item.organizer.getAsync(
  function(asyncResult) {
    console.log(JSON.stringify(asyncResult));
  }
);
```

##### <a name="type"></a><span data-ttu-id="df815-519">Тип</span><span class="sxs-lookup"><span data-stu-id="df815-519">Type</span></span>

*   <span data-ttu-id="df815-520">[](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7) | [Организатор](/javascript/api/outlook/office.organizer?view=outlook-js-1.7) EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="df815-520">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7) | [Organizer](/javascript/api/outlook/office.organizer?view=outlook-js-1.7)</span></span>

##### <a name="requirements"></a><span data-ttu-id="df815-521">Требования</span><span class="sxs-lookup"><span data-stu-id="df815-521">Requirements</span></span>

|<span data-ttu-id="df815-522">Требование</span><span class="sxs-lookup"><span data-stu-id="df815-522">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="df815-523">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="df815-523">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="df815-524">1.0</span><span class="sxs-lookup"><span data-stu-id="df815-524">1.0</span></span>|<span data-ttu-id="df815-525">1.7</span><span class="sxs-lookup"><span data-stu-id="df815-525">1.7</span></span>|
|[<span data-ttu-id="df815-526">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="df815-526">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="df815-527">ReadItem</span><span class="sxs-lookup"><span data-stu-id="df815-527">ReadItem</span></span>|<span data-ttu-id="df815-528">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="df815-528">ReadWriteItem</span></span>|
|[<span data-ttu-id="df815-529">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="df815-529">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="df815-530">Чтение</span><span class="sxs-lookup"><span data-stu-id="df815-530">Read</span></span>|<span data-ttu-id="df815-531">Создание</span><span class="sxs-lookup"><span data-stu-id="df815-531">Compose</span></span>|

<br>

---
---

#### <a name="nullable-recurrence-recurrencejavascriptapioutlookofficerecurrenceviewoutlook-js-17"></a><span data-ttu-id="df815-532">(Nullable) повторение [](/javascript/api/outlook/office.recurrence?view=outlook-js-1.7) : повторение</span><span class="sxs-lookup"><span data-stu-id="df815-532">(nullable) recurrence: [Recurrence](/javascript/api/outlook/office.recurrence?view=outlook-js-1.7)</span></span>

<span data-ttu-id="df815-533">Получает или задает шаблон повторения встречи.</span><span class="sxs-lookup"><span data-stu-id="df815-533">Gets or sets the recurrence pattern of an appointment.</span></span> <span data-ttu-id="df815-534">Получает шаблон повторения приглашения на собрание.</span><span class="sxs-lookup"><span data-stu-id="df815-534">Gets the recurrence pattern of a meeting request.</span></span> <span data-ttu-id="df815-535">Режимы чтения и создания для элементов встречи.</span><span class="sxs-lookup"><span data-stu-id="df815-535">Read and compose modes for appointment items.</span></span> <span data-ttu-id="df815-536">Режим чтения для элементов приглашения на собрания.</span><span class="sxs-lookup"><span data-stu-id="df815-536">Read mode for meeting request items.</span></span>

<span data-ttu-id="df815-537">`recurrence` Свойство возвращает объект [повторения](/javascript/api/outlook/office.recurrence?view=outlook-js-1.7) для повторяющихся встреч или приглашений на собрания, если элемент представляет собой серию или экземпляр в ряду.</span><span class="sxs-lookup"><span data-stu-id="df815-537">The `recurrence` property returns a [recurrence](/javascript/api/outlook/office.recurrence?view=outlook-js-1.7) object for recurring appointments or meetings requests if an item is a series or an instance in a series.</span></span> <span data-ttu-id="df815-538">`null`возвращается для отдельных встреч и приглашений на собрание для отдельных встреч.</span><span class="sxs-lookup"><span data-stu-id="df815-538">`null` is returned for single appointments and meeting requests of single appointments.</span></span> <span data-ttu-id="df815-539">`undefined`возвращается для сообщений, которые не являются приглашениями на собрания.</span><span class="sxs-lookup"><span data-stu-id="df815-539">`undefined` is returned for messages that are not meeting requests.</span></span>

> <span data-ttu-id="df815-540">Note: приглашения на `itemClass` собрания имеют значение IPM. Schedule. Meeting. Request.</span><span class="sxs-lookup"><span data-stu-id="df815-540">Note: Meeting requests have an `itemClass` value of IPM.Schedule.Meeting.Request.</span></span>

> <span data-ttu-id="df815-541">Note: при наличии объекта `null`повторения это указывает на то, что объект является одной встречей или приглашением на собрание одной встречи, а не частью ряда.</span><span class="sxs-lookup"><span data-stu-id="df815-541">Note: If the recurrence object is `null`, this indicates that the object is a single appointment or a meeting request of a single appointment and NOT a part of a series.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="df815-542">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="df815-542">Read mode</span></span>

<span data-ttu-id="df815-543">`recurrence` Свойство возвращает объект [повторения](/javascript/api/outlook/office.recurrence?view=outlook-js-1.7) , представляющий повторение встречи.</span><span class="sxs-lookup"><span data-stu-id="df815-543">The `recurrence` property returns a [Recurrence](/javascript/api/outlook/office.recurrence?view=outlook-js-1.7) object that represents the appointment recurrence.</span></span> <span data-ttu-id="df815-544">Оно доступно для встреч и приглашений на собрания.</span><span class="sxs-lookup"><span data-stu-id="df815-544">This is available for appointments and meeting requests.</span></span>

```js
var recurrence = Office.context.mailbox.item.recurrence;
console.log("Recurrence: " + JSON.stringify(recurrence));
```

##### <a name="compose-mode"></a><span data-ttu-id="df815-545">Режим создания</span><span class="sxs-lookup"><span data-stu-id="df815-545">Compose mode</span></span>

<span data-ttu-id="df815-546">`recurrence` Свойство возвращает объект [повторения](/javascript/api/outlook/office.recurrence?view=outlook-js-1.7) , который предоставляет методы для управления повторением встречи.</span><span class="sxs-lookup"><span data-stu-id="df815-546">The `recurrence` property returns a [Recurrence](/javascript/api/outlook/office.recurrence?view=outlook-js-1.7) object that provides methods to manage the appointment recurrence.</span></span> <span data-ttu-id="df815-547">Оно доступно для встреч.</span><span class="sxs-lookup"><span data-stu-id="df815-547">This is available for appointments.</span></span>

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

##### <a name="type"></a><span data-ttu-id="df815-548">Тип</span><span class="sxs-lookup"><span data-stu-id="df815-548">Type</span></span>

* [<span data-ttu-id="df815-549">Повторения</span><span class="sxs-lookup"><span data-stu-id="df815-549">Recurrence</span></span>](/javascript/api/outlook/office.recurrence?view=outlook-js-1.7)

|<span data-ttu-id="df815-550">Требование</span><span class="sxs-lookup"><span data-stu-id="df815-550">Requirement</span></span>|<span data-ttu-id="df815-551">Значение</span><span class="sxs-lookup"><span data-stu-id="df815-551">Value</span></span>|
|---|---|
|[<span data-ttu-id="df815-552">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="df815-552">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="df815-553">1.7</span><span class="sxs-lookup"><span data-stu-id="df815-553">1.7</span></span>|
|[<span data-ttu-id="df815-554">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="df815-554">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="df815-555">ReadItem</span><span class="sxs-lookup"><span data-stu-id="df815-555">ReadItem</span></span>|
|[<span data-ttu-id="df815-556">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="df815-556">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="df815-557">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="df815-557">Compose or Read</span></span>|

<br>

---
---

#### <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-17recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-17"></a><span data-ttu-id="df815-558">requiredAttendees: Array. <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)>|[получатели](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="df815-558">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span></span>

<span data-ttu-id="df815-559">Предоставляет доступ к обязательным участникам события.</span><span class="sxs-lookup"><span data-stu-id="df815-559">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="df815-560">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="df815-560">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="df815-561">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="df815-561">Read mode</span></span>

<span data-ttu-id="df815-562">Свойство `requiredAttendees` возвращает массив, содержащий объект `EmailAddressDetails` для каждого обязательного участника собрания.</span><span class="sxs-lookup"><span data-stu-id="df815-562">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

```js
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="df815-563">Режим создания</span><span class="sxs-lookup"><span data-stu-id="df815-563">Compose mode</span></span>

<span data-ttu-id="df815-564">Свойство `requiredAttendees` возвращает объект `Recipients`, предоставляющий методы, с помощью которых можно получить или обновить сведения об обязательных участниках собрания.</span><span class="sxs-lookup"><span data-stu-id="df815-564">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

```js
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="df815-565">Тип</span><span class="sxs-lookup"><span data-stu-id="df815-565">Type</span></span>

*   <span data-ttu-id="df815-566">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="df815-566">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span></span>

##### <a name="requirements"></a><span data-ttu-id="df815-567">Требования</span><span class="sxs-lookup"><span data-stu-id="df815-567">Requirements</span></span>

|<span data-ttu-id="df815-568">Требование</span><span class="sxs-lookup"><span data-stu-id="df815-568">Requirement</span></span>|<span data-ttu-id="df815-569">Значение</span><span class="sxs-lookup"><span data-stu-id="df815-569">Value</span></span>|
|---|---|
|[<span data-ttu-id="df815-570">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="df815-570">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="df815-571">1.0</span><span class="sxs-lookup"><span data-stu-id="df815-571">1.0</span></span>|
|[<span data-ttu-id="df815-572">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="df815-572">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="df815-573">ReadItem</span><span class="sxs-lookup"><span data-stu-id="df815-573">ReadItem</span></span>|
|[<span data-ttu-id="df815-574">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="df815-574">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="df815-575">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="df815-575">Compose or Read</span></span>|

<br>

---
---

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-17"></a><span data-ttu-id="df815-576">Отправитель: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="df815-576">sender: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)</span></span>

<span data-ttu-id="df815-p128">Получает электронный адрес отправителя электронного сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="df815-p128">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="df815-p129">Свойства [`from`](#from-emailaddressdetailsfrom) и `sender` представляют одно лицо, если сообщение не отправлено представителем. В противном случае свойство `from` представляет лицо, делегировавшее полномочия, а свойство sender — представителя.</span><span class="sxs-lookup"><span data-stu-id="df815-p129">The [`from`](#from-emailaddressdetailsfrom) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="df815-581">Свойству `recipientType`, принадлежащему объекту `EmailAddressDetails` в свойстве `sender`, задано значение `undefined`.</span><span class="sxs-lookup"><span data-stu-id="df815-581">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="df815-582">Тип</span><span class="sxs-lookup"><span data-stu-id="df815-582">Type</span></span>

*   [<span data-ttu-id="df815-583">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="df815-583">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)

##### <a name="requirements"></a><span data-ttu-id="df815-584">Требования</span><span class="sxs-lookup"><span data-stu-id="df815-584">Requirements</span></span>

|<span data-ttu-id="df815-585">Требование</span><span class="sxs-lookup"><span data-stu-id="df815-585">Requirement</span></span>|<span data-ttu-id="df815-586">Значение</span><span class="sxs-lookup"><span data-stu-id="df815-586">Value</span></span>|
|---|---|
|[<span data-ttu-id="df815-587">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="df815-587">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="df815-588">1.0</span><span class="sxs-lookup"><span data-stu-id="df815-588">1.0</span></span>|
|[<span data-ttu-id="df815-589">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="df815-589">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="df815-590">ReadItem</span><span class="sxs-lookup"><span data-stu-id="df815-590">ReadItem</span></span>|
|[<span data-ttu-id="df815-591">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="df815-591">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="df815-592">Чтение</span><span class="sxs-lookup"><span data-stu-id="df815-592">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="df815-593">Пример</span><span class="sxs-lookup"><span data-stu-id="df815-593">Example</span></span>

```js
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

<br>

---
---

#### <a name="nullable-seriesid-string"></a><span data-ttu-id="df815-594">(Nullable) seriesId: строка</span><span class="sxs-lookup"><span data-stu-id="df815-594">(nullable) seriesId: String</span></span>

<span data-ttu-id="df815-595">Получает идентификатор ряда, к которому принадлежит экземпляр.</span><span class="sxs-lookup"><span data-stu-id="df815-595">Gets the id of the series that an instance belongs to.</span></span>

<span data-ttu-id="df815-596">В Outlook в Интернете и на настольных клиентах `seriesId` возвращается идентификатор веб-служб Exchange (EWS) родительского элемента (ряда), к которому принадлежит этот элемент.</span><span class="sxs-lookup"><span data-stu-id="df815-596">In Outlook on the web and desktop clients, the `seriesId` returns the Exchange Web Services (EWS) ID of the parent (series) item that this item belongs to.</span></span> <span data-ttu-id="df815-597">Однако в iOS и Android `seriesId` возвращается идентификатор REST родительского элемента.</span><span class="sxs-lookup"><span data-stu-id="df815-597">However, in iOS and Android, the `seriesId` returns the REST ID of the parent item.</span></span>

> [!NOTE]
> <span data-ttu-id="df815-598">Идентификатор, возвращаемый свойством `seriesId`, совпадает с идентификатором элемента веб-служб Exchange.</span><span class="sxs-lookup"><span data-stu-id="df815-598">The identifier returned by the `seriesId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="df815-599">`seriesId` Свойство не совпадает с идентификаторами Outlook, используемыми в REST API Outlook.</span><span class="sxs-lookup"><span data-stu-id="df815-599">The `seriesId` property is not identical to the Outlook IDs used by the Outlook REST API.</span></span> <span data-ttu-id="df815-600">Прежде чем совершать вызовы REST API, используя это значение, его необходимо преобразовать с помощью [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="df815-600">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="df815-601">Дополнительные сведения см. в статье [Использование REST API Outlook из надстройки Outlook](/outlook/add-ins/use-rest-api).</span><span class="sxs-lookup"><span data-stu-id="df815-601">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api).</span></span>

<span data-ttu-id="df815-602">`seriesId` Свойство возвращает `null` элементы, у которых нет родительских элементов, таких как одиночные встречи, элементы ряда или приглашения на собрание, `undefined` и возвращаемые для других элементов, не являющиеся приглашениями на собрания.</span><span class="sxs-lookup"><span data-stu-id="df815-602">The `seriesId` property returns `null` for items that do not have parent items such as single appointments, series items, or meeting requests and returns `undefined` for any other items that are not meeting requests.</span></span>

##### <a name="type"></a><span data-ttu-id="df815-603">Тип</span><span class="sxs-lookup"><span data-stu-id="df815-603">Type</span></span>

* <span data-ttu-id="df815-604">String</span><span class="sxs-lookup"><span data-stu-id="df815-604">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="df815-605">Требования</span><span class="sxs-lookup"><span data-stu-id="df815-605">Requirements</span></span>

|<span data-ttu-id="df815-606">Требование</span><span class="sxs-lookup"><span data-stu-id="df815-606">Requirement</span></span>|<span data-ttu-id="df815-607">Значение</span><span class="sxs-lookup"><span data-stu-id="df815-607">Value</span></span>|
|---|---|
|[<span data-ttu-id="df815-608">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="df815-608">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="df815-609">1.7</span><span class="sxs-lookup"><span data-stu-id="df815-609">1.7</span></span>|
|[<span data-ttu-id="df815-610">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="df815-610">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="df815-611">ReadItem</span><span class="sxs-lookup"><span data-stu-id="df815-611">ReadItem</span></span>|
|[<span data-ttu-id="df815-612">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="df815-612">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="df815-613">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="df815-613">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="df815-614">Пример</span><span class="sxs-lookup"><span data-stu-id="df815-614">Example</span></span>

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

#### <a name="start-datetimejavascriptapioutlookofficetimeviewoutlook-js-17"></a><span data-ttu-id="df815-615">Начало: Дата | [Time (время](/javascript/api/outlook/office.time?view=outlook-js-1.7) )</span><span class="sxs-lookup"><span data-stu-id="df815-615">start: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.7)</span></span>

<span data-ttu-id="df815-616">Получает или задает дату и время начала встречи.</span><span class="sxs-lookup"><span data-stu-id="df815-616">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="df815-p132">Свойство `start` представлено в виде значения даты и времени в формате UTC. Это значение можно преобразовать в местные значения даты и времени клиента с помощью метода [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime).</span><span class="sxs-lookup"><span data-stu-id="df815-p132">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="df815-619">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="df815-619">Read mode</span></span>

<span data-ttu-id="df815-620">Свойство `start` возвращает объект `Date`.</span><span class="sxs-lookup"><span data-stu-id="df815-620">The `start` property returns a `Date` object.</span></span>

```js
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="df815-621">Режим создания</span><span class="sxs-lookup"><span data-stu-id="df815-621">Compose mode</span></span>

<span data-ttu-id="df815-622">Свойство `start` возвращает объект `Time`.</span><span class="sxs-lookup"><span data-stu-id="df815-622">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="df815-623">Если вы задаете время начала с помощью метода [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.7#setasync-datetime--options--callback-), необходимо использовать метод [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) для преобразования местного времени на клиенте в формат UTC для сервера.</span><span class="sxs-lookup"><span data-stu-id="df815-623">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.7#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="df815-624">В примере ниже с помощью метода [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.7#setasync-datetime--options--callback-) объекта `Time` задается время начала встречи в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="df815-624">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.7#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="df815-625">Тип</span><span class="sxs-lookup"><span data-stu-id="df815-625">Type</span></span>

*   <span data-ttu-id="df815-626">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="df815-626">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.7)</span></span>

##### <a name="requirements"></a><span data-ttu-id="df815-627">Требования</span><span class="sxs-lookup"><span data-stu-id="df815-627">Requirements</span></span>

|<span data-ttu-id="df815-628">Требование</span><span class="sxs-lookup"><span data-stu-id="df815-628">Requirement</span></span>|<span data-ttu-id="df815-629">Значение</span><span class="sxs-lookup"><span data-stu-id="df815-629">Value</span></span>|
|---|---|
|[<span data-ttu-id="df815-630">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="df815-630">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="df815-631">1.0</span><span class="sxs-lookup"><span data-stu-id="df815-631">1.0</span></span>|
|[<span data-ttu-id="df815-632">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="df815-632">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="df815-633">ReadItem</span><span class="sxs-lookup"><span data-stu-id="df815-633">ReadItem</span></span>|
|[<span data-ttu-id="df815-634">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="df815-634">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="df815-635">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="df815-635">Compose or Read</span></span>|

<br>

---
---

#### <a name="subject-stringsubjectjavascriptapioutlookofficesubjectviewoutlook-js-17"></a><span data-ttu-id="df815-636">Тема: строка | [Subject (тема](/javascript/api/outlook/office.subject?view=outlook-js-1.7) )</span><span class="sxs-lookup"><span data-stu-id="df815-636">subject: String|[Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.7)</span></span>

<span data-ttu-id="df815-637">Получает или задает описание, которое отображается в поле темы элемента.</span><span class="sxs-lookup"><span data-stu-id="df815-637">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="df815-638">Свойство `subject` получает или задает всю тему элемента для отправки с почтового сервера.</span><span class="sxs-lookup"><span data-stu-id="df815-638">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="df815-639">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="df815-639">Read mode</span></span>

<span data-ttu-id="df815-p133">Свойство `subject` возвращает строку. С помощью свойства [`normalizedSubject`](#normalizedsubject-string) можно получить тему без начальных префиксов, таких как `RE:` и `FW:`.</span><span class="sxs-lookup"><span data-stu-id="df815-p133">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

<span data-ttu-id="df815-642">В примере кода JavaScript, приведенном ниже, показано, как получить доступ к свойству `subject` текущего элемента в Outlook.</span><span class="sxs-lookup"><span data-stu-id="df815-642">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

```js
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="df815-643">Режим создания</span><span class="sxs-lookup"><span data-stu-id="df815-643">Compose mode</span></span>

<span data-ttu-id="df815-644">Свойство `subject` возвращает объект `Subject`, который предоставляет методы для получения и задания темы.</span><span class="sxs-lookup"><span data-stu-id="df815-644">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```js
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="df815-645">Тип</span><span class="sxs-lookup"><span data-stu-id="df815-645">Type</span></span>

*   <span data-ttu-id="df815-646">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="df815-646">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.7)</span></span>

##### <a name="requirements"></a><span data-ttu-id="df815-647">Требования</span><span class="sxs-lookup"><span data-stu-id="df815-647">Requirements</span></span>

|<span data-ttu-id="df815-648">Требование</span><span class="sxs-lookup"><span data-stu-id="df815-648">Requirement</span></span>|<span data-ttu-id="df815-649">Значение</span><span class="sxs-lookup"><span data-stu-id="df815-649">Value</span></span>|
|---|---|
|[<span data-ttu-id="df815-650">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="df815-650">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="df815-651">1.0</span><span class="sxs-lookup"><span data-stu-id="df815-651">1.0</span></span>|
|[<span data-ttu-id="df815-652">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="df815-652">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="df815-653">ReadItem</span><span class="sxs-lookup"><span data-stu-id="df815-653">ReadItem</span></span>|
|[<span data-ttu-id="df815-654">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="df815-654">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="df815-655">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="df815-655">Compose or Read</span></span>|

<br>

---
---

#### <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-17recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-17"></a><span data-ttu-id="df815-656">Кому: Array. <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)>|[получатели](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="df815-656">to: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span></span>

<span data-ttu-id="df815-657">Предоставляет доступ к получателям, указанным в строке **Кому** сообщения.</span><span class="sxs-lookup"><span data-stu-id="df815-657">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="df815-658">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="df815-658">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="df815-659">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="df815-659">Read mode</span></span>

<span data-ttu-id="df815-p135">Свойство `to` возвращает массив, содержащий объект `EmailAddressDetails` для каждого получателя в строке **Кому** сообщения. Коллекция может включать не более 100 элементов.</span><span class="sxs-lookup"><span data-stu-id="df815-p135">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="df815-662">Режим создания</span><span class="sxs-lookup"><span data-stu-id="df815-662">Compose mode</span></span>

<span data-ttu-id="df815-663">Свойство `to` возвращает объект `Recipients`, предоставляющий методы для получения или обновления получателей, которые указаны в строке **Кому** сообщения.</span><span class="sxs-lookup"><span data-stu-id="df815-663">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

```js
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="df815-664">Тип</span><span class="sxs-lookup"><span data-stu-id="df815-664">Type</span></span>

*   <span data-ttu-id="df815-665">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="df815-665">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span></span>

##### <a name="requirements"></a><span data-ttu-id="df815-666">Требования</span><span class="sxs-lookup"><span data-stu-id="df815-666">Requirements</span></span>

|<span data-ttu-id="df815-667">Требование</span><span class="sxs-lookup"><span data-stu-id="df815-667">Requirement</span></span>|<span data-ttu-id="df815-668">Значение</span><span class="sxs-lookup"><span data-stu-id="df815-668">Value</span></span>|
|---|---|
|[<span data-ttu-id="df815-669">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="df815-669">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="df815-670">1.0</span><span class="sxs-lookup"><span data-stu-id="df815-670">1.0</span></span>|
|[<span data-ttu-id="df815-671">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="df815-671">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="df815-672">ReadItem</span><span class="sxs-lookup"><span data-stu-id="df815-672">ReadItem</span></span>|
|[<span data-ttu-id="df815-673">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="df815-673">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="df815-674">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="df815-674">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="df815-675">Методы</span><span class="sxs-lookup"><span data-stu-id="df815-675">Methods</span></span>

#### <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="df815-676">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="df815-676">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="df815-677">Добавляет файл в сообщение или встречу в качестве вложения.</span><span class="sxs-lookup"><span data-stu-id="df815-677">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="df815-678">Метод `addFileAttachmentAsync` передает файл по указанному универсальному коду ресурса (URI) и вкладывает его в элемент в форме создания.</span><span class="sxs-lookup"><span data-stu-id="df815-678">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="df815-679">Идентификатор можно использовать с методом [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback), чтобы удалить вложение, добавленное во время текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="df815-679">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="df815-680">Параметры</span><span class="sxs-lookup"><span data-stu-id="df815-680">Parameters</span></span>
|<span data-ttu-id="df815-681">Имя</span><span class="sxs-lookup"><span data-stu-id="df815-681">Name</span></span>|<span data-ttu-id="df815-682">Тип</span><span class="sxs-lookup"><span data-stu-id="df815-682">Type</span></span>|<span data-ttu-id="df815-683">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="df815-683">Attributes</span></span>|<span data-ttu-id="df815-684">Описание</span><span class="sxs-lookup"><span data-stu-id="df815-684">Description</span></span>|
|---|---|---|---|
|`uri`|<span data-ttu-id="df815-685">String</span><span class="sxs-lookup"><span data-stu-id="df815-685">String</span></span>||<span data-ttu-id="df815-p136">Универсальный код ресурса (URI), представляющий расположение файла, который нужно вложить в сообщение или встречу. Максимальная длина — 2048 символов.</span><span class="sxs-lookup"><span data-stu-id="df815-p136">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="df815-688">String</span><span class="sxs-lookup"><span data-stu-id="df815-688">String</span></span>||<span data-ttu-id="df815-p137">Имя вложения, которое отображается при передаче вложения. Максимальная длина — 255 символов.</span><span class="sxs-lookup"><span data-stu-id="df815-p137">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="df815-691">Объект</span><span class="sxs-lookup"><span data-stu-id="df815-691">Object</span></span>|<span data-ttu-id="df815-692">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="df815-692">&lt;optional&gt;</span></span>|<span data-ttu-id="df815-693">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="df815-693">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="df815-694">Object</span><span class="sxs-lookup"><span data-stu-id="df815-694">Object</span></span>|<span data-ttu-id="df815-695">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="df815-695">&lt;optional&gt;</span></span>|<span data-ttu-id="df815-696">В методе обратного вызова разработчики могут указать любой объект, к которому необходимо получить доступ.</span><span class="sxs-lookup"><span data-stu-id="df815-696">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="df815-697">Boolean</span><span class="sxs-lookup"><span data-stu-id="df815-697">Boolean</span></span>|<span data-ttu-id="df815-698">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="df815-698">&lt;optional&gt;</span></span>|<span data-ttu-id="df815-699">Значение `true` указывает на то, что вложение будет встроено в текст сообщения и не должно отображаться в списке вложений.</span><span class="sxs-lookup"><span data-stu-id="df815-699">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="df815-700">function</span><span class="sxs-lookup"><span data-stu-id="df815-700">function</span></span>|<span data-ttu-id="df815-701">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="df815-701">&lt;optional&gt;</span></span>|<span data-ttu-id="df815-702">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="df815-702">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="df815-703">После успешного выполнения идентификатор вложения будет представлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="df815-703">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="df815-704">Если добавить вложение не удастся, объект `asyncResult` будет содержать объект `Error` с описанием ошибки.</span><span class="sxs-lookup"><span data-stu-id="df815-704">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="df815-705">Ошибки</span><span class="sxs-lookup"><span data-stu-id="df815-705">Errors</span></span>

|<span data-ttu-id="df815-706">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="df815-706">Error code</span></span>|<span data-ttu-id="df815-707">Описание</span><span class="sxs-lookup"><span data-stu-id="df815-707">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="df815-708">Вложение превышает максимальный размер.</span><span class="sxs-lookup"><span data-stu-id="df815-708">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="df815-709">Расширение вложения не поддерживается.</span><span class="sxs-lookup"><span data-stu-id="df815-709">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="df815-710">Сообщение или встреча содержат слишком много вложений.</span><span class="sxs-lookup"><span data-stu-id="df815-710">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="df815-711">Требования</span><span class="sxs-lookup"><span data-stu-id="df815-711">Requirements</span></span>

|<span data-ttu-id="df815-712">Требование</span><span class="sxs-lookup"><span data-stu-id="df815-712">Requirement</span></span>|<span data-ttu-id="df815-713">Значение</span><span class="sxs-lookup"><span data-stu-id="df815-713">Value</span></span>|
|---|---|
|[<span data-ttu-id="df815-714">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="df815-714">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="df815-715">1.1</span><span class="sxs-lookup"><span data-stu-id="df815-715">1.1</span></span>|
|[<span data-ttu-id="df815-716">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="df815-716">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="df815-717">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="df815-717">ReadWriteItem</span></span>|
|[<span data-ttu-id="df815-718">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="df815-718">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="df815-719">Создание</span><span class="sxs-lookup"><span data-stu-id="df815-719">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="df815-720">Примеры</span><span class="sxs-lookup"><span data-stu-id="df815-720">Examples</span></span>

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

<span data-ttu-id="df815-721">В приведенном ниже примере файл изображения добавляется в качестве встроенного вложения, а ссылка на вложение добавляется в текст сообщения.</span><span class="sxs-lookup"><span data-stu-id="df815-721">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

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

#### <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="df815-722">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="df815-722">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="df815-723">Добавляет обработчик для поддерживаемого события.</span><span class="sxs-lookup"><span data-stu-id="df815-723">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="df815-724">В настоящее время поддерживаются типы `Office.EventType.AppointmentTimeChanged`событий `Office.EventType.RecipientsChanged`, и`Office.EventType.RecurrenceChanged`</span><span class="sxs-lookup"><span data-stu-id="df815-724">Currently the supported event types are `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`</span></span>

##### <a name="parameters"></a><span data-ttu-id="df815-725">Параметры</span><span class="sxs-lookup"><span data-stu-id="df815-725">Parameters</span></span>

| <span data-ttu-id="df815-726">Имя</span><span class="sxs-lookup"><span data-stu-id="df815-726">Name</span></span> | <span data-ttu-id="df815-727">Тип</span><span class="sxs-lookup"><span data-stu-id="df815-727">Type</span></span> | <span data-ttu-id="df815-728">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="df815-728">Attributes</span></span> | <span data-ttu-id="df815-729">Описание</span><span class="sxs-lookup"><span data-stu-id="df815-729">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="df815-730">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="df815-730">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="df815-731">Событие, которое должно вызвать обработчик.</span><span class="sxs-lookup"><span data-stu-id="df815-731">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="df815-732">Function</span><span class="sxs-lookup"><span data-stu-id="df815-732">Function</span></span> || <span data-ttu-id="df815-p138">Функция для обработки события. Функция должна принимать один параметр, представляющий собой объектный литерал. Значение свойства `type` параметра совпадет со значением параметра `eventType`, переданного методу `addHandlerAsync`.</span><span class="sxs-lookup"><span data-stu-id="df815-p138">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="df815-736">Объект</span><span class="sxs-lookup"><span data-stu-id="df815-736">Object</span></span> | <span data-ttu-id="df815-737">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="df815-737">&lt;optional&gt;</span></span> | <span data-ttu-id="df815-738">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="df815-738">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="df815-739">Объект</span><span class="sxs-lookup"><span data-stu-id="df815-739">Object</span></span> | <span data-ttu-id="df815-740">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="df815-740">&lt;optional&gt;</span></span> | <span data-ttu-id="df815-741">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="df815-741">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="df815-742">функция</span><span class="sxs-lookup"><span data-stu-id="df815-742">function</span></span>| <span data-ttu-id="df815-743">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="df815-743">&lt;optional&gt;</span></span>|<span data-ttu-id="df815-744">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="df815-744">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="df815-745">Требования</span><span class="sxs-lookup"><span data-stu-id="df815-745">Requirements</span></span>

|<span data-ttu-id="df815-746">Требование</span><span class="sxs-lookup"><span data-stu-id="df815-746">Requirement</span></span>| <span data-ttu-id="df815-747">Значение</span><span class="sxs-lookup"><span data-stu-id="df815-747">Value</span></span>|
|---|---|
|[<span data-ttu-id="df815-748">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="df815-748">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="df815-749">1.7</span><span class="sxs-lookup"><span data-stu-id="df815-749">1.7</span></span> |
|[<span data-ttu-id="df815-750">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="df815-750">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="df815-751">ReadItem</span><span class="sxs-lookup"><span data-stu-id="df815-751">ReadItem</span></span> |
|[<span data-ttu-id="df815-752">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="df815-752">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="df815-753">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="df815-753">Compose or Read</span></span> |

##### <a name="example"></a><span data-ttu-id="df815-754">Пример</span><span class="sxs-lookup"><span data-stu-id="df815-754">Example</span></span>

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

#### <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="df815-755">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="df815-755">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="df815-756">Добавляет к сообщению элемент Exchange, например сообщение, в виде вложения.</span><span class="sxs-lookup"><span data-stu-id="df815-756">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="df815-p139">С помощью метода `addItemAttachmentAsync` можно в элемент формы создания вложить элемент с указанным идентификатором Exchange. Если указать метод обратного вызова, то этот метод вызывается с помощью параметра `asyncResult`, который содержит идентификатор вложения или код, указывающий на ошибки, которые произошли при вложении элемента. При необходимости можно использовать параметр `options` для передачи сведений о состоянии методу обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="df815-p139">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="df815-760">Идентификатор можно использовать с методом [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback), чтобы удалить вложение, добавленное во время текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="df815-760">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="df815-761">Если ваша надстройка Office работает в Outlook в Интернете, `addItemAttachmentAsync` метод может присоединять элементы к элементам, отличным от редактируемого элемента; Однако это не поддерживается и не рекомендуется.</span><span class="sxs-lookup"><span data-stu-id="df815-761">If your Office Add-in is running in Outlook on the web, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="df815-762">Параметры</span><span class="sxs-lookup"><span data-stu-id="df815-762">Parameters</span></span>

|<span data-ttu-id="df815-763">Имя</span><span class="sxs-lookup"><span data-stu-id="df815-763">Name</span></span>|<span data-ttu-id="df815-764">Тип</span><span class="sxs-lookup"><span data-stu-id="df815-764">Type</span></span>|<span data-ttu-id="df815-765">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="df815-765">Attributes</span></span>|<span data-ttu-id="df815-766">Описание</span><span class="sxs-lookup"><span data-stu-id="df815-766">Description</span></span>|
|---|---|---|---|
|`itemId`|<span data-ttu-id="df815-767">String</span><span class="sxs-lookup"><span data-stu-id="df815-767">String</span></span>||<span data-ttu-id="df815-p140">Идентификатор Exchange для вкладываемого элемента. Максимальная длина — 100 символов.</span><span class="sxs-lookup"><span data-stu-id="df815-p140">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="df815-770">String</span><span class="sxs-lookup"><span data-stu-id="df815-770">String</span></span>||<span data-ttu-id="df815-771">Тема вкладываемого элемента.</span><span class="sxs-lookup"><span data-stu-id="df815-771">The subject of the item to be attached.</span></span> <span data-ttu-id="df815-772">Максимальная длина: 255 символов.</span><span class="sxs-lookup"><span data-stu-id="df815-772">The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="df815-773">Object</span><span class="sxs-lookup"><span data-stu-id="df815-773">Object</span></span>|<span data-ttu-id="df815-774">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="df815-774">&lt;optional&gt;</span></span>|<span data-ttu-id="df815-775">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="df815-775">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="df815-776">Объект</span><span class="sxs-lookup"><span data-stu-id="df815-776">Object</span></span>|<span data-ttu-id="df815-777">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="df815-777">&lt;optional&gt;</span></span>|<span data-ttu-id="df815-778">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="df815-778">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="df815-779">функция</span><span class="sxs-lookup"><span data-stu-id="df815-779">function</span></span>|<span data-ttu-id="df815-780">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="df815-780">&lt;optional&gt;</span></span>|<span data-ttu-id="df815-781">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="df815-781">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="df815-782">После успешного выполнения идентификатор вложения будет представлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="df815-782">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="df815-783">Если добавить вложение не удастся, объект `asyncResult` будет содержать объект `Error` с описанием ошибки.</span><span class="sxs-lookup"><span data-stu-id="df815-783">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="df815-784">Ошибки</span><span class="sxs-lookup"><span data-stu-id="df815-784">Errors</span></span>

|<span data-ttu-id="df815-785">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="df815-785">Error code</span></span>|<span data-ttu-id="df815-786">Описание</span><span class="sxs-lookup"><span data-stu-id="df815-786">Description</span></span>|
|------------|-------------|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="df815-787">Сообщение или встреча содержат слишком много вложений.</span><span class="sxs-lookup"><span data-stu-id="df815-787">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="df815-788">Требования</span><span class="sxs-lookup"><span data-stu-id="df815-788">Requirements</span></span>

|<span data-ttu-id="df815-789">Требование</span><span class="sxs-lookup"><span data-stu-id="df815-789">Requirement</span></span>|<span data-ttu-id="df815-790">Значение</span><span class="sxs-lookup"><span data-stu-id="df815-790">Value</span></span>|
|---|---|
|[<span data-ttu-id="df815-791">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="df815-791">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="df815-792">1.1</span><span class="sxs-lookup"><span data-stu-id="df815-792">1.1</span></span>|
|[<span data-ttu-id="df815-793">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="df815-793">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="df815-794">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="df815-794">ReadWriteItem</span></span>|
|[<span data-ttu-id="df815-795">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="df815-795">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="df815-796">Создание</span><span class="sxs-lookup"><span data-stu-id="df815-796">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="df815-797">Пример</span><span class="sxs-lookup"><span data-stu-id="df815-797">Example</span></span>

<span data-ttu-id="df815-798">В следующем примере существующий элемент Outlook добавляется в виде вложения с именем `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="df815-798">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

#### <a name="close"></a><span data-ttu-id="df815-799">close()</span><span class="sxs-lookup"><span data-stu-id="df815-799">close()</span></span>

<span data-ttu-id="df815-800">Закрывает текущий создаваемый элемент.</span><span class="sxs-lookup"><span data-stu-id="df815-800">Closes the current item that is being composed.</span></span>

<span data-ttu-id="df815-p142">Работа метода `close` зависит от текущего состояния создаваемого элемента. Если элемент содержит несохраненные изменения, клиент предложит пользователю сохранить или отклонить их либо отменить действие закрытия.</span><span class="sxs-lookup"><span data-stu-id="df815-p142">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="df815-803">Если в Outlook в Интернете элемент представляет собой встречу, ранее сохраненную с помощью метода `saveAsync`, пользователю предлагается сохранить, отклонить или отменить действие, даже если с момента последнего сохранения элемента не вносились какие-либо изменения.</span><span class="sxs-lookup"><span data-stu-id="df815-803">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="df815-804">Если в клиенте Outlook для настольных ПК сообщение представляет собой ответ в тексте, метод `close` не работает.</span><span class="sxs-lookup"><span data-stu-id="df815-804">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="df815-805">Требования</span><span class="sxs-lookup"><span data-stu-id="df815-805">Requirements</span></span>

|<span data-ttu-id="df815-806">Требование</span><span class="sxs-lookup"><span data-stu-id="df815-806">Requirement</span></span>|<span data-ttu-id="df815-807">Значение</span><span class="sxs-lookup"><span data-stu-id="df815-807">Value</span></span>|
|---|---|
|[<span data-ttu-id="df815-808">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="df815-808">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="df815-809">1.3</span><span class="sxs-lookup"><span data-stu-id="df815-809">1.3</span></span>|
|[<span data-ttu-id="df815-810">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="df815-810">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="df815-811">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="df815-811">Restricted</span></span>|
|[<span data-ttu-id="df815-812">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="df815-812">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="df815-813">Создание</span><span class="sxs-lookup"><span data-stu-id="df815-813">Compose</span></span>|

<br>

---
---

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="df815-814">displayReplyAllForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="df815-814">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="df815-815">Отображает форму ответа, включающую отправителя и всех получателей выбранного сообщения или организатора и всех участников выбранной встречи.</span><span class="sxs-lookup"><span data-stu-id="df815-815">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="df815-816">Этот метод не поддерживается в Outlook на iOS или Android.</span><span class="sxs-lookup"><span data-stu-id="df815-816">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="df815-817">В Outlook в Интернете форма ответа отображается в виде всплывающей формы в представлении из трех столбцов и всплывающей формы в представлении с 2 или 1 столбца.</span><span class="sxs-lookup"><span data-stu-id="df815-817">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="df815-818">Если любой строковый параметр превышает указанные для него ограничения, `displayReplyAllForm` возвращает исключение.</span><span class="sxs-lookup"><span data-stu-id="df815-818">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="df815-819">Если в `formData.attachments` параметре указаны вложения, Outlook в Интернете и клиенте для настольных компьютеров пытаются скачать все вложения и присоединить их к форме ответа.</span><span class="sxs-lookup"><span data-stu-id="df815-819">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form.</span></span> <span data-ttu-id="df815-820">Если какие-либо вложения добавить не удается, в форме отображается сообщение об ошибке.</span><span class="sxs-lookup"><span data-stu-id="df815-820">If any attachments fail to be added, an error is shown in the form UI.</span></span> <span data-ttu-id="df815-821">Если сообщения об ошибках не предусмотрены, то они не отображаются.</span><span class="sxs-lookup"><span data-stu-id="df815-821">If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="df815-822">Параметры</span><span class="sxs-lookup"><span data-stu-id="df815-822">Parameters</span></span>

|<span data-ttu-id="df815-823">Имя</span><span class="sxs-lookup"><span data-stu-id="df815-823">Name</span></span>|<span data-ttu-id="df815-824">Тип</span><span class="sxs-lookup"><span data-stu-id="df815-824">Type</span></span>|<span data-ttu-id="df815-825">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="df815-825">Attributes</span></span>|<span data-ttu-id="df815-826">Описание</span><span class="sxs-lookup"><span data-stu-id="df815-826">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="df815-827">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="df815-827">String &#124; Object</span></span>||<span data-ttu-id="df815-p144">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="df815-p144">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="df815-830">**ИЛИ**</span><span class="sxs-lookup"><span data-stu-id="df815-830">**OR**</span></span><br/><span data-ttu-id="df815-p145">Объект, который содержит текст или данные вложения и функцию обратного вызова. Ниже представлено определение этого объекта.</span><span class="sxs-lookup"><span data-stu-id="df815-p145">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="df815-833">String.</span><span class="sxs-lookup"><span data-stu-id="df815-833">String</span></span>|<span data-ttu-id="df815-834">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="df815-834">&lt;optional&gt;</span></span>|<span data-ttu-id="df815-p146">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="df815-p146">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="df815-837">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="df815-837">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="df815-838">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="df815-838">&lt;optional&gt;</span></span>|<span data-ttu-id="df815-839">Массив объектов JSON, представляющих собой вложенные файлы или элементы.</span><span class="sxs-lookup"><span data-stu-id="df815-839">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="df815-840">String.</span><span class="sxs-lookup"><span data-stu-id="df815-840">String</span></span>||<span data-ttu-id="df815-p147">Указывает тип вложения. Допустимые значения: `file` для вложенного файла и `item` для вложенного элемента.</span><span class="sxs-lookup"><span data-stu-id="df815-p147">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="df815-843">Строка</span><span class="sxs-lookup"><span data-stu-id="df815-843">String</span></span>||<span data-ttu-id="df815-844">Строка, содержащая имя вложения, длиной до 255 символов.</span><span class="sxs-lookup"><span data-stu-id="df815-844">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="df815-845">Строка</span><span class="sxs-lookup"><span data-stu-id="df815-845">String</span></span>||<span data-ttu-id="df815-p148">Используется, только если свойству `type` задано значение `file`. URI расположения файла.</span><span class="sxs-lookup"><span data-stu-id="df815-p148">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="df815-848">Логический</span><span class="sxs-lookup"><span data-stu-id="df815-848">Boolean</span></span>||<span data-ttu-id="df815-p149">Используется, только если свойству `type` задано значение `file`. Значение `true` указывает на то, что вложение будет встроено в текст сообщения и не должно отображаться в списке вложений.</span><span class="sxs-lookup"><span data-stu-id="df815-p149">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="df815-851">String</span><span class="sxs-lookup"><span data-stu-id="df815-851">String</span></span>||<span data-ttu-id="df815-p150">Используется, только если свойству `type` задано значение `item`. Идентификатор вложения EWS. Это строка длиной до 100 символов.</span><span class="sxs-lookup"><span data-stu-id="df815-p150">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="df815-855">function</span><span class="sxs-lookup"><span data-stu-id="df815-855">function</span></span>|<span data-ttu-id="df815-856">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="df815-856">&lt;optional&gt;</span></span>|<span data-ttu-id="df815-857">По завершении работы метода функция, переданная параметру `callback`, вызывается с помощью одного параметра `asyncResult`, представляющего собой объект [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="df815-857">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="df815-858">Требования</span><span class="sxs-lookup"><span data-stu-id="df815-858">Requirements</span></span>

|<span data-ttu-id="df815-859">Требование</span><span class="sxs-lookup"><span data-stu-id="df815-859">Requirement</span></span>|<span data-ttu-id="df815-860">Значение</span><span class="sxs-lookup"><span data-stu-id="df815-860">Value</span></span>|
|---|---|
|[<span data-ttu-id="df815-861">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="df815-861">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="df815-862">1.0</span><span class="sxs-lookup"><span data-stu-id="df815-862">1.0</span></span>|
|[<span data-ttu-id="df815-863">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="df815-863">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="df815-864">ReadItem</span><span class="sxs-lookup"><span data-stu-id="df815-864">ReadItem</span></span>|
|[<span data-ttu-id="df815-865">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="df815-865">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="df815-866">Чтение</span><span class="sxs-lookup"><span data-stu-id="df815-866">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="df815-867">Примеры</span><span class="sxs-lookup"><span data-stu-id="df815-867">Examples</span></span>

<span data-ttu-id="df815-868">Приведенный ниже код передает строку в функцию `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="df815-868">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="df815-869">Ответ с пустым текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="df815-869">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="df815-870">Ответ только с текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="df815-870">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="df815-871">Ответ с текстом сообщения и вложенным файлом.</span><span class="sxs-lookup"><span data-stu-id="df815-871">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="df815-872">Ответ с текстом сообщения и вложенным элементом.</span><span class="sxs-lookup"><span data-stu-id="df815-872">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="df815-873">Ответ с текстом сообщения, вложенным файлом, вложенным элементом и обратным вызовом.</span><span class="sxs-lookup"><span data-stu-id="df815-873">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="df815-874">displayReplyForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="df815-874">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="df815-875">Отображает форму ответа, включающую только отправителя выбранного сообщения или организатора выбранной встречи.</span><span class="sxs-lookup"><span data-stu-id="df815-875">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="df815-876">Этот метод не поддерживается в Outlook на iOS или Android.</span><span class="sxs-lookup"><span data-stu-id="df815-876">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="df815-877">В Outlook в Интернете форма ответа отображается в виде всплывающей формы в представлении из трех столбцов и всплывающей формы в представлении с 2 или 1 столбца.</span><span class="sxs-lookup"><span data-stu-id="df815-877">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="df815-878">Если любой строковый параметр превышает указанные для него ограничения, `displayReplyForm` возвращает исключение.</span><span class="sxs-lookup"><span data-stu-id="df815-878">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="df815-879">Если в `formData.attachments` параметре указаны вложения, Outlook в Интернете и клиенте для настольных компьютеров пытаются скачать все вложения и присоединить их к форме ответа.</span><span class="sxs-lookup"><span data-stu-id="df815-879">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form.</span></span> <span data-ttu-id="df815-880">Если какие-либо вложения добавить не удается, в форме отображается сообщение об ошибке.</span><span class="sxs-lookup"><span data-stu-id="df815-880">If any attachments fail to be added, an error is shown in the form UI.</span></span> <span data-ttu-id="df815-881">Если сообщения об ошибках не предусмотрены, то они не отображаются.</span><span class="sxs-lookup"><span data-stu-id="df815-881">If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="df815-882">Параметры</span><span class="sxs-lookup"><span data-stu-id="df815-882">Parameters</span></span>

|<span data-ttu-id="df815-883">Имя</span><span class="sxs-lookup"><span data-stu-id="df815-883">Name</span></span>|<span data-ttu-id="df815-884">Тип</span><span class="sxs-lookup"><span data-stu-id="df815-884">Type</span></span>|<span data-ttu-id="df815-885">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="df815-885">Attributes</span></span>|<span data-ttu-id="df815-886">Описание</span><span class="sxs-lookup"><span data-stu-id="df815-886">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="df815-887">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="df815-887">String &#124; Object</span></span>||<span data-ttu-id="df815-p152">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="df815-p152">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="df815-890">**ИЛИ**</span><span class="sxs-lookup"><span data-stu-id="df815-890">**OR**</span></span><br/><span data-ttu-id="df815-p153">Объект, который содержит текст или данные вложения и функцию обратного вызова. Ниже представлено определение этого объекта.</span><span class="sxs-lookup"><span data-stu-id="df815-p153">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="df815-893">String.</span><span class="sxs-lookup"><span data-stu-id="df815-893">String</span></span>|<span data-ttu-id="df815-894">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="df815-894">&lt;optional&gt;</span></span>|<span data-ttu-id="df815-p154">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="df815-p154">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="df815-897">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="df815-897">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="df815-898">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="df815-898">&lt;optional&gt;</span></span>|<span data-ttu-id="df815-899">Массив объектов JSON, представляющих собой вложенные файлы или элементы.</span><span class="sxs-lookup"><span data-stu-id="df815-899">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="df815-900">String</span><span class="sxs-lookup"><span data-stu-id="df815-900">String</span></span>||<span data-ttu-id="df815-p155">Указывает тип вложения. Допустимые значения: `file` для вложенного файла и `item` для вложенного элемента.</span><span class="sxs-lookup"><span data-stu-id="df815-p155">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="df815-903">Строка</span><span class="sxs-lookup"><span data-stu-id="df815-903">String</span></span>||<span data-ttu-id="df815-904">Строка, содержащая имя вложения, длиной до 255 символов.</span><span class="sxs-lookup"><span data-stu-id="df815-904">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="df815-905">Строка</span><span class="sxs-lookup"><span data-stu-id="df815-905">String</span></span>||<span data-ttu-id="df815-p156">Используется, только если свойству `type` задано значение `file`. URI расположения файла.</span><span class="sxs-lookup"><span data-stu-id="df815-p156">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="df815-908">Логический</span><span class="sxs-lookup"><span data-stu-id="df815-908">Boolean</span></span>||<span data-ttu-id="df815-p157">Используется, только если свойству `type` задано значение `file`. Значение `true` указывает на то, что вложение будет встроено в текст сообщения и не должно отображаться в списке вложений.</span><span class="sxs-lookup"><span data-stu-id="df815-p157">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="df815-911">String</span><span class="sxs-lookup"><span data-stu-id="df815-911">String</span></span>||<span data-ttu-id="df815-p158">Используется, только если свойству `type` задано значение `item`. Идентификатор вложения EWS. Это строка длиной до 100 символов.</span><span class="sxs-lookup"><span data-stu-id="df815-p158">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="df815-915">function</span><span class="sxs-lookup"><span data-stu-id="df815-915">function</span></span>|<span data-ttu-id="df815-916">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="df815-916">&lt;optional&gt;</span></span>|<span data-ttu-id="df815-917">По завершении работы метода функция, переданная параметру `callback`, вызывается с помощью одного параметра `asyncResult`, представляющего собой объект [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="df815-917">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="df815-918">Требования</span><span class="sxs-lookup"><span data-stu-id="df815-918">Requirements</span></span>

|<span data-ttu-id="df815-919">Требование</span><span class="sxs-lookup"><span data-stu-id="df815-919">Requirement</span></span>|<span data-ttu-id="df815-920">Значение</span><span class="sxs-lookup"><span data-stu-id="df815-920">Value</span></span>|
|---|---|
|[<span data-ttu-id="df815-921">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="df815-921">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="df815-922">1.0</span><span class="sxs-lookup"><span data-stu-id="df815-922">1.0</span></span>|
|[<span data-ttu-id="df815-923">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="df815-923">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="df815-924">ReadItem</span><span class="sxs-lookup"><span data-stu-id="df815-924">ReadItem</span></span>|
|[<span data-ttu-id="df815-925">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="df815-925">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="df815-926">Чтение</span><span class="sxs-lookup"><span data-stu-id="df815-926">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="df815-927">Примеры</span><span class="sxs-lookup"><span data-stu-id="df815-927">Examples</span></span>

<span data-ttu-id="df815-928">Приведенный ниже код передает строку в функцию `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="df815-928">The following code passes a string to the `displayReplyForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="df815-929">Ответ с пустым текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="df815-929">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="df815-930">Ответ только с текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="df815-930">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="df815-931">Ответ с текстом сообщения и вложенным файлом.</span><span class="sxs-lookup"><span data-stu-id="df815-931">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="df815-932">Ответ с текстом сообщения и вложенным элементом.</span><span class="sxs-lookup"><span data-stu-id="df815-932">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="df815-933">Ответ с текстом сообщения, вложенным файлом, вложенным элементом и обратным вызовом.</span><span class="sxs-lookup"><span data-stu-id="df815-933">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlookofficeentitiesviewoutlook-js-17"></a><span data-ttu-id="df815-934">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.7)}</span><span class="sxs-lookup"><span data-stu-id="df815-934">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.7)}</span></span>

<span data-ttu-id="df815-935">Получает сущности, обнаруженные в теле выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="df815-935">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="df815-936">Этот метод не поддерживается в Outlook на iOS или Android.</span><span class="sxs-lookup"><span data-stu-id="df815-936">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="df815-937">Требования</span><span class="sxs-lookup"><span data-stu-id="df815-937">Requirements</span></span>

|<span data-ttu-id="df815-938">Требование</span><span class="sxs-lookup"><span data-stu-id="df815-938">Requirement</span></span>|<span data-ttu-id="df815-939">Значение</span><span class="sxs-lookup"><span data-stu-id="df815-939">Value</span></span>|
|---|---|
|[<span data-ttu-id="df815-940">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="df815-940">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="df815-941">1.0</span><span class="sxs-lookup"><span data-stu-id="df815-941">1.0</span></span>|
|[<span data-ttu-id="df815-942">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="df815-942">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="df815-943">ReadItem</span><span class="sxs-lookup"><span data-stu-id="df815-943">ReadItem</span></span>|
|[<span data-ttu-id="df815-944">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="df815-944">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="df815-945">Чтение</span><span class="sxs-lookup"><span data-stu-id="df815-945">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="df815-946">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="df815-946">Returns:</span></span>

<span data-ttu-id="df815-947">Тип: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="df815-947">Type: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.7)</span></span>

##### <a name="example"></a><span data-ttu-id="df815-948">Пример</span><span class="sxs-lookup"><span data-stu-id="df815-948">Example</span></span>

<span data-ttu-id="df815-949">Ниже приведен пример получения доступа к сущностям контактов в теле текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="df815-949">The following example accesses the contacts entities in the current item's body.</span></span>

```js
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

<br>

---
---

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-17meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-17phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-17tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-17"></a><span data-ttu-id="df815-950">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.7)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.7)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.7)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.7))>}</span><span class="sxs-lookup"><span data-stu-id="df815-950">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.7)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.7)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.7)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.7))>}</span></span>

<span data-ttu-id="df815-951">Получает массив всех сущностей указанного типа, обнаруженных в теле выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="df815-951">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="df815-952">Этот метод не поддерживается в Outlook на iOS или Android.</span><span class="sxs-lookup"><span data-stu-id="df815-952">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="df815-953">Параметры</span><span class="sxs-lookup"><span data-stu-id="df815-953">Parameters</span></span>

|<span data-ttu-id="df815-954">Имя</span><span class="sxs-lookup"><span data-stu-id="df815-954">Name</span></span>|<span data-ttu-id="df815-955">Тип</span><span class="sxs-lookup"><span data-stu-id="df815-955">Type</span></span>|<span data-ttu-id="df815-956">Описание</span><span class="sxs-lookup"><span data-stu-id="df815-956">Description</span></span>|
|---|---|---|
|`entityType`|[<span data-ttu-id="df815-957">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="df815-957">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook/office.mailboxenums.entitytype?view=outlook-js-1.7)|<span data-ttu-id="df815-958">Одно из значений перечисления EntityType.</span><span class="sxs-lookup"><span data-stu-id="df815-958">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="df815-959">Требования</span><span class="sxs-lookup"><span data-stu-id="df815-959">Requirements</span></span>

|<span data-ttu-id="df815-960">Требование</span><span class="sxs-lookup"><span data-stu-id="df815-960">Requirement</span></span>|<span data-ttu-id="df815-961">Значение</span><span class="sxs-lookup"><span data-stu-id="df815-961">Value</span></span>|
|---|---|
|[<span data-ttu-id="df815-962">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="df815-962">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="df815-963">1.0</span><span class="sxs-lookup"><span data-stu-id="df815-963">1.0</span></span>|
|[<span data-ttu-id="df815-964">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="df815-964">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="df815-965">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="df815-965">Restricted</span></span>|
|[<span data-ttu-id="df815-966">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="df815-966">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="df815-967">Чтение</span><span class="sxs-lookup"><span data-stu-id="df815-967">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="df815-968">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="df815-968">Returns:</span></span>

<span data-ttu-id="df815-969">Если значение, переданное в `entityType`, не является допустимым членом перечисления `EntityType`, метод возвращает значение NULL.</span><span class="sxs-lookup"><span data-stu-id="df815-969">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="df815-970">Если в теле элемента отсутствуют сущности указанного типа, метод возвращает пустой массив.</span><span class="sxs-lookup"><span data-stu-id="df815-970">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="df815-971">В противном случае тип объектов в возвращаемом массиве зависит от типа сущности, запрошенной в параметре `entityType`.</span><span class="sxs-lookup"><span data-stu-id="df815-971">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="df815-972">Хотя минимальный уровень разрешений для использования этого метода — **Restricted**, для некоторых типов сущностей требуется доступ на уровне **ReadItem**, как указано в приведенной ниже таблице.</span><span class="sxs-lookup"><span data-stu-id="df815-972">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

|<span data-ttu-id="df815-973">Значение параметра `entityType`</span><span class="sxs-lookup"><span data-stu-id="df815-973">Value of `entityType`</span></span>|<span data-ttu-id="df815-974">Тип объектов в возвращаемом массиве</span><span class="sxs-lookup"><span data-stu-id="df815-974">Type of objects in returned array</span></span>|<span data-ttu-id="df815-975">Необходимый уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="df815-975">Required Permission Level</span></span>|
|---|---|---|
|`Address`|<span data-ttu-id="df815-976">String</span><span class="sxs-lookup"><span data-stu-id="df815-976">String</span></span>|<span data-ttu-id="df815-977">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="df815-977">**Restricted**</span></span>|
|`Contact`|<span data-ttu-id="df815-978">Contact</span><span class="sxs-lookup"><span data-stu-id="df815-978">Contact</span></span>|<span data-ttu-id="df815-979">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="df815-979">**ReadItem**</span></span>|
|`EmailAddress`|<span data-ttu-id="df815-980">String</span><span class="sxs-lookup"><span data-stu-id="df815-980">String</span></span>|<span data-ttu-id="df815-981">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="df815-981">**ReadItem**</span></span>|
|`MeetingSuggestion`|<span data-ttu-id="df815-982">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="df815-982">MeetingSuggestion</span></span>|<span data-ttu-id="df815-983">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="df815-983">**ReadItem**</span></span>|
|`PhoneNumber`|<span data-ttu-id="df815-984">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="df815-984">PhoneNumber</span></span>|<span data-ttu-id="df815-985">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="df815-985">**Restricted**</span></span>|
|`TaskSuggestion`|<span data-ttu-id="df815-986">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="df815-986">TaskSuggestion</span></span>|<span data-ttu-id="df815-987">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="df815-987">**ReadItem**</span></span>|
|`URL`|<span data-ttu-id="df815-988">String</span><span class="sxs-lookup"><span data-stu-id="df815-988">String</span></span>|<span data-ttu-id="df815-989">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="df815-989">**Restricted**</span></span>|

<span data-ttu-id="df815-990">Тип: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.7)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.7)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.7)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.7))></span><span class="sxs-lookup"><span data-stu-id="df815-990">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.7)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.7)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.7)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.7))></span></span>

##### <a name="example"></a><span data-ttu-id="df815-991">Пример</span><span class="sxs-lookup"><span data-stu-id="df815-991">Example</span></span>

<span data-ttu-id="df815-992">В примере ниже показано, как получить доступ к массиву строк, которые представляют собой почтовые адреса в теле текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="df815-992">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-17meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-17phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-17tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-17"></a><span data-ttu-id="df815-993">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.7)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.7)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.7)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.7))>}</span><span class="sxs-lookup"><span data-stu-id="df815-993">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.7)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.7)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.7)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.7))>}</span></span>

<span data-ttu-id="df815-994">Возвращает известные сущности в выбранном элементе, которые проходят через именованный фильтр, определяемый в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="df815-994">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="df815-995">Этот метод не поддерживается в Outlook на iOS или Android.</span><span class="sxs-lookup"><span data-stu-id="df815-995">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="df815-996">Метод `getFilteredEntitiesByName` возвращает сущности, соответствующие регулярному выражению, которое определяется в элементе правила [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) в XML-файле манифеста, с использованием указанного значения элемента `FilterName`.</span><span class="sxs-lookup"><span data-stu-id="df815-996">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="df815-997">Параметры</span><span class="sxs-lookup"><span data-stu-id="df815-997">Parameters</span></span>

|<span data-ttu-id="df815-998">Имя</span><span class="sxs-lookup"><span data-stu-id="df815-998">Name</span></span>|<span data-ttu-id="df815-999">Тип</span><span class="sxs-lookup"><span data-stu-id="df815-999">Type</span></span>|<span data-ttu-id="df815-1000">Описание</span><span class="sxs-lookup"><span data-stu-id="df815-1000">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="df815-1001">String</span><span class="sxs-lookup"><span data-stu-id="df815-1001">String</span></span>|<span data-ttu-id="df815-1002">Имя элемента правила `ItemHasKnownEntity`, определяющее соответствующий фильтр.</span><span class="sxs-lookup"><span data-stu-id="df815-1002">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="df815-1003">Требования</span><span class="sxs-lookup"><span data-stu-id="df815-1003">Requirements</span></span>

|<span data-ttu-id="df815-1004">Требование</span><span class="sxs-lookup"><span data-stu-id="df815-1004">Requirement</span></span>|<span data-ttu-id="df815-1005">Значение</span><span class="sxs-lookup"><span data-stu-id="df815-1005">Value</span></span>|
|---|---|
|[<span data-ttu-id="df815-1006">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="df815-1006">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="df815-1007">1.0</span><span class="sxs-lookup"><span data-stu-id="df815-1007">1.0</span></span>|
|[<span data-ttu-id="df815-1008">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="df815-1008">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="df815-1009">ReadItem</span><span class="sxs-lookup"><span data-stu-id="df815-1009">ReadItem</span></span>|
|[<span data-ttu-id="df815-1010">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="df815-1010">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="df815-1011">Чтение</span><span class="sxs-lookup"><span data-stu-id="df815-1011">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="df815-1012">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="df815-1012">Returns:</span></span>

<span data-ttu-id="df815-p160">Если в манифесте нет элемента `ItemHasKnownEntity` со значением `FilterName`, соответствующим параметру `name`, метод возвращает `null`. Если параметр `name` соответствует элементу `ItemHasKnownEntity` в манифесте, но при этом в текущем элементе нет соответствующих сущностей, метод возвращает пустой массив.</span><span class="sxs-lookup"><span data-stu-id="df815-p160">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="df815-1015">Тип: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.7)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.7)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.7)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.7))></span><span class="sxs-lookup"><span data-stu-id="df815-1015">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.7)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.7)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.7)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.7))></span></span>

<br>

---
---

#### <a name="getregexmatches--object"></a><span data-ttu-id="df815-1016">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="df815-1016">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="df815-1017">Возвращает строковые значения в выбранном элементе, которые соответствуют регулярным выражениям, определенным в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="df815-1017">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="df815-1018">Этот метод не поддерживается в Outlook на iOS или Android.</span><span class="sxs-lookup"><span data-stu-id="df815-1018">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="df815-p161">Метод `getRegExMatches` возвращает строки, соответствующие регулярному выражению, которое определяется в каждом элементе правила `ItemHasRegularExpressionMatch` или `ItemHasKnownEntity` в XML-файле манифеста. Для правила `ItemHasRegularExpressionMatch` соответствующую строку должно содержать свойство элемента, указанного этим правилом. Простой тип `PropertyName` определяет поддерживаемые свойства.</span><span class="sxs-lookup"><span data-stu-id="df815-p161">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="df815-1022">Например, рассмотрим манифест надстройки, который содержит указанный ниже элемент `Rule`.</span><span class="sxs-lookup"><span data-stu-id="df815-1022">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="df815-1023">Объект, возвращаемый методом `getRegExMatches`, будет содержать два свойства: `fruits` и `veggies`.</span><span class="sxs-lookup"><span data-stu-id="df815-1023">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="df815-p162">Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты. Лучше используйте метод [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.7#getasync-coerciontype--options--callback-) для этого.</span><span class="sxs-lookup"><span data-stu-id="df815-p162">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.7#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="df815-1027">Requirements</span><span class="sxs-lookup"><span data-stu-id="df815-1027">Requirements</span></span>

|<span data-ttu-id="df815-1028">Требование</span><span class="sxs-lookup"><span data-stu-id="df815-1028">Requirement</span></span>|<span data-ttu-id="df815-1029">Значение</span><span class="sxs-lookup"><span data-stu-id="df815-1029">Value</span></span>|
|---|---|
|[<span data-ttu-id="df815-1030">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="df815-1030">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="df815-1031">1.0</span><span class="sxs-lookup"><span data-stu-id="df815-1031">1.0</span></span>|
|[<span data-ttu-id="df815-1032">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="df815-1032">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="df815-1033">ReadItem</span><span class="sxs-lookup"><span data-stu-id="df815-1033">ReadItem</span></span>|
|[<span data-ttu-id="df815-1034">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="df815-1034">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="df815-1035">Чтение</span><span class="sxs-lookup"><span data-stu-id="df815-1035">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="df815-1036">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="df815-1036">Returns:</span></span>

<span data-ttu-id="df815-p163">Объект, содержащий массив строк, которые соответствуют регулярным выражениям, определяемым в XML-файле манифеста. Имя каждого массива равно соответствующему значению атрибута `RegExName` подходящего правила `ItemHasRegularExpressionMatch` или атрибута `FilterName` соответствующего правила `ItemHasKnownEntity`.</span><span class="sxs-lookup"><span data-stu-id="df815-p163">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<span data-ttu-id="df815-1039">Тип: Object</span><span class="sxs-lookup"><span data-stu-id="df815-1039">Type: Object</span></span>

##### <a name="example"></a><span data-ttu-id="df815-1040">Пример</span><span class="sxs-lookup"><span data-stu-id="df815-1040">Example</span></span>

<span data-ttu-id="df815-1041">В приведенном ниже примере показано, как получить доступ к массиву совпадений с элементами `fruits` и `veggies` правил активации регулярных выражений, указанными в манифесте.</span><span class="sxs-lookup"><span data-stu-id="df815-1041">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```js
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

<br>

---
---

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="df815-1042">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="df815-1042">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="df815-1043">Возвращает строковые значения в выбранном элементе, которые соответствуют именованному регулярному выражению, определенному в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="df815-1043">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="df815-1044">Этот метод не поддерживается в Outlook на iOS или Android.</span><span class="sxs-lookup"><span data-stu-id="df815-1044">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="df815-1045">Метод `getRegExMatchesByName` возвращает строки, соответствующие регулярному выражению, которое определяется в элементе правила `ItemHasRegularExpressionMatch` в XML-файле манифеста, с использованием указанного значения элемента `RegExName`.</span><span class="sxs-lookup"><span data-stu-id="df815-1045">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="df815-p164">Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты.</span><span class="sxs-lookup"><span data-stu-id="df815-p164">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="df815-1048">Параметры</span><span class="sxs-lookup"><span data-stu-id="df815-1048">Parameters</span></span>

|<span data-ttu-id="df815-1049">Имя</span><span class="sxs-lookup"><span data-stu-id="df815-1049">Name</span></span>|<span data-ttu-id="df815-1050">Тип</span><span class="sxs-lookup"><span data-stu-id="df815-1050">Type</span></span>|<span data-ttu-id="df815-1051">Описание</span><span class="sxs-lookup"><span data-stu-id="df815-1051">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="df815-1052">String</span><span class="sxs-lookup"><span data-stu-id="df815-1052">String</span></span>|<span data-ttu-id="df815-1053">Имя элемента правила `ItemHasRegularExpressionMatch`, определяющее соответствующий фильтр.</span><span class="sxs-lookup"><span data-stu-id="df815-1053">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="df815-1054">Требования</span><span class="sxs-lookup"><span data-stu-id="df815-1054">Requirements</span></span>

|<span data-ttu-id="df815-1055">Требование</span><span class="sxs-lookup"><span data-stu-id="df815-1055">Requirement</span></span>|<span data-ttu-id="df815-1056">Значение</span><span class="sxs-lookup"><span data-stu-id="df815-1056">Value</span></span>|
|---|---|
|[<span data-ttu-id="df815-1057">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="df815-1057">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="df815-1058">1.0</span><span class="sxs-lookup"><span data-stu-id="df815-1058">1.0</span></span>|
|[<span data-ttu-id="df815-1059">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="df815-1059">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="df815-1060">ReadItem</span><span class="sxs-lookup"><span data-stu-id="df815-1060">ReadItem</span></span>|
|[<span data-ttu-id="df815-1061">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="df815-1061">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="df815-1062">Чтение</span><span class="sxs-lookup"><span data-stu-id="df815-1062">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="df815-1063">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="df815-1063">Returns:</span></span>

<span data-ttu-id="df815-1064">Массив строк, соответствующих регулярному выражению, определяемому в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="df815-1064">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<span data-ttu-id="df815-1065">Тип: Array. < String ></span><span class="sxs-lookup"><span data-stu-id="df815-1065">Type: Array.< String ></span></span>

##### <a name="example"></a><span data-ttu-id="df815-1066">Пример</span><span class="sxs-lookup"><span data-stu-id="df815-1066">Example</span></span>

```js
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

<br>

---
---

#### <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="df815-1067">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="df815-1067">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="df815-1068">Асинхронно возвращает данные, выбранные в теме или тексте сообщения.</span><span class="sxs-lookup"><span data-stu-id="df815-1068">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="df815-p165">Если выделенный фрагмент отсутствует, но курсор находится в тексте или теме, метод возвращает значение NULL для выбранных данных. Если выбраны не текст и не тема, метод возвращает ошибку `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="df815-p165">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="df815-1071">Параметры</span><span class="sxs-lookup"><span data-stu-id="df815-1071">Parameters</span></span>

|<span data-ttu-id="df815-1072">Имя</span><span class="sxs-lookup"><span data-stu-id="df815-1072">Name</span></span>|<span data-ttu-id="df815-1073">Тип</span><span class="sxs-lookup"><span data-stu-id="df815-1073">Type</span></span>|<span data-ttu-id="df815-1074">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="df815-1074">Attributes</span></span>|<span data-ttu-id="df815-1075">Описание</span><span class="sxs-lookup"><span data-stu-id="df815-1075">Description</span></span>|
|---|---|---|---|
|`coercionType`|[<span data-ttu-id="df815-1076">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="df815-1076">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="df815-p166">Запрашивает формат данных. Если задано значение Text, метод возвращает обычный текст как строку, удаляя все имеющиеся HTML-теги. Если задано значение HTML, метод возвращает выделенный текст (обычный текст или HTML).</span><span class="sxs-lookup"><span data-stu-id="df815-p166">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`|<span data-ttu-id="df815-1080">Object</span><span class="sxs-lookup"><span data-stu-id="df815-1080">Object</span></span>|<span data-ttu-id="df815-1081">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="df815-1081">&lt;optional&gt;</span></span>|<span data-ttu-id="df815-1082">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="df815-1082">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="df815-1083">Объект</span><span class="sxs-lookup"><span data-stu-id="df815-1083">Object</span></span>|<span data-ttu-id="df815-1084">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="df815-1084">&lt;optional&gt;</span></span>|<span data-ttu-id="df815-1085">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="df815-1085">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="df815-1086">функция</span><span class="sxs-lookup"><span data-stu-id="df815-1086">function</span></span>||<span data-ttu-id="df815-1087">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="df815-1087">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="df815-1088">Чтобы получить доступ к выбранным данным из метода обратного вызова, вызовите `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="df815-1088">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="df815-1089">Для доступа к исходному свойству, представляющему собой источник выбранных данных, вызовите параметр `asyncResult.value.sourceProperty`, который может иметь значение `body` или `subject`.</span><span class="sxs-lookup"><span data-stu-id="df815-1089">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="df815-1090">Требования</span><span class="sxs-lookup"><span data-stu-id="df815-1090">Requirements</span></span>

|<span data-ttu-id="df815-1091">Требование</span><span class="sxs-lookup"><span data-stu-id="df815-1091">Requirement</span></span>|<span data-ttu-id="df815-1092">Значение</span><span class="sxs-lookup"><span data-stu-id="df815-1092">Value</span></span>|
|---|---|
|[<span data-ttu-id="df815-1093">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="df815-1093">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="df815-1094">1.2</span><span class="sxs-lookup"><span data-stu-id="df815-1094">1.2</span></span>|
|[<span data-ttu-id="df815-1095">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="df815-1095">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="df815-1096">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="df815-1096">ReadWriteItem</span></span>|
|[<span data-ttu-id="df815-1097">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="df815-1097">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="df815-1098">Создание</span><span class="sxs-lookup"><span data-stu-id="df815-1098">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="df815-1099">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="df815-1099">Returns:</span></span>

<span data-ttu-id="df815-1100">Выбранные данные в виде строки с форматом, определенным в параметре `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="df815-1100">The selected data as a string with format determined by `coercionType`.</span></span>

<span data-ttu-id="df815-1101">Тип: String</span><span class="sxs-lookup"><span data-stu-id="df815-1101">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="df815-1102">Пример</span><span class="sxs-lookup"><span data-stu-id="df815-1102">Example</span></span>

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

#### <a name="getselectedentities--entitiesjavascriptapioutlookofficeentitiesviewoutlook-js-17"></a><span data-ttu-id="df815-1103">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.7)}</span><span class="sxs-lookup"><span data-stu-id="df815-1103">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.7)}</span></span>

<span data-ttu-id="df815-1104">Возвращает сущности, найденные в выделенном совпадении, выбранном пользователем.</span><span class="sxs-lookup"><span data-stu-id="df815-1104">Gets the entities found in a highlighted match a user has selected.</span></span> <span data-ttu-id="df815-1105">Выделенные совпадения применяются к [контекстным надстройкам](/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="df815-1105">Highlighted matches apply to [contextual add-ins](/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="df815-1106">Этот метод не поддерживается в Outlook на iOS или Android.</span><span class="sxs-lookup"><span data-stu-id="df815-1106">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="df815-1107">Требования</span><span class="sxs-lookup"><span data-stu-id="df815-1107">Requirements</span></span>

|<span data-ttu-id="df815-1108">Требование</span><span class="sxs-lookup"><span data-stu-id="df815-1108">Requirement</span></span>|<span data-ttu-id="df815-1109">Значение</span><span class="sxs-lookup"><span data-stu-id="df815-1109">Value</span></span>|
|---|---|
|[<span data-ttu-id="df815-1110">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="df815-1110">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="df815-1111">1.6</span><span class="sxs-lookup"><span data-stu-id="df815-1111">1.6</span></span>|
|[<span data-ttu-id="df815-1112">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="df815-1112">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="df815-1113">ReadItem</span><span class="sxs-lookup"><span data-stu-id="df815-1113">ReadItem</span></span>|
|[<span data-ttu-id="df815-1114">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="df815-1114">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="df815-1115">Чтение</span><span class="sxs-lookup"><span data-stu-id="df815-1115">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="df815-1116">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="df815-1116">Returns:</span></span>

<span data-ttu-id="df815-1117">Тип: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="df815-1117">Type: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.7)</span></span>

##### <a name="example"></a><span data-ttu-id="df815-1118">Пример</span><span class="sxs-lookup"><span data-stu-id="df815-1118">Example</span></span>

<span data-ttu-id="df815-1119">В приведенном ниже примере показано, как получить доступ к сущностям адресов в выделенном совпадении, выбранном пользователем.</span><span class="sxs-lookup"><span data-stu-id="df815-1119">The following example accesses the addresses entities in the highlighted match selected by the user.</span></span>

```js
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

<br>

---
---

#### <a name="getselectedregexmatches--object"></a><span data-ttu-id="df815-1120">getSelectedRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="df815-1120">getSelectedRegExMatches() → {Object}</span></span>

<span data-ttu-id="df815-p169">Возвращает строковые значения в выделенном совпадении, которые соответствуют регулярным выражениям, определенным в XML-файле манифеста. Выделенные совпадения применяются к [контекстным надстройкам](/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="df815-p169">Returns string values in a highlighted match that match the regular expressions defined in the manifest XML file. Highlighted matches apply to [contextual add-ins](/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="df815-1123">Этот метод не поддерживается в Outlook на iOS или Android.</span><span class="sxs-lookup"><span data-stu-id="df815-1123">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="df815-p170">Метод `getSelectedRegExMatches` возвращает строки, соответствующие регулярному выражению, которое определяется в каждом элементе правила `ItemHasRegularExpressionMatch` или `ItemHasKnownEntity` в XML-файле манифеста. Для правила `ItemHasRegularExpressionMatch` соответствующую строку должно содержать свойство элемента, указанного этим правилом. Простой тип `PropertyName` определяет поддерживаемые свойства.</span><span class="sxs-lookup"><span data-stu-id="df815-p170">The `getSelectedRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="df815-1127">Например, рассмотрим манифест надстройки, который содержит указанный ниже элемент `Rule`.</span><span class="sxs-lookup"><span data-stu-id="df815-1127">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="df815-1128">Объект, возвращаемый методом `getRegExMatches`, будет содержать два свойства: `fruits` и `veggies`.</span><span class="sxs-lookup"><span data-stu-id="df815-1128">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="df815-p171">Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты. Лучше используйте метод [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.7#getasync-coerciontype--options--callback-) для этого.</span><span class="sxs-lookup"><span data-stu-id="df815-p171">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.7#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="df815-1132">Requirements</span><span class="sxs-lookup"><span data-stu-id="df815-1132">Requirements</span></span>

|<span data-ttu-id="df815-1133">Требование</span><span class="sxs-lookup"><span data-stu-id="df815-1133">Requirement</span></span>|<span data-ttu-id="df815-1134">Значение</span><span class="sxs-lookup"><span data-stu-id="df815-1134">Value</span></span>|
|---|---|
|[<span data-ttu-id="df815-1135">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="df815-1135">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="df815-1136">1.6</span><span class="sxs-lookup"><span data-stu-id="df815-1136">1.6</span></span>|
|[<span data-ttu-id="df815-1137">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="df815-1137">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="df815-1138">ReadItem</span><span class="sxs-lookup"><span data-stu-id="df815-1138">ReadItem</span></span>|
|[<span data-ttu-id="df815-1139">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="df815-1139">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="df815-1140">Чтение</span><span class="sxs-lookup"><span data-stu-id="df815-1140">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="df815-1141">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="df815-1141">Returns:</span></span>

<span data-ttu-id="df815-p172">Объект, содержащий массив строк, которые соответствуют регулярным выражениям, определяемым в XML-файле манифеста. Имя каждого массива равно соответствующему значению атрибута `RegExName` подходящего правила `ItemHasRegularExpressionMatch` или атрибута `FilterName` соответствующего правила `ItemHasKnownEntity`.</span><span class="sxs-lookup"><span data-stu-id="df815-p172">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

##### <a name="example"></a><span data-ttu-id="df815-1144">Пример</span><span class="sxs-lookup"><span data-stu-id="df815-1144">Example</span></span>

<span data-ttu-id="df815-1145">В приведенном ниже примере показано, как получить доступ к массиву совпадений с элементами `fruits` и `veggies` правил активации регулярных выражений, указанными в манифесте.</span><span class="sxs-lookup"><span data-stu-id="df815-1145">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```js
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

<br>

---
---

#### <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="df815-1146">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="df815-1146">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="df815-1147">Асинхронно загружает настраиваемые свойства для надстройки для выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="df815-1147">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="df815-p173">Настраиваемые свойства сохраняются в виде пар "ключ-значение" для каждого приложения и каждого элемента. Этот метод возвращает объект `CustomProperties` при обратном вызове, который предоставляет методы для доступа к настраиваемым свойствам, характерным для текущего элемента и текущей надстройки. Настраиваемые свойства не шифруются для элемента, поэтому этот способ хранения не является безопасным.</span><span class="sxs-lookup"><span data-stu-id="df815-p173">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="df815-1151">Параметры</span><span class="sxs-lookup"><span data-stu-id="df815-1151">Parameters</span></span>

|<span data-ttu-id="df815-1152">Имя</span><span class="sxs-lookup"><span data-stu-id="df815-1152">Name</span></span>|<span data-ttu-id="df815-1153">Тип</span><span class="sxs-lookup"><span data-stu-id="df815-1153">Type</span></span>|<span data-ttu-id="df815-1154">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="df815-1154">Attributes</span></span>|<span data-ttu-id="df815-1155">Описание</span><span class="sxs-lookup"><span data-stu-id="df815-1155">Description</span></span>|
|---|---|---|---|
|`callback`|<span data-ttu-id="df815-1156">function</span><span class="sxs-lookup"><span data-stu-id="df815-1156">function</span></span>||<span data-ttu-id="df815-1157">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="df815-1157">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="df815-1158">Настраиваемые свойства предоставляются в виде объекта [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.7) в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="df815-1158">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.7) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="df815-1159">Этот объект позволяет получить, задать и удалить настраиваемые свойства для элемента, а также сохранить изменения, внесенные в набор настраиваемых свойств, на сервере.</span><span class="sxs-lookup"><span data-stu-id="df815-1159">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`|<span data-ttu-id="df815-1160">Объект</span><span class="sxs-lookup"><span data-stu-id="df815-1160">Object</span></span>|<span data-ttu-id="df815-1161">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="df815-1161">&lt;optional&gt;</span></span>|<span data-ttu-id="df815-1162">Разработчики могут указать любой объект, к которому необходимо получить доступ, в функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="df815-1162">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="df815-1163">Доступ к этому объекту можно получить с помощью свойства `asyncResult.asyncContext` в функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="df815-1163">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="df815-1164">Требования</span><span class="sxs-lookup"><span data-stu-id="df815-1164">Requirements</span></span>

|<span data-ttu-id="df815-1165">Требование</span><span class="sxs-lookup"><span data-stu-id="df815-1165">Requirement</span></span>|<span data-ttu-id="df815-1166">Значение</span><span class="sxs-lookup"><span data-stu-id="df815-1166">Value</span></span>|
|---|---|
|[<span data-ttu-id="df815-1167">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="df815-1167">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="df815-1168">1.0</span><span class="sxs-lookup"><span data-stu-id="df815-1168">1.0</span></span>|
|[<span data-ttu-id="df815-1169">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="df815-1169">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="df815-1170">ReadItem</span><span class="sxs-lookup"><span data-stu-id="df815-1170">ReadItem</span></span>|
|[<span data-ttu-id="df815-1171">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="df815-1171">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="df815-1172">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="df815-1172">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="df815-1173">Пример</span><span class="sxs-lookup"><span data-stu-id="df815-1173">Example</span></span>

<span data-ttu-id="df815-p176">Приведенный ниже пример кода показывает, как асинхронно загружать настраиваемые свойства, характерные для текущего элемента, с помощью метода `loadCustomPropertiesAsync`. Этот пример также показывает, как сохранять эти свойства на сервере с помощью метода `CustomProperties.saveAsync`. После загрузки настраиваемых свойств в этом примере кода метод `CustomProperties.get` используется для считывания настраиваемого свойства `myProp`, метод `CustomProperties.set` — для записи настраиваемого свойства `otherProp`, а метод `saveAsync` — для сохранения настраиваемых свойств.</span><span class="sxs-lookup"><span data-stu-id="df815-p176">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

#### <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="df815-1177">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="df815-1177">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="df815-1178">Удаляет вложение из сообщения или встречи.</span><span class="sxs-lookup"><span data-stu-id="df815-1178">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="df815-1179">Метод `removeAttachmentAsync` удаляет из элемента вложение с указанным идентификатором.</span><span class="sxs-lookup"><span data-stu-id="df815-1179">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="df815-1180">Идентификатор вложения рекомендуется использовать для удаления вложения, только если оно добавлено тем же почтовым приложением в ходе текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="df815-1180">As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session.</span></span> <span data-ttu-id="df815-1181">В Outlook в Интернете и мобильных устройствах идентификатор вложения действителен только в рамках одного сеанса.</span><span class="sxs-lookup"><span data-stu-id="df815-1181">In Outlook on the web and mobile devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="df815-1182">Сеанс завершается, когда пользователь закрывает приложение или начинает создавать элемент во встроенной форме, а затем переходит из формы в отдельное окно.</span><span class="sxs-lookup"><span data-stu-id="df815-1182">A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="df815-1183">Параметры</span><span class="sxs-lookup"><span data-stu-id="df815-1183">Parameters</span></span>

|<span data-ttu-id="df815-1184">Имя</span><span class="sxs-lookup"><span data-stu-id="df815-1184">Name</span></span>|<span data-ttu-id="df815-1185">Тип</span><span class="sxs-lookup"><span data-stu-id="df815-1185">Type</span></span>|<span data-ttu-id="df815-1186">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="df815-1186">Attributes</span></span>|<span data-ttu-id="df815-1187">Описание</span><span class="sxs-lookup"><span data-stu-id="df815-1187">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="df815-1188">String</span><span class="sxs-lookup"><span data-stu-id="df815-1188">String</span></span>||<span data-ttu-id="df815-1189">Идентификатор удаляемого вложения.</span><span class="sxs-lookup"><span data-stu-id="df815-1189">The identifier of the attachment to remove.</span></span>|
|`options`|<span data-ttu-id="df815-1190">Объект</span><span class="sxs-lookup"><span data-stu-id="df815-1190">Object</span></span>|<span data-ttu-id="df815-1191">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="df815-1191">&lt;optional&gt;</span></span>|<span data-ttu-id="df815-1192">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="df815-1192">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="df815-1193">Объект</span><span class="sxs-lookup"><span data-stu-id="df815-1193">Object</span></span>|<span data-ttu-id="df815-1194">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="df815-1194">&lt;optional&gt;</span></span>|<span data-ttu-id="df815-1195">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="df815-1195">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="df815-1196">функция</span><span class="sxs-lookup"><span data-stu-id="df815-1196">function</span></span>|<span data-ttu-id="df815-1197">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="df815-1197">&lt;optional&gt;</span></span>|<span data-ttu-id="df815-1198">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="df815-1198">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="df815-1199">Если удалить вложение не удается, свойство `asyncResult.error` содержит код ошибки с указанием ее причины.</span><span class="sxs-lookup"><span data-stu-id="df815-1199">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="df815-1200">Ошибки</span><span class="sxs-lookup"><span data-stu-id="df815-1200">Errors</span></span>

|<span data-ttu-id="df815-1201">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="df815-1201">Error code</span></span>|<span data-ttu-id="df815-1202">Описание</span><span class="sxs-lookup"><span data-stu-id="df815-1202">Description</span></span>|
|------------|-------------|
|`InvalidAttachmentId`|<span data-ttu-id="df815-1203">Идентификатор вложения не существует.</span><span class="sxs-lookup"><span data-stu-id="df815-1203">The attachment identifier does not exist.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="df815-1204">Требования</span><span class="sxs-lookup"><span data-stu-id="df815-1204">Requirements</span></span>

|<span data-ttu-id="df815-1205">Требование</span><span class="sxs-lookup"><span data-stu-id="df815-1205">Requirement</span></span>|<span data-ttu-id="df815-1206">Значение</span><span class="sxs-lookup"><span data-stu-id="df815-1206">Value</span></span>|
|---|---|
|[<span data-ttu-id="df815-1207">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="df815-1207">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="df815-1208">1.1</span><span class="sxs-lookup"><span data-stu-id="df815-1208">1.1</span></span>|
|[<span data-ttu-id="df815-1209">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="df815-1209">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="df815-1210">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="df815-1210">ReadWriteItem</span></span>|
|[<span data-ttu-id="df815-1211">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="df815-1211">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="df815-1212">Создание</span><span class="sxs-lookup"><span data-stu-id="df815-1212">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="df815-1213">Пример</span><span class="sxs-lookup"><span data-stu-id="df815-1213">Example</span></span>

<span data-ttu-id="df815-1214">Указанный ниже код удаляет вложение с идентификатором "0".</span><span class="sxs-lookup"><span data-stu-id="df815-1214">The following code removes an attachment with an identifier of '0'.</span></span>

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

#### <a name="removehandlerasynceventtype-options-callback"></a><span data-ttu-id="df815-1215">removeHandlerAsync(eventType, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="df815-1215">removeHandlerAsync(eventType, [options], [callback])</span></span>

<span data-ttu-id="df815-1216">Удаляет обработчиков для поддерживаемого типа события.</span><span class="sxs-lookup"><span data-stu-id="df815-1216">Removes the event handlers for a supported event type.</span></span>

<span data-ttu-id="df815-1217">В настоящее время поддерживаются типы `Office.EventType.AppointmentTimeChanged`событий `Office.EventType.RecipientsChanged`, и`Office.EventType.RecurrenceChanged`</span><span class="sxs-lookup"><span data-stu-id="df815-1217">Currently the supported event types are `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`</span></span>

##### <a name="parameters"></a><span data-ttu-id="df815-1218">Параметры</span><span class="sxs-lookup"><span data-stu-id="df815-1218">Parameters</span></span>

| <span data-ttu-id="df815-1219">Имя</span><span class="sxs-lookup"><span data-stu-id="df815-1219">Name</span></span> | <span data-ttu-id="df815-1220">Тип</span><span class="sxs-lookup"><span data-stu-id="df815-1220">Type</span></span> | <span data-ttu-id="df815-1221">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="df815-1221">Attributes</span></span> | <span data-ttu-id="df815-1222">Описание</span><span class="sxs-lookup"><span data-stu-id="df815-1222">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="df815-1223">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="df815-1223">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="df815-1224">Событие, которое должно вызвать обработчик.</span><span class="sxs-lookup"><span data-stu-id="df815-1224">The event that should invoke the handler.</span></span> |
| `options` | <span data-ttu-id="df815-1225">Объект</span><span class="sxs-lookup"><span data-stu-id="df815-1225">Object</span></span> | <span data-ttu-id="df815-1226">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="df815-1226">&lt;optional&gt;</span></span> | <span data-ttu-id="df815-1227">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="df815-1227">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="df815-1228">Объект</span><span class="sxs-lookup"><span data-stu-id="df815-1228">Object</span></span> | <span data-ttu-id="df815-1229">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="df815-1229">&lt;optional&gt;</span></span> | <span data-ttu-id="df815-1230">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="df815-1230">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="df815-1231">функция</span><span class="sxs-lookup"><span data-stu-id="df815-1231">function</span></span>| <span data-ttu-id="df815-1232">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="df815-1232">&lt;optional&gt;</span></span>|<span data-ttu-id="df815-1233">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="df815-1233">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="df815-1234">Требования</span><span class="sxs-lookup"><span data-stu-id="df815-1234">Requirements</span></span>

|<span data-ttu-id="df815-1235">Требование</span><span class="sxs-lookup"><span data-stu-id="df815-1235">Requirement</span></span>| <span data-ttu-id="df815-1236">Значение</span><span class="sxs-lookup"><span data-stu-id="df815-1236">Value</span></span>|
|---|---|
|[<span data-ttu-id="df815-1237">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="df815-1237">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="df815-1238">1.7</span><span class="sxs-lookup"><span data-stu-id="df815-1238">1.7</span></span> |
|[<span data-ttu-id="df815-1239">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="df815-1239">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="df815-1240">ReadItem</span><span class="sxs-lookup"><span data-stu-id="df815-1240">ReadItem</span></span> |
|[<span data-ttu-id="df815-1241">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="df815-1241">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="df815-1242">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="df815-1242">Compose or Read</span></span> |

##### <a name="example"></a><span data-ttu-id="df815-1243">Пример</span><span class="sxs-lookup"><span data-stu-id="df815-1243">Example</span></span>

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

#### <a name="saveasyncoptions-callback"></a><span data-ttu-id="df815-1244">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="df815-1244">saveAsync([options], callback)</span></span>

<span data-ttu-id="df815-1245">Асинхронно сохраняет элемент.</span><span class="sxs-lookup"><span data-stu-id="df815-1245">Asynchronously saves an item.</span></span>

<span data-ttu-id="df815-1246">При вызове этот метод сохраняет текущее сообщение в виде черновика и возвращает идентификатор элемента с помощью метода обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="df815-1246">When invoked, this method saves the current message as a draft and returns the item id via the callback method.</span></span> <span data-ttu-id="df815-1247">В Outlook в Интернете или Outlook в интерактивном режиме элемент сохраняется на сервере.</span><span class="sxs-lookup"><span data-stu-id="df815-1247">In Outlook on the web or Outlook in online mode, the item is saved to the server.</span></span> <span data-ttu-id="df815-1248">В Outlook в режиме кэширования этот элемент сохраняется в локальном кэше.</span><span class="sxs-lookup"><span data-stu-id="df815-1248">In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="df815-1249">Если в режиме создания надстройка вызывает для элемента метод `saveAsync`, чтобы получить параметр `itemId` для использования с EWS или REST API, необходимо помнить, что синхронизация элемента с сервером может занять много времени, если Outlook работает в режиме кэширования данных.</span><span class="sxs-lookup"><span data-stu-id="df815-1249">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="df815-1250">До окончания синхронизации элемента применение параметра `itemId` будет приводить к ошибке.</span><span class="sxs-lookup"><span data-stu-id="df815-1250">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="df815-p180">Если метод `saveAsync` вызывается для встречи в режиме создания, она сохраняется как обычная встреча в календаре пользователя, а не как черновик. При сохранении новой встречи приглашения не отправляются. При сохранении существующей встречи уведомления отправляются добавленным или удаленным участникам.</span><span class="sxs-lookup"><span data-stu-id="df815-p180">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="df815-1254">Следующие клиенты отличаются другим поведением `saveAsync` в отношении встреч в режиме создания:</span><span class="sxs-lookup"><span data-stu-id="df815-1254">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="df815-1255">Outlook в Mac не поддерживает сохранение собраний.</span><span class="sxs-lookup"><span data-stu-id="df815-1255">Outlook on Mac does not support saving a meeting.</span></span> <span data-ttu-id="df815-1256">`saveAsync` Метод завершается с ошибкой при вызове из собрания в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="df815-1256">The `saveAsync` method fails when called from a meeting in compose mode.</span></span> <span data-ttu-id="df815-1257">Просмотреть [не удается сохранить собрание в виде черновика в Outlook для Mac с помощью API Office JS](https://support.microsoft.com/help/4505745) для обхода.</span><span class="sxs-lookup"><span data-stu-id="df815-1257">See [Cannot save a meeting as a draft in Outlook for Mac by using Office JS API](https://support.microsoft.com/help/4505745) for a workaround.</span></span>
> - <span data-ttu-id="df815-1258">Outlook в Интернете всегда отправляет приглашение или обновление при вызове `saveAsync` для встречи в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="df815-1258">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="df815-1259">Параметры</span><span class="sxs-lookup"><span data-stu-id="df815-1259">Parameters</span></span>

|<span data-ttu-id="df815-1260">Имя</span><span class="sxs-lookup"><span data-stu-id="df815-1260">Name</span></span>|<span data-ttu-id="df815-1261">Тип</span><span class="sxs-lookup"><span data-stu-id="df815-1261">Type</span></span>|<span data-ttu-id="df815-1262">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="df815-1262">Attributes</span></span>|<span data-ttu-id="df815-1263">Описание</span><span class="sxs-lookup"><span data-stu-id="df815-1263">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="df815-1264">Объект</span><span class="sxs-lookup"><span data-stu-id="df815-1264">Object</span></span>|<span data-ttu-id="df815-1265">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="df815-1265">&lt;optional&gt;</span></span>|<span data-ttu-id="df815-1266">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="df815-1266">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="df815-1267">Объект</span><span class="sxs-lookup"><span data-stu-id="df815-1267">Object</span></span>|<span data-ttu-id="df815-1268">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="df815-1268">&lt;optional&gt;</span></span>|<span data-ttu-id="df815-1269">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="df815-1269">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="df815-1270">функция</span><span class="sxs-lookup"><span data-stu-id="df815-1270">function</span></span>||<span data-ttu-id="df815-1271">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="df815-1271">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="df815-1272">После успешного выполнения идентификатор элемента будет предоставлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="df815-1272">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="df815-1273">Требования</span><span class="sxs-lookup"><span data-stu-id="df815-1273">Requirements</span></span>

|<span data-ttu-id="df815-1274">Требование</span><span class="sxs-lookup"><span data-stu-id="df815-1274">Requirement</span></span>|<span data-ttu-id="df815-1275">Значение</span><span class="sxs-lookup"><span data-stu-id="df815-1275">Value</span></span>|
|---|---|
|[<span data-ttu-id="df815-1276">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="df815-1276">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="df815-1277">1.3</span><span class="sxs-lookup"><span data-stu-id="df815-1277">1.3</span></span>|
|[<span data-ttu-id="df815-1278">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="df815-1278">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="df815-1279">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="df815-1279">ReadWriteItem</span></span>|
|[<span data-ttu-id="df815-1280">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="df815-1280">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="df815-1281">Создание</span><span class="sxs-lookup"><span data-stu-id="df815-1281">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="df815-1282">Примеры</span><span class="sxs-lookup"><span data-stu-id="df815-1282">Examples</span></span>

```js
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result.
  });
```

<span data-ttu-id="df815-p182">Ниже приведен пример параметра `result`, переданного функции обратного вызова. Свойство `value` содержит идентификатор элемента.</span><span class="sxs-lookup"><span data-stu-id="df815-p182">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

<br>

---
---

#### <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="df815-1285">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="df815-1285">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="df815-1286">Асинхронно вставляет данные в текст или тему сообщения.</span><span class="sxs-lookup"><span data-stu-id="df815-1286">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="df815-p183">Метод `setSelectedDataAsync` вставляет указанную строку в местоположение курсора в теме или тексте элемента либо, если текст выделен в редакторе, он заменяет выделенный текст. Если курсор находится вне текста или темы элемента, возвращается ошибка. После вставки курсор помещается в конец вставленного содержимого.</span><span class="sxs-lookup"><span data-stu-id="df815-p183">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="df815-1290">Параметры</span><span class="sxs-lookup"><span data-stu-id="df815-1290">Parameters</span></span>

|<span data-ttu-id="df815-1291">Имя</span><span class="sxs-lookup"><span data-stu-id="df815-1291">Name</span></span>|<span data-ttu-id="df815-1292">Тип</span><span class="sxs-lookup"><span data-stu-id="df815-1292">Type</span></span>|<span data-ttu-id="df815-1293">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="df815-1293">Attributes</span></span>|<span data-ttu-id="df815-1294">Описание</span><span class="sxs-lookup"><span data-stu-id="df815-1294">Description</span></span>|
|---|---|---|---|
|`data`|<span data-ttu-id="df815-1295">String</span><span class="sxs-lookup"><span data-stu-id="df815-1295">String</span></span>||<span data-ttu-id="df815-p184">Вставляемые данные. Объем данных не должен превышать 1 000 000 символов. Если передано больше 1 000 000 символов, возвращается исключение `ArgumentOutOfRange`.</span><span class="sxs-lookup"><span data-stu-id="df815-p184">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`|<span data-ttu-id="df815-1299">Object</span><span class="sxs-lookup"><span data-stu-id="df815-1299">Object</span></span>|<span data-ttu-id="df815-1300">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="df815-1300">&lt;optional&gt;</span></span>|<span data-ttu-id="df815-1301">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="df815-1301">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="df815-1302">Object</span><span class="sxs-lookup"><span data-stu-id="df815-1302">Object</span></span>|<span data-ttu-id="df815-1303">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="df815-1303">&lt;optional&gt;</span></span>|<span data-ttu-id="df815-1304">В методе обратного вызова разработчики могут указать любой объект, к которому необходимо получить доступ.</span><span class="sxs-lookup"><span data-stu-id="df815-1304">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="df815-1305">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="df815-1305">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="df815-1306">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="df815-1306">&lt;optional&gt;</span></span>|<span data-ttu-id="df815-1307">Если `text`текущий стиль применяется в Outlook для веб-клиентов и клиентов для настольных ПК.</span><span class="sxs-lookup"><span data-stu-id="df815-1307">If `text`, the current style is applied in Outlook on the web and desktop clients.</span></span> <span data-ttu-id="df815-1308">Если поле представляет собой редактор HTML, вставляются только текстовые данные, даже если они имеют формат HTML.</span><span class="sxs-lookup"><span data-stu-id="df815-1308">If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="df815-1309">Если `html` и поле поддерживает HTML (тема не используется), текущий стиль применяется в Outlook в Интернете, а в настольных клиентах Outlook применяется стиль по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="df815-1309">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook on the web and the default style is applied in Outlook desktop clients.</span></span> <span data-ttu-id="df815-1310">Если поле является текстовым, возвращается ошибка `InvalidDataFormat`.</span><span class="sxs-lookup"><span data-stu-id="df815-1310">If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="df815-1311">Если свойство `coercionType` не задано, результат зависит от поля: если поле имеет формат HTML, используется текст в формате HTML, а если поле текстовое, применяется обычный текст.</span><span class="sxs-lookup"><span data-stu-id="df815-1311">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`|<span data-ttu-id="df815-1312">функция</span><span class="sxs-lookup"><span data-stu-id="df815-1312">function</span></span>||<span data-ttu-id="df815-1313">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="df815-1313">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="df815-1314">Требования</span><span class="sxs-lookup"><span data-stu-id="df815-1314">Requirements</span></span>

|<span data-ttu-id="df815-1315">Требование</span><span class="sxs-lookup"><span data-stu-id="df815-1315">Requirement</span></span>|<span data-ttu-id="df815-1316">Значение</span><span class="sxs-lookup"><span data-stu-id="df815-1316">Value</span></span>|
|---|---|
|[<span data-ttu-id="df815-1317">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="df815-1317">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="df815-1318">1.2</span><span class="sxs-lookup"><span data-stu-id="df815-1318">1.2</span></span>|
|[<span data-ttu-id="df815-1319">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="df815-1319">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="df815-1320">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="df815-1320">ReadWriteItem</span></span>|
|[<span data-ttu-id="df815-1321">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="df815-1321">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="df815-1322">Создание</span><span class="sxs-lookup"><span data-stu-id="df815-1322">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="df815-1323">Пример</span><span class="sxs-lookup"><span data-stu-id="df815-1323">Example</span></span>

```js
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
