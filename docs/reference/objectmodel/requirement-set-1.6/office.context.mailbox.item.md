---
title: Office. Context. Mailbox. Item — набор требований 1,6
description: ''
ms.date: 04/12/2019
localization_priority: Normal
ms.openlocfilehash: cc7897f791c5a07ed5c17a686b6601a1a7633f00
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/24/2019
ms.locfileid: "32451775"
---
# <a name="item"></a><span data-ttu-id="9af0b-102">item</span><span class="sxs-lookup"><span data-stu-id="9af0b-102">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="9af0b-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="9af0b-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="9af0b-p101">Пространство имен `item` используется для доступа к выбранному в данный момент сообщению, приглашению на собрание или описанию встречи. Вы можете определить тип пространства имен `item` с помощью свойства [itemType](#itemtype-officemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="9af0b-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="9af0b-106">Requirements</span><span class="sxs-lookup"><span data-stu-id="9af0b-106">Requirements</span></span>

|<span data-ttu-id="9af0b-107">Требование</span><span class="sxs-lookup"><span data-stu-id="9af0b-107">Requirement</span></span>| <span data-ttu-id="9af0b-108">Значение</span><span class="sxs-lookup"><span data-stu-id="9af0b-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="9af0b-109">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="9af0b-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9af0b-110">1.0</span><span class="sxs-lookup"><span data-stu-id="9af0b-110">1.0</span></span>|
|[<span data-ttu-id="9af0b-111">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="9af0b-111">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9af0b-112">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="9af0b-112">Restricted</span></span>|
|[<span data-ttu-id="9af0b-113">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9af0b-113">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9af0b-114">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="9af0b-114">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="9af0b-115">Элементы и методы</span><span class="sxs-lookup"><span data-stu-id="9af0b-115">Members and methods</span></span>

| <span data-ttu-id="9af0b-116">Элемент</span><span class="sxs-lookup"><span data-stu-id="9af0b-116">Member</span></span> | <span data-ttu-id="9af0b-117">Тип</span><span class="sxs-lookup"><span data-stu-id="9af0b-117">Type</span></span> |
|--------|------|
| [<span data-ttu-id="9af0b-118">attachments</span><span class="sxs-lookup"><span data-stu-id="9af0b-118">attachments</span></span>](#attachments-arrayattachmentdetails) | <span data-ttu-id="9af0b-119">Элемент</span><span class="sxs-lookup"><span data-stu-id="9af0b-119">Member</span></span> |
| [<span data-ttu-id="9af0b-120">bcc</span><span class="sxs-lookup"><span data-stu-id="9af0b-120">bcc</span></span>](#bcc-recipients) | <span data-ttu-id="9af0b-121">Элемент</span><span class="sxs-lookup"><span data-stu-id="9af0b-121">Member</span></span> |
| [<span data-ttu-id="9af0b-122">body</span><span class="sxs-lookup"><span data-stu-id="9af0b-122">body</span></span>](#body-body) | <span data-ttu-id="9af0b-123">Элемент</span><span class="sxs-lookup"><span data-stu-id="9af0b-123">Member</span></span> |
| [<span data-ttu-id="9af0b-124">cc</span><span class="sxs-lookup"><span data-stu-id="9af0b-124">cc</span></span>](#cc-arrayemailaddressdetailsrecipients) | <span data-ttu-id="9af0b-125">Элемент</span><span class="sxs-lookup"><span data-stu-id="9af0b-125">Member</span></span> |
| [<span data-ttu-id="9af0b-126">conversationId</span><span class="sxs-lookup"><span data-stu-id="9af0b-126">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="9af0b-127">Элемент</span><span class="sxs-lookup"><span data-stu-id="9af0b-127">Member</span></span> |
| [<span data-ttu-id="9af0b-128">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="9af0b-128">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="9af0b-129">Элемент</span><span class="sxs-lookup"><span data-stu-id="9af0b-129">Member</span></span> |
| [<span data-ttu-id="9af0b-130">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="9af0b-130">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="9af0b-131">Элемент</span><span class="sxs-lookup"><span data-stu-id="9af0b-131">Member</span></span> |
| [<span data-ttu-id="9af0b-132">end</span><span class="sxs-lookup"><span data-stu-id="9af0b-132">end</span></span>](#end-datetime) | <span data-ttu-id="9af0b-133">Элемент</span><span class="sxs-lookup"><span data-stu-id="9af0b-133">Member</span></span> |
| [<span data-ttu-id="9af0b-134">from</span><span class="sxs-lookup"><span data-stu-id="9af0b-134">from</span></span>](#from-emailaddressdetails) | <span data-ttu-id="9af0b-135">Элемент</span><span class="sxs-lookup"><span data-stu-id="9af0b-135">Member</span></span> |
| [<span data-ttu-id="9af0b-136">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="9af0b-136">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="9af0b-137">Элемент</span><span class="sxs-lookup"><span data-stu-id="9af0b-137">Member</span></span> |
| [<span data-ttu-id="9af0b-138">itemClass</span><span class="sxs-lookup"><span data-stu-id="9af0b-138">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="9af0b-139">Элемент</span><span class="sxs-lookup"><span data-stu-id="9af0b-139">Member</span></span> |
| [<span data-ttu-id="9af0b-140">itemId</span><span class="sxs-lookup"><span data-stu-id="9af0b-140">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="9af0b-141">Элемент</span><span class="sxs-lookup"><span data-stu-id="9af0b-141">Member</span></span> |
| [<span data-ttu-id="9af0b-142">itemType</span><span class="sxs-lookup"><span data-stu-id="9af0b-142">itemType</span></span>](#itemtype-officemailboxenumsitemtype) | <span data-ttu-id="9af0b-143">Элемент</span><span class="sxs-lookup"><span data-stu-id="9af0b-143">Member</span></span> |
| [<span data-ttu-id="9af0b-144">location</span><span class="sxs-lookup"><span data-stu-id="9af0b-144">location</span></span>](#location-stringlocation) | <span data-ttu-id="9af0b-145">Элемент</span><span class="sxs-lookup"><span data-stu-id="9af0b-145">Member</span></span> |
| [<span data-ttu-id="9af0b-146">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="9af0b-146">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="9af0b-147">Элемент</span><span class="sxs-lookup"><span data-stu-id="9af0b-147">Member</span></span> |
| [<span data-ttu-id="9af0b-148">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="9af0b-148">notificationMessages</span></span>](#notificationmessages-notificationmessages) | <span data-ttu-id="9af0b-149">Элемент</span><span class="sxs-lookup"><span data-stu-id="9af0b-149">Member</span></span> |
| [<span data-ttu-id="9af0b-150">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="9af0b-150">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="9af0b-151">Элемент</span><span class="sxs-lookup"><span data-stu-id="9af0b-151">Member</span></span> |
| [<span data-ttu-id="9af0b-152">organizer</span><span class="sxs-lookup"><span data-stu-id="9af0b-152">organizer</span></span>](#organizer-emailaddressdetails) | <span data-ttu-id="9af0b-153">Элемент</span><span class="sxs-lookup"><span data-stu-id="9af0b-153">Member</span></span> |
| [<span data-ttu-id="9af0b-154">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="9af0b-154">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="9af0b-155">Member</span><span class="sxs-lookup"><span data-stu-id="9af0b-155">Member</span></span> |
| [<span data-ttu-id="9af0b-156">sender</span><span class="sxs-lookup"><span data-stu-id="9af0b-156">sender</span></span>](#sender-emailaddressdetails) | <span data-ttu-id="9af0b-157">Элемент</span><span class="sxs-lookup"><span data-stu-id="9af0b-157">Member</span></span> |
| [<span data-ttu-id="9af0b-158">start</span><span class="sxs-lookup"><span data-stu-id="9af0b-158">start</span></span>](#start-datetime) | <span data-ttu-id="9af0b-159">Элемент</span><span class="sxs-lookup"><span data-stu-id="9af0b-159">Member</span></span> |
| [<span data-ttu-id="9af0b-160">subject</span><span class="sxs-lookup"><span data-stu-id="9af0b-160">subject</span></span>](#subject-stringsubject) | <span data-ttu-id="9af0b-161">Элемент</span><span class="sxs-lookup"><span data-stu-id="9af0b-161">Member</span></span> |
| [<span data-ttu-id="9af0b-162">to</span><span class="sxs-lookup"><span data-stu-id="9af0b-162">to</span></span>](#to-arrayemailaddressdetailsrecipients) | <span data-ttu-id="9af0b-163">Элемент</span><span class="sxs-lookup"><span data-stu-id="9af0b-163">Member</span></span> |
| [<span data-ttu-id="9af0b-164">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="9af0b-164">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="9af0b-165">Метод</span><span class="sxs-lookup"><span data-stu-id="9af0b-165">Method</span></span> |
| [<span data-ttu-id="9af0b-166">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="9af0b-166">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="9af0b-167">Метод</span><span class="sxs-lookup"><span data-stu-id="9af0b-167">Method</span></span> |
| [<span data-ttu-id="9af0b-168">close</span><span class="sxs-lookup"><span data-stu-id="9af0b-168">close</span></span>](#close) | <span data-ttu-id="9af0b-169">Метод</span><span class="sxs-lookup"><span data-stu-id="9af0b-169">Method</span></span> |
| [<span data-ttu-id="9af0b-170">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="9af0b-170">displayReplyAllForm</span></span>](#displayreplyallformformdata-callback) | <span data-ttu-id="9af0b-171">Метод</span><span class="sxs-lookup"><span data-stu-id="9af0b-171">Method</span></span> |
| [<span data-ttu-id="9af0b-172">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="9af0b-172">displayReplyForm</span></span>](#displayreplyformformdata-callback) | <span data-ttu-id="9af0b-173">Метод</span><span class="sxs-lookup"><span data-stu-id="9af0b-173">Method</span></span> |
| [<span data-ttu-id="9af0b-174">getEntities</span><span class="sxs-lookup"><span data-stu-id="9af0b-174">getEntities</span></span>](#getentities--entities) | <span data-ttu-id="9af0b-175">Метод</span><span class="sxs-lookup"><span data-stu-id="9af0b-175">Method</span></span> |
| [<span data-ttu-id="9af0b-176">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="9af0b-176">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="9af0b-177">Метод</span><span class="sxs-lookup"><span data-stu-id="9af0b-177">Method</span></span> |
| [<span data-ttu-id="9af0b-178">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="9af0b-178">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="9af0b-179">Метод</span><span class="sxs-lookup"><span data-stu-id="9af0b-179">Method</span></span> |
| [<span data-ttu-id="9af0b-180">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="9af0b-180">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="9af0b-181">Метод</span><span class="sxs-lookup"><span data-stu-id="9af0b-181">Method</span></span> |
| [<span data-ttu-id="9af0b-182">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="9af0b-182">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="9af0b-183">Метод</span><span class="sxs-lookup"><span data-stu-id="9af0b-183">Method</span></span> |
| [<span data-ttu-id="9af0b-184">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="9af0b-184">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="9af0b-185">Метод</span><span class="sxs-lookup"><span data-stu-id="9af0b-185">Method</span></span> |
| [<span data-ttu-id="9af0b-186">Жетселектедентитиес</span><span class="sxs-lookup"><span data-stu-id="9af0b-186">getSelectedEntities</span></span>](#getselectedentities--entities) | <span data-ttu-id="9af0b-187">Метод</span><span class="sxs-lookup"><span data-stu-id="9af0b-187">Method</span></span> |
| [<span data-ttu-id="9af0b-188">Жетселектедрежексматчес</span><span class="sxs-lookup"><span data-stu-id="9af0b-188">getSelectedRegExMatches</span></span>](#getselectedregexmatches--object) | <span data-ttu-id="9af0b-189">Метод</span><span class="sxs-lookup"><span data-stu-id="9af0b-189">Method</span></span> |
| [<span data-ttu-id="9af0b-190">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="9af0b-190">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="9af0b-191">Метод</span><span class="sxs-lookup"><span data-stu-id="9af0b-191">Method</span></span> |
| [<span data-ttu-id="9af0b-192">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="9af0b-192">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="9af0b-193">Метод</span><span class="sxs-lookup"><span data-stu-id="9af0b-193">Method</span></span> |
| [<span data-ttu-id="9af0b-194">saveAsync</span><span class="sxs-lookup"><span data-stu-id="9af0b-194">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="9af0b-195">Метод</span><span class="sxs-lookup"><span data-stu-id="9af0b-195">Method</span></span> |
| [<span data-ttu-id="9af0b-196">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="9af0b-196">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="9af0b-197">Метод</span><span class="sxs-lookup"><span data-stu-id="9af0b-197">Method</span></span> |

### <a name="example"></a><span data-ttu-id="9af0b-198">Пример</span><span class="sxs-lookup"><span data-stu-id="9af0b-198">Example</span></span>

<span data-ttu-id="9af0b-199">В примере кода JavaScript, приведенном ниже, показано, как получить доступ к свойству `subject` текущего элемента в Outlook.</span><span class="sxs-lookup"><span data-stu-id="9af0b-199">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="9af0b-200">Элементы</span><span class="sxs-lookup"><span data-stu-id="9af0b-200">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook16officeattachmentdetails"></a><span data-ttu-id="9af0b-201">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_6/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="9af0b-201">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_6/office.attachmentdetails)></span></span>

<span data-ttu-id="9af0b-p102">Получает массив вложений для элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="9af0b-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="9af0b-204">Outlook блокирует определенные типы файлов из-за потенциальных проблем с безопасностью, поэтому они не возвращаются.</span><span class="sxs-lookup"><span data-stu-id="9af0b-204">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="9af0b-205">Дополнительные сведения см. в статье [Блокировка вложений в Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="9af0b-205">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="9af0b-206">Type</span><span class="sxs-lookup"><span data-stu-id="9af0b-206">Type</span></span>

*   <span data-ttu-id="9af0b-207">Array.<[AttachmentDetails](/javascript/api/outlook_1_6/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="9af0b-207">Array.<[AttachmentDetails](/javascript/api/outlook_1_6/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="9af0b-208">Требования</span><span class="sxs-lookup"><span data-stu-id="9af0b-208">Requirements</span></span>

|<span data-ttu-id="9af0b-209">Требование</span><span class="sxs-lookup"><span data-stu-id="9af0b-209">Requirement</span></span>| <span data-ttu-id="9af0b-210">Значение</span><span class="sxs-lookup"><span data-stu-id="9af0b-210">Value</span></span>|
|---|---|
|[<span data-ttu-id="9af0b-211">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="9af0b-211">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9af0b-212">1.0</span><span class="sxs-lookup"><span data-stu-id="9af0b-212">1.0</span></span>|
|[<span data-ttu-id="9af0b-213">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="9af0b-213">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9af0b-214">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9af0b-214">ReadItem</span></span>|
|[<span data-ttu-id="9af0b-215">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9af0b-215">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9af0b-216">Чтение</span><span class="sxs-lookup"><span data-stu-id="9af0b-216">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9af0b-217">Пример</span><span class="sxs-lookup"><span data-stu-id="9af0b-217">Example</span></span>

<span data-ttu-id="9af0b-218">С помощью приведенного ниже кода можно создать HTML-строку с подробными сведениями обо всех вложениях для текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="9af0b-218">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

####  <a name="bcc-recipientsjavascriptapioutlook16officerecipients"></a><span data-ttu-id="9af0b-219">bcc :[Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="9af0b-219">bcc :[Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

<span data-ttu-id="9af0b-220">Получает объект, который предоставляет методы для получения или обновления получателей скрытой копии сообщения.</span><span class="sxs-lookup"><span data-stu-id="9af0b-220">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="9af0b-221">Только в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="9af0b-221">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="9af0b-222">Type</span><span class="sxs-lookup"><span data-stu-id="9af0b-222">Type</span></span>

*   [<span data-ttu-id="9af0b-223">Получатели</span><span class="sxs-lookup"><span data-stu-id="9af0b-223">Recipients</span></span>](/javascript/api/outlook_1_6/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="9af0b-224">Требования</span><span class="sxs-lookup"><span data-stu-id="9af0b-224">Requirements</span></span>

|<span data-ttu-id="9af0b-225">Требование</span><span class="sxs-lookup"><span data-stu-id="9af0b-225">Requirement</span></span>| <span data-ttu-id="9af0b-226">Значение</span><span class="sxs-lookup"><span data-stu-id="9af0b-226">Value</span></span>|
|---|---|
|[<span data-ttu-id="9af0b-227">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="9af0b-227">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9af0b-228">1.1</span><span class="sxs-lookup"><span data-stu-id="9af0b-228">1.1</span></span>|
|[<span data-ttu-id="9af0b-229">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="9af0b-229">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9af0b-230">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9af0b-230">ReadItem</span></span>|
|[<span data-ttu-id="9af0b-231">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9af0b-231">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9af0b-232">Создание</span><span class="sxs-lookup"><span data-stu-id="9af0b-232">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="9af0b-233">Пример</span><span class="sxs-lookup"><span data-stu-id="9af0b-233">Example</span></span>

```javascript
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlook16officebody"></a><span data-ttu-id="9af0b-234">body :[Body](/javascript/api/outlook_1_6/office.body)</span><span class="sxs-lookup"><span data-stu-id="9af0b-234">body :[Body](/javascript/api/outlook_1_6/office.body)</span></span>

<span data-ttu-id="9af0b-235">Получает объект, предоставляющий методы для работы с основным текстом элемента.</span><span class="sxs-lookup"><span data-stu-id="9af0b-235">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="9af0b-236">Type</span><span class="sxs-lookup"><span data-stu-id="9af0b-236">Type</span></span>

*   [<span data-ttu-id="9af0b-237">Body</span><span class="sxs-lookup"><span data-stu-id="9af0b-237">Body</span></span>](/javascript/api/outlook_1_6/office.body)

##### <a name="requirements"></a><span data-ttu-id="9af0b-238">Требования</span><span class="sxs-lookup"><span data-stu-id="9af0b-238">Requirements</span></span>

|<span data-ttu-id="9af0b-239">Требование</span><span class="sxs-lookup"><span data-stu-id="9af0b-239">Requirement</span></span>| <span data-ttu-id="9af0b-240">Значение</span><span class="sxs-lookup"><span data-stu-id="9af0b-240">Value</span></span>|
|---|---|
|[<span data-ttu-id="9af0b-241">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="9af0b-241">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9af0b-242">1.1</span><span class="sxs-lookup"><span data-stu-id="9af0b-242">1.1</span></span>|
|[<span data-ttu-id="9af0b-243">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="9af0b-243">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9af0b-244">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9af0b-244">ReadItem</span></span>|
|[<span data-ttu-id="9af0b-245">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9af0b-245">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9af0b-246">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="9af0b-246">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9af0b-247">Пример</span><span class="sxs-lookup"><span data-stu-id="9af0b-247">Example</span></span>

<span data-ttu-id="9af0b-248">В этом примере возвращается текст сообщения в формате обычного текста.</span><span class="sxs-lookup"><span data-stu-id="9af0b-248">This example gets the body of the message in plain text.</span></span>

```javascript
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="9af0b-249">Ниже приведен пример итогового параметра, переданного функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="9af0b-249">The following is an example of the result parameter passed to the callback function.</span></span>

```json
{
  "value": "TEXT of whole body (including threads below)",
  "status": "succeeded",
  "asyncContext": "This is passed to the callback"
}
```

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients"></a><span data-ttu-id="9af0b-250">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="9af0b-250">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

<span data-ttu-id="9af0b-251">Предоставляет доступ к получателям копии сообщения.</span><span class="sxs-lookup"><span data-stu-id="9af0b-251">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="9af0b-252">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="9af0b-252">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="9af0b-253">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="9af0b-253">Read mode</span></span>

<span data-ttu-id="9af0b-p106">Свойство `cc` возвращает массив, который содержит объект `EmailAddressDetails` для каждого получателя, указанного в строке **Копия** сообщения. Коллекция может включать не более 100 элементов.</span><span class="sxs-lookup"><span data-stu-id="9af0b-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="9af0b-256">Режим создания</span><span class="sxs-lookup"><span data-stu-id="9af0b-256">Compose mode</span></span>

<span data-ttu-id="9af0b-257">Свойство `cc` возвращает объект `Recipients`, предоставляющий методы для получения или обновления получателей, которые указаны в строке **Копия** сообщения.</span><span class="sxs-lookup"><span data-stu-id="9af0b-257">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

```javascript
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="9af0b-258">Type</span><span class="sxs-lookup"><span data-stu-id="9af0b-258">Type</span></span>

*   <span data-ttu-id="9af0b-259">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="9af0b-259">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="9af0b-260">Требования</span><span class="sxs-lookup"><span data-stu-id="9af0b-260">Requirements</span></span>

|<span data-ttu-id="9af0b-261">Требование</span><span class="sxs-lookup"><span data-stu-id="9af0b-261">Requirement</span></span>| <span data-ttu-id="9af0b-262">Значение</span><span class="sxs-lookup"><span data-stu-id="9af0b-262">Value</span></span>|
|---|---|
|[<span data-ttu-id="9af0b-263">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="9af0b-263">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9af0b-264">1.0</span><span class="sxs-lookup"><span data-stu-id="9af0b-264">1.0</span></span>|
|[<span data-ttu-id="9af0b-265">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="9af0b-265">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9af0b-266">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9af0b-266">ReadItem</span></span>|
|[<span data-ttu-id="9af0b-267">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9af0b-267">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9af0b-268">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="9af0b-268">Compose or Read</span></span>|

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="9af0b-269">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="9af0b-269">(nullable) conversationId :String</span></span>

<span data-ttu-id="9af0b-270">Получает идентификатор разговора по электронной почте, содержащего конкретное сообщение.</span><span class="sxs-lookup"><span data-stu-id="9af0b-270">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="9af0b-p107">Вы можете получить целочисленное значение этого свойства, если ваше почтовое приложение активируется в формах просмотра или формах создания ответов. Если пользователь изменит тему ответа, после его отправки идентификатор беседы будет изменен, и полученное ранее значение будет недействительным.</span><span class="sxs-lookup"><span data-stu-id="9af0b-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="9af0b-p108">Это свойство имеет значение NULL для нового элемента в форме создания. Свойство `conversationId` вернет значение, если пользователь задаст тему и сохранит элемент.</span><span class="sxs-lookup"><span data-stu-id="9af0b-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="9af0b-275">Type</span><span class="sxs-lookup"><span data-stu-id="9af0b-275">Type</span></span>

*   <span data-ttu-id="9af0b-276">String</span><span class="sxs-lookup"><span data-stu-id="9af0b-276">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="9af0b-277">Требования</span><span class="sxs-lookup"><span data-stu-id="9af0b-277">Requirements</span></span>

|<span data-ttu-id="9af0b-278">Требование</span><span class="sxs-lookup"><span data-stu-id="9af0b-278">Requirement</span></span>| <span data-ttu-id="9af0b-279">Значение</span><span class="sxs-lookup"><span data-stu-id="9af0b-279">Value</span></span>|
|---|---|
|[<span data-ttu-id="9af0b-280">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="9af0b-280">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9af0b-281">1.0</span><span class="sxs-lookup"><span data-stu-id="9af0b-281">1.0</span></span>|
|[<span data-ttu-id="9af0b-282">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="9af0b-282">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9af0b-283">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9af0b-283">ReadItem</span></span>|
|[<span data-ttu-id="9af0b-284">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9af0b-284">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9af0b-285">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="9af0b-285">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9af0b-286">Пример</span><span class="sxs-lookup"><span data-stu-id="9af0b-286">Example</span></span>

```javascript
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

#### <a name="datetimecreated-date"></a><span data-ttu-id="9af0b-287">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="9af0b-287">dateTimeCreated :Date</span></span>

<span data-ttu-id="9af0b-p109">Получает дату и время создания элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="9af0b-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="9af0b-290">Тип</span><span class="sxs-lookup"><span data-stu-id="9af0b-290">Type</span></span>

*   <span data-ttu-id="9af0b-291">Дата</span><span class="sxs-lookup"><span data-stu-id="9af0b-291">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="9af0b-292">Требования</span><span class="sxs-lookup"><span data-stu-id="9af0b-292">Requirements</span></span>

|<span data-ttu-id="9af0b-293">Требование</span><span class="sxs-lookup"><span data-stu-id="9af0b-293">Requirement</span></span>| <span data-ttu-id="9af0b-294">Значение</span><span class="sxs-lookup"><span data-stu-id="9af0b-294">Value</span></span>|
|---|---|
|[<span data-ttu-id="9af0b-295">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="9af0b-295">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9af0b-296">1.0</span><span class="sxs-lookup"><span data-stu-id="9af0b-296">1.0</span></span>|
|[<span data-ttu-id="9af0b-297">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="9af0b-297">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9af0b-298">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9af0b-298">ReadItem</span></span>|
|[<span data-ttu-id="9af0b-299">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9af0b-299">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9af0b-300">Чтение</span><span class="sxs-lookup"><span data-stu-id="9af0b-300">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9af0b-301">Пример</span><span class="sxs-lookup"><span data-stu-id="9af0b-301">Example</span></span>

```javascript
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="9af0b-302">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="9af0b-302">dateTimeModified :Date</span></span>

<span data-ttu-id="9af0b-p110">Получает дату и время последнего изменения элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="9af0b-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="9af0b-305">Этот элемент не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="9af0b-305">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="9af0b-306">Type</span><span class="sxs-lookup"><span data-stu-id="9af0b-306">Type</span></span>

*   <span data-ttu-id="9af0b-307">Дата</span><span class="sxs-lookup"><span data-stu-id="9af0b-307">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="9af0b-308">Требования</span><span class="sxs-lookup"><span data-stu-id="9af0b-308">Requirements</span></span>

|<span data-ttu-id="9af0b-309">Требование</span><span class="sxs-lookup"><span data-stu-id="9af0b-309">Requirement</span></span>| <span data-ttu-id="9af0b-310">Значение</span><span class="sxs-lookup"><span data-stu-id="9af0b-310">Value</span></span>|
|---|---|
|[<span data-ttu-id="9af0b-311">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="9af0b-311">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9af0b-312">1.0</span><span class="sxs-lookup"><span data-stu-id="9af0b-312">1.0</span></span>|
|[<span data-ttu-id="9af0b-313">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="9af0b-313">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9af0b-314">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9af0b-314">ReadItem</span></span>|
|[<span data-ttu-id="9af0b-315">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9af0b-315">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9af0b-316">Чтение</span><span class="sxs-lookup"><span data-stu-id="9af0b-316">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9af0b-317">Пример</span><span class="sxs-lookup"><span data-stu-id="9af0b-317">Example</span></span>

```javascript
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

####  <a name="end-datetimejavascriptapioutlook16officetime"></a><span data-ttu-id="9af0b-318">end :Date|[Time](/javascript/api/outlook_1_6/office.time)</span><span class="sxs-lookup"><span data-stu-id="9af0b-318">end :Date|[Time](/javascript/api/outlook_1_6/office.time)</span></span>

<span data-ttu-id="9af0b-319">Получает или задает дату и время окончания встречи.</span><span class="sxs-lookup"><span data-stu-id="9af0b-319">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="9af0b-p111">Свойство `end` представлено в виде значения даты и времени в формате UTC. Преобразовать значение свойства end в местные значения даты и времени клиента можно с помощью метода [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime).</span><span class="sxs-lookup"><span data-stu-id="9af0b-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="9af0b-322">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="9af0b-322">Read mode</span></span>

<span data-ttu-id="9af0b-323">Свойство `end` возвращает объект `Date`.</span><span class="sxs-lookup"><span data-stu-id="9af0b-323">The `end` property returns a `Date` object.</span></span>

```javascript
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="9af0b-324">Режим создания</span><span class="sxs-lookup"><span data-stu-id="9af0b-324">Compose mode</span></span>

<span data-ttu-id="9af0b-325">Свойство `end` возвращает объект `Time`.</span><span class="sxs-lookup"><span data-stu-id="9af0b-325">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="9af0b-326">Если вы задаете время окончания с помощью метода [`Time.setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-), необходимо использовать метод [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) для преобразования местного времени на клиенте в формат UTC для сервера.</span><span class="sxs-lookup"><span data-stu-id="9af0b-326">When you use the [`Time.setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="9af0b-327">В примере ниже показано, как с помощью метода [`setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) объекта `Time` задать время окончания встречи.</span><span class="sxs-lookup"><span data-stu-id="9af0b-327">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="9af0b-328">Type</span><span class="sxs-lookup"><span data-stu-id="9af0b-328">Type</span></span>

*   <span data-ttu-id="9af0b-329">Date | [Time](/javascript/api/outlook_1_6/office.time)</span><span class="sxs-lookup"><span data-stu-id="9af0b-329">Date | [Time](/javascript/api/outlook_1_6/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="9af0b-330">Требования</span><span class="sxs-lookup"><span data-stu-id="9af0b-330">Requirements</span></span>

|<span data-ttu-id="9af0b-331">Требование</span><span class="sxs-lookup"><span data-stu-id="9af0b-331">Requirement</span></span>| <span data-ttu-id="9af0b-332">Значение</span><span class="sxs-lookup"><span data-stu-id="9af0b-332">Value</span></span>|
|---|---|
|[<span data-ttu-id="9af0b-333">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="9af0b-333">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9af0b-334">1.0</span><span class="sxs-lookup"><span data-stu-id="9af0b-334">1.0</span></span>|
|[<span data-ttu-id="9af0b-335">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="9af0b-335">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9af0b-336">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9af0b-336">ReadItem</span></span>|
|[<span data-ttu-id="9af0b-337">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9af0b-337">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9af0b-338">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="9af0b-338">Compose or Read</span></span>|

#### <a name="from-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails"></a><span data-ttu-id="9af0b-339">from :[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="9af0b-339">from :[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)</span></span>

<span data-ttu-id="9af0b-p112">Получает электронный адрес отправителя сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="9af0b-p112">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="9af0b-p113">Свойства `from` и [`sender`](#sender-emailaddressdetails) представляют одно лицо, если сообщение не отправлено представителем. В противном случае свойство `from` представляет лицо, делегировавшее полномочия, а свойство sender — представителя.</span><span class="sxs-lookup"><span data-stu-id="9af0b-p113">The `from` and [`sender`](#sender-emailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="9af0b-344">Свойству `recipientType`, принадлежащему объекту `EmailAddressDetails` в свойстве `from`, задано значение `undefined`.</span><span class="sxs-lookup"><span data-stu-id="9af0b-344">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="9af0b-345">Тип</span><span class="sxs-lookup"><span data-stu-id="9af0b-345">Type</span></span>

*   [<span data-ttu-id="9af0b-346">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="9af0b-346">EmailAddressDetails</span></span>](/javascript/api/outlook_1_6/office.emailaddressdetails)

##### <a name="example"></a><span data-ttu-id="9af0b-347">Пример</span><span class="sxs-lookup"><span data-stu-id="9af0b-347">Example</span></span>

```javascript
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

##### <a name="requirements"></a><span data-ttu-id="9af0b-348">Требования</span><span class="sxs-lookup"><span data-stu-id="9af0b-348">Requirements</span></span>

|<span data-ttu-id="9af0b-349">Требование</span><span class="sxs-lookup"><span data-stu-id="9af0b-349">Requirement</span></span>| <span data-ttu-id="9af0b-350">Значение</span><span class="sxs-lookup"><span data-stu-id="9af0b-350">Value</span></span>|
|---|---|
|[<span data-ttu-id="9af0b-351">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="9af0b-351">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9af0b-352">1.0</span><span class="sxs-lookup"><span data-stu-id="9af0b-352">1.0</span></span>|
|[<span data-ttu-id="9af0b-353">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="9af0b-353">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9af0b-354">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9af0b-354">ReadItem</span></span>|
|[<span data-ttu-id="9af0b-355">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9af0b-355">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9af0b-356">Чтение</span><span class="sxs-lookup"><span data-stu-id="9af0b-356">Read</span></span>|

#### <a name="internetmessageid-string"></a><span data-ttu-id="9af0b-357">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="9af0b-357">internetMessageId :String</span></span>

<span data-ttu-id="9af0b-p114">Получает идентификатор интернет-сообщения для электронного сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="9af0b-p114">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="9af0b-360">Type</span><span class="sxs-lookup"><span data-stu-id="9af0b-360">Type</span></span>

*   <span data-ttu-id="9af0b-361">String</span><span class="sxs-lookup"><span data-stu-id="9af0b-361">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="9af0b-362">Требования</span><span class="sxs-lookup"><span data-stu-id="9af0b-362">Requirements</span></span>

|<span data-ttu-id="9af0b-363">Требование</span><span class="sxs-lookup"><span data-stu-id="9af0b-363">Requirement</span></span>| <span data-ttu-id="9af0b-364">Значение</span><span class="sxs-lookup"><span data-stu-id="9af0b-364">Value</span></span>|
|---|---|
|[<span data-ttu-id="9af0b-365">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="9af0b-365">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9af0b-366">1.0</span><span class="sxs-lookup"><span data-stu-id="9af0b-366">1.0</span></span>|
|[<span data-ttu-id="9af0b-367">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="9af0b-367">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9af0b-368">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9af0b-368">ReadItem</span></span>|
|[<span data-ttu-id="9af0b-369">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9af0b-369">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9af0b-370">Чтение</span><span class="sxs-lookup"><span data-stu-id="9af0b-370">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9af0b-371">Пример</span><span class="sxs-lookup"><span data-stu-id="9af0b-371">Example</span></span>

```javascript
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="9af0b-372">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="9af0b-372">itemClass :String</span></span>

<span data-ttu-id="9af0b-p115">Получает класс элемента веб-служб Exchange для выбранного элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="9af0b-p115">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="9af0b-p116">Свойство `itemClass` указывает класс сообщения выбранного элемента. Ниже приводятся классы сообщения по умолчанию для элемента сообщения или встречи.</span><span class="sxs-lookup"><span data-stu-id="9af0b-p116">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="9af0b-377">Тип</span><span class="sxs-lookup"><span data-stu-id="9af0b-377">Type</span></span> | <span data-ttu-id="9af0b-378">Описание</span><span class="sxs-lookup"><span data-stu-id="9af0b-378">Description</span></span> | <span data-ttu-id="9af0b-379">Класс элемента</span><span class="sxs-lookup"><span data-stu-id="9af0b-379">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="9af0b-380">Элементы встречи</span><span class="sxs-lookup"><span data-stu-id="9af0b-380">Appointment items</span></span> | <span data-ttu-id="9af0b-381">Это элементы календаря для класса элемента `IPM.Appointment` или `IPM.Appointment.Occurrence`.</span><span class="sxs-lookup"><span data-stu-id="9af0b-381">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurrence` |
| <span data-ttu-id="9af0b-382">Элементы сообщения</span><span class="sxs-lookup"><span data-stu-id="9af0b-382">Message items</span></span> | <span data-ttu-id="9af0b-383">Сюда входят электронные сообщения, для которых по умолчанию задан класс сообщения `IPM.Note`, а также приглашения на собрания, ответы на них и уведомления об их отмене, использующие `IPM.Schedule.Meeting` в качестве базового класса сообщения.</span><span class="sxs-lookup"><span data-stu-id="9af0b-383">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="9af0b-384">Можно создавать настраиваемые классы сообщения, расширяющие классы сообщения по умолчанию, например настраиваемый класс сообщения о встрече `IPM.Appointment.Contoso`.</span><span class="sxs-lookup"><span data-stu-id="9af0b-384">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="9af0b-385">Type</span><span class="sxs-lookup"><span data-stu-id="9af0b-385">Type</span></span>

*   <span data-ttu-id="9af0b-386">String</span><span class="sxs-lookup"><span data-stu-id="9af0b-386">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="9af0b-387">Требования</span><span class="sxs-lookup"><span data-stu-id="9af0b-387">Requirements</span></span>

|<span data-ttu-id="9af0b-388">Требование</span><span class="sxs-lookup"><span data-stu-id="9af0b-388">Requirement</span></span>| <span data-ttu-id="9af0b-389">Значение</span><span class="sxs-lookup"><span data-stu-id="9af0b-389">Value</span></span>|
|---|---|
|[<span data-ttu-id="9af0b-390">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="9af0b-390">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9af0b-391">1.0</span><span class="sxs-lookup"><span data-stu-id="9af0b-391">1.0</span></span>|
|[<span data-ttu-id="9af0b-392">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="9af0b-392">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9af0b-393">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9af0b-393">ReadItem</span></span>|
|[<span data-ttu-id="9af0b-394">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9af0b-394">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9af0b-395">Чтение</span><span class="sxs-lookup"><span data-stu-id="9af0b-395">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9af0b-396">Пример</span><span class="sxs-lookup"><span data-stu-id="9af0b-396">Example</span></span>

```javascript
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="9af0b-397">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="9af0b-397">(nullable) itemId :String</span></span>

<span data-ttu-id="9af0b-p117">Получает идентификатор элемента веб-служб Exchange для текущего элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="9af0b-p117">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="9af0b-400">Идентификатор, возвращаемый свойством `itemId`, совпадает с идентификатором элемента веб-служб Exchange.</span><span class="sxs-lookup"><span data-stu-id="9af0b-400">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="9af0b-401">Свойство `itemId` не совпадает с идентификатором записи Outlook, а также идентификатором, который используется REST API Outlook.</span><span class="sxs-lookup"><span data-stu-id="9af0b-401">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="9af0b-402">Прежде чем совершать вызовы REST API, используя это значение, его необходимо преобразовать с помощью [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="9af0b-402">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="9af0b-403">Дополнительные сведения см. в статье [Использование REST API Outlook из надстройки Outlook](/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="9af0b-403">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="9af0b-p119">Свойство `itemId` недоступно в режиме создания. Если требуется идентификатор элемента, с помощью метода [`saveAsync`](#saveasyncoptions-callback) можно сохранить элемент в хранилище. При этом в параметре [`AsyncResult.value`](/javascript/api/office/office.asyncresult) функции обратного вызова возвращается идентификатор элемента.</span><span class="sxs-lookup"><span data-stu-id="9af0b-p119">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="9af0b-406">Тип</span><span class="sxs-lookup"><span data-stu-id="9af0b-406">Type</span></span>

*   <span data-ttu-id="9af0b-407">String</span><span class="sxs-lookup"><span data-stu-id="9af0b-407">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="9af0b-408">Требования</span><span class="sxs-lookup"><span data-stu-id="9af0b-408">Requirements</span></span>

|<span data-ttu-id="9af0b-409">Требование</span><span class="sxs-lookup"><span data-stu-id="9af0b-409">Requirement</span></span>| <span data-ttu-id="9af0b-410">Значение</span><span class="sxs-lookup"><span data-stu-id="9af0b-410">Value</span></span>|
|---|---|
|[<span data-ttu-id="9af0b-411">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="9af0b-411">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9af0b-412">1.0</span><span class="sxs-lookup"><span data-stu-id="9af0b-412">1.0</span></span>|
|[<span data-ttu-id="9af0b-413">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="9af0b-413">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9af0b-414">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9af0b-414">ReadItem</span></span>|
|[<span data-ttu-id="9af0b-415">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9af0b-415">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9af0b-416">Чтение</span><span class="sxs-lookup"><span data-stu-id="9af0b-416">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9af0b-417">Пример</span><span class="sxs-lookup"><span data-stu-id="9af0b-417">Example</span></span>

<span data-ttu-id="9af0b-p120">Указанный ниже код проверяет наличие идентификатора элемента. Если свойство `itemId` возвращает значение `null` или `undefined`, элемент будет сохранен в хранилище, а из асинхронного результата будет получен идентификатор элемента.</span><span class="sxs-lookup"><span data-stu-id="9af0b-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```javascript
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result) {
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook16officemailboxenumsitemtype"></a><span data-ttu-id="9af0b-420">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_6/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="9af0b-420">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_6/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="9af0b-421">Получает тип элемента, который представляет экземпляр.</span><span class="sxs-lookup"><span data-stu-id="9af0b-421">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="9af0b-422">Свойство `itemType` возвращает одно из значений перечисления `ItemType`, которое указывает, является ли экземпляр объекта `item` сообщением или встречей.</span><span class="sxs-lookup"><span data-stu-id="9af0b-422">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="9af0b-423">Type</span><span class="sxs-lookup"><span data-stu-id="9af0b-423">Type</span></span>

*   [<span data-ttu-id="9af0b-424">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="9af0b-424">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook_1_6/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="9af0b-425">Требования</span><span class="sxs-lookup"><span data-stu-id="9af0b-425">Requirements</span></span>

|<span data-ttu-id="9af0b-426">Требование</span><span class="sxs-lookup"><span data-stu-id="9af0b-426">Requirement</span></span>| <span data-ttu-id="9af0b-427">Значение</span><span class="sxs-lookup"><span data-stu-id="9af0b-427">Value</span></span>|
|---|---|
|[<span data-ttu-id="9af0b-428">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="9af0b-428">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9af0b-429">1.0</span><span class="sxs-lookup"><span data-stu-id="9af0b-429">1.0</span></span>|
|[<span data-ttu-id="9af0b-430">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="9af0b-430">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9af0b-431">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9af0b-431">ReadItem</span></span>|
|[<span data-ttu-id="9af0b-432">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9af0b-432">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9af0b-433">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="9af0b-433">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9af0b-434">Пример</span><span class="sxs-lookup"><span data-stu-id="9af0b-434">Example</span></span>

```javascript
if (Office.context.mailbox.item.itemType === Office.MailboxEnums.ItemType.Message) {
  // Do something.
} else {
  // Do something else.
}
```

####  <a name="location-stringlocationjavascriptapioutlook16officelocation"></a><span data-ttu-id="9af0b-435">location :String|[Location](/javascript/api/outlook_1_6/office.location)</span><span class="sxs-lookup"><span data-stu-id="9af0b-435">location :String|[Location](/javascript/api/outlook_1_6/office.location)</span></span>

<span data-ttu-id="9af0b-436">Получает или задает место встречи.</span><span class="sxs-lookup"><span data-stu-id="9af0b-436">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="9af0b-437">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="9af0b-437">Read mode</span></span>

<span data-ttu-id="9af0b-438">Свойство `location` возвращает строку, содержащую сведения о месте встречи.</span><span class="sxs-lookup"><span data-stu-id="9af0b-438">The `location` property returns a string that contains the location of the appointment.</span></span>

```javascript
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="9af0b-439">Режим создания</span><span class="sxs-lookup"><span data-stu-id="9af0b-439">Compose mode</span></span>

<span data-ttu-id="9af0b-440">Свойство `location` возвращает объект `Location`, предоставляющий методы, которые используются для получения и задания места встречи.</span><span class="sxs-lookup"><span data-stu-id="9af0b-440">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```javascript
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="9af0b-441">Type</span><span class="sxs-lookup"><span data-stu-id="9af0b-441">Type</span></span>

*   <span data-ttu-id="9af0b-442">String | [Location](/javascript/api/outlook_1_6/office.location)</span><span class="sxs-lookup"><span data-stu-id="9af0b-442">String | [Location](/javascript/api/outlook_1_6/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="9af0b-443">Требования</span><span class="sxs-lookup"><span data-stu-id="9af0b-443">Requirements</span></span>

|<span data-ttu-id="9af0b-444">Требование</span><span class="sxs-lookup"><span data-stu-id="9af0b-444">Requirement</span></span>| <span data-ttu-id="9af0b-445">Значение</span><span class="sxs-lookup"><span data-stu-id="9af0b-445">Value</span></span>|
|---|---|
|[<span data-ttu-id="9af0b-446">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="9af0b-446">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9af0b-447">1.0</span><span class="sxs-lookup"><span data-stu-id="9af0b-447">1.0</span></span>|
|[<span data-ttu-id="9af0b-448">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="9af0b-448">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9af0b-449">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9af0b-449">ReadItem</span></span>|
|[<span data-ttu-id="9af0b-450">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9af0b-450">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9af0b-451">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="9af0b-451">Compose or Read</span></span>|

#### <a name="normalizedsubject-string"></a><span data-ttu-id="9af0b-452">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="9af0b-452">normalizedSubject :String</span></span>

<span data-ttu-id="9af0b-p121">Получает тему элемента со всеми удаленными префиксами (включая `RE:` и `FWD:`). Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="9af0b-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="9af0b-p122">Свойство normalizedSubject получает тему элемента со стандартными префиксами (такими как `RE:` и `FW:`), добавляемыми почтовыми программами. Для получения темы элемента с неизмененными префиксами используйте свойство [`subject`](#subject-stringsubject).</span><span class="sxs-lookup"><span data-stu-id="9af0b-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="9af0b-457">Type</span><span class="sxs-lookup"><span data-stu-id="9af0b-457">Type</span></span>

*   <span data-ttu-id="9af0b-458">String</span><span class="sxs-lookup"><span data-stu-id="9af0b-458">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="9af0b-459">Требования</span><span class="sxs-lookup"><span data-stu-id="9af0b-459">Requirements</span></span>

|<span data-ttu-id="9af0b-460">Требование</span><span class="sxs-lookup"><span data-stu-id="9af0b-460">Requirement</span></span>| <span data-ttu-id="9af0b-461">Значение</span><span class="sxs-lookup"><span data-stu-id="9af0b-461">Value</span></span>|
|---|---|
|[<span data-ttu-id="9af0b-462">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="9af0b-462">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9af0b-463">1.0</span><span class="sxs-lookup"><span data-stu-id="9af0b-463">1.0</span></span>|
|[<span data-ttu-id="9af0b-464">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="9af0b-464">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9af0b-465">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9af0b-465">ReadItem</span></span>|
|[<span data-ttu-id="9af0b-466">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9af0b-466">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9af0b-467">Чтение</span><span class="sxs-lookup"><span data-stu-id="9af0b-467">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9af0b-468">Пример</span><span class="sxs-lookup"><span data-stu-id="9af0b-468">Example</span></span>

```javascript
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

####  <a name="notificationmessages-notificationmessagesjavascriptapioutlook16officenotificationmessages"></a><span data-ttu-id="9af0b-469">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_6/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="9af0b-469">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_6/office.notificationmessages)</span></span>

<span data-ttu-id="9af0b-470">Получает сообщения уведомления для элемента.</span><span class="sxs-lookup"><span data-stu-id="9af0b-470">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="9af0b-471">Тип</span><span class="sxs-lookup"><span data-stu-id="9af0b-471">Type</span></span>

*   [<span data-ttu-id="9af0b-472">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="9af0b-472">NotificationMessages</span></span>](/javascript/api/outlook_1_6/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="9af0b-473">Требования</span><span class="sxs-lookup"><span data-stu-id="9af0b-473">Requirements</span></span>

|<span data-ttu-id="9af0b-474">Требование</span><span class="sxs-lookup"><span data-stu-id="9af0b-474">Requirement</span></span>| <span data-ttu-id="9af0b-475">Значение</span><span class="sxs-lookup"><span data-stu-id="9af0b-475">Value</span></span>|
|---|---|
|[<span data-ttu-id="9af0b-476">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="9af0b-476">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9af0b-477">1.3</span><span class="sxs-lookup"><span data-stu-id="9af0b-477">1.3</span></span>|
|[<span data-ttu-id="9af0b-478">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="9af0b-478">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9af0b-479">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9af0b-479">ReadItem</span></span>|
|[<span data-ttu-id="9af0b-480">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9af0b-480">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9af0b-481">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="9af0b-481">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9af0b-482">Пример</span><span class="sxs-lookup"><span data-stu-id="9af0b-482">Example</span></span>

```javascript
// Get all notifications.
Office.context.mailbox.item.notificationMessages.getAllAsync(
  function (asyncResult) {
    console.log(JSON.stringify(asyncResult));
  }
);
```

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients"></a><span data-ttu-id="9af0b-483">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="9af0b-483">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

<span data-ttu-id="9af0b-484">Предоставляет доступ к необязательным участникам события.</span><span class="sxs-lookup"><span data-stu-id="9af0b-484">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="9af0b-485">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="9af0b-485">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="9af0b-486">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="9af0b-486">Read mode</span></span>

<span data-ttu-id="9af0b-487">Свойство `optionalAttendees` возвращает массив, содержащий объект `EmailAddressDetails` для каждого необязательного участника собрания.</span><span class="sxs-lookup"><span data-stu-id="9af0b-487">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

```javascript
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="9af0b-488">Режим создания</span><span class="sxs-lookup"><span data-stu-id="9af0b-488">Compose mode</span></span>

<span data-ttu-id="9af0b-489">Свойство `optionalAttendees` возвращает объект `Recipients`, который предоставляет методы для получения или обновления необязательных участников собрания.</span><span class="sxs-lookup"><span data-stu-id="9af0b-489">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="9af0b-490">Type</span><span class="sxs-lookup"><span data-stu-id="9af0b-490">Type</span></span>

*   <span data-ttu-id="9af0b-491">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="9af0b-491">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="9af0b-492">Требования</span><span class="sxs-lookup"><span data-stu-id="9af0b-492">Requirements</span></span>

|<span data-ttu-id="9af0b-493">Требование</span><span class="sxs-lookup"><span data-stu-id="9af0b-493">Requirement</span></span>| <span data-ttu-id="9af0b-494">Значение</span><span class="sxs-lookup"><span data-stu-id="9af0b-494">Value</span></span>|
|---|---|
|[<span data-ttu-id="9af0b-495">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="9af0b-495">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9af0b-496">1.0</span><span class="sxs-lookup"><span data-stu-id="9af0b-496">1.0</span></span>|
|[<span data-ttu-id="9af0b-497">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="9af0b-497">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9af0b-498">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9af0b-498">ReadItem</span></span>|
|[<span data-ttu-id="9af0b-499">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9af0b-499">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9af0b-500">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="9af0b-500">Compose or Read</span></span>|

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails"></a><span data-ttu-id="9af0b-501">organizer :[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="9af0b-501">organizer :[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)</span></span>

<span data-ttu-id="9af0b-p124">Получает электронный адрес организатора указанного собрания. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="9af0b-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="9af0b-504">Type</span><span class="sxs-lookup"><span data-stu-id="9af0b-504">Type</span></span>

*   [<span data-ttu-id="9af0b-505">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="9af0b-505">EmailAddressDetails</span></span>](/javascript/api/outlook_1_6/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="9af0b-506">Требования</span><span class="sxs-lookup"><span data-stu-id="9af0b-506">Requirements</span></span>

|<span data-ttu-id="9af0b-507">Требование</span><span class="sxs-lookup"><span data-stu-id="9af0b-507">Requirement</span></span>| <span data-ttu-id="9af0b-508">Значение</span><span class="sxs-lookup"><span data-stu-id="9af0b-508">Value</span></span>|
|---|---|
|[<span data-ttu-id="9af0b-509">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="9af0b-509">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9af0b-510">1.0</span><span class="sxs-lookup"><span data-stu-id="9af0b-510">1.0</span></span>|
|[<span data-ttu-id="9af0b-511">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="9af0b-511">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9af0b-512">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9af0b-512">ReadItem</span></span>|
|[<span data-ttu-id="9af0b-513">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9af0b-513">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9af0b-514">Чтение</span><span class="sxs-lookup"><span data-stu-id="9af0b-514">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9af0b-515">Пример</span><span class="sxs-lookup"><span data-stu-id="9af0b-515">Example</span></span>

```javascript
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients"></a><span data-ttu-id="9af0b-516">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="9af0b-516">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

<span data-ttu-id="9af0b-517">Предоставляет доступ к обязательным участникам события.</span><span class="sxs-lookup"><span data-stu-id="9af0b-517">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="9af0b-518">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="9af0b-518">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="9af0b-519">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="9af0b-519">Read mode</span></span>

<span data-ttu-id="9af0b-520">Свойство `requiredAttendees` возвращает массив, содержащий объект `EmailAddressDetails` для каждого обязательного участника собрания.</span><span class="sxs-lookup"><span data-stu-id="9af0b-520">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

```javascript
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="9af0b-521">Режим создания</span><span class="sxs-lookup"><span data-stu-id="9af0b-521">Compose mode</span></span>

<span data-ttu-id="9af0b-522">Свойство `requiredAttendees` возвращает объект `Recipients`, предоставляющий методы, с помощью которых можно получить или обновить сведения об обязательных участниках собрания.</span><span class="sxs-lookup"><span data-stu-id="9af0b-522">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="9af0b-523">Тип</span><span class="sxs-lookup"><span data-stu-id="9af0b-523">Type</span></span>

*   <span data-ttu-id="9af0b-524">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="9af0b-524">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="9af0b-525">Требования</span><span class="sxs-lookup"><span data-stu-id="9af0b-525">Requirements</span></span>

|<span data-ttu-id="9af0b-526">Требование</span><span class="sxs-lookup"><span data-stu-id="9af0b-526">Requirement</span></span>| <span data-ttu-id="9af0b-527">Значение</span><span class="sxs-lookup"><span data-stu-id="9af0b-527">Value</span></span>|
|---|---|
|[<span data-ttu-id="9af0b-528">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="9af0b-528">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9af0b-529">1.0</span><span class="sxs-lookup"><span data-stu-id="9af0b-529">1.0</span></span>|
|[<span data-ttu-id="9af0b-530">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="9af0b-530">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9af0b-531">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9af0b-531">ReadItem</span></span>|
|[<span data-ttu-id="9af0b-532">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9af0b-532">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9af0b-533">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="9af0b-533">Compose or Read</span></span>|

#### <a name="sender-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails"></a><span data-ttu-id="9af0b-534">sender :[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="9af0b-534">sender :[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)</span></span>

<span data-ttu-id="9af0b-p126">Получает электронный адрес отправителя электронного сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="9af0b-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="9af0b-p127">Свойства [`from`](#from-emailaddressdetails) и `sender` представляют одно лицо, если сообщение не отправлено представителем. В противном случае свойство `from` представляет лицо, делегировавшее полномочия, а свойство sender — представителя.</span><span class="sxs-lookup"><span data-stu-id="9af0b-p127">The [`from`](#from-emailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="9af0b-539">Свойству `recipientType`, принадлежащему объекту `EmailAddressDetails` в свойстве `sender`, задано значение `undefined`.</span><span class="sxs-lookup"><span data-stu-id="9af0b-539">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="9af0b-540">Type</span><span class="sxs-lookup"><span data-stu-id="9af0b-540">Type</span></span>

*   [<span data-ttu-id="9af0b-541">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="9af0b-541">EmailAddressDetails</span></span>](/javascript/api/outlook_1_6/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="9af0b-542">Требования</span><span class="sxs-lookup"><span data-stu-id="9af0b-542">Requirements</span></span>

|<span data-ttu-id="9af0b-543">Требование</span><span class="sxs-lookup"><span data-stu-id="9af0b-543">Requirement</span></span>| <span data-ttu-id="9af0b-544">Значение</span><span class="sxs-lookup"><span data-stu-id="9af0b-544">Value</span></span>|
|---|---|
|[<span data-ttu-id="9af0b-545">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="9af0b-545">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9af0b-546">1.0</span><span class="sxs-lookup"><span data-stu-id="9af0b-546">1.0</span></span>|
|[<span data-ttu-id="9af0b-547">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="9af0b-547">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9af0b-548">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9af0b-548">ReadItem</span></span>|
|[<span data-ttu-id="9af0b-549">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9af0b-549">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9af0b-550">Чтение</span><span class="sxs-lookup"><span data-stu-id="9af0b-550">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9af0b-551">Пример</span><span class="sxs-lookup"><span data-stu-id="9af0b-551">Example</span></span>

```javascript
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

####  <a name="start-datetimejavascriptapioutlook16officetime"></a><span data-ttu-id="9af0b-552">start :Date|[Time](/javascript/api/outlook_1_6/office.time)</span><span class="sxs-lookup"><span data-stu-id="9af0b-552">start :Date|[Time](/javascript/api/outlook_1_6/office.time)</span></span>

<span data-ttu-id="9af0b-553">Получает или задает дату и время начала встречи.</span><span class="sxs-lookup"><span data-stu-id="9af0b-553">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="9af0b-p128">Свойство `start` представлено в виде значения даты и времени в формате UTC. Это значение можно преобразовать в местные значения даты и времени клиента с помощью метода [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime).</span><span class="sxs-lookup"><span data-stu-id="9af0b-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="9af0b-556">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="9af0b-556">Read mode</span></span>

<span data-ttu-id="9af0b-557">Свойство `start` возвращает объект `Date`.</span><span class="sxs-lookup"><span data-stu-id="9af0b-557">The `start` property returns a `Date` object.</span></span>

```javascript
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="9af0b-558">Режим создания</span><span class="sxs-lookup"><span data-stu-id="9af0b-558">Compose mode</span></span>

<span data-ttu-id="9af0b-559">Свойство `start` возвращает объект `Time`.</span><span class="sxs-lookup"><span data-stu-id="9af0b-559">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="9af0b-560">Если вы задаете время начала с помощью метода [`Time.setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-), необходимо использовать метод [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) для преобразования местного времени на клиенте в формат UTC для сервера.</span><span class="sxs-lookup"><span data-stu-id="9af0b-560">When you use the [`Time.setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="9af0b-561">В примере ниже с помощью метода [`setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) объекта `Time` задается время начала встречи в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="9af0b-561">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="9af0b-562">Type</span><span class="sxs-lookup"><span data-stu-id="9af0b-562">Type</span></span>

*   <span data-ttu-id="9af0b-563">Date | [Time](/javascript/api/outlook_1_6/office.time)</span><span class="sxs-lookup"><span data-stu-id="9af0b-563">Date | [Time](/javascript/api/outlook_1_6/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="9af0b-564">Требования</span><span class="sxs-lookup"><span data-stu-id="9af0b-564">Requirements</span></span>

|<span data-ttu-id="9af0b-565">Требование</span><span class="sxs-lookup"><span data-stu-id="9af0b-565">Requirement</span></span>| <span data-ttu-id="9af0b-566">Значение</span><span class="sxs-lookup"><span data-stu-id="9af0b-566">Value</span></span>|
|---|---|
|[<span data-ttu-id="9af0b-567">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="9af0b-567">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9af0b-568">1.0</span><span class="sxs-lookup"><span data-stu-id="9af0b-568">1.0</span></span>|
|[<span data-ttu-id="9af0b-569">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="9af0b-569">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9af0b-570">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9af0b-570">ReadItem</span></span>|
|[<span data-ttu-id="9af0b-571">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9af0b-571">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9af0b-572">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="9af0b-572">Compose or Read</span></span>|

####  <a name="subject-stringsubjectjavascriptapioutlook16officesubject"></a><span data-ttu-id="9af0b-573">subject :String|[Subject](/javascript/api/outlook_1_6/office.subject)</span><span class="sxs-lookup"><span data-stu-id="9af0b-573">subject :String|[Subject](/javascript/api/outlook_1_6/office.subject)</span></span>

<span data-ttu-id="9af0b-574">Получает или задает описание, которое отображается в поле темы элемента.</span><span class="sxs-lookup"><span data-stu-id="9af0b-574">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="9af0b-575">Свойство `subject` получает или задает всю тему элемента для отправки с почтового сервера.</span><span class="sxs-lookup"><span data-stu-id="9af0b-575">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="9af0b-576">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="9af0b-576">Read mode</span></span>

<span data-ttu-id="9af0b-p129">Свойство `subject` возвращает строку. С помощью свойства [`normalizedSubject`](#normalizedsubject-string) можно получить тему без начальных префиксов, таких как `RE:` и `FW:`.</span><span class="sxs-lookup"><span data-stu-id="9af0b-p129">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```javascript
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="9af0b-579">Режим создания</span><span class="sxs-lookup"><span data-stu-id="9af0b-579">Compose mode</span></span>

<span data-ttu-id="9af0b-580">Свойство `subject` возвращает объект `Subject`, который предоставляет методы для получения и задания темы.</span><span class="sxs-lookup"><span data-stu-id="9af0b-580">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```javascript
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="9af0b-581">Тип</span><span class="sxs-lookup"><span data-stu-id="9af0b-581">Type</span></span>

*   <span data-ttu-id="9af0b-582">String | [Subject](/javascript/api/outlook_1_6/office.subject)</span><span class="sxs-lookup"><span data-stu-id="9af0b-582">String | [Subject](/javascript/api/outlook_1_6/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="9af0b-583">Требования</span><span class="sxs-lookup"><span data-stu-id="9af0b-583">Requirements</span></span>

|<span data-ttu-id="9af0b-584">Требование</span><span class="sxs-lookup"><span data-stu-id="9af0b-584">Requirement</span></span>| <span data-ttu-id="9af0b-585">Значение</span><span class="sxs-lookup"><span data-stu-id="9af0b-585">Value</span></span>|
|---|---|
|[<span data-ttu-id="9af0b-586">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="9af0b-586">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9af0b-587">1.0</span><span class="sxs-lookup"><span data-stu-id="9af0b-587">1.0</span></span>|
|[<span data-ttu-id="9af0b-588">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="9af0b-588">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9af0b-589">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9af0b-589">ReadItem</span></span>|
|[<span data-ttu-id="9af0b-590">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9af0b-590">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9af0b-591">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="9af0b-591">Compose or Read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients"></a><span data-ttu-id="9af0b-592">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="9af0b-592">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

<span data-ttu-id="9af0b-593">Предоставляет доступ к получателям, указанным в строке **Кому** сообщения.</span><span class="sxs-lookup"><span data-stu-id="9af0b-593">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="9af0b-594">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="9af0b-594">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="9af0b-595">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="9af0b-595">Read mode</span></span>

<span data-ttu-id="9af0b-p131">Свойство `to` возвращает массив, содержащий объект `EmailAddressDetails` для каждого получателя в строке **Кому** сообщения. Коллекция может включать не более 100 элементов.</span><span class="sxs-lookup"><span data-stu-id="9af0b-p131">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="9af0b-598">Режим создания</span><span class="sxs-lookup"><span data-stu-id="9af0b-598">Compose mode</span></span>

<span data-ttu-id="9af0b-599">Свойство `to` возвращает объект `Recipients`, предоставляющий методы для получения или обновления получателей, которые указаны в строке **Кому** сообщения.</span><span class="sxs-lookup"><span data-stu-id="9af0b-599">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

```javascript
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="9af0b-600">Тип</span><span class="sxs-lookup"><span data-stu-id="9af0b-600">Type</span></span>

*   <span data-ttu-id="9af0b-601">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="9af0b-601">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="9af0b-602">Требования</span><span class="sxs-lookup"><span data-stu-id="9af0b-602">Requirements</span></span>

|<span data-ttu-id="9af0b-603">Требование</span><span class="sxs-lookup"><span data-stu-id="9af0b-603">Requirement</span></span>| <span data-ttu-id="9af0b-604">Значение</span><span class="sxs-lookup"><span data-stu-id="9af0b-604">Value</span></span>|
|---|---|
|[<span data-ttu-id="9af0b-605">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="9af0b-605">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9af0b-606">1.0</span><span class="sxs-lookup"><span data-stu-id="9af0b-606">1.0</span></span>|
|[<span data-ttu-id="9af0b-607">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="9af0b-607">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9af0b-608">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9af0b-608">ReadItem</span></span>|
|[<span data-ttu-id="9af0b-609">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9af0b-609">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9af0b-610">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="9af0b-610">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="9af0b-611">Методы</span><span class="sxs-lookup"><span data-stu-id="9af0b-611">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="9af0b-612">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="9af0b-612">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="9af0b-613">Добавляет файл в сообщение или встречу в качестве вложения.</span><span class="sxs-lookup"><span data-stu-id="9af0b-613">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="9af0b-614">Метод `addFileAttachmentAsync` передает файл по указанному универсальному коду ресурса (URI) и вкладывает его в элемент в форме создания.</span><span class="sxs-lookup"><span data-stu-id="9af0b-614">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="9af0b-615">Идентификатор можно использовать с методом [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback), чтобы удалить вложение, добавленное во время текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="9af0b-615">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9af0b-616">Параметры</span><span class="sxs-lookup"><span data-stu-id="9af0b-616">Parameters</span></span>

|<span data-ttu-id="9af0b-617">Имя</span><span class="sxs-lookup"><span data-stu-id="9af0b-617">Name</span></span>| <span data-ttu-id="9af0b-618">Тип</span><span class="sxs-lookup"><span data-stu-id="9af0b-618">Type</span></span>| <span data-ttu-id="9af0b-619">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="9af0b-619">Attributes</span></span>| <span data-ttu-id="9af0b-620">Описание</span><span class="sxs-lookup"><span data-stu-id="9af0b-620">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="9af0b-621">Строка</span><span class="sxs-lookup"><span data-stu-id="9af0b-621">String</span></span>||<span data-ttu-id="9af0b-p132">Универсальный код ресурса (URI), представляющий расположение файла, который нужно вложить в сообщение или встречу. Максимальная длина — 2048 символов.</span><span class="sxs-lookup"><span data-stu-id="9af0b-p132">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="9af0b-624">String</span><span class="sxs-lookup"><span data-stu-id="9af0b-624">String</span></span>||<span data-ttu-id="9af0b-p133">Имя вложения, которое отображается при передаче вложения. Максимальная длина — 255 символов.</span><span class="sxs-lookup"><span data-stu-id="9af0b-p133">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="9af0b-627">Object</span><span class="sxs-lookup"><span data-stu-id="9af0b-627">Object</span></span>| <span data-ttu-id="9af0b-628">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="9af0b-628">&lt;optional&gt;</span></span>|<span data-ttu-id="9af0b-629">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="9af0b-629">An object literal that contains one or more of the following properties.</span></span>|
| `options.asyncContext` | <span data-ttu-id="9af0b-630">Object</span><span class="sxs-lookup"><span data-stu-id="9af0b-630">Object</span></span> | <span data-ttu-id="9af0b-631">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="9af0b-631">&lt;optional&gt;</span></span> | <span data-ttu-id="9af0b-632">В методе обратного вызова разработчики могут указать любой объект, к которому необходимо получить доступ.</span><span class="sxs-lookup"><span data-stu-id="9af0b-632">Developers can provide any object they wish to access in the callback method.</span></span> |
| `options.isInline` | <span data-ttu-id="9af0b-633">Boolean</span><span class="sxs-lookup"><span data-stu-id="9af0b-633">Boolean</span></span> | <span data-ttu-id="9af0b-634">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="9af0b-634">&lt;optional&gt;</span></span> | <span data-ttu-id="9af0b-635">Значение `true` указывает на то, что вложение будет встроено в текст сообщения и не должно отображаться в списке вложений.</span><span class="sxs-lookup"><span data-stu-id="9af0b-635">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
|`callback`| <span data-ttu-id="9af0b-636">function</span><span class="sxs-lookup"><span data-stu-id="9af0b-636">function</span></span>| <span data-ttu-id="9af0b-637">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="9af0b-637">&lt;optional&gt;</span></span>|<span data-ttu-id="9af0b-638">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="9af0b-638">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="9af0b-639">После успешного выполнения идентификатор вложения будет представлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="9af0b-639">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="9af0b-640">Если добавить вложение не удастся, объект `asyncResult` будет содержать объект `Error` с описанием ошибки.</span><span class="sxs-lookup"><span data-stu-id="9af0b-640">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="9af0b-641">Ошибки</span><span class="sxs-lookup"><span data-stu-id="9af0b-641">Errors</span></span>

| <span data-ttu-id="9af0b-642">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="9af0b-642">Error code</span></span> | <span data-ttu-id="9af0b-643">Описание</span><span class="sxs-lookup"><span data-stu-id="9af0b-643">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="9af0b-644">Вложение превышает максимальный размер.</span><span class="sxs-lookup"><span data-stu-id="9af0b-644">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="9af0b-645">Расширение вложения не поддерживается.</span><span class="sxs-lookup"><span data-stu-id="9af0b-645">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="9af0b-646">Сообщение или встреча содержат слишком много вложений.</span><span class="sxs-lookup"><span data-stu-id="9af0b-646">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="9af0b-647">Требования</span><span class="sxs-lookup"><span data-stu-id="9af0b-647">Requirements</span></span>

|<span data-ttu-id="9af0b-648">Требование</span><span class="sxs-lookup"><span data-stu-id="9af0b-648">Requirement</span></span>| <span data-ttu-id="9af0b-649">Значение</span><span class="sxs-lookup"><span data-stu-id="9af0b-649">Value</span></span>|
|---|---|
|[<span data-ttu-id="9af0b-650">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="9af0b-650">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9af0b-651">1.1</span><span class="sxs-lookup"><span data-stu-id="9af0b-651">1.1</span></span>|
|[<span data-ttu-id="9af0b-652">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="9af0b-652">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9af0b-653">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="9af0b-653">ReadWriteItem</span></span>|
|[<span data-ttu-id="9af0b-654">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9af0b-654">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9af0b-655">Создание</span><span class="sxs-lookup"><span data-stu-id="9af0b-655">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="9af0b-656">Примеры</span><span class="sxs-lookup"><span data-stu-id="9af0b-656">Examples</span></span>

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

<span data-ttu-id="9af0b-657">В приведенном ниже примере файл изображения добавляется в качестве встроенного вложения, а ссылка на вложение добавляется в текст сообщения.</span><span class="sxs-lookup"><span data-stu-id="9af0b-657">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

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

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="9af0b-658">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="9af0b-658">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="9af0b-659">Добавляет к сообщению элемент Exchange, например сообщение, в виде вложения.</span><span class="sxs-lookup"><span data-stu-id="9af0b-659">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="9af0b-p134">С помощью метода `addItemAttachmentAsync` можно в элемент формы создания вложить элемент с указанным идентификатором Exchange. Если указать метод обратного вызова, то этот метод вызывается с помощью параметра `asyncResult`, который содержит идентификатор вложения или код, указывающий на ошибки, которые произошли при вложении элемента. При необходимости можно использовать параметр `options` для передачи сведений о состоянии методу обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="9af0b-p134">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="9af0b-663">Идентификатор можно использовать с методом [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback), чтобы удалить вложение, добавленное во время текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="9af0b-663">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="9af0b-664">Если ваша надстройка Office выполняется в Outlook Web App, метод `addItemAttachmentAsync` обеспечивает вложение элементов в элементы, отличные от редактируемого. Однако это действие не рекомендуем выполнять, так как оно не поддерживается.</span><span class="sxs-lookup"><span data-stu-id="9af0b-664">If your Office Add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9af0b-665">Параметры</span><span class="sxs-lookup"><span data-stu-id="9af0b-665">Parameters</span></span>

|<span data-ttu-id="9af0b-666">Имя</span><span class="sxs-lookup"><span data-stu-id="9af0b-666">Name</span></span>| <span data-ttu-id="9af0b-667">Тип</span><span class="sxs-lookup"><span data-stu-id="9af0b-667">Type</span></span>| <span data-ttu-id="9af0b-668">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="9af0b-668">Attributes</span></span>| <span data-ttu-id="9af0b-669">Описание</span><span class="sxs-lookup"><span data-stu-id="9af0b-669">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="9af0b-670">Строка</span><span class="sxs-lookup"><span data-stu-id="9af0b-670">String</span></span>||<span data-ttu-id="9af0b-p135">Идентификатор Exchange для вкладываемого элемента. Максимальная длина — 100 символов.</span><span class="sxs-lookup"><span data-stu-id="9af0b-p135">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="9af0b-673">String</span><span class="sxs-lookup"><span data-stu-id="9af0b-673">String</span></span>||<span data-ttu-id="9af0b-674">Тема вкладываемого элемента.</span><span class="sxs-lookup"><span data-stu-id="9af0b-674">The subject of the item to be attached.</span></span> <span data-ttu-id="9af0b-675">Максимальная длина: 255 символов.</span><span class="sxs-lookup"><span data-stu-id="9af0b-675">The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="9af0b-676">Object</span><span class="sxs-lookup"><span data-stu-id="9af0b-676">Object</span></span>| <span data-ttu-id="9af0b-677">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="9af0b-677">&lt;optional&gt;</span></span>|<span data-ttu-id="9af0b-678">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="9af0b-678">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="9af0b-679">Object</span><span class="sxs-lookup"><span data-stu-id="9af0b-679">Object</span></span>| <span data-ttu-id="9af0b-680">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="9af0b-680">&lt;optional&gt;</span></span>|<span data-ttu-id="9af0b-681">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="9af0b-681">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="9af0b-682">функция</span><span class="sxs-lookup"><span data-stu-id="9af0b-682">function</span></span>| <span data-ttu-id="9af0b-683">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="9af0b-683">&lt;optional&gt;</span></span>|<span data-ttu-id="9af0b-684">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="9af0b-684">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="9af0b-685">После успешного выполнения идентификатор вложения будет представлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="9af0b-685">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="9af0b-686">Если добавить вложение не удастся, объект `asyncResult` будет содержать объект `Error` с описанием ошибки.</span><span class="sxs-lookup"><span data-stu-id="9af0b-686">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="9af0b-687">Ошибки</span><span class="sxs-lookup"><span data-stu-id="9af0b-687">Errors</span></span>

| <span data-ttu-id="9af0b-688">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="9af0b-688">Error code</span></span> | <span data-ttu-id="9af0b-689">Описание</span><span class="sxs-lookup"><span data-stu-id="9af0b-689">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="9af0b-690">Сообщение или встреча содержат слишком много вложений.</span><span class="sxs-lookup"><span data-stu-id="9af0b-690">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="9af0b-691">Требования</span><span class="sxs-lookup"><span data-stu-id="9af0b-691">Requirements</span></span>

|<span data-ttu-id="9af0b-692">Требование</span><span class="sxs-lookup"><span data-stu-id="9af0b-692">Requirement</span></span>| <span data-ttu-id="9af0b-693">Значение</span><span class="sxs-lookup"><span data-stu-id="9af0b-693">Value</span></span>|
|---|---|
|[<span data-ttu-id="9af0b-694">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="9af0b-694">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9af0b-695">1.1</span><span class="sxs-lookup"><span data-stu-id="9af0b-695">1.1</span></span>|
|[<span data-ttu-id="9af0b-696">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="9af0b-696">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9af0b-697">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="9af0b-697">ReadWriteItem</span></span>|
|[<span data-ttu-id="9af0b-698">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9af0b-698">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9af0b-699">Создание</span><span class="sxs-lookup"><span data-stu-id="9af0b-699">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="9af0b-700">Пример</span><span class="sxs-lookup"><span data-stu-id="9af0b-700">Example</span></span>

<span data-ttu-id="9af0b-701">В следующем примере существующий элемент Outlook добавляется в виде вложения с именем `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="9af0b-701">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

####  <a name="close"></a><span data-ttu-id="9af0b-702">close()</span><span class="sxs-lookup"><span data-stu-id="9af0b-702">close()</span></span>

<span data-ttu-id="9af0b-703">Закрывает текущий создаваемый элемент.</span><span class="sxs-lookup"><span data-stu-id="9af0b-703">Closes the current item that is being composed.</span></span>

<span data-ttu-id="9af0b-p137">Работа метода `close` зависит от текущего состояния создаваемого элемента. Если элемент содержит несохраненные изменения, клиент предложит пользователю сохранить или отклонить их либо отменить действие закрытия.</span><span class="sxs-lookup"><span data-stu-id="9af0b-p137">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="9af0b-706">Если в Outlook в Интернете элемент представляет собой встречу, ранее сохраненную с помощью метода `saveAsync`, пользователю предлагается сохранить, отклонить или отменить действие, даже если с момента последнего сохранения элемента не вносились какие-либо изменения.</span><span class="sxs-lookup"><span data-stu-id="9af0b-706">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="9af0b-707">Если в клиенте Outlook для настольных ПК сообщение представляет собой ответ в тексте, метод `close` не работает.</span><span class="sxs-lookup"><span data-stu-id="9af0b-707">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="9af0b-708">Требования</span><span class="sxs-lookup"><span data-stu-id="9af0b-708">Requirements</span></span>

|<span data-ttu-id="9af0b-709">Требование</span><span class="sxs-lookup"><span data-stu-id="9af0b-709">Requirement</span></span>| <span data-ttu-id="9af0b-710">Значение</span><span class="sxs-lookup"><span data-stu-id="9af0b-710">Value</span></span>|
|---|---|
|[<span data-ttu-id="9af0b-711">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="9af0b-711">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9af0b-712">1.3</span><span class="sxs-lookup"><span data-stu-id="9af0b-712">1.3</span></span>|
|[<span data-ttu-id="9af0b-713">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="9af0b-713">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9af0b-714">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="9af0b-714">Restricted</span></span>|
|[<span data-ttu-id="9af0b-715">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9af0b-715">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9af0b-716">Создание</span><span class="sxs-lookup"><span data-stu-id="9af0b-716">Compose</span></span>|

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="9af0b-717">displayReplyAllForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="9af0b-717">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="9af0b-718">Отображает форму ответа, включающую отправителя и всех получателей выбранного сообщения или организатора и всех участников выбранной встречи.</span><span class="sxs-lookup"><span data-stu-id="9af0b-718">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="9af0b-719">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="9af0b-719">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="9af0b-720">В Outlook Web App форма ответа отображается в виде всплывающей формы в представлении с 3 либо 1 или 2 колонками.</span><span class="sxs-lookup"><span data-stu-id="9af0b-720">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="9af0b-721">Если любой строковый параметр превышает указанные для него ограничения, `displayReplyAllForm` возвращает исключение.</span><span class="sxs-lookup"><span data-stu-id="9af0b-721">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="9af0b-p138">Если в параметре `formData.attachments` указаны вложения, Outlook и Outlook Web App пытаются скачать их и вложить в форму ответа. Если какие-либо вложения добавить не удается, в форме отображается сообщение об ошибке. Если сообщения об ошибках не предусмотрены, то они не отображаются.</span><span class="sxs-lookup"><span data-stu-id="9af0b-p138">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9af0b-725">Параметры</span><span class="sxs-lookup"><span data-stu-id="9af0b-725">Parameters</span></span>

| <span data-ttu-id="9af0b-726">Имя</span><span class="sxs-lookup"><span data-stu-id="9af0b-726">Name</span></span> | <span data-ttu-id="9af0b-727">Тип</span><span class="sxs-lookup"><span data-stu-id="9af0b-727">Type</span></span> | <span data-ttu-id="9af0b-728">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="9af0b-728">Attributes</span></span> | <span data-ttu-id="9af0b-729">Описание</span><span class="sxs-lookup"><span data-stu-id="9af0b-729">Description</span></span> |
|---|---|---|---|
|`formData`| <span data-ttu-id="9af0b-730">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="9af0b-730">String &#124; Object</span></span>| |<span data-ttu-id="9af0b-p139">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="9af0b-p139">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="9af0b-733">**ИЛИ**</span><span class="sxs-lookup"><span data-stu-id="9af0b-733">**OR**</span></span><br/><span data-ttu-id="9af0b-p140">Объект, который содержит текст или данные вложения и функцию обратного вызова. Ниже представлено определение этого объекта.</span><span class="sxs-lookup"><span data-stu-id="9af0b-p140">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="9af0b-736">Строка</span><span class="sxs-lookup"><span data-stu-id="9af0b-736">String</span></span> | <span data-ttu-id="9af0b-737">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="9af0b-737">&lt;optional&gt;</span></span> | <span data-ttu-id="9af0b-p141">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="9af0b-p141">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="9af0b-740">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="9af0b-740">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="9af0b-741">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="9af0b-741">&lt;optional&gt;</span></span> | <span data-ttu-id="9af0b-742">Массив объектов JSON, представляющих собой вложенные файлы или элементы.</span><span class="sxs-lookup"><span data-stu-id="9af0b-742">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="9af0b-743">Строка</span><span class="sxs-lookup"><span data-stu-id="9af0b-743">String</span></span> | | <span data-ttu-id="9af0b-p142">Указывает тип вложения. Допустимые значения: `file` для вложенного файла и `item` для вложенного элемента.</span><span class="sxs-lookup"><span data-stu-id="9af0b-p142">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="9af0b-746">Строка</span><span class="sxs-lookup"><span data-stu-id="9af0b-746">String</span></span> | | <span data-ttu-id="9af0b-747">Строка, содержащая имя вложения, длиной до 255 символов.</span><span class="sxs-lookup"><span data-stu-id="9af0b-747">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="9af0b-748">Строка</span><span class="sxs-lookup"><span data-stu-id="9af0b-748">String</span></span> | | <span data-ttu-id="9af0b-p143">Используется, только если свойству `type` задано значение `file`. URI расположения файла.</span><span class="sxs-lookup"><span data-stu-id="9af0b-p143">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.isInline` | <span data-ttu-id="9af0b-751">Логический</span><span class="sxs-lookup"><span data-stu-id="9af0b-751">Boolean</span></span> | | <span data-ttu-id="9af0b-p144">Используется, только если свойству `type` задано значение `file`. Значение `true` указывает на то, что вложение будет встроено в текст сообщения и не должно отображаться в списке вложений.</span><span class="sxs-lookup"><span data-stu-id="9af0b-p144">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="9af0b-754">String</span><span class="sxs-lookup"><span data-stu-id="9af0b-754">String</span></span> | | <span data-ttu-id="9af0b-p145">Используется, только если свойству `type` задано значение `item`. Идентификатор вложения EWS. Это строка длиной до 100 символов.</span><span class="sxs-lookup"><span data-stu-id="9af0b-p145">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="9af0b-758">function</span><span class="sxs-lookup"><span data-stu-id="9af0b-758">function</span></span> | <span data-ttu-id="9af0b-759">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="9af0b-759">&lt;optional&gt;</span></span> | <span data-ttu-id="9af0b-760">По завершении работы метода функция, переданная параметру `callback`, вызывается с помощью одного параметра `asyncResult`, представляющего собой объект [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="9af0b-760">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="9af0b-761">Требования</span><span class="sxs-lookup"><span data-stu-id="9af0b-761">Requirements</span></span>

|<span data-ttu-id="9af0b-762">Требование</span><span class="sxs-lookup"><span data-stu-id="9af0b-762">Requirement</span></span>| <span data-ttu-id="9af0b-763">Значение</span><span class="sxs-lookup"><span data-stu-id="9af0b-763">Value</span></span>|
|---|---|
|[<span data-ttu-id="9af0b-764">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="9af0b-764">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9af0b-765">1.0</span><span class="sxs-lookup"><span data-stu-id="9af0b-765">1.0</span></span>|
|[<span data-ttu-id="9af0b-766">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="9af0b-766">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9af0b-767">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9af0b-767">ReadItem</span></span>|
|[<span data-ttu-id="9af0b-768">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9af0b-768">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9af0b-769">Чтение</span><span class="sxs-lookup"><span data-stu-id="9af0b-769">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="9af0b-770">Примеры</span><span class="sxs-lookup"><span data-stu-id="9af0b-770">Examples</span></span>

<span data-ttu-id="9af0b-771">Приведенный ниже код передает строку в функцию `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="9af0b-771">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="9af0b-772">Ответ с пустым текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="9af0b-772">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="9af0b-773">Ответ только с текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="9af0b-773">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="9af0b-774">Ответ с текстом сообщения и вложенным файлом.</span><span class="sxs-lookup"><span data-stu-id="9af0b-774">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="9af0b-775">Ответ с текстом сообщения и вложенным элементом.</span><span class="sxs-lookup"><span data-stu-id="9af0b-775">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="9af0b-776">Ответ с текстом сообщения, вложенным файлом, вложенным элементом и обратным вызовом.</span><span class="sxs-lookup"><span data-stu-id="9af0b-776">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="9af0b-777">displayReplyForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="9af0b-777">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="9af0b-778">Отображает форму ответа, включающую только отправителя выбранного сообщения или организатора выбранной встречи.</span><span class="sxs-lookup"><span data-stu-id="9af0b-778">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="9af0b-779">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="9af0b-779">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="9af0b-780">В Outlook Web App форма ответа отображается в виде всплывающей формы в представлении с 3 либо 1 или 2 колонками.</span><span class="sxs-lookup"><span data-stu-id="9af0b-780">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="9af0b-781">Если любой строковый параметр превышает указанные для него ограничения, `displayReplyForm` возвращает исключение.</span><span class="sxs-lookup"><span data-stu-id="9af0b-781">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="9af0b-p146">Если в параметре `formData.attachments` указаны вложения, Outlook и Outlook Web App пытаются скачать их и вложить в форму ответа. Если какие-либо вложения добавить не удается, в форме отображается сообщение об ошибке. Если сообщения об ошибках не предусмотрены, то они не отображаются.</span><span class="sxs-lookup"><span data-stu-id="9af0b-p146">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9af0b-785">Параметры</span><span class="sxs-lookup"><span data-stu-id="9af0b-785">Parameters</span></span>

| <span data-ttu-id="9af0b-786">Имя</span><span class="sxs-lookup"><span data-stu-id="9af0b-786">Name</span></span> | <span data-ttu-id="9af0b-787">Тип</span><span class="sxs-lookup"><span data-stu-id="9af0b-787">Type</span></span> | <span data-ttu-id="9af0b-788">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="9af0b-788">Attributes</span></span> | <span data-ttu-id="9af0b-789">Описание</span><span class="sxs-lookup"><span data-stu-id="9af0b-789">Description</span></span> |
|---|---|---|---|
|`formData`| <span data-ttu-id="9af0b-790">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="9af0b-790">String &#124; Object</span></span>| | <span data-ttu-id="9af0b-p147">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="9af0b-p147">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="9af0b-793">**ИЛИ**</span><span class="sxs-lookup"><span data-stu-id="9af0b-793">**OR**</span></span><br/><span data-ttu-id="9af0b-p148">Объект, который содержит текст или данные вложения и функцию обратного вызова. Ниже представлено определение этого объекта.</span><span class="sxs-lookup"><span data-stu-id="9af0b-p148">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="9af0b-796">Строка</span><span class="sxs-lookup"><span data-stu-id="9af0b-796">String</span></span> | <span data-ttu-id="9af0b-797">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="9af0b-797">&lt;optional&gt;</span></span> | <span data-ttu-id="9af0b-p149">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="9af0b-p149">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="9af0b-800">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="9af0b-800">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="9af0b-801">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="9af0b-801">&lt;optional&gt;</span></span> | <span data-ttu-id="9af0b-802">Массив объектов JSON, представляющих собой вложенные файлы или элементы.</span><span class="sxs-lookup"><span data-stu-id="9af0b-802">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="9af0b-803">Строка</span><span class="sxs-lookup"><span data-stu-id="9af0b-803">String</span></span> | | <span data-ttu-id="9af0b-p150">Указывает тип вложения. Допустимые значения: `file` для вложенного файла и `item` для вложенного элемента.</span><span class="sxs-lookup"><span data-stu-id="9af0b-p150">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="9af0b-806">Строка</span><span class="sxs-lookup"><span data-stu-id="9af0b-806">String</span></span> | | <span data-ttu-id="9af0b-807">Строка, содержащая имя вложения, длиной до 255 символов.</span><span class="sxs-lookup"><span data-stu-id="9af0b-807">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="9af0b-808">Строка</span><span class="sxs-lookup"><span data-stu-id="9af0b-808">String</span></span> | | <span data-ttu-id="9af0b-p151">Используется, только если свойству `type` задано значение `file`. URI расположения файла.</span><span class="sxs-lookup"><span data-stu-id="9af0b-p151">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.isInline` | <span data-ttu-id="9af0b-811">Логический</span><span class="sxs-lookup"><span data-stu-id="9af0b-811">Boolean</span></span> | | <span data-ttu-id="9af0b-p152">Используется, только если свойству `type` задано значение `file`. Значение `true` указывает на то, что вложение будет встроено в текст сообщения и не должно отображаться в списке вложений.</span><span class="sxs-lookup"><span data-stu-id="9af0b-p152">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="9af0b-814">String</span><span class="sxs-lookup"><span data-stu-id="9af0b-814">String</span></span> | | <span data-ttu-id="9af0b-p153">Используется, только если свойству `type` задано значение `item`. Идентификатор вложения EWS. Это строка длиной до 100 символов.</span><span class="sxs-lookup"><span data-stu-id="9af0b-p153">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="9af0b-818">function</span><span class="sxs-lookup"><span data-stu-id="9af0b-818">function</span></span> | <span data-ttu-id="9af0b-819">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="9af0b-819">&lt;optional&gt;</span></span> | <span data-ttu-id="9af0b-820">По завершении работы метода функция, переданная параметру `callback`, вызывается с помощью одного параметра `asyncResult`, представляющего собой объект [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="9af0b-820">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="9af0b-821">Требования</span><span class="sxs-lookup"><span data-stu-id="9af0b-821">Requirements</span></span>

|<span data-ttu-id="9af0b-822">Требование</span><span class="sxs-lookup"><span data-stu-id="9af0b-822">Requirement</span></span>| <span data-ttu-id="9af0b-823">Значение</span><span class="sxs-lookup"><span data-stu-id="9af0b-823">Value</span></span>|
|---|---|
|[<span data-ttu-id="9af0b-824">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="9af0b-824">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9af0b-825">1.0</span><span class="sxs-lookup"><span data-stu-id="9af0b-825">1.0</span></span>|
|[<span data-ttu-id="9af0b-826">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="9af0b-826">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9af0b-827">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9af0b-827">ReadItem</span></span>|
|[<span data-ttu-id="9af0b-828">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9af0b-828">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9af0b-829">Чтение</span><span class="sxs-lookup"><span data-stu-id="9af0b-829">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="9af0b-830">Примеры</span><span class="sxs-lookup"><span data-stu-id="9af0b-830">Examples</span></span>

<span data-ttu-id="9af0b-831">Приведенный ниже код передает строку в функцию `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="9af0b-831">The following code passes a string to the `displayReplyForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="9af0b-832">Ответ с пустым текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="9af0b-832">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="9af0b-833">Ответ только с текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="9af0b-833">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="9af0b-834">Ответ с текстом сообщения и вложенным файлом.</span><span class="sxs-lookup"><span data-stu-id="9af0b-834">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="9af0b-835">Ответ с текстом сообщения и вложенным элементом.</span><span class="sxs-lookup"><span data-stu-id="9af0b-835">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="9af0b-836">Ответ с текстом сообщения, вложенным файлом, вложенным элементом и обратным вызовом.</span><span class="sxs-lookup"><span data-stu-id="9af0b-836">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlook16officeentities"></a><span data-ttu-id="9af0b-837">getEntities() → {[Entities](/javascript/api/outlook_1_6/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="9af0b-837">getEntities() → {[Entities](/javascript/api/outlook_1_6/office.entities)}</span></span>

<span data-ttu-id="9af0b-838">Получает сущности, обнаруженные в теле выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="9af0b-838">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="9af0b-839">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="9af0b-839">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="9af0b-840">Требования</span><span class="sxs-lookup"><span data-stu-id="9af0b-840">Requirements</span></span>

|<span data-ttu-id="9af0b-841">Требование</span><span class="sxs-lookup"><span data-stu-id="9af0b-841">Requirement</span></span>| <span data-ttu-id="9af0b-842">Значение</span><span class="sxs-lookup"><span data-stu-id="9af0b-842">Value</span></span>|
|---|---|
|[<span data-ttu-id="9af0b-843">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="9af0b-843">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9af0b-844">1.0</span><span class="sxs-lookup"><span data-stu-id="9af0b-844">1.0</span></span>|
|[<span data-ttu-id="9af0b-845">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="9af0b-845">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9af0b-846">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9af0b-846">ReadItem</span></span>|
|[<span data-ttu-id="9af0b-847">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9af0b-847">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9af0b-848">Чтение</span><span class="sxs-lookup"><span data-stu-id="9af0b-848">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="9af0b-849">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="9af0b-849">Returns:</span></span>

<span data-ttu-id="9af0b-850">Тип: [Entities](/javascript/api/outlook_1_6/office.entities)</span><span class="sxs-lookup"><span data-stu-id="9af0b-850">Type: [Entities](/javascript/api/outlook_1_6/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="9af0b-851">Пример</span><span class="sxs-lookup"><span data-stu-id="9af0b-851">Example</span></span>

<span data-ttu-id="9af0b-852">Ниже приведен пример получения доступа к сущностям контактов в теле текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="9af0b-852">The following example accesses the contacts entities in the current item's body.</span></span>

```javascript
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook16officecontactmeetingsuggestionjavascriptapioutlook16officemeetingsuggestionphonenumberjavascriptapioutlook16officephonenumbertasksuggestionjavascriptapioutlook16officetasksuggestion"></a><span data-ttu-id="9af0b-853">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="9af0b-853">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))>}</span></span>

<span data-ttu-id="9af0b-854">Получает массив всех сущностей указанного типа, обнаруженных в теле выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="9af0b-854">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="9af0b-855">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="9af0b-855">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9af0b-856">Параметры</span><span class="sxs-lookup"><span data-stu-id="9af0b-856">Parameters</span></span>

|<span data-ttu-id="9af0b-857">Имя</span><span class="sxs-lookup"><span data-stu-id="9af0b-857">Name</span></span>| <span data-ttu-id="9af0b-858">Тип</span><span class="sxs-lookup"><span data-stu-id="9af0b-858">Type</span></span>| <span data-ttu-id="9af0b-859">Описание</span><span class="sxs-lookup"><span data-stu-id="9af0b-859">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="9af0b-860">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="9af0b-860">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook_1_6/office.mailboxenums.entitytype)|<span data-ttu-id="9af0b-861">Одно из значений перечисления EntityType.</span><span class="sxs-lookup"><span data-stu-id="9af0b-861">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9af0b-862">Требования</span><span class="sxs-lookup"><span data-stu-id="9af0b-862">Requirements</span></span>

|<span data-ttu-id="9af0b-863">Требование</span><span class="sxs-lookup"><span data-stu-id="9af0b-863">Requirement</span></span>| <span data-ttu-id="9af0b-864">Значение</span><span class="sxs-lookup"><span data-stu-id="9af0b-864">Value</span></span>|
|---|---|
|[<span data-ttu-id="9af0b-865">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="9af0b-865">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9af0b-866">1.0</span><span class="sxs-lookup"><span data-stu-id="9af0b-866">1.0</span></span>|
|[<span data-ttu-id="9af0b-867">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="9af0b-867">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9af0b-868">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="9af0b-868">Restricted</span></span>|
|[<span data-ttu-id="9af0b-869">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9af0b-869">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9af0b-870">Чтение</span><span class="sxs-lookup"><span data-stu-id="9af0b-870">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="9af0b-871">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="9af0b-871">Returns:</span></span>

<span data-ttu-id="9af0b-872">Если значение, переданное в `entityType`, не является допустимым членом перечисления `EntityType`, метод возвращает значение NULL.</span><span class="sxs-lookup"><span data-stu-id="9af0b-872">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="9af0b-873">Если в теле элемента отсутствуют сущности указанного типа, метод возвращает пустой массив.</span><span class="sxs-lookup"><span data-stu-id="9af0b-873">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="9af0b-874">В противном случае тип объектов в возвращаемом массиве зависит от типа сущности, запрошенной в параметре `entityType`.</span><span class="sxs-lookup"><span data-stu-id="9af0b-874">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="9af0b-875">Хотя минимальный уровень разрешений для использования этого метода — **Restricted**, для некоторых типов сущностей требуется доступ на уровне **ReadItem**, как указано в приведенной ниже таблице.</span><span class="sxs-lookup"><span data-stu-id="9af0b-875">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="9af0b-876">Значение параметра `entityType`</span><span class="sxs-lookup"><span data-stu-id="9af0b-876">Value of `entityType`</span></span> | <span data-ttu-id="9af0b-877">Тип объектов в возвращаемом массиве</span><span class="sxs-lookup"><span data-stu-id="9af0b-877">Type of objects in returned array</span></span> | <span data-ttu-id="9af0b-878">Необходимый уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="9af0b-878">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="9af0b-879">String</span><span class="sxs-lookup"><span data-stu-id="9af0b-879">String</span></span> | <span data-ttu-id="9af0b-880">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="9af0b-880">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="9af0b-881">Contact</span><span class="sxs-lookup"><span data-stu-id="9af0b-881">Contact</span></span> | <span data-ttu-id="9af0b-882">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="9af0b-882">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="9af0b-883">String</span><span class="sxs-lookup"><span data-stu-id="9af0b-883">String</span></span> | <span data-ttu-id="9af0b-884">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="9af0b-884">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="9af0b-885">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="9af0b-885">MeetingSuggestion</span></span> | <span data-ttu-id="9af0b-886">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="9af0b-886">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="9af0b-887">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="9af0b-887">PhoneNumber</span></span> | <span data-ttu-id="9af0b-888">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="9af0b-888">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="9af0b-889">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="9af0b-889">TaskSuggestion</span></span> | <span data-ttu-id="9af0b-890">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="9af0b-890">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="9af0b-891">String</span><span class="sxs-lookup"><span data-stu-id="9af0b-891">String</span></span> | <span data-ttu-id="9af0b-892">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="9af0b-892">**Restricted**</span></span> |

<span data-ttu-id="9af0b-893">Тип: Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="9af0b-893">Type: Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="9af0b-894">Пример</span><span class="sxs-lookup"><span data-stu-id="9af0b-894">Example</span></span>

<span data-ttu-id="9af0b-895">В примере ниже показано, как получить доступ к массиву строк, которые представляют собой почтовые адреса в теле текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="9af0b-895">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook16officecontactmeetingsuggestionjavascriptapioutlook16officemeetingsuggestionphonenumberjavascriptapioutlook16officephonenumbertasksuggestionjavascriptapioutlook16officetasksuggestion"></a><span data-ttu-id="9af0b-896">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="9af0b-896">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))>}</span></span>

<span data-ttu-id="9af0b-897">Возвращает известные сущности в выбранном элементе, которые проходят через именованный фильтр, определяемый в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="9af0b-897">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="9af0b-898">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="9af0b-898">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="9af0b-899">Метод `getFilteredEntitiesByName` возвращает сущности, соответствующие регулярному выражению, которое определяется в элементе правила [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) в XML-файле манифеста, с использованием указанного значения элемента `FilterName`.</span><span class="sxs-lookup"><span data-stu-id="9af0b-899">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9af0b-900">Параметры</span><span class="sxs-lookup"><span data-stu-id="9af0b-900">Parameters</span></span>

|<span data-ttu-id="9af0b-901">Имя</span><span class="sxs-lookup"><span data-stu-id="9af0b-901">Name</span></span>| <span data-ttu-id="9af0b-902">Тип</span><span class="sxs-lookup"><span data-stu-id="9af0b-902">Type</span></span>| <span data-ttu-id="9af0b-903">Описание</span><span class="sxs-lookup"><span data-stu-id="9af0b-903">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="9af0b-904">Строка</span><span class="sxs-lookup"><span data-stu-id="9af0b-904">String</span></span>|<span data-ttu-id="9af0b-905">Имя элемента правила `ItemHasKnownEntity`, определяющее соответствующий фильтр.</span><span class="sxs-lookup"><span data-stu-id="9af0b-905">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9af0b-906">Требования</span><span class="sxs-lookup"><span data-stu-id="9af0b-906">Requirements</span></span>

|<span data-ttu-id="9af0b-907">Требование</span><span class="sxs-lookup"><span data-stu-id="9af0b-907">Requirement</span></span>| <span data-ttu-id="9af0b-908">Значение</span><span class="sxs-lookup"><span data-stu-id="9af0b-908">Value</span></span>|
|---|---|
|[<span data-ttu-id="9af0b-909">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="9af0b-909">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9af0b-910">1.0</span><span class="sxs-lookup"><span data-stu-id="9af0b-910">1.0</span></span>|
|[<span data-ttu-id="9af0b-911">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="9af0b-911">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9af0b-912">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9af0b-912">ReadItem</span></span>|
|[<span data-ttu-id="9af0b-913">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9af0b-913">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9af0b-914">Чтение</span><span class="sxs-lookup"><span data-stu-id="9af0b-914">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="9af0b-915">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="9af0b-915">Returns:</span></span>

<span data-ttu-id="9af0b-p155">Если в манифесте нет элемента `ItemHasKnownEntity` со значением `FilterName`, соответствующим параметру `name`, метод возвращает `null`. Если параметр `name` соответствует элементу `ItemHasKnownEntity` в манифесте, но при этом в текущем элементе нет соответствующих сущностей, метод возвращает пустой массив.</span><span class="sxs-lookup"><span data-stu-id="9af0b-p155">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="9af0b-918">Тип: Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="9af0b-918">Type: Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))></span></span>

#### <a name="getregexmatches--object"></a><span data-ttu-id="9af0b-919">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="9af0b-919">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="9af0b-920">Возвращает строковые значения в выбранном элементе, которые соответствуют регулярным выражениям, определенным в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="9af0b-920">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="9af0b-921">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="9af0b-921">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="9af0b-p156">Метод `getRegExMatches` возвращает строки, соответствующие регулярному выражению, которое определяется в каждом элементе правила `ItemHasRegularExpressionMatch` или `ItemHasKnownEntity` в XML-файле манифеста. Для правила `ItemHasRegularExpressionMatch` соответствующую строку должно содержать свойство элемента, указанного этим правилом. Простой тип `PropertyName` определяет поддерживаемые свойства.</span><span class="sxs-lookup"><span data-stu-id="9af0b-p156">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="9af0b-925">Например, рассмотрим манифест надстройки, который содержит указанный ниже элемент `Rule`.</span><span class="sxs-lookup"><span data-stu-id="9af0b-925">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="9af0b-926">Объект, возвращаемый методом `getRegExMatches`, будет содержать два свойства: `fruits` и `veggies`.</span><span class="sxs-lookup"><span data-stu-id="9af0b-926">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="9af0b-p157">Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты. Лучше используйте метод [`Body.getAsync`](/javascript/api/outlook_1_6/office.body#getasync-coerciontype--options--callback-) для этого.</span><span class="sxs-lookup"><span data-stu-id="9af0b-p157">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook_1_6/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="9af0b-930">Requirements</span><span class="sxs-lookup"><span data-stu-id="9af0b-930">Requirements</span></span>

|<span data-ttu-id="9af0b-931">Требование</span><span class="sxs-lookup"><span data-stu-id="9af0b-931">Requirement</span></span>| <span data-ttu-id="9af0b-932">Значение</span><span class="sxs-lookup"><span data-stu-id="9af0b-932">Value</span></span>|
|---|---|
|[<span data-ttu-id="9af0b-933">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="9af0b-933">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9af0b-934">1.0</span><span class="sxs-lookup"><span data-stu-id="9af0b-934">1.0</span></span>|
|[<span data-ttu-id="9af0b-935">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="9af0b-935">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9af0b-936">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9af0b-936">ReadItem</span></span>|
|[<span data-ttu-id="9af0b-937">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9af0b-937">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9af0b-938">Чтение</span><span class="sxs-lookup"><span data-stu-id="9af0b-938">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="9af0b-939">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="9af0b-939">Returns:</span></span>

<span data-ttu-id="9af0b-p158">Объект, содержащий массив строк, которые соответствуют регулярным выражениям, определяемым в XML-файле манифеста. Имя каждого массива равно соответствующему значению атрибута `RegExName` подходящего правила `ItemHasRegularExpressionMatch` или атрибута `FilterName` соответствующего правила `ItemHasKnownEntity`.</span><span class="sxs-lookup"><span data-stu-id="9af0b-p158">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="9af0b-942">Тип:</span><span class="sxs-lookup"><span data-stu-id="9af0b-942">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="9af0b-943">Object</span><span class="sxs-lookup"><span data-stu-id="9af0b-943">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="9af0b-944">Пример</span><span class="sxs-lookup"><span data-stu-id="9af0b-944">Example</span></span>

<span data-ttu-id="9af0b-945">В приведенном ниже примере показано, как получить доступ к массиву совпадений с элементами `fruits` и `veggies` правил активации регулярных выражений, указанными в манифесте.</span><span class="sxs-lookup"><span data-stu-id="9af0b-945">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```javascript
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="9af0b-946">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="9af0b-946">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="9af0b-947">Возвращает строковые значения в выбранном элементе, которые соответствуют именованному регулярному выражению, определенному в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="9af0b-947">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="9af0b-948">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="9af0b-948">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="9af0b-949">Метод `getRegExMatchesByName` возвращает строки, соответствующие регулярному выражению, которое определяется в элементе правила `ItemHasRegularExpressionMatch` в XML-файле манифеста, с использованием указанного значения элемента `RegExName`.</span><span class="sxs-lookup"><span data-stu-id="9af0b-949">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="9af0b-p159">Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты.</span><span class="sxs-lookup"><span data-stu-id="9af0b-p159">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9af0b-952">Параметры</span><span class="sxs-lookup"><span data-stu-id="9af0b-952">Parameters</span></span>

|<span data-ttu-id="9af0b-953">Имя</span><span class="sxs-lookup"><span data-stu-id="9af0b-953">Name</span></span>| <span data-ttu-id="9af0b-954">Тип</span><span class="sxs-lookup"><span data-stu-id="9af0b-954">Type</span></span>| <span data-ttu-id="9af0b-955">Описание</span><span class="sxs-lookup"><span data-stu-id="9af0b-955">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="9af0b-956">Строка</span><span class="sxs-lookup"><span data-stu-id="9af0b-956">String</span></span>|<span data-ttu-id="9af0b-957">Имя элемента правила `ItemHasRegularExpressionMatch`, определяющее соответствующий фильтр.</span><span class="sxs-lookup"><span data-stu-id="9af0b-957">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9af0b-958">Требования</span><span class="sxs-lookup"><span data-stu-id="9af0b-958">Requirements</span></span>

|<span data-ttu-id="9af0b-959">Требование</span><span class="sxs-lookup"><span data-stu-id="9af0b-959">Requirement</span></span>| <span data-ttu-id="9af0b-960">Значение</span><span class="sxs-lookup"><span data-stu-id="9af0b-960">Value</span></span>|
|---|---|
|[<span data-ttu-id="9af0b-961">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="9af0b-961">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9af0b-962">1.0</span><span class="sxs-lookup"><span data-stu-id="9af0b-962">1.0</span></span>|
|[<span data-ttu-id="9af0b-963">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="9af0b-963">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9af0b-964">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9af0b-964">ReadItem</span></span>|
|[<span data-ttu-id="9af0b-965">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9af0b-965">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9af0b-966">Чтение</span><span class="sxs-lookup"><span data-stu-id="9af0b-966">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="9af0b-967">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="9af0b-967">Returns:</span></span>

<span data-ttu-id="9af0b-968">Массив строк, соответствующих регулярному выражению, определяемому в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="9af0b-968">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="9af0b-969">

<dt>Тип</dt>

</span><span class="sxs-lookup"><span data-stu-id="9af0b-969">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="9af0b-970">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="9af0b-970">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="9af0b-971">Пример</span><span class="sxs-lookup"><span data-stu-id="9af0b-971">Example</span></span>

```javascript
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="9af0b-972">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="9af0b-972">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="9af0b-973">Асинхронно возвращает данные, выбранные в теме или тексте сообщения.</span><span class="sxs-lookup"><span data-stu-id="9af0b-973">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="9af0b-p160">Если выделенный фрагмент отсутствует, но курсор находится в тексте или теме, метод возвращает значение NULL для выбранных данных. Если выбраны не текст и не тема, метод возвращает ошибку `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="9af0b-p160">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9af0b-976">Параметры</span><span class="sxs-lookup"><span data-stu-id="9af0b-976">Parameters</span></span>

|<span data-ttu-id="9af0b-977">Имя</span><span class="sxs-lookup"><span data-stu-id="9af0b-977">Name</span></span>| <span data-ttu-id="9af0b-978">Тип</span><span class="sxs-lookup"><span data-stu-id="9af0b-978">Type</span></span>| <span data-ttu-id="9af0b-979">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="9af0b-979">Attributes</span></span>| <span data-ttu-id="9af0b-980">Описание</span><span class="sxs-lookup"><span data-stu-id="9af0b-980">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="9af0b-981">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="9af0b-981">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="9af0b-p161">Запрашивает формат данных. Если задано значение Text, метод возвращает обычный текст как строку, удаляя все имеющиеся HTML-теги. Если задано значение HTML, метод возвращает выделенный текст (обычный текст или HTML).</span><span class="sxs-lookup"><span data-stu-id="9af0b-p161">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="9af0b-985">Object</span><span class="sxs-lookup"><span data-stu-id="9af0b-985">Object</span></span>| <span data-ttu-id="9af0b-986">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="9af0b-986">&lt;optional&gt;</span></span>|<span data-ttu-id="9af0b-987">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="9af0b-987">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="9af0b-988">Объект</span><span class="sxs-lookup"><span data-stu-id="9af0b-988">Object</span></span>| <span data-ttu-id="9af0b-989">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="9af0b-989">&lt;optional&gt;</span></span>|<span data-ttu-id="9af0b-990">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="9af0b-990">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="9af0b-991">функция</span><span class="sxs-lookup"><span data-stu-id="9af0b-991">function</span></span>||<span data-ttu-id="9af0b-992">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="9af0b-992">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="9af0b-993">Чтобы получить доступ к выбранным данным из метода обратного вызова, вызовите `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="9af0b-993">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="9af0b-994">Для доступа к исходному свойству, представляющему собой источник выбранных данных, вызовите параметр `asyncResult.value.sourceProperty`, который может иметь значение `body` или `subject`.</span><span class="sxs-lookup"><span data-stu-id="9af0b-994">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9af0b-995">Требования</span><span class="sxs-lookup"><span data-stu-id="9af0b-995">Requirements</span></span>

|<span data-ttu-id="9af0b-996">Требование</span><span class="sxs-lookup"><span data-stu-id="9af0b-996">Requirement</span></span>| <span data-ttu-id="9af0b-997">Значение</span><span class="sxs-lookup"><span data-stu-id="9af0b-997">Value</span></span>|
|---|---|
|[<span data-ttu-id="9af0b-998">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="9af0b-998">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9af0b-999">1.2</span><span class="sxs-lookup"><span data-stu-id="9af0b-999">1.2</span></span>|
|[<span data-ttu-id="9af0b-1000">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="9af0b-1000">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9af0b-1001">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="9af0b-1001">ReadWriteItem</span></span>|
|[<span data-ttu-id="9af0b-1002">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9af0b-1002">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9af0b-1003">Создание</span><span class="sxs-lookup"><span data-stu-id="9af0b-1003">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="9af0b-1004">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="9af0b-1004">Returns:</span></span>

<span data-ttu-id="9af0b-1005">Выбранные данные в виде строки с форматом, определенным в параметре `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="9af0b-1005">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="9af0b-1006">

<dt>Тип</dt>

</span><span class="sxs-lookup"><span data-stu-id="9af0b-1006">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="9af0b-1007">String</span><span class="sxs-lookup"><span data-stu-id="9af0b-1007">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="9af0b-1008">Пример</span><span class="sxs-lookup"><span data-stu-id="9af0b-1008">Example</span></span>

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

#### <a name="getselectedentities--entitiesjavascriptapioutlook16officeentities"></a><span data-ttu-id="9af0b-1009">getSelectedEntities() → {[Entities](/javascript/api/outlook_1_6/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="9af0b-1009">getSelectedEntities() → {[Entities](/javascript/api/outlook_1_6/office.entities)}</span></span>

<span data-ttu-id="9af0b-1010">Возвращает сущности, найденные в выделенном совпадении, выбранном пользователем.</span><span class="sxs-lookup"><span data-stu-id="9af0b-1010">Gets the entities found in a highlighted match a user has selected.</span></span> <span data-ttu-id="9af0b-1011">Выделенные совпадения применяются к [контекстным надстройкам](/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="9af0b-1011">Highlighted matches apply to [contextual add-ins](/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="9af0b-1012">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="9af0b-1012">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="9af0b-1013">Требования</span><span class="sxs-lookup"><span data-stu-id="9af0b-1013">Requirements</span></span>

|<span data-ttu-id="9af0b-1014">Требование</span><span class="sxs-lookup"><span data-stu-id="9af0b-1014">Requirement</span></span>| <span data-ttu-id="9af0b-1015">Значение</span><span class="sxs-lookup"><span data-stu-id="9af0b-1015">Value</span></span>|
|---|---|
|[<span data-ttu-id="9af0b-1016">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="9af0b-1016">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9af0b-1017">1.6</span><span class="sxs-lookup"><span data-stu-id="9af0b-1017">1.6</span></span> |
|[<span data-ttu-id="9af0b-1018">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="9af0b-1018">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9af0b-1019">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9af0b-1019">ReadItem</span></span>|
|[<span data-ttu-id="9af0b-1020">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9af0b-1020">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9af0b-1021">Чтение</span><span class="sxs-lookup"><span data-stu-id="9af0b-1021">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="9af0b-1022">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="9af0b-1022">Returns:</span></span>

<span data-ttu-id="9af0b-1023">Тип: [Entities](/javascript/api/outlook_1_6/office.entities)</span><span class="sxs-lookup"><span data-stu-id="9af0b-1023">Type: [Entities](/javascript/api/outlook_1_6/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="9af0b-1024">Пример</span><span class="sxs-lookup"><span data-stu-id="9af0b-1024">Example</span></span>

<span data-ttu-id="9af0b-1025">В приведенном ниже примере показано, как получить доступ к сущностям адресов в выделенном совпадении, выбранном пользователем.</span><span class="sxs-lookup"><span data-stu-id="9af0b-1025">The following example accesses the addresses entities in the highlighted match selected by the user.</span></span>

```javascript
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

#### <a name="getselectedregexmatches--object"></a><span data-ttu-id="9af0b-1026">getSelectedRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="9af0b-1026">getSelectedRegExMatches() → {Object}</span></span>

<span data-ttu-id="9af0b-p164">Возвращает строковые значения в выделенном совпадении, которые соответствуют регулярным выражениям, определенным в XML-файле манифеста. Выделенные совпадения применяются к [контекстным надстройкам](/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="9af0b-p164">Returns string values in a highlighted match that match the regular expressions defined in the manifest XML file. Highlighted matches apply to [contextual add-ins](/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="9af0b-1029">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="9af0b-1029">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="9af0b-p165">Метод `getSelectedRegExMatches` возвращает строки, соответствующие регулярному выражению, которое определяется в каждом элементе правила `ItemHasRegularExpressionMatch` или `ItemHasKnownEntity` в XML-файле манифеста. Для правила `ItemHasRegularExpressionMatch` соответствующую строку должно содержать свойство элемента, указанного этим правилом. Простой тип `PropertyName` определяет поддерживаемые свойства.</span><span class="sxs-lookup"><span data-stu-id="9af0b-p165">The `getSelectedRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="9af0b-1033">Например, рассмотрим манифест надстройки, который содержит указанный ниже элемент `Rule`.</span><span class="sxs-lookup"><span data-stu-id="9af0b-1033">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="9af0b-1034">Объект, возвращаемый методом `getRegExMatches`, будет содержать два свойства: `fruits` и `veggies`.</span><span class="sxs-lookup"><span data-stu-id="9af0b-1034">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="9af0b-p166">Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты. Лучше используйте метод [`Body.getAsync`](/javascript/api/outlook_1_6/office.body#getasync-coerciontype--options--callback-) для этого.</span><span class="sxs-lookup"><span data-stu-id="9af0b-p166">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook_1_6/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="9af0b-1038">Requirements</span><span class="sxs-lookup"><span data-stu-id="9af0b-1038">Requirements</span></span>

|<span data-ttu-id="9af0b-1039">Требование</span><span class="sxs-lookup"><span data-stu-id="9af0b-1039">Requirement</span></span>| <span data-ttu-id="9af0b-1040">Значение</span><span class="sxs-lookup"><span data-stu-id="9af0b-1040">Value</span></span>|
|---|---|
|[<span data-ttu-id="9af0b-1041">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="9af0b-1041">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9af0b-1042">1.6</span><span class="sxs-lookup"><span data-stu-id="9af0b-1042">1.6</span></span> |
|[<span data-ttu-id="9af0b-1043">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="9af0b-1043">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9af0b-1044">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9af0b-1044">ReadItem</span></span>|
|[<span data-ttu-id="9af0b-1045">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9af0b-1045">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9af0b-1046">Чтение</span><span class="sxs-lookup"><span data-stu-id="9af0b-1046">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="9af0b-1047">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="9af0b-1047">Returns:</span></span>

<span data-ttu-id="9af0b-p167">Объект, содержащий массив строк, которые соответствуют регулярным выражениям, определяемым в XML-файле манифеста. Имя каждого массива равно соответствующему значению атрибута `RegExName` подходящего правила `ItemHasRegularExpressionMatch` или атрибута `FilterName` соответствующего правила `ItemHasKnownEntity`.</span><span class="sxs-lookup"><span data-stu-id="9af0b-p167">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

##### <a name="example"></a><span data-ttu-id="9af0b-1050">Пример</span><span class="sxs-lookup"><span data-stu-id="9af0b-1050">Example</span></span>

<span data-ttu-id="9af0b-1051">В приведенном ниже примере показано, как получить доступ к массиву совпадений с элементами `fruits` и `veggies` правил активации регулярных выражений, указанными в манифесте.</span><span class="sxs-lookup"><span data-stu-id="9af0b-1051">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```javascript
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="9af0b-1052">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="9af0b-1052">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="9af0b-1053">Асинхронно загружает настраиваемые свойства для надстройки для выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="9af0b-1053">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="9af0b-p168">Настраиваемые свойства сохраняются в виде пар "ключ-значение" для каждого приложения и каждого элемента. Этот метод возвращает объект `CustomProperties` при обратном вызове, который предоставляет методы для доступа к настраиваемым свойствам, характерным для текущего элемента и текущей надстройки. Настраиваемые свойства не шифруются для элемента, поэтому этот способ хранения не является безопасным.</span><span class="sxs-lookup"><span data-stu-id="9af0b-p168">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9af0b-1057">Параметры</span><span class="sxs-lookup"><span data-stu-id="9af0b-1057">Parameters</span></span>

|<span data-ttu-id="9af0b-1058">Имя</span><span class="sxs-lookup"><span data-stu-id="9af0b-1058">Name</span></span>| <span data-ttu-id="9af0b-1059">Тип</span><span class="sxs-lookup"><span data-stu-id="9af0b-1059">Type</span></span>| <span data-ttu-id="9af0b-1060">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="9af0b-1060">Attributes</span></span>| <span data-ttu-id="9af0b-1061">Описание</span><span class="sxs-lookup"><span data-stu-id="9af0b-1061">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="9af0b-1062">function</span><span class="sxs-lookup"><span data-stu-id="9af0b-1062">function</span></span>||<span data-ttu-id="9af0b-1063">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="9af0b-1063">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="9af0b-1064">Настраиваемые свойства предоставляются в виде объекта [`CustomProperties`](/javascript/api/outlook_1_6/office.customproperties) в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="9af0b-1064">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook_1_6/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="9af0b-1065">Этот объект позволяет получить, задать и удалить настраиваемые свойства для элемента, а также сохранить изменения, внесенные в набор настраиваемых свойств, на сервере.</span><span class="sxs-lookup"><span data-stu-id="9af0b-1065">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="9af0b-1066">Объект</span><span class="sxs-lookup"><span data-stu-id="9af0b-1066">Object</span></span>| <span data-ttu-id="9af0b-1067">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="9af0b-1067">&lt;optional&gt;</span></span>|<span data-ttu-id="9af0b-1068">Разработчики могут указать любой объект, к которому необходимо получить доступ, в функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="9af0b-1068">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="9af0b-1069">Доступ к этому объекту можно получить с помощью свойства `asyncResult.asyncContext` в функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="9af0b-1069">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9af0b-1070">Требования</span><span class="sxs-lookup"><span data-stu-id="9af0b-1070">Requirements</span></span>

|<span data-ttu-id="9af0b-1071">Требование</span><span class="sxs-lookup"><span data-stu-id="9af0b-1071">Requirement</span></span>| <span data-ttu-id="9af0b-1072">Значение</span><span class="sxs-lookup"><span data-stu-id="9af0b-1072">Value</span></span>|
|---|---|
|[<span data-ttu-id="9af0b-1073">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="9af0b-1073">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9af0b-1074">1.0</span><span class="sxs-lookup"><span data-stu-id="9af0b-1074">1.0</span></span>|
|[<span data-ttu-id="9af0b-1075">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="9af0b-1075">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9af0b-1076">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9af0b-1076">ReadItem</span></span>|
|[<span data-ttu-id="9af0b-1077">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9af0b-1077">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9af0b-1078">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="9af0b-1078">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9af0b-1079">Пример</span><span class="sxs-lookup"><span data-stu-id="9af0b-1079">Example</span></span>

<span data-ttu-id="9af0b-p171">Приведенный ниже пример кода показывает, как асинхронно загружать настраиваемые свойства, характерные для текущего элемента, с помощью метода `loadCustomPropertiesAsync`. Этот пример также показывает, как сохранять эти свойства на сервере с помощью метода `CustomProperties.saveAsync`. После загрузки настраиваемых свойств в этом примере кода метод `CustomProperties.get` используется для считывания настраиваемого свойства `myProp`, метод `CustomProperties.set` — для записи настраиваемого свойства `otherProp`, а метод `saveAsync` — для сохранения настраиваемых свойств.</span><span class="sxs-lookup"><span data-stu-id="9af0b-p171">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="9af0b-1083">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="9af0b-1083">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="9af0b-1084">Удаляет вложение из сообщения или встречи.</span><span class="sxs-lookup"><span data-stu-id="9af0b-1084">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="9af0b-p172">Метод `removeAttachmentAsync` удаляет из элемента вложение с указанным идентификатором. Идентификатор вложения рекомендуется использовать для удаления вложения, только если оно добавлено тем же почтовым приложением в ходе текущего сеанса. В Outlook Web App и Outlook Web App для устройств идентификатор вложения действителен только в рамках одного сеанса. Сеанс завершается, когда пользователь закрывает приложение или начинает создавать элемент во встроенной форме, а затем переходит из формы в отдельное окно.</span><span class="sxs-lookup"><span data-stu-id="9af0b-p172">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9af0b-1089">Параметры</span><span class="sxs-lookup"><span data-stu-id="9af0b-1089">Parameters</span></span>

|<span data-ttu-id="9af0b-1090">Имя</span><span class="sxs-lookup"><span data-stu-id="9af0b-1090">Name</span></span>| <span data-ttu-id="9af0b-1091">Тип</span><span class="sxs-lookup"><span data-stu-id="9af0b-1091">Type</span></span>| <span data-ttu-id="9af0b-1092">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="9af0b-1092">Attributes</span></span>| <span data-ttu-id="9af0b-1093">Описание</span><span class="sxs-lookup"><span data-stu-id="9af0b-1093">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="9af0b-1094">Строка</span><span class="sxs-lookup"><span data-stu-id="9af0b-1094">String</span></span>||<span data-ttu-id="9af0b-1095">Идентификатор удаляемого вложения.</span><span class="sxs-lookup"><span data-stu-id="9af0b-1095">The identifier of the attachment to remove.</span></span>|
|`options`| <span data-ttu-id="9af0b-1096">Object</span><span class="sxs-lookup"><span data-stu-id="9af0b-1096">Object</span></span>| <span data-ttu-id="9af0b-1097">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="9af0b-1097">&lt;optional&gt;</span></span>|<span data-ttu-id="9af0b-1098">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="9af0b-1098">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="9af0b-1099">Object</span><span class="sxs-lookup"><span data-stu-id="9af0b-1099">Object</span></span>| <span data-ttu-id="9af0b-1100">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="9af0b-1100">&lt;optional&gt;</span></span>|<span data-ttu-id="9af0b-1101">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="9af0b-1101">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="9af0b-1102">функция</span><span class="sxs-lookup"><span data-stu-id="9af0b-1102">function</span></span>| <span data-ttu-id="9af0b-1103">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="9af0b-1103">&lt;optional&gt;</span></span>|<span data-ttu-id="9af0b-1104">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="9af0b-1104">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="9af0b-1105">Если удалить вложение не удается, свойство `asyncResult.error` содержит код ошибки с указанием ее причины.</span><span class="sxs-lookup"><span data-stu-id="9af0b-1105">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="9af0b-1106">Ошибки</span><span class="sxs-lookup"><span data-stu-id="9af0b-1106">Errors</span></span>

| <span data-ttu-id="9af0b-1107">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="9af0b-1107">Error code</span></span> | <span data-ttu-id="9af0b-1108">Описание</span><span class="sxs-lookup"><span data-stu-id="9af0b-1108">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="9af0b-1109">Идентификатор вложения не существует.</span><span class="sxs-lookup"><span data-stu-id="9af0b-1109">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="9af0b-1110">Требования</span><span class="sxs-lookup"><span data-stu-id="9af0b-1110">Requirements</span></span>

|<span data-ttu-id="9af0b-1111">Требование</span><span class="sxs-lookup"><span data-stu-id="9af0b-1111">Requirement</span></span>| <span data-ttu-id="9af0b-1112">Значение</span><span class="sxs-lookup"><span data-stu-id="9af0b-1112">Value</span></span>|
|---|---|
|[<span data-ttu-id="9af0b-1113">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="9af0b-1113">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9af0b-1114">1.1</span><span class="sxs-lookup"><span data-stu-id="9af0b-1114">1.1</span></span>|
|[<span data-ttu-id="9af0b-1115">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="9af0b-1115">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9af0b-1116">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="9af0b-1116">ReadWriteItem</span></span>|
|[<span data-ttu-id="9af0b-1117">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9af0b-1117">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9af0b-1118">Создание</span><span class="sxs-lookup"><span data-stu-id="9af0b-1118">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="9af0b-1119">Пример</span><span class="sxs-lookup"><span data-stu-id="9af0b-1119">Example</span></span>

<span data-ttu-id="9af0b-1120">Указанный ниже код удаляет вложение с идентификатором "0".</span><span class="sxs-lookup"><span data-stu-id="9af0b-1120">The following code removes an attachment with an identifier of '0'.</span></span>

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

####  <a name="saveasyncoptions-callback"></a><span data-ttu-id="9af0b-1121">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="9af0b-1121">saveAsync([options], callback)</span></span>

<span data-ttu-id="9af0b-1122">Асинхронно сохраняет элемент.</span><span class="sxs-lookup"><span data-stu-id="9af0b-1122">Asynchronously saves an item.</span></span>

<span data-ttu-id="9af0b-p173">При вызове этот метод сохраняет текущее сообщение в виде черновика и возвращает идентификатор элемента с помощью метода обратного вызова. В Outlook Web App или интерактивном режиме Outlook этот элемент сохраняется на сервере. В Outlook в режиме кэширования этот элемент сохраняется в локальном кэше.</span><span class="sxs-lookup"><span data-stu-id="9af0b-p173">When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="9af0b-1126">Если в режиме создания надстройка вызывает для элемента метод `saveAsync`, чтобы получить параметр `itemId` для использования с EWS или REST API, необходимо помнить, что синхронизация элемента с сервером может занять много времени, если Outlook работает в режиме кэширования данных.</span><span class="sxs-lookup"><span data-stu-id="9af0b-1126">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="9af0b-1127">До окончания синхронизации элемента применение параметра `itemId` будет приводить к ошибке.</span><span class="sxs-lookup"><span data-stu-id="9af0b-1127">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="9af0b-p175">Если метод `saveAsync` вызывается для встречи в режиме создания, она сохраняется как обычная встреча в календаре пользователя, а не как черновик. При сохранении новой встречи приглашения не отправляются. При сохранении существующей встречи уведомления отправляются добавленным или удаленным участникам.</span><span class="sxs-lookup"><span data-stu-id="9af0b-p175">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="9af0b-1131">Следующие клиенты отличаются другим поведением `saveAsync` в отношении встреч в режиме создания:</span><span class="sxs-lookup"><span data-stu-id="9af0b-1131">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="9af0b-1132">Outlook для Mac не поддерживает `saveAsync` для собраний в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="9af0b-1132">Mac Outlook does not support `saveAsync` on a meeting in compose mode.</span></span> <span data-ttu-id="9af0b-1133">При вызове `saveAsync` для собрания в Outlook для Mac возвращается ошибка.</span><span class="sxs-lookup"><span data-stu-id="9af0b-1133">Calling `saveAsync` on a meeting in Mac Outlook will return an error.</span></span>
> - <span data-ttu-id="9af0b-1134">Outlook в Интернете всегда отправляет приглашение или обновление при вызове `saveAsync` для встречи в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="9af0b-1134">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9af0b-1135">Параметры</span><span class="sxs-lookup"><span data-stu-id="9af0b-1135">Parameters</span></span>

|<span data-ttu-id="9af0b-1136">Имя</span><span class="sxs-lookup"><span data-stu-id="9af0b-1136">Name</span></span>| <span data-ttu-id="9af0b-1137">Тип</span><span class="sxs-lookup"><span data-stu-id="9af0b-1137">Type</span></span>| <span data-ttu-id="9af0b-1138">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="9af0b-1138">Attributes</span></span>| <span data-ttu-id="9af0b-1139">Описание</span><span class="sxs-lookup"><span data-stu-id="9af0b-1139">Description</span></span>|
|---|---|---|---|
|`options`| <span data-ttu-id="9af0b-1140">Object</span><span class="sxs-lookup"><span data-stu-id="9af0b-1140">Object</span></span>| <span data-ttu-id="9af0b-1141">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="9af0b-1141">&lt;optional&gt;</span></span>|<span data-ttu-id="9af0b-1142">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="9af0b-1142">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="9af0b-1143">Object</span><span class="sxs-lookup"><span data-stu-id="9af0b-1143">Object</span></span>| <span data-ttu-id="9af0b-1144">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="9af0b-1144">&lt;optional&gt;</span></span>|<span data-ttu-id="9af0b-1145">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="9af0b-1145">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="9af0b-1146">функция</span><span class="sxs-lookup"><span data-stu-id="9af0b-1146">function</span></span>||<span data-ttu-id="9af0b-1147">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="9af0b-1147">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="9af0b-1148">После успешного выполнения идентификатор элемента будет предоставлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="9af0b-1148">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9af0b-1149">Требования</span><span class="sxs-lookup"><span data-stu-id="9af0b-1149">Requirements</span></span>

|<span data-ttu-id="9af0b-1150">Требование</span><span class="sxs-lookup"><span data-stu-id="9af0b-1150">Requirement</span></span>| <span data-ttu-id="9af0b-1151">Значение</span><span class="sxs-lookup"><span data-stu-id="9af0b-1151">Value</span></span>|
|---|---|
|[<span data-ttu-id="9af0b-1152">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="9af0b-1152">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9af0b-1153">1.3</span><span class="sxs-lookup"><span data-stu-id="9af0b-1153">1.3</span></span>|
|[<span data-ttu-id="9af0b-1154">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="9af0b-1154">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9af0b-1155">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="9af0b-1155">ReadWriteItem</span></span>|
|[<span data-ttu-id="9af0b-1156">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9af0b-1156">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9af0b-1157">Создание</span><span class="sxs-lookup"><span data-stu-id="9af0b-1157">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="9af0b-1158">Примеры</span><span class="sxs-lookup"><span data-stu-id="9af0b-1158">Examples</span></span>

```javascript
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result.
  });
```

<span data-ttu-id="9af0b-p177">Ниже приведен пример параметра `result`, переданного функции обратного вызова. Свойство `value` содержит идентификатор элемента.</span><span class="sxs-lookup"><span data-stu-id="9af0b-p177">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

####  <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="9af0b-1161">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="9af0b-1161">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="9af0b-1162">Асинхронно вставляет данные в текст или тему сообщения.</span><span class="sxs-lookup"><span data-stu-id="9af0b-1162">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="9af0b-p178">Метод `setSelectedDataAsync` вставляет указанную строку в местоположение курсора в теме или тексте элемента либо, если текст выделен в редакторе, он заменяет выделенный текст. Если курсор находится вне текста или темы элемента, возвращается ошибка. После вставки курсор помещается в конец вставленного содержимого.</span><span class="sxs-lookup"><span data-stu-id="9af0b-p178">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9af0b-1166">Параметры</span><span class="sxs-lookup"><span data-stu-id="9af0b-1166">Parameters</span></span>

|<span data-ttu-id="9af0b-1167">Имя</span><span class="sxs-lookup"><span data-stu-id="9af0b-1167">Name</span></span>| <span data-ttu-id="9af0b-1168">Тип</span><span class="sxs-lookup"><span data-stu-id="9af0b-1168">Type</span></span>| <span data-ttu-id="9af0b-1169">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="9af0b-1169">Attributes</span></span>| <span data-ttu-id="9af0b-1170">Описание</span><span class="sxs-lookup"><span data-stu-id="9af0b-1170">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="9af0b-1171">String</span><span class="sxs-lookup"><span data-stu-id="9af0b-1171">String</span></span>||<span data-ttu-id="9af0b-p179">Вставляемые данные. Объем данных не должен превышать 1 000 000 символов. Если передано больше 1 000 000 символов, возвращается исключение `ArgumentOutOfRange`.</span><span class="sxs-lookup"><span data-stu-id="9af0b-p179">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="9af0b-1175">Object</span><span class="sxs-lookup"><span data-stu-id="9af0b-1175">Object</span></span>| <span data-ttu-id="9af0b-1176">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="9af0b-1176">&lt;optional&gt;</span></span>|<span data-ttu-id="9af0b-1177">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="9af0b-1177">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="9af0b-1178">Object</span><span class="sxs-lookup"><span data-stu-id="9af0b-1178">Object</span></span>| <span data-ttu-id="9af0b-1179">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="9af0b-1179">&lt;optional&gt;</span></span>|<span data-ttu-id="9af0b-1180">В методе обратного вызова разработчики могут указать любой объект, к которому необходимо получить доступ.</span><span class="sxs-lookup"><span data-stu-id="9af0b-1180">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`| [<span data-ttu-id="9af0b-1181">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="9af0b-1181">Office.CoercionType</span></span>](office.md#coerciontype-string)| <span data-ttu-id="9af0b-1182">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="9af0b-1182">&lt;optional&gt;</span></span>|<span data-ttu-id="9af0b-p180">Если задано значение `text`, текущий стиль применяется в Outlook Web App и Outlook. Если поле представляет собой редактор HTML, вставляются только текстовые данные, даже если они имеют формат HTML.</span><span class="sxs-lookup"><span data-stu-id="9af0b-p180">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="9af0b-p181">Если задано значение `html` и поле (не тема) поддерживает HTML, в Outlook Web App применяется текущий стиль, а в Outlook — стиль по умолчанию. Если поле является текстовым, возвращается ошибка `InvalidDataFormat`.</span><span class="sxs-lookup"><span data-stu-id="9af0b-p181">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="9af0b-1187">Если свойство `coercionType` не задано, результат зависит от поля: если поле имеет формат HTML, используется текст в формате HTML, а если поле текстовое, применяется обычный текст.</span><span class="sxs-lookup"><span data-stu-id="9af0b-1187">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="9af0b-1188">функция</span><span class="sxs-lookup"><span data-stu-id="9af0b-1188">function</span></span>||<span data-ttu-id="9af0b-1189">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="9af0b-1189">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="9af0b-1190">Требования</span><span class="sxs-lookup"><span data-stu-id="9af0b-1190">Requirements</span></span>

|<span data-ttu-id="9af0b-1191">Требование</span><span class="sxs-lookup"><span data-stu-id="9af0b-1191">Requirement</span></span>| <span data-ttu-id="9af0b-1192">Значение</span><span class="sxs-lookup"><span data-stu-id="9af0b-1192">Value</span></span>|
|---|---|
|[<span data-ttu-id="9af0b-1193">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="9af0b-1193">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9af0b-1194">1.2</span><span class="sxs-lookup"><span data-stu-id="9af0b-1194">1.2</span></span>|
|[<span data-ttu-id="9af0b-1195">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="9af0b-1195">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9af0b-1196">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="9af0b-1196">ReadWriteItem</span></span>|
|[<span data-ttu-id="9af0b-1197">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9af0b-1197">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9af0b-1198">Создание</span><span class="sxs-lookup"><span data-stu-id="9af0b-1198">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="9af0b-1199">Пример</span><span class="sxs-lookup"><span data-stu-id="9af0b-1199">Example</span></span>

```javascript
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
