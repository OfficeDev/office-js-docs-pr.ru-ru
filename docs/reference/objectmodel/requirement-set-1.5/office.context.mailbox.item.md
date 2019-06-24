---
title: Office.context.mailbox.item — набор обязательных элементов 1.5
description: ''
ms.date: 06/20/2019
localization_priority: Priority
ms.openlocfilehash: 5d7388960a7eae2d66b3f55429a71717f25b803e
ms.sourcegitcommit: 382e2735a1295da914f2bfc38883e518070cec61
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/21/2019
ms.locfileid: "35127361"
---
# <a name="item"></a><span data-ttu-id="ce20c-102">item</span><span class="sxs-lookup"><span data-stu-id="ce20c-102">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="ce20c-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="ce20c-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="ce20c-p101">Пространство имен `item` используется для доступа к выбранному в данный момент сообщению, приглашению на собрание или описанию встречи. Вы можете определить тип пространства имен `item` с помощью свойства [itemType](#itemtype-officemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="ce20c-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="ce20c-106">Requirements</span><span class="sxs-lookup"><span data-stu-id="ce20c-106">Requirements</span></span>

|<span data-ttu-id="ce20c-107">Требование</span><span class="sxs-lookup"><span data-stu-id="ce20c-107">Requirement</span></span>| <span data-ttu-id="ce20c-108">Значение</span><span class="sxs-lookup"><span data-stu-id="ce20c-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="ce20c-109">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="ce20c-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ce20c-110">1.0</span><span class="sxs-lookup"><span data-stu-id="ce20c-110">1.0</span></span>|
|[<span data-ttu-id="ce20c-111">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ce20c-111">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ce20c-112">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="ce20c-112">Restricted</span></span>|
|[<span data-ttu-id="ce20c-113">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ce20c-113">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ce20c-114">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="ce20c-114">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="ce20c-115">Элементы и методы</span><span class="sxs-lookup"><span data-stu-id="ce20c-115">Members and methods</span></span>

| <span data-ttu-id="ce20c-116">Элемент</span><span class="sxs-lookup"><span data-stu-id="ce20c-116">Member</span></span> | <span data-ttu-id="ce20c-117">Тип</span><span class="sxs-lookup"><span data-stu-id="ce20c-117">Type</span></span> |
|--------|------|
| [<span data-ttu-id="ce20c-118">attachments</span><span class="sxs-lookup"><span data-stu-id="ce20c-118">attachments</span></span>](#attachments-arrayattachmentdetails) | <span data-ttu-id="ce20c-119">Элемент</span><span class="sxs-lookup"><span data-stu-id="ce20c-119">Member</span></span> |
| [<span data-ttu-id="ce20c-120">bcc</span><span class="sxs-lookup"><span data-stu-id="ce20c-120">bcc</span></span>](#bcc-recipients) | <span data-ttu-id="ce20c-121">Элемент</span><span class="sxs-lookup"><span data-stu-id="ce20c-121">Member</span></span> |
| [<span data-ttu-id="ce20c-122">body</span><span class="sxs-lookup"><span data-stu-id="ce20c-122">body</span></span>](#body-body) | <span data-ttu-id="ce20c-123">Элемент</span><span class="sxs-lookup"><span data-stu-id="ce20c-123">Member</span></span> |
| [<span data-ttu-id="ce20c-124">cc</span><span class="sxs-lookup"><span data-stu-id="ce20c-124">cc</span></span>](#cc-arrayemailaddressdetailsrecipients) | <span data-ttu-id="ce20c-125">Элемент</span><span class="sxs-lookup"><span data-stu-id="ce20c-125">Member</span></span> |
| [<span data-ttu-id="ce20c-126">conversationId</span><span class="sxs-lookup"><span data-stu-id="ce20c-126">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="ce20c-127">Элемент</span><span class="sxs-lookup"><span data-stu-id="ce20c-127">Member</span></span> |
| [<span data-ttu-id="ce20c-128">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="ce20c-128">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="ce20c-129">Элемент</span><span class="sxs-lookup"><span data-stu-id="ce20c-129">Member</span></span> |
| [<span data-ttu-id="ce20c-130">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="ce20c-130">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="ce20c-131">Элемент</span><span class="sxs-lookup"><span data-stu-id="ce20c-131">Member</span></span> |
| [<span data-ttu-id="ce20c-132">end</span><span class="sxs-lookup"><span data-stu-id="ce20c-132">end</span></span>](#end-datetime) | <span data-ttu-id="ce20c-133">Элемент</span><span class="sxs-lookup"><span data-stu-id="ce20c-133">Member</span></span> |
| [<span data-ttu-id="ce20c-134">from</span><span class="sxs-lookup"><span data-stu-id="ce20c-134">from</span></span>](#from-emailaddressdetails) | <span data-ttu-id="ce20c-135">Элемент</span><span class="sxs-lookup"><span data-stu-id="ce20c-135">Member</span></span> |
| [<span data-ttu-id="ce20c-136">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="ce20c-136">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="ce20c-137">Элемент</span><span class="sxs-lookup"><span data-stu-id="ce20c-137">Member</span></span> |
| [<span data-ttu-id="ce20c-138">itemClass</span><span class="sxs-lookup"><span data-stu-id="ce20c-138">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="ce20c-139">Элемент</span><span class="sxs-lookup"><span data-stu-id="ce20c-139">Member</span></span> |
| [<span data-ttu-id="ce20c-140">itemId</span><span class="sxs-lookup"><span data-stu-id="ce20c-140">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="ce20c-141">Элемент</span><span class="sxs-lookup"><span data-stu-id="ce20c-141">Member</span></span> |
| [<span data-ttu-id="ce20c-142">itemType</span><span class="sxs-lookup"><span data-stu-id="ce20c-142">itemType</span></span>](#itemtype-officemailboxenumsitemtype) | <span data-ttu-id="ce20c-143">Элемент</span><span class="sxs-lookup"><span data-stu-id="ce20c-143">Member</span></span> |
| [<span data-ttu-id="ce20c-144">location</span><span class="sxs-lookup"><span data-stu-id="ce20c-144">location</span></span>](#location-stringlocation) | <span data-ttu-id="ce20c-145">Элемент</span><span class="sxs-lookup"><span data-stu-id="ce20c-145">Member</span></span> |
| [<span data-ttu-id="ce20c-146">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="ce20c-146">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="ce20c-147">Элемент</span><span class="sxs-lookup"><span data-stu-id="ce20c-147">Member</span></span> |
| [<span data-ttu-id="ce20c-148">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="ce20c-148">notificationMessages</span></span>](#notificationmessages-notificationmessages) | <span data-ttu-id="ce20c-149">Элемент</span><span class="sxs-lookup"><span data-stu-id="ce20c-149">Member</span></span> |
| [<span data-ttu-id="ce20c-150">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="ce20c-150">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="ce20c-151">Элемент</span><span class="sxs-lookup"><span data-stu-id="ce20c-151">Member</span></span> |
| [<span data-ttu-id="ce20c-152">organizer</span><span class="sxs-lookup"><span data-stu-id="ce20c-152">organizer</span></span>](#organizer-emailaddressdetails) | <span data-ttu-id="ce20c-153">Элемент</span><span class="sxs-lookup"><span data-stu-id="ce20c-153">Member</span></span> |
| [<span data-ttu-id="ce20c-154">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="ce20c-154">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="ce20c-155">Member</span><span class="sxs-lookup"><span data-stu-id="ce20c-155">Member</span></span> |
| [<span data-ttu-id="ce20c-156">sender</span><span class="sxs-lookup"><span data-stu-id="ce20c-156">sender</span></span>](#sender-emailaddressdetails) | <span data-ttu-id="ce20c-157">Элемент</span><span class="sxs-lookup"><span data-stu-id="ce20c-157">Member</span></span> |
| [<span data-ttu-id="ce20c-158">start</span><span class="sxs-lookup"><span data-stu-id="ce20c-158">start</span></span>](#start-datetime) | <span data-ttu-id="ce20c-159">Элемент</span><span class="sxs-lookup"><span data-stu-id="ce20c-159">Member</span></span> |
| [<span data-ttu-id="ce20c-160">subject</span><span class="sxs-lookup"><span data-stu-id="ce20c-160">subject</span></span>](#subject-stringsubject) | <span data-ttu-id="ce20c-161">Элемент</span><span class="sxs-lookup"><span data-stu-id="ce20c-161">Member</span></span> |
| [<span data-ttu-id="ce20c-162">to</span><span class="sxs-lookup"><span data-stu-id="ce20c-162">to</span></span>](#to-arrayemailaddressdetailsrecipients) | <span data-ttu-id="ce20c-163">Элемент</span><span class="sxs-lookup"><span data-stu-id="ce20c-163">Member</span></span> |
| [<span data-ttu-id="ce20c-164">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="ce20c-164">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="ce20c-165">Метод</span><span class="sxs-lookup"><span data-stu-id="ce20c-165">Method</span></span> |
| [<span data-ttu-id="ce20c-166">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="ce20c-166">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="ce20c-167">Метод</span><span class="sxs-lookup"><span data-stu-id="ce20c-167">Method</span></span> |
| [<span data-ttu-id="ce20c-168">close</span><span class="sxs-lookup"><span data-stu-id="ce20c-168">close</span></span>](#close) | <span data-ttu-id="ce20c-169">Метод</span><span class="sxs-lookup"><span data-stu-id="ce20c-169">Method</span></span> |
| [<span data-ttu-id="ce20c-170">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="ce20c-170">displayReplyAllForm</span></span>](#displayreplyallformformdata-callback) | <span data-ttu-id="ce20c-171">Метод</span><span class="sxs-lookup"><span data-stu-id="ce20c-171">Method</span></span> |
| [<span data-ttu-id="ce20c-172">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="ce20c-172">displayReplyForm</span></span>](#displayreplyformformdata-callback) | <span data-ttu-id="ce20c-173">Метод</span><span class="sxs-lookup"><span data-stu-id="ce20c-173">Method</span></span> |
| [<span data-ttu-id="ce20c-174">getEntities</span><span class="sxs-lookup"><span data-stu-id="ce20c-174">getEntities</span></span>](#getentities--entities) | <span data-ttu-id="ce20c-175">Метод</span><span class="sxs-lookup"><span data-stu-id="ce20c-175">Method</span></span> |
| [<span data-ttu-id="ce20c-176">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="ce20c-176">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="ce20c-177">Метод</span><span class="sxs-lookup"><span data-stu-id="ce20c-177">Method</span></span> |
| [<span data-ttu-id="ce20c-178">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="ce20c-178">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="ce20c-179">Метод</span><span class="sxs-lookup"><span data-stu-id="ce20c-179">Method</span></span> |
| [<span data-ttu-id="ce20c-180">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="ce20c-180">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="ce20c-181">Метод</span><span class="sxs-lookup"><span data-stu-id="ce20c-181">Method</span></span> |
| [<span data-ttu-id="ce20c-182">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="ce20c-182">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="ce20c-183">Метод</span><span class="sxs-lookup"><span data-stu-id="ce20c-183">Method</span></span> |
| [<span data-ttu-id="ce20c-184">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="ce20c-184">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="ce20c-185">Метод</span><span class="sxs-lookup"><span data-stu-id="ce20c-185">Method</span></span> |
| [<span data-ttu-id="ce20c-186">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="ce20c-186">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="ce20c-187">Метод</span><span class="sxs-lookup"><span data-stu-id="ce20c-187">Method</span></span> |
| [<span data-ttu-id="ce20c-188">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="ce20c-188">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="ce20c-189">Метод</span><span class="sxs-lookup"><span data-stu-id="ce20c-189">Method</span></span> |
| [<span data-ttu-id="ce20c-190">saveAsync</span><span class="sxs-lookup"><span data-stu-id="ce20c-190">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="ce20c-191">Метод</span><span class="sxs-lookup"><span data-stu-id="ce20c-191">Method</span></span> |
| [<span data-ttu-id="ce20c-192">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="ce20c-192">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="ce20c-193">Метод</span><span class="sxs-lookup"><span data-stu-id="ce20c-193">Method</span></span> |

### <a name="example"></a><span data-ttu-id="ce20c-194">Пример</span><span class="sxs-lookup"><span data-stu-id="ce20c-194">Example</span></span>

<span data-ttu-id="ce20c-195">В примере кода JavaScript, приведенном ниже, показано, как получить доступ к свойству `subject` текущего элемента в Outlook.</span><span class="sxs-lookup"><span data-stu-id="ce20c-195">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="ce20c-196">Members</span><span class="sxs-lookup"><span data-stu-id="ce20c-196">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook15officeattachmentdetails"></a><span data-ttu-id="ce20c-197">attachments: Array.<[AttachmentDetails](/javascript/api/outlook_1_5/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="ce20c-197">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_5/office.attachmentdetails)></span></span>

<span data-ttu-id="ce20c-p102">Получает массив вложений для элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="ce20c-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="ce20c-200">Outlook блокирует определенные типы файлов из-за потенциальных проблем с безопасностью, поэтому они не возвращаются.</span><span class="sxs-lookup"><span data-stu-id="ce20c-200">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="ce20c-201">Дополнительные сведения см. в статье [Блокировка вложений в Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="ce20c-201">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="ce20c-202">Тип</span><span class="sxs-lookup"><span data-stu-id="ce20c-202">Type</span></span>

*   <span data-ttu-id="ce20c-203">Array.<[AttachmentDetails](/javascript/api/outlook_1_5/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="ce20c-203">Array.<[AttachmentDetails](/javascript/api/outlook_1_5/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="ce20c-204">Требования</span><span class="sxs-lookup"><span data-stu-id="ce20c-204">Requirements</span></span>

|<span data-ttu-id="ce20c-205">Требование</span><span class="sxs-lookup"><span data-stu-id="ce20c-205">Requirement</span></span>| <span data-ttu-id="ce20c-206">Значение</span><span class="sxs-lookup"><span data-stu-id="ce20c-206">Value</span></span>|
|---|---|
|[<span data-ttu-id="ce20c-207">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="ce20c-207">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ce20c-208">1.0</span><span class="sxs-lookup"><span data-stu-id="ce20c-208">1.0</span></span>|
|[<span data-ttu-id="ce20c-209">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ce20c-209">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ce20c-210">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ce20c-210">ReadItem</span></span>|
|[<span data-ttu-id="ce20c-211">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ce20c-211">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ce20c-212">Чтение</span><span class="sxs-lookup"><span data-stu-id="ce20c-212">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ce20c-213">Пример</span><span class="sxs-lookup"><span data-stu-id="ce20c-213">Example</span></span>

<span data-ttu-id="ce20c-214">С помощью приведенного ниже кода можно создать HTML-строку с подробными сведениями обо всех вложениях для текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="ce20c-214">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

#### <a name="bcc-recipientsjavascriptapioutlook15officerecipients"></a><span data-ttu-id="ce20c-215">bcc: [Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="ce20c-215">bcc :[Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

<span data-ttu-id="ce20c-216">Получает объект, который предоставляет методы для получения или обновления получателей скрытой копии сообщения.</span><span class="sxs-lookup"><span data-stu-id="ce20c-216">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="ce20c-217">Только в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="ce20c-217">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="ce20c-218">Тип</span><span class="sxs-lookup"><span data-stu-id="ce20c-218">Type</span></span>

*   [<span data-ttu-id="ce20c-219">Получатели</span><span class="sxs-lookup"><span data-stu-id="ce20c-219">Recipients</span></span>](/javascript/api/outlook_1_5/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="ce20c-220">Требования</span><span class="sxs-lookup"><span data-stu-id="ce20c-220">Requirements</span></span>

|<span data-ttu-id="ce20c-221">Требование</span><span class="sxs-lookup"><span data-stu-id="ce20c-221">Requirement</span></span>| <span data-ttu-id="ce20c-222">Значение</span><span class="sxs-lookup"><span data-stu-id="ce20c-222">Value</span></span>|
|---|---|
|[<span data-ttu-id="ce20c-223">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="ce20c-223">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ce20c-224">1.1</span><span class="sxs-lookup"><span data-stu-id="ce20c-224">1.1</span></span>|
|[<span data-ttu-id="ce20c-225">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ce20c-225">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ce20c-226">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ce20c-226">ReadItem</span></span>|
|[<span data-ttu-id="ce20c-227">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ce20c-227">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ce20c-228">Создание</span><span class="sxs-lookup"><span data-stu-id="ce20c-228">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="ce20c-229">Пример</span><span class="sxs-lookup"><span data-stu-id="ce20c-229">Example</span></span>

```javascript
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

#### <a name="body-bodyjavascriptapioutlook15officebody"></a><span data-ttu-id="ce20c-230">body: [Body](/javascript/api/outlook_1_5/office.body)</span><span class="sxs-lookup"><span data-stu-id="ce20c-230">body :[Body](/javascript/api/outlook_1_5/office.body)</span></span>

<span data-ttu-id="ce20c-231">Получает объект, предоставляющий методы для работы с основным текстом элемента.</span><span class="sxs-lookup"><span data-stu-id="ce20c-231">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="ce20c-232">Тип</span><span class="sxs-lookup"><span data-stu-id="ce20c-232">Type</span></span>

*   [<span data-ttu-id="ce20c-233">Body</span><span class="sxs-lookup"><span data-stu-id="ce20c-233">Body</span></span>](/javascript/api/outlook_1_5/office.body)

##### <a name="requirements"></a><span data-ttu-id="ce20c-234">Требования</span><span class="sxs-lookup"><span data-stu-id="ce20c-234">Requirements</span></span>

|<span data-ttu-id="ce20c-235">Требование</span><span class="sxs-lookup"><span data-stu-id="ce20c-235">Requirement</span></span>| <span data-ttu-id="ce20c-236">Значение</span><span class="sxs-lookup"><span data-stu-id="ce20c-236">Value</span></span>|
|---|---|
|[<span data-ttu-id="ce20c-237">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="ce20c-237">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ce20c-238">1.1</span><span class="sxs-lookup"><span data-stu-id="ce20c-238">1.1</span></span>|
|[<span data-ttu-id="ce20c-239">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ce20c-239">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ce20c-240">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ce20c-240">ReadItem</span></span>|
|[<span data-ttu-id="ce20c-241">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ce20c-241">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ce20c-242">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="ce20c-242">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ce20c-243">Пример</span><span class="sxs-lookup"><span data-stu-id="ce20c-243">Example</span></span>

<span data-ttu-id="ce20c-244">В этом примере возвращается текст сообщения в формате обычного текста.</span><span class="sxs-lookup"><span data-stu-id="ce20c-244">This example gets the body of the message in plain text.</span></span>

```javascript
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="ce20c-245">Ниже приведен пример итогового параметра, переданного функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="ce20c-245">The following is an example of the result parameter passed to the callback function.</span></span>

```json
{
  "value": "TEXT of whole body (including threads below)",
  "status": "succeeded",
  "asyncContext": "This is passed to the callback"
}
```

#### <a name="cc-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients"></a><span data-ttu-id="ce20c-246">cc: Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="ce20c-246">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

<span data-ttu-id="ce20c-247">Предоставляет доступ к получателям копии сообщения.</span><span class="sxs-lookup"><span data-stu-id="ce20c-247">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="ce20c-248">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="ce20c-248">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="ce20c-249">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="ce20c-249">Read mode</span></span>

<span data-ttu-id="ce20c-p106">Свойство `cc` возвращает массив, который содержит объект `EmailAddressDetails` для каждого получателя, указанного в строке **Копия** сообщения. Коллекция может включать не более 100 элементов.</span><span class="sxs-lookup"><span data-stu-id="ce20c-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="ce20c-252">Режим создания</span><span class="sxs-lookup"><span data-stu-id="ce20c-252">Compose mode</span></span>

<span data-ttu-id="ce20c-253">Свойство `cc` возвращает объект `Recipients`, предоставляющий методы для получения или обновления получателей, которые указаны в строке **Копия** сообщения.</span><span class="sxs-lookup"><span data-stu-id="ce20c-253">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

```javascript
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="ce20c-254">Тип</span><span class="sxs-lookup"><span data-stu-id="ce20c-254">Type</span></span>

*   <span data-ttu-id="ce20c-255">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="ce20c-255">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="ce20c-256">Требования</span><span class="sxs-lookup"><span data-stu-id="ce20c-256">Requirements</span></span>

|<span data-ttu-id="ce20c-257">Требование</span><span class="sxs-lookup"><span data-stu-id="ce20c-257">Requirement</span></span>| <span data-ttu-id="ce20c-258">Значение</span><span class="sxs-lookup"><span data-stu-id="ce20c-258">Value</span></span>|
|---|---|
|[<span data-ttu-id="ce20c-259">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="ce20c-259">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ce20c-260">1.0</span><span class="sxs-lookup"><span data-stu-id="ce20c-260">1.0</span></span>|
|[<span data-ttu-id="ce20c-261">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ce20c-261">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ce20c-262">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ce20c-262">ReadItem</span></span>|
|[<span data-ttu-id="ce20c-263">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ce20c-263">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ce20c-264">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="ce20c-264">Compose or Read</span></span>|

#### <a name="nullable-conversationid-string"></a><span data-ttu-id="ce20c-265">(nullable) conversationId: String</span><span class="sxs-lookup"><span data-stu-id="ce20c-265">(nullable) conversationId :String</span></span>

<span data-ttu-id="ce20c-266">Получает идентификатор разговора по электронной почте, содержащего конкретное сообщение.</span><span class="sxs-lookup"><span data-stu-id="ce20c-266">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="ce20c-p107">Вы можете получить целочисленное значение этого свойства, если ваше почтовое приложение активируется в формах просмотра или формах создания ответов. Если пользователь изменит тему ответа, после его отправки идентификатор беседы будет изменен, и полученное ранее значение будет недействительным.</span><span class="sxs-lookup"><span data-stu-id="ce20c-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="ce20c-p108">Это свойство имеет значение NULL для нового элемента в форме создания. Свойство `conversationId` вернет значение, если пользователь задаст тему и сохранит элемент.</span><span class="sxs-lookup"><span data-stu-id="ce20c-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="ce20c-271">Тип</span><span class="sxs-lookup"><span data-stu-id="ce20c-271">Type</span></span>

*   <span data-ttu-id="ce20c-272">String</span><span class="sxs-lookup"><span data-stu-id="ce20c-272">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="ce20c-273">Требования</span><span class="sxs-lookup"><span data-stu-id="ce20c-273">Requirements</span></span>

|<span data-ttu-id="ce20c-274">Требование</span><span class="sxs-lookup"><span data-stu-id="ce20c-274">Requirement</span></span>| <span data-ttu-id="ce20c-275">Значение</span><span class="sxs-lookup"><span data-stu-id="ce20c-275">Value</span></span>|
|---|---|
|[<span data-ttu-id="ce20c-276">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="ce20c-276">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ce20c-277">1.0</span><span class="sxs-lookup"><span data-stu-id="ce20c-277">1.0</span></span>|
|[<span data-ttu-id="ce20c-278">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ce20c-278">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ce20c-279">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ce20c-279">ReadItem</span></span>|
|[<span data-ttu-id="ce20c-280">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ce20c-280">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ce20c-281">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="ce20c-281">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ce20c-282">Пример</span><span class="sxs-lookup"><span data-stu-id="ce20c-282">Example</span></span>

```javascript
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

#### <a name="datetimecreated-date"></a><span data-ttu-id="ce20c-283">dateTimeCreated: Date</span><span class="sxs-lookup"><span data-stu-id="ce20c-283">dateTimeCreated :Date</span></span>

<span data-ttu-id="ce20c-p109">Получает дату и время создания элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="ce20c-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="ce20c-286">Тип</span><span class="sxs-lookup"><span data-stu-id="ce20c-286">Type</span></span>

*   <span data-ttu-id="ce20c-287">Дата</span><span class="sxs-lookup"><span data-stu-id="ce20c-287">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="ce20c-288">Требования</span><span class="sxs-lookup"><span data-stu-id="ce20c-288">Requirements</span></span>

|<span data-ttu-id="ce20c-289">Требование</span><span class="sxs-lookup"><span data-stu-id="ce20c-289">Requirement</span></span>| <span data-ttu-id="ce20c-290">Значение</span><span class="sxs-lookup"><span data-stu-id="ce20c-290">Value</span></span>|
|---|---|
|[<span data-ttu-id="ce20c-291">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="ce20c-291">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ce20c-292">1.0</span><span class="sxs-lookup"><span data-stu-id="ce20c-292">1.0</span></span>|
|[<span data-ttu-id="ce20c-293">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ce20c-293">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ce20c-294">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ce20c-294">ReadItem</span></span>|
|[<span data-ttu-id="ce20c-295">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ce20c-295">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ce20c-296">Чтение</span><span class="sxs-lookup"><span data-stu-id="ce20c-296">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ce20c-297">Пример</span><span class="sxs-lookup"><span data-stu-id="ce20c-297">Example</span></span>

```javascript
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="ce20c-298">dateTimeModified: Date</span><span class="sxs-lookup"><span data-stu-id="ce20c-298">dateTimeModified :Date</span></span>

<span data-ttu-id="ce20c-p110">Получает дату и время последнего изменения элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="ce20c-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="ce20c-301">Этот элемент не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="ce20c-301">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="ce20c-302">Тип</span><span class="sxs-lookup"><span data-stu-id="ce20c-302">Type</span></span>

*   <span data-ttu-id="ce20c-303">Дата</span><span class="sxs-lookup"><span data-stu-id="ce20c-303">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="ce20c-304">Требования</span><span class="sxs-lookup"><span data-stu-id="ce20c-304">Requirements</span></span>

|<span data-ttu-id="ce20c-305">Требование</span><span class="sxs-lookup"><span data-stu-id="ce20c-305">Requirement</span></span>| <span data-ttu-id="ce20c-306">Значение</span><span class="sxs-lookup"><span data-stu-id="ce20c-306">Value</span></span>|
|---|---|
|[<span data-ttu-id="ce20c-307">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="ce20c-307">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ce20c-308">1.0</span><span class="sxs-lookup"><span data-stu-id="ce20c-308">1.0</span></span>|
|[<span data-ttu-id="ce20c-309">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ce20c-309">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ce20c-310">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ce20c-310">ReadItem</span></span>|
|[<span data-ttu-id="ce20c-311">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ce20c-311">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ce20c-312">Чтение</span><span class="sxs-lookup"><span data-stu-id="ce20c-312">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ce20c-313">Пример</span><span class="sxs-lookup"><span data-stu-id="ce20c-313">Example</span></span>

```javascript
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

#### <a name="end-datetimejavascriptapioutlook15officetime"></a><span data-ttu-id="ce20c-314">end: Date|[Time](/javascript/api/outlook_1_5/office.time)</span><span class="sxs-lookup"><span data-stu-id="ce20c-314">end :Date|[Time](/javascript/api/outlook_1_5/office.time)</span></span>

<span data-ttu-id="ce20c-315">Получает или задает дату и время окончания встречи.</span><span class="sxs-lookup"><span data-stu-id="ce20c-315">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="ce20c-p111">Свойство `end` представлено в виде значения даты и времени в формате UTC. Преобразовать значение свойства end в местные значения даты и времени клиента можно с помощью метода [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime).</span><span class="sxs-lookup"><span data-stu-id="ce20c-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="ce20c-318">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="ce20c-318">Read mode</span></span>

<span data-ttu-id="ce20c-319">Свойство `end` возвращает объект `Date`.</span><span class="sxs-lookup"><span data-stu-id="ce20c-319">The `end` property returns a `Date` object.</span></span>

```javascript
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="ce20c-320">Режим создания</span><span class="sxs-lookup"><span data-stu-id="ce20c-320">Compose mode</span></span>

<span data-ttu-id="ce20c-321">Свойство `end` возвращает объект `Time`.</span><span class="sxs-lookup"><span data-stu-id="ce20c-321">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="ce20c-322">Если вы задаете время окончания с помощью метода [`Time.setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-), необходимо использовать метод [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) для преобразования местного времени на клиенте в формат UTC для сервера.</span><span class="sxs-lookup"><span data-stu-id="ce20c-322">When you use the [`Time.setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="ce20c-323">В примере ниже показано, как с помощью метода [`setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) объекта `Time` задать время окончания встречи.</span><span class="sxs-lookup"><span data-stu-id="ce20c-323">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="ce20c-324">Тип</span><span class="sxs-lookup"><span data-stu-id="ce20c-324">Type</span></span>

*   <span data-ttu-id="ce20c-325">Date | [Time](/javascript/api/outlook_1_5/office.time)</span><span class="sxs-lookup"><span data-stu-id="ce20c-325">Date | [Time](/javascript/api/outlook_1_5/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="ce20c-326">Требования</span><span class="sxs-lookup"><span data-stu-id="ce20c-326">Requirements</span></span>

|<span data-ttu-id="ce20c-327">Требование</span><span class="sxs-lookup"><span data-stu-id="ce20c-327">Requirement</span></span>| <span data-ttu-id="ce20c-328">Значение</span><span class="sxs-lookup"><span data-stu-id="ce20c-328">Value</span></span>|
|---|---|
|[<span data-ttu-id="ce20c-329">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="ce20c-329">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ce20c-330">1.0</span><span class="sxs-lookup"><span data-stu-id="ce20c-330">1.0</span></span>|
|[<span data-ttu-id="ce20c-331">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ce20c-331">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ce20c-332">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ce20c-332">ReadItem</span></span>|
|[<span data-ttu-id="ce20c-333">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ce20c-333">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ce20c-334">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="ce20c-334">Compose or Read</span></span>|

#### <a name="from-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails"></a><span data-ttu-id="ce20c-335">from: [EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="ce20c-335">from :[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)</span></span>

<span data-ttu-id="ce20c-p112">Получает электронный адрес отправителя сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="ce20c-p112">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="ce20c-p113">Свойства `from` и [`sender`](#sender-emailaddressdetails) представляют одно лицо, если сообщение не отправлено представителем. В противном случае свойство `from` представляет лицо, делегировавшее полномочия, а свойство sender — представителя.</span><span class="sxs-lookup"><span data-stu-id="ce20c-p113">The `from` and [`sender`](#sender-emailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="ce20c-340">Свойству `recipientType`, принадлежащему объекту `EmailAddressDetails` в свойстве `from`, задано значение `undefined`.</span><span class="sxs-lookup"><span data-stu-id="ce20c-340">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="ce20c-341">Тип</span><span class="sxs-lookup"><span data-stu-id="ce20c-341">Type</span></span>

*   [<span data-ttu-id="ce20c-342">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="ce20c-342">EmailAddressDetails</span></span>](/javascript/api/outlook_1_5/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="ce20c-343">Требования</span><span class="sxs-lookup"><span data-stu-id="ce20c-343">Requirements</span></span>

|<span data-ttu-id="ce20c-344">Требование</span><span class="sxs-lookup"><span data-stu-id="ce20c-344">Requirement</span></span>| <span data-ttu-id="ce20c-345">Значение</span><span class="sxs-lookup"><span data-stu-id="ce20c-345">Value</span></span>|
|---|---|
|[<span data-ttu-id="ce20c-346">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="ce20c-346">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ce20c-347">1.0</span><span class="sxs-lookup"><span data-stu-id="ce20c-347">1.0</span></span>|
|[<span data-ttu-id="ce20c-348">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ce20c-348">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ce20c-349">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ce20c-349">ReadItem</span></span>|
|[<span data-ttu-id="ce20c-350">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ce20c-350">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ce20c-351">Чтение</span><span class="sxs-lookup"><span data-stu-id="ce20c-351">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ce20c-352">Пример</span><span class="sxs-lookup"><span data-stu-id="ce20c-352">Example</span></span>

```javascript
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

#### <a name="internetmessageid-string"></a><span data-ttu-id="ce20c-353">internetMessageId: String</span><span class="sxs-lookup"><span data-stu-id="ce20c-353">internetMessageId :String</span></span>

<span data-ttu-id="ce20c-p114">Получает идентификатор интернет-сообщения для электронного сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="ce20c-p114">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="ce20c-356">Тип</span><span class="sxs-lookup"><span data-stu-id="ce20c-356">Type</span></span>

*   <span data-ttu-id="ce20c-357">String</span><span class="sxs-lookup"><span data-stu-id="ce20c-357">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="ce20c-358">Требования</span><span class="sxs-lookup"><span data-stu-id="ce20c-358">Requirements</span></span>

|<span data-ttu-id="ce20c-359">Требование</span><span class="sxs-lookup"><span data-stu-id="ce20c-359">Requirement</span></span>| <span data-ttu-id="ce20c-360">Значение</span><span class="sxs-lookup"><span data-stu-id="ce20c-360">Value</span></span>|
|---|---|
|[<span data-ttu-id="ce20c-361">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="ce20c-361">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ce20c-362">1.0</span><span class="sxs-lookup"><span data-stu-id="ce20c-362">1.0</span></span>|
|[<span data-ttu-id="ce20c-363">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ce20c-363">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ce20c-364">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ce20c-364">ReadItem</span></span>|
|[<span data-ttu-id="ce20c-365">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ce20c-365">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ce20c-366">Чтение</span><span class="sxs-lookup"><span data-stu-id="ce20c-366">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ce20c-367">Пример</span><span class="sxs-lookup"><span data-stu-id="ce20c-367">Example</span></span>

```javascript
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="ce20c-368">itemClass: String</span><span class="sxs-lookup"><span data-stu-id="ce20c-368">itemClass :String</span></span>

<span data-ttu-id="ce20c-p115">Получает класс элемента веб-служб Exchange для выбранного элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="ce20c-p115">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="ce20c-p116">Свойство `itemClass` указывает класс сообщения выбранного элемента. Ниже приводятся классы сообщения по умолчанию для элемента сообщения или встречи.</span><span class="sxs-lookup"><span data-stu-id="ce20c-p116">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="ce20c-373">Тип</span><span class="sxs-lookup"><span data-stu-id="ce20c-373">Type</span></span> | <span data-ttu-id="ce20c-374">Описание</span><span class="sxs-lookup"><span data-stu-id="ce20c-374">Description</span></span> | <span data-ttu-id="ce20c-375">Класс элемента</span><span class="sxs-lookup"><span data-stu-id="ce20c-375">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="ce20c-376">Элементы встречи</span><span class="sxs-lookup"><span data-stu-id="ce20c-376">Appointment items</span></span> | <span data-ttu-id="ce20c-377">Это элементы календаря для класса элемента `IPM.Appointment` или `IPM.Appointment.Occurrence`.</span><span class="sxs-lookup"><span data-stu-id="ce20c-377">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurrence` |
| <span data-ttu-id="ce20c-378">Элементы сообщения</span><span class="sxs-lookup"><span data-stu-id="ce20c-378">Message items</span></span> | <span data-ttu-id="ce20c-379">Сюда входят электронные сообщения, для которых по умолчанию задан класс сообщения `IPM.Note`, а также приглашения на собрания, ответы на них и уведомления об их отмене, использующие `IPM.Schedule.Meeting` в качестве базового класса сообщения.</span><span class="sxs-lookup"><span data-stu-id="ce20c-379">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="ce20c-380">Можно создавать настраиваемые классы сообщения, расширяющие классы сообщения по умолчанию, например настраиваемый класс сообщения о встрече `IPM.Appointment.Contoso`.</span><span class="sxs-lookup"><span data-stu-id="ce20c-380">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="ce20c-381">Тип</span><span class="sxs-lookup"><span data-stu-id="ce20c-381">Type</span></span>

*   <span data-ttu-id="ce20c-382">String</span><span class="sxs-lookup"><span data-stu-id="ce20c-382">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="ce20c-383">Требования</span><span class="sxs-lookup"><span data-stu-id="ce20c-383">Requirements</span></span>

|<span data-ttu-id="ce20c-384">Требование</span><span class="sxs-lookup"><span data-stu-id="ce20c-384">Requirement</span></span>| <span data-ttu-id="ce20c-385">Значение</span><span class="sxs-lookup"><span data-stu-id="ce20c-385">Value</span></span>|
|---|---|
|[<span data-ttu-id="ce20c-386">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="ce20c-386">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ce20c-387">1.0</span><span class="sxs-lookup"><span data-stu-id="ce20c-387">1.0</span></span>|
|[<span data-ttu-id="ce20c-388">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ce20c-388">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ce20c-389">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ce20c-389">ReadItem</span></span>|
|[<span data-ttu-id="ce20c-390">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ce20c-390">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ce20c-391">Чтение</span><span class="sxs-lookup"><span data-stu-id="ce20c-391">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ce20c-392">Пример</span><span class="sxs-lookup"><span data-stu-id="ce20c-392">Example</span></span>

```javascript
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="ce20c-393">(nullable) itemId: String</span><span class="sxs-lookup"><span data-stu-id="ce20c-393">(nullable) itemId :String</span></span>

<span data-ttu-id="ce20c-p117">Получает идентификатор элемента веб-служб Exchange для текущего элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="ce20c-p117">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="ce20c-396">Идентификатор, возвращаемый свойством `itemId`, совпадает с идентификатором элемента веб-служб Exchange.</span><span class="sxs-lookup"><span data-stu-id="ce20c-396">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="ce20c-397">Свойство `itemId` не совпадает с идентификатором записи Outlook, а также идентификатором, который используется REST API Outlook.</span><span class="sxs-lookup"><span data-stu-id="ce20c-397">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="ce20c-398">Прежде чем совершать вызовы REST API, используя это значение, его необходимо преобразовать с помощью [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="ce20c-398">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="ce20c-399">Дополнительные сведения см. в статье [Использование REST API Outlook из надстройки Outlook](/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="ce20c-399">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="ce20c-p119">Свойство `itemId` недоступно в режиме создания. Если требуется идентификатор элемента, с помощью метода [`saveAsync`](#saveasyncoptions-callback) можно сохранить элемент в хранилище. При этом в параметре [`AsyncResult.value`](/javascript/api/office/office.asyncresult) функции обратного вызова возвращается идентификатор элемента.</span><span class="sxs-lookup"><span data-stu-id="ce20c-p119">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="ce20c-402">Тип</span><span class="sxs-lookup"><span data-stu-id="ce20c-402">Type</span></span>

*   <span data-ttu-id="ce20c-403">String</span><span class="sxs-lookup"><span data-stu-id="ce20c-403">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="ce20c-404">Требования</span><span class="sxs-lookup"><span data-stu-id="ce20c-404">Requirements</span></span>

|<span data-ttu-id="ce20c-405">Требование</span><span class="sxs-lookup"><span data-stu-id="ce20c-405">Requirement</span></span>| <span data-ttu-id="ce20c-406">Значение</span><span class="sxs-lookup"><span data-stu-id="ce20c-406">Value</span></span>|
|---|---|
|[<span data-ttu-id="ce20c-407">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="ce20c-407">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ce20c-408">1.0</span><span class="sxs-lookup"><span data-stu-id="ce20c-408">1.0</span></span>|
|[<span data-ttu-id="ce20c-409">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ce20c-409">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ce20c-410">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ce20c-410">ReadItem</span></span>|
|[<span data-ttu-id="ce20c-411">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ce20c-411">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ce20c-412">Чтение</span><span class="sxs-lookup"><span data-stu-id="ce20c-412">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ce20c-413">Пример</span><span class="sxs-lookup"><span data-stu-id="ce20c-413">Example</span></span>

<span data-ttu-id="ce20c-p120">Указанный ниже код проверяет наличие идентификатора элемента. Если свойство `itemId` возвращает значение `null` или `undefined`, элемент будет сохранен в хранилище, а из асинхронного результата будет получен идентификатор элемента.</span><span class="sxs-lookup"><span data-stu-id="ce20c-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```javascript
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result) {
    itemId = result.value;
  });
}
```

#### <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook15officemailboxenumsitemtype"></a><span data-ttu-id="ce20c-416">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook_1_5/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="ce20c-416">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_5/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="ce20c-417">Получает тип элемента, который представляет экземпляр.</span><span class="sxs-lookup"><span data-stu-id="ce20c-417">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="ce20c-418">Свойство `itemType` возвращает одно из значений перечисления `ItemType`, которое указывает, является ли экземпляр объекта `item` сообщением или встречей.</span><span class="sxs-lookup"><span data-stu-id="ce20c-418">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="ce20c-419">Тип</span><span class="sxs-lookup"><span data-stu-id="ce20c-419">Type</span></span>

*   [<span data-ttu-id="ce20c-420">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="ce20c-420">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook_1_5/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="ce20c-421">Требования</span><span class="sxs-lookup"><span data-stu-id="ce20c-421">Requirements</span></span>

|<span data-ttu-id="ce20c-422">Требование</span><span class="sxs-lookup"><span data-stu-id="ce20c-422">Requirement</span></span>| <span data-ttu-id="ce20c-423">Значение</span><span class="sxs-lookup"><span data-stu-id="ce20c-423">Value</span></span>|
|---|---|
|[<span data-ttu-id="ce20c-424">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="ce20c-424">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ce20c-425">1.0</span><span class="sxs-lookup"><span data-stu-id="ce20c-425">1.0</span></span>|
|[<span data-ttu-id="ce20c-426">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ce20c-426">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ce20c-427">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ce20c-427">ReadItem</span></span>|
|[<span data-ttu-id="ce20c-428">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ce20c-428">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ce20c-429">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="ce20c-429">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ce20c-430">Пример</span><span class="sxs-lookup"><span data-stu-id="ce20c-430">Example</span></span>

```javascript
if (Office.context.mailbox.item.itemType === Office.MailboxEnums.ItemType.Message) {
  // Do something.
} else {
  // Do something else.
}
```

#### <a name="location-stringlocationjavascriptapioutlook15officelocation"></a><span data-ttu-id="ce20c-431">location: String|[Location](/javascript/api/outlook_1_5/office.location)</span><span class="sxs-lookup"><span data-stu-id="ce20c-431">location :String|[Location](/javascript/api/outlook_1_5/office.location)</span></span>

<span data-ttu-id="ce20c-432">Получает или задает место встречи.</span><span class="sxs-lookup"><span data-stu-id="ce20c-432">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="ce20c-433">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="ce20c-433">Read mode</span></span>

<span data-ttu-id="ce20c-434">Свойство `location` возвращает строку, содержащую сведения о месте встречи.</span><span class="sxs-lookup"><span data-stu-id="ce20c-434">The `location` property returns a string that contains the location of the appointment.</span></span>

```javascript
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="ce20c-435">Режим создания</span><span class="sxs-lookup"><span data-stu-id="ce20c-435">Compose mode</span></span>

<span data-ttu-id="ce20c-436">Свойство `location` возвращает объект `Location`, предоставляющий методы, которые используются для получения и задания места встречи.</span><span class="sxs-lookup"><span data-stu-id="ce20c-436">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```javascript
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="ce20c-437">Тип</span><span class="sxs-lookup"><span data-stu-id="ce20c-437">Type</span></span>

*   <span data-ttu-id="ce20c-438">String | [Location](/javascript/api/outlook_1_5/office.location)</span><span class="sxs-lookup"><span data-stu-id="ce20c-438">String | [Location](/javascript/api/outlook_1_5/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="ce20c-439">Требования</span><span class="sxs-lookup"><span data-stu-id="ce20c-439">Requirements</span></span>

|<span data-ttu-id="ce20c-440">Требование</span><span class="sxs-lookup"><span data-stu-id="ce20c-440">Requirement</span></span>| <span data-ttu-id="ce20c-441">Значение</span><span class="sxs-lookup"><span data-stu-id="ce20c-441">Value</span></span>|
|---|---|
|[<span data-ttu-id="ce20c-442">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="ce20c-442">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ce20c-443">1.0</span><span class="sxs-lookup"><span data-stu-id="ce20c-443">1.0</span></span>|
|[<span data-ttu-id="ce20c-444">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ce20c-444">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ce20c-445">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ce20c-445">ReadItem</span></span>|
|[<span data-ttu-id="ce20c-446">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ce20c-446">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ce20c-447">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="ce20c-447">Compose or Read</span></span>|

#### <a name="normalizedsubject-string"></a><span data-ttu-id="ce20c-448">normalizedSubject: String</span><span class="sxs-lookup"><span data-stu-id="ce20c-448">normalizedSubject :String</span></span>

<span data-ttu-id="ce20c-p121">Получает тему элемента со всеми удаленными префиксами (включая `RE:` и `FWD:`). Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="ce20c-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="ce20c-p122">Свойство normalizedSubject получает тему элемента со стандартными префиксами (такими как `RE:` и `FW:`), добавляемыми почтовыми программами. Для получения темы элемента с неизмененными префиксами используйте свойство [`subject`](#subject-stringsubject).</span><span class="sxs-lookup"><span data-stu-id="ce20c-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="ce20c-453">Тип</span><span class="sxs-lookup"><span data-stu-id="ce20c-453">Type</span></span>

*   <span data-ttu-id="ce20c-454">String</span><span class="sxs-lookup"><span data-stu-id="ce20c-454">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="ce20c-455">Требования</span><span class="sxs-lookup"><span data-stu-id="ce20c-455">Requirements</span></span>

|<span data-ttu-id="ce20c-456">Требование</span><span class="sxs-lookup"><span data-stu-id="ce20c-456">Requirement</span></span>| <span data-ttu-id="ce20c-457">Значение</span><span class="sxs-lookup"><span data-stu-id="ce20c-457">Value</span></span>|
|---|---|
|[<span data-ttu-id="ce20c-458">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="ce20c-458">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ce20c-459">1.0</span><span class="sxs-lookup"><span data-stu-id="ce20c-459">1.0</span></span>|
|[<span data-ttu-id="ce20c-460">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ce20c-460">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ce20c-461">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ce20c-461">ReadItem</span></span>|
|[<span data-ttu-id="ce20c-462">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ce20c-462">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ce20c-463">Чтение</span><span class="sxs-lookup"><span data-stu-id="ce20c-463">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ce20c-464">Пример</span><span class="sxs-lookup"><span data-stu-id="ce20c-464">Example</span></span>

```javascript
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

#### <a name="notificationmessages-notificationmessagesjavascriptapioutlook15officenotificationmessages"></a><span data-ttu-id="ce20c-465">notificationMessages: [NotificationMessages](/javascript/api/outlook_1_5/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="ce20c-465">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_5/office.notificationmessages)</span></span>

<span data-ttu-id="ce20c-466">Получает сообщения уведомления для элемента.</span><span class="sxs-lookup"><span data-stu-id="ce20c-466">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="ce20c-467">Тип</span><span class="sxs-lookup"><span data-stu-id="ce20c-467">Type</span></span>

*   [<span data-ttu-id="ce20c-468">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="ce20c-468">NotificationMessages</span></span>](/javascript/api/outlook_1_5/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="ce20c-469">Требования</span><span class="sxs-lookup"><span data-stu-id="ce20c-469">Requirements</span></span>

|<span data-ttu-id="ce20c-470">Требование</span><span class="sxs-lookup"><span data-stu-id="ce20c-470">Requirement</span></span>| <span data-ttu-id="ce20c-471">Значение</span><span class="sxs-lookup"><span data-stu-id="ce20c-471">Value</span></span>|
|---|---|
|[<span data-ttu-id="ce20c-472">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="ce20c-472">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ce20c-473">1.3</span><span class="sxs-lookup"><span data-stu-id="ce20c-473">1.3</span></span>|
|[<span data-ttu-id="ce20c-474">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ce20c-474">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ce20c-475">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ce20c-475">ReadItem</span></span>|
|[<span data-ttu-id="ce20c-476">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ce20c-476">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ce20c-477">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="ce20c-477">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ce20c-478">Пример</span><span class="sxs-lookup"><span data-stu-id="ce20c-478">Example</span></span>

```javascript
// Get all notifications.
Office.context.mailbox.item.notificationMessages.getAllAsync(
  function (asyncResult) {
    console.log(JSON.stringify(asyncResult));
  }
);
```

#### <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients"></a><span data-ttu-id="ce20c-479">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="ce20c-479">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

<span data-ttu-id="ce20c-480">Предоставляет доступ к необязательным участникам события.</span><span class="sxs-lookup"><span data-stu-id="ce20c-480">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="ce20c-481">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="ce20c-481">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="ce20c-482">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="ce20c-482">Read mode</span></span>

<span data-ttu-id="ce20c-483">Свойство `optionalAttendees` возвращает массив, содержащий объект `EmailAddressDetails` для каждого необязательного участника собрания.</span><span class="sxs-lookup"><span data-stu-id="ce20c-483">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

```javascript
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="ce20c-484">Режим создания</span><span class="sxs-lookup"><span data-stu-id="ce20c-484">Compose mode</span></span>

<span data-ttu-id="ce20c-485">Свойство `optionalAttendees` возвращает объект `Recipients`, который предоставляет методы для получения или обновления необязательных участников собрания.</span><span class="sxs-lookup"><span data-stu-id="ce20c-485">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="ce20c-486">Тип</span><span class="sxs-lookup"><span data-stu-id="ce20c-486">Type</span></span>

*   <span data-ttu-id="ce20c-487">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="ce20c-487">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="ce20c-488">Требования</span><span class="sxs-lookup"><span data-stu-id="ce20c-488">Requirements</span></span>

|<span data-ttu-id="ce20c-489">Требование</span><span class="sxs-lookup"><span data-stu-id="ce20c-489">Requirement</span></span>| <span data-ttu-id="ce20c-490">Значение</span><span class="sxs-lookup"><span data-stu-id="ce20c-490">Value</span></span>|
|---|---|
|[<span data-ttu-id="ce20c-491">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="ce20c-491">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ce20c-492">1.0</span><span class="sxs-lookup"><span data-stu-id="ce20c-492">1.0</span></span>|
|[<span data-ttu-id="ce20c-493">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ce20c-493">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ce20c-494">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ce20c-494">ReadItem</span></span>|
|[<span data-ttu-id="ce20c-495">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ce20c-495">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ce20c-496">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="ce20c-496">Compose or Read</span></span>|

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails"></a><span data-ttu-id="ce20c-497">organizer: [EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="ce20c-497">organizer :[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)</span></span>

<span data-ttu-id="ce20c-p124">Получает электронный адрес организатора указанного собрания. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="ce20c-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="ce20c-500">Тип</span><span class="sxs-lookup"><span data-stu-id="ce20c-500">Type</span></span>

*   [<span data-ttu-id="ce20c-501">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="ce20c-501">EmailAddressDetails</span></span>](/javascript/api/outlook_1_5/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="ce20c-502">Требования</span><span class="sxs-lookup"><span data-stu-id="ce20c-502">Requirements</span></span>

|<span data-ttu-id="ce20c-503">Требование</span><span class="sxs-lookup"><span data-stu-id="ce20c-503">Requirement</span></span>| <span data-ttu-id="ce20c-504">Значение</span><span class="sxs-lookup"><span data-stu-id="ce20c-504">Value</span></span>|
|---|---|
|[<span data-ttu-id="ce20c-505">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="ce20c-505">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ce20c-506">1.0</span><span class="sxs-lookup"><span data-stu-id="ce20c-506">1.0</span></span>|
|[<span data-ttu-id="ce20c-507">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ce20c-507">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ce20c-508">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ce20c-508">ReadItem</span></span>|
|[<span data-ttu-id="ce20c-509">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ce20c-509">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ce20c-510">Чтение</span><span class="sxs-lookup"><span data-stu-id="ce20c-510">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ce20c-511">Пример</span><span class="sxs-lookup"><span data-stu-id="ce20c-511">Example</span></span>

```javascript
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

#### <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients"></a><span data-ttu-id="ce20c-512">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="ce20c-512">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

<span data-ttu-id="ce20c-513">Предоставляет доступ к обязательным участникам события.</span><span class="sxs-lookup"><span data-stu-id="ce20c-513">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="ce20c-514">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="ce20c-514">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="ce20c-515">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="ce20c-515">Read mode</span></span>

<span data-ttu-id="ce20c-516">Свойство `requiredAttendees` возвращает массив, содержащий объект `EmailAddressDetails` для каждого обязательного участника собрания.</span><span class="sxs-lookup"><span data-stu-id="ce20c-516">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

```javascript
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="ce20c-517">Режим создания</span><span class="sxs-lookup"><span data-stu-id="ce20c-517">Compose mode</span></span>

<span data-ttu-id="ce20c-518">Свойство `requiredAttendees` возвращает объект `Recipients`, предоставляющий методы, с помощью которых можно получить или обновить сведения об обязательных участниках собрания.</span><span class="sxs-lookup"><span data-stu-id="ce20c-518">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="ce20c-519">Тип</span><span class="sxs-lookup"><span data-stu-id="ce20c-519">Type</span></span>

*   <span data-ttu-id="ce20c-520">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="ce20c-520">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="ce20c-521">Требования</span><span class="sxs-lookup"><span data-stu-id="ce20c-521">Requirements</span></span>

|<span data-ttu-id="ce20c-522">Требование</span><span class="sxs-lookup"><span data-stu-id="ce20c-522">Requirement</span></span>| <span data-ttu-id="ce20c-523">Значение</span><span class="sxs-lookup"><span data-stu-id="ce20c-523">Value</span></span>|
|---|---|
|[<span data-ttu-id="ce20c-524">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="ce20c-524">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ce20c-525">1.0</span><span class="sxs-lookup"><span data-stu-id="ce20c-525">1.0</span></span>|
|[<span data-ttu-id="ce20c-526">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ce20c-526">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ce20c-527">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ce20c-527">ReadItem</span></span>|
|[<span data-ttu-id="ce20c-528">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ce20c-528">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ce20c-529">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="ce20c-529">Compose or Read</span></span>|

#### <a name="sender-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails"></a><span data-ttu-id="ce20c-530">sender: [EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="ce20c-530">sender :[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)</span></span>

<span data-ttu-id="ce20c-p126">Получает электронный адрес отправителя электронного сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="ce20c-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="ce20c-p127">Свойства [`from`](#from-emailaddressdetails) и `sender` представляют одно лицо, если сообщение не отправлено представителем. В противном случае свойство `from` представляет лицо, делегировавшее полномочия, а свойство sender — представителя.</span><span class="sxs-lookup"><span data-stu-id="ce20c-p127">The [`from`](#from-emailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="ce20c-535">Свойству `recipientType`, принадлежащему объекту `EmailAddressDetails` в свойстве `sender`, задано значение `undefined`.</span><span class="sxs-lookup"><span data-stu-id="ce20c-535">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="ce20c-536">Тип</span><span class="sxs-lookup"><span data-stu-id="ce20c-536">Type</span></span>

*   [<span data-ttu-id="ce20c-537">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="ce20c-537">EmailAddressDetails</span></span>](/javascript/api/outlook_1_5/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="ce20c-538">Требования</span><span class="sxs-lookup"><span data-stu-id="ce20c-538">Requirements</span></span>

|<span data-ttu-id="ce20c-539">Требование</span><span class="sxs-lookup"><span data-stu-id="ce20c-539">Requirement</span></span>| <span data-ttu-id="ce20c-540">Значение</span><span class="sxs-lookup"><span data-stu-id="ce20c-540">Value</span></span>|
|---|---|
|[<span data-ttu-id="ce20c-541">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="ce20c-541">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ce20c-542">1.0</span><span class="sxs-lookup"><span data-stu-id="ce20c-542">1.0</span></span>|
|[<span data-ttu-id="ce20c-543">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ce20c-543">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ce20c-544">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ce20c-544">ReadItem</span></span>|
|[<span data-ttu-id="ce20c-545">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ce20c-545">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ce20c-546">Чтение</span><span class="sxs-lookup"><span data-stu-id="ce20c-546">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ce20c-547">Пример</span><span class="sxs-lookup"><span data-stu-id="ce20c-547">Example</span></span>

```javascript
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

#### <a name="start-datetimejavascriptapioutlook15officetime"></a><span data-ttu-id="ce20c-548">start: Date|[Time](/javascript/api/outlook_1_5/office.time)</span><span class="sxs-lookup"><span data-stu-id="ce20c-548">start :Date|[Time](/javascript/api/outlook_1_5/office.time)</span></span>

<span data-ttu-id="ce20c-549">Получает или задает дату и время начала встречи.</span><span class="sxs-lookup"><span data-stu-id="ce20c-549">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="ce20c-p128">Свойство `start` представлено в виде значения даты и времени в формате UTC. Это значение можно преобразовать в местные значения даты и времени клиента с помощью метода [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime).</span><span class="sxs-lookup"><span data-stu-id="ce20c-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="ce20c-552">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="ce20c-552">Read mode</span></span>

<span data-ttu-id="ce20c-553">Свойство `start` возвращает объект `Date`.</span><span class="sxs-lookup"><span data-stu-id="ce20c-553">The `start` property returns a `Date` object.</span></span>

```javascript
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="ce20c-554">Режим создания</span><span class="sxs-lookup"><span data-stu-id="ce20c-554">Compose mode</span></span>

<span data-ttu-id="ce20c-555">Свойство `start` возвращает объект `Time`.</span><span class="sxs-lookup"><span data-stu-id="ce20c-555">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="ce20c-556">Если вы задаете время начала с помощью метода [`Time.setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-), необходимо использовать метод [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) для преобразования местного времени на клиенте в формат UTC для сервера.</span><span class="sxs-lookup"><span data-stu-id="ce20c-556">When you use the [`Time.setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="ce20c-557">В примере ниже с помощью метода [`setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) объекта `Time` задается время начала встречи в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="ce20c-557">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="ce20c-558">Тип</span><span class="sxs-lookup"><span data-stu-id="ce20c-558">Type</span></span>

*   <span data-ttu-id="ce20c-559">Date | [Time](/javascript/api/outlook_1_5/office.time)</span><span class="sxs-lookup"><span data-stu-id="ce20c-559">Date | [Time](/javascript/api/outlook_1_5/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="ce20c-560">Требования</span><span class="sxs-lookup"><span data-stu-id="ce20c-560">Requirements</span></span>

|<span data-ttu-id="ce20c-561">Требование</span><span class="sxs-lookup"><span data-stu-id="ce20c-561">Requirement</span></span>| <span data-ttu-id="ce20c-562">Значение</span><span class="sxs-lookup"><span data-stu-id="ce20c-562">Value</span></span>|
|---|---|
|[<span data-ttu-id="ce20c-563">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="ce20c-563">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ce20c-564">1.0</span><span class="sxs-lookup"><span data-stu-id="ce20c-564">1.0</span></span>|
|[<span data-ttu-id="ce20c-565">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ce20c-565">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ce20c-566">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ce20c-566">ReadItem</span></span>|
|[<span data-ttu-id="ce20c-567">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ce20c-567">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ce20c-568">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="ce20c-568">Compose or Read</span></span>|

#### <a name="subject-stringsubjectjavascriptapioutlook15officesubject"></a><span data-ttu-id="ce20c-569">subject: String|[Subject](/javascript/api/outlook_1_5/office.subject)</span><span class="sxs-lookup"><span data-stu-id="ce20c-569">subject :String|[Subject](/javascript/api/outlook_1_5/office.subject)</span></span>

<span data-ttu-id="ce20c-570">Получает или задает описание, которое отображается в поле темы элемента.</span><span class="sxs-lookup"><span data-stu-id="ce20c-570">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="ce20c-571">Свойство `subject` получает или задает всю тему элемента для отправки с почтового сервера.</span><span class="sxs-lookup"><span data-stu-id="ce20c-571">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="ce20c-572">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="ce20c-572">Read mode</span></span>

<span data-ttu-id="ce20c-p129">Свойство `subject` возвращает строку. С помощью свойства [`normalizedSubject`](#normalizedsubject-string) можно получить тему без начальных префиксов, таких как `RE:` и `FW:`.</span><span class="sxs-lookup"><span data-stu-id="ce20c-p129">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```javascript
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="ce20c-575">Режим создания</span><span class="sxs-lookup"><span data-stu-id="ce20c-575">Compose mode</span></span>

<span data-ttu-id="ce20c-576">Свойство `subject` возвращает объект `Subject`, который предоставляет методы для получения и задания темы.</span><span class="sxs-lookup"><span data-stu-id="ce20c-576">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```javascript
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="ce20c-577">Тип</span><span class="sxs-lookup"><span data-stu-id="ce20c-577">Type</span></span>

*   <span data-ttu-id="ce20c-578">String | [Subject](/javascript/api/outlook_1_5/office.subject)</span><span class="sxs-lookup"><span data-stu-id="ce20c-578">String | [Subject](/javascript/api/outlook_1_5/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="ce20c-579">Требования</span><span class="sxs-lookup"><span data-stu-id="ce20c-579">Requirements</span></span>

|<span data-ttu-id="ce20c-580">Требование</span><span class="sxs-lookup"><span data-stu-id="ce20c-580">Requirement</span></span>| <span data-ttu-id="ce20c-581">Значение</span><span class="sxs-lookup"><span data-stu-id="ce20c-581">Value</span></span>|
|---|---|
|[<span data-ttu-id="ce20c-582">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="ce20c-582">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ce20c-583">1.0</span><span class="sxs-lookup"><span data-stu-id="ce20c-583">1.0</span></span>|
|[<span data-ttu-id="ce20c-584">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ce20c-584">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ce20c-585">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ce20c-585">ReadItem</span></span>|
|[<span data-ttu-id="ce20c-586">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ce20c-586">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ce20c-587">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="ce20c-587">Compose or Read</span></span>|

#### <a name="to-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients"></a><span data-ttu-id="ce20c-588">to: Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="ce20c-588">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

<span data-ttu-id="ce20c-589">Предоставляет доступ к получателям, указанным в строке **Кому** сообщения.</span><span class="sxs-lookup"><span data-stu-id="ce20c-589">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="ce20c-590">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="ce20c-590">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="ce20c-591">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="ce20c-591">Read mode</span></span>

<span data-ttu-id="ce20c-p131">Свойство `to` возвращает массив, содержащий объект `EmailAddressDetails` для каждого получателя в строке **Кому** сообщения. Коллекция может включать не более 100 элементов.</span><span class="sxs-lookup"><span data-stu-id="ce20c-p131">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="ce20c-594">Режим создания</span><span class="sxs-lookup"><span data-stu-id="ce20c-594">Compose mode</span></span>

<span data-ttu-id="ce20c-595">Свойство `to` возвращает объект `Recipients`, предоставляющий методы для получения или обновления получателей, которые указаны в строке **Кому** сообщения.</span><span class="sxs-lookup"><span data-stu-id="ce20c-595">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

```javascript
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="ce20c-596">Тип</span><span class="sxs-lookup"><span data-stu-id="ce20c-596">Type</span></span>

*   <span data-ttu-id="ce20c-597">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="ce20c-597">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="ce20c-598">Требования</span><span class="sxs-lookup"><span data-stu-id="ce20c-598">Requirements</span></span>

|<span data-ttu-id="ce20c-599">Требование</span><span class="sxs-lookup"><span data-stu-id="ce20c-599">Requirement</span></span>| <span data-ttu-id="ce20c-600">Значение</span><span class="sxs-lookup"><span data-stu-id="ce20c-600">Value</span></span>|
|---|---|
|[<span data-ttu-id="ce20c-601">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="ce20c-601">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ce20c-602">1.0</span><span class="sxs-lookup"><span data-stu-id="ce20c-602">1.0</span></span>|
|[<span data-ttu-id="ce20c-603">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ce20c-603">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ce20c-604">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ce20c-604">ReadItem</span></span>|
|[<span data-ttu-id="ce20c-605">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ce20c-605">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ce20c-606">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="ce20c-606">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="ce20c-607">Методы</span><span class="sxs-lookup"><span data-stu-id="ce20c-607">Methods</span></span>

#### <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="ce20c-608">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="ce20c-608">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="ce20c-609">Добавляет файл в сообщение или встречу в качестве вложения.</span><span class="sxs-lookup"><span data-stu-id="ce20c-609">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="ce20c-610">Метод `addFileAttachmentAsync` передает файл по указанному универсальному коду ресурса (URI) и вкладывает его в элемент в форме создания.</span><span class="sxs-lookup"><span data-stu-id="ce20c-610">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="ce20c-611">Идентификатор можно использовать с методом [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback), чтобы удалить вложение, добавленное во время текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="ce20c-611">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ce20c-612">Параметры</span><span class="sxs-lookup"><span data-stu-id="ce20c-612">Parameters</span></span>

|<span data-ttu-id="ce20c-613">Имя</span><span class="sxs-lookup"><span data-stu-id="ce20c-613">Name</span></span>| <span data-ttu-id="ce20c-614">Тип</span><span class="sxs-lookup"><span data-stu-id="ce20c-614">Type</span></span>| <span data-ttu-id="ce20c-615">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="ce20c-615">Attributes</span></span>| <span data-ttu-id="ce20c-616">Описание</span><span class="sxs-lookup"><span data-stu-id="ce20c-616">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="ce20c-617">String</span><span class="sxs-lookup"><span data-stu-id="ce20c-617">String</span></span>||<span data-ttu-id="ce20c-p132">Универсальный код ресурса (URI), представляющий расположение файла, который нужно вложить в сообщение или встречу. Максимальная длина — 2048 символов.</span><span class="sxs-lookup"><span data-stu-id="ce20c-p132">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="ce20c-620">String</span><span class="sxs-lookup"><span data-stu-id="ce20c-620">String</span></span>||<span data-ttu-id="ce20c-p133">Имя вложения, которое отображается при передаче вложения. Максимальная длина — 255 символов.</span><span class="sxs-lookup"><span data-stu-id="ce20c-p133">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="ce20c-623">Object</span><span class="sxs-lookup"><span data-stu-id="ce20c-623">Object</span></span>| <span data-ttu-id="ce20c-624">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="ce20c-624">&lt;optional&gt;</span></span>|<span data-ttu-id="ce20c-625">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="ce20c-625">An object literal that contains one or more of the following properties.</span></span>|
| `options.asyncContext` | <span data-ttu-id="ce20c-626">Object</span><span class="sxs-lookup"><span data-stu-id="ce20c-626">Object</span></span> | <span data-ttu-id="ce20c-627">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="ce20c-627">&lt;optional&gt;</span></span> | <span data-ttu-id="ce20c-628">В методе обратного вызова разработчики могут указать любой объект, к которому необходимо получить доступ.</span><span class="sxs-lookup"><span data-stu-id="ce20c-628">Developers can provide any object they wish to access in the callback method.</span></span> |
| `options.isInline` | <span data-ttu-id="ce20c-629">Boolean</span><span class="sxs-lookup"><span data-stu-id="ce20c-629">Boolean</span></span> | <span data-ttu-id="ce20c-630">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="ce20c-630">&lt;optional&gt;</span></span> | <span data-ttu-id="ce20c-631">Значение `true` указывает на то, что вложение будет встроено в текст сообщения и не должно отображаться в списке вложений.</span><span class="sxs-lookup"><span data-stu-id="ce20c-631">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
|`callback`| <span data-ttu-id="ce20c-632">function</span><span class="sxs-lookup"><span data-stu-id="ce20c-632">function</span></span>| <span data-ttu-id="ce20c-633">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="ce20c-633">&lt;optional&gt;</span></span>|<span data-ttu-id="ce20c-634">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="ce20c-634">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="ce20c-635">После успешного выполнения идентификатор вложения будет представлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="ce20c-635">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="ce20c-636">Если добавить вложение не удастся, объект `asyncResult` будет содержать объект `Error` с описанием ошибки.</span><span class="sxs-lookup"><span data-stu-id="ce20c-636">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="ce20c-637">Ошибки</span><span class="sxs-lookup"><span data-stu-id="ce20c-637">Errors</span></span>

| <span data-ttu-id="ce20c-638">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="ce20c-638">Error code</span></span> | <span data-ttu-id="ce20c-639">Описание</span><span class="sxs-lookup"><span data-stu-id="ce20c-639">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="ce20c-640">Вложение превышает максимальный размер.</span><span class="sxs-lookup"><span data-stu-id="ce20c-640">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="ce20c-641">Расширение вложения не поддерживается.</span><span class="sxs-lookup"><span data-stu-id="ce20c-641">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="ce20c-642">Сообщение или встреча содержат слишком много вложений.</span><span class="sxs-lookup"><span data-stu-id="ce20c-642">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="ce20c-643">Требования</span><span class="sxs-lookup"><span data-stu-id="ce20c-643">Requirements</span></span>

|<span data-ttu-id="ce20c-644">Требование</span><span class="sxs-lookup"><span data-stu-id="ce20c-644">Requirement</span></span>| <span data-ttu-id="ce20c-645">Значение</span><span class="sxs-lookup"><span data-stu-id="ce20c-645">Value</span></span>|
|---|---|
|[<span data-ttu-id="ce20c-646">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="ce20c-646">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ce20c-647">1.1</span><span class="sxs-lookup"><span data-stu-id="ce20c-647">1.1</span></span>|
|[<span data-ttu-id="ce20c-648">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ce20c-648">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ce20c-649">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="ce20c-649">ReadWriteItem</span></span>|
|[<span data-ttu-id="ce20c-650">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ce20c-650">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ce20c-651">Создание</span><span class="sxs-lookup"><span data-stu-id="ce20c-651">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="ce20c-652">Примеры</span><span class="sxs-lookup"><span data-stu-id="ce20c-652">Examples</span></span>

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

<span data-ttu-id="ce20c-653">В приведенном ниже примере файл изображения добавляется в качестве встроенного вложения, а ссылка на вложение добавляется в текст сообщения.</span><span class="sxs-lookup"><span data-stu-id="ce20c-653">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

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

#### <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="ce20c-654">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="ce20c-654">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="ce20c-655">Добавляет к сообщению элемент Exchange, например сообщение, в виде вложения.</span><span class="sxs-lookup"><span data-stu-id="ce20c-655">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="ce20c-p134">С помощью метода `addItemAttachmentAsync` можно в элемент формы создания вложить элемент с указанным идентификатором Exchange. Если указать метод обратного вызова, то этот метод вызывается с помощью параметра `asyncResult`, который содержит идентификатор вложения или код, указывающий на ошибки, которые произошли при вложении элемента. При необходимости можно использовать параметр `options` для передачи сведений о состоянии методу обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="ce20c-p134">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="ce20c-659">Идентификатор можно использовать с методом [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback), чтобы удалить вложение, добавленное во время текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="ce20c-659">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="ce20c-660">Если ваша надстройка Office выполняется в Outlook в Интернете, метод `addItemAttachmentAsync` обеспечивает вложение элементов в элементы, отличные от редактируемого. Однако это действие не рекомендуется выполнять, так как оно не поддерживается.</span><span class="sxs-lookup"><span data-stu-id="ce20c-660">If your Office Add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ce20c-661">Параметры</span><span class="sxs-lookup"><span data-stu-id="ce20c-661">Parameters</span></span>

|<span data-ttu-id="ce20c-662">Имя</span><span class="sxs-lookup"><span data-stu-id="ce20c-662">Name</span></span>| <span data-ttu-id="ce20c-663">Тип</span><span class="sxs-lookup"><span data-stu-id="ce20c-663">Type</span></span>| <span data-ttu-id="ce20c-664">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="ce20c-664">Attributes</span></span>| <span data-ttu-id="ce20c-665">Описание</span><span class="sxs-lookup"><span data-stu-id="ce20c-665">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="ce20c-666">String</span><span class="sxs-lookup"><span data-stu-id="ce20c-666">String</span></span>||<span data-ttu-id="ce20c-p135">Идентификатор Exchange для вкладываемого элемента. Максимальная длина — 100 символов.</span><span class="sxs-lookup"><span data-stu-id="ce20c-p135">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="ce20c-669">String</span><span class="sxs-lookup"><span data-stu-id="ce20c-669">String</span></span>||<span data-ttu-id="ce20c-670">Тема вкладываемого элемента.</span><span class="sxs-lookup"><span data-stu-id="ce20c-670">The subject of the item to be attached.</span></span> <span data-ttu-id="ce20c-671">Максимальная длина: 255 символов.</span><span class="sxs-lookup"><span data-stu-id="ce20c-671">The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="ce20c-672">Object</span><span class="sxs-lookup"><span data-stu-id="ce20c-672">Object</span></span>| <span data-ttu-id="ce20c-673">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="ce20c-673">&lt;optional&gt;</span></span>|<span data-ttu-id="ce20c-674">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="ce20c-674">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="ce20c-675">Object</span><span class="sxs-lookup"><span data-stu-id="ce20c-675">Object</span></span>| <span data-ttu-id="ce20c-676">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="ce20c-676">&lt;optional&gt;</span></span>|<span data-ttu-id="ce20c-677">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="ce20c-677">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="ce20c-678">функция</span><span class="sxs-lookup"><span data-stu-id="ce20c-678">function</span></span>| <span data-ttu-id="ce20c-679">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="ce20c-679">&lt;optional&gt;</span></span>|<span data-ttu-id="ce20c-680">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="ce20c-680">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="ce20c-681">После успешного выполнения идентификатор вложения будет представлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="ce20c-681">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="ce20c-682">Если добавить вложение не удастся, объект `asyncResult` будет содержать объект `Error` с описанием ошибки.</span><span class="sxs-lookup"><span data-stu-id="ce20c-682">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="ce20c-683">Ошибки</span><span class="sxs-lookup"><span data-stu-id="ce20c-683">Errors</span></span>

| <span data-ttu-id="ce20c-684">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="ce20c-684">Error code</span></span> | <span data-ttu-id="ce20c-685">Описание</span><span class="sxs-lookup"><span data-stu-id="ce20c-685">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="ce20c-686">Сообщение или встреча содержат слишком много вложений.</span><span class="sxs-lookup"><span data-stu-id="ce20c-686">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="ce20c-687">Требования</span><span class="sxs-lookup"><span data-stu-id="ce20c-687">Requirements</span></span>

|<span data-ttu-id="ce20c-688">Требование</span><span class="sxs-lookup"><span data-stu-id="ce20c-688">Requirement</span></span>| <span data-ttu-id="ce20c-689">Значение</span><span class="sxs-lookup"><span data-stu-id="ce20c-689">Value</span></span>|
|---|---|
|[<span data-ttu-id="ce20c-690">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="ce20c-690">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ce20c-691">1.1</span><span class="sxs-lookup"><span data-stu-id="ce20c-691">1.1</span></span>|
|[<span data-ttu-id="ce20c-692">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ce20c-692">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ce20c-693">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="ce20c-693">ReadWriteItem</span></span>|
|[<span data-ttu-id="ce20c-694">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ce20c-694">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ce20c-695">Создание</span><span class="sxs-lookup"><span data-stu-id="ce20c-695">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="ce20c-696">Пример</span><span class="sxs-lookup"><span data-stu-id="ce20c-696">Example</span></span>

<span data-ttu-id="ce20c-697">В следующем примере существующий элемент Outlook добавляется в виде вложения с именем `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="ce20c-697">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

#### <a name="close"></a><span data-ttu-id="ce20c-698">close()</span><span class="sxs-lookup"><span data-stu-id="ce20c-698">close()</span></span>

<span data-ttu-id="ce20c-699">Закрывает текущий создаваемый элемент.</span><span class="sxs-lookup"><span data-stu-id="ce20c-699">Closes the current item that is being composed.</span></span>

<span data-ttu-id="ce20c-p137">Работа метода `close` зависит от текущего состояния создаваемого элемента. Если элемент содержит несохраненные изменения, клиент предложит пользователю сохранить или отклонить их либо отменить действие закрытия.</span><span class="sxs-lookup"><span data-stu-id="ce20c-p137">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="ce20c-702">Если в Outlook в Интернете элемент представляет собой встречу, ранее сохраненную с помощью метода `saveAsync`, пользователю предлагается сохранить, отклонить или отменить действие, даже если с момента последнего сохранения элемента не вносились какие-либо изменения.</span><span class="sxs-lookup"><span data-stu-id="ce20c-702">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="ce20c-703">Если в клиенте Outlook для настольных ПК сообщение представляет собой ответ в тексте, метод `close` не работает.</span><span class="sxs-lookup"><span data-stu-id="ce20c-703">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="ce20c-704">Требования</span><span class="sxs-lookup"><span data-stu-id="ce20c-704">Requirements</span></span>

|<span data-ttu-id="ce20c-705">Требование</span><span class="sxs-lookup"><span data-stu-id="ce20c-705">Requirement</span></span>| <span data-ttu-id="ce20c-706">Значение</span><span class="sxs-lookup"><span data-stu-id="ce20c-706">Value</span></span>|
|---|---|
|[<span data-ttu-id="ce20c-707">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="ce20c-707">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ce20c-708">1.3</span><span class="sxs-lookup"><span data-stu-id="ce20c-708">1.3</span></span>|
|[<span data-ttu-id="ce20c-709">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ce20c-709">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ce20c-710">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="ce20c-710">Restricted</span></span>|
|[<span data-ttu-id="ce20c-711">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ce20c-711">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ce20c-712">Создание</span><span class="sxs-lookup"><span data-stu-id="ce20c-712">Compose</span></span>|

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="ce20c-713">displayReplyAllForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="ce20c-713">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="ce20c-714">Отображает форму ответа, включающую отправителя и всех получателей выбранного сообщения или организатора и всех участников выбранной встречи.</span><span class="sxs-lookup"><span data-stu-id="ce20c-714">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="ce20c-715">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="ce20c-715">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="ce20c-716">В Outlook в Интернете форма ответа отображается в виде всплывающей формы в представлении с 3 либо 1 или 2 столбцами.</span><span class="sxs-lookup"><span data-stu-id="ce20c-716">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="ce20c-717">Если любой строковый параметр превышает указанные для него ограничения, `displayReplyAllForm` возвращает исключение.</span><span class="sxs-lookup"><span data-stu-id="ce20c-717">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="ce20c-p138">Если в параметре `formData.attachments` указаны вложения, Outlook в Интернете и классические клиенты пытаются скачать их и вложить в форму ответа. Если какие-либо вложения добавить не удается, в форме отображается сообщение об ошибке. Если сообщения об ошибках не предусмотрены, то они не отображаются.</span><span class="sxs-lookup"><span data-stu-id="ce20c-p138">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ce20c-721">Параметры</span><span class="sxs-lookup"><span data-stu-id="ce20c-721">Parameters</span></span>

| <span data-ttu-id="ce20c-722">Имя</span><span class="sxs-lookup"><span data-stu-id="ce20c-722">Name</span></span> | <span data-ttu-id="ce20c-723">Тип</span><span class="sxs-lookup"><span data-stu-id="ce20c-723">Type</span></span> | <span data-ttu-id="ce20c-724">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="ce20c-724">Attributes</span></span> | <span data-ttu-id="ce20c-725">Описание</span><span class="sxs-lookup"><span data-stu-id="ce20c-725">Description</span></span> |
|---|---|---|---|
|`formData`| <span data-ttu-id="ce20c-726">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="ce20c-726">String &#124; Object</span></span>| |<span data-ttu-id="ce20c-p139">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="ce20c-p139">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="ce20c-729">**ИЛИ**</span><span class="sxs-lookup"><span data-stu-id="ce20c-729">**OR**</span></span><br/><span data-ttu-id="ce20c-p140">Объект, который содержит текст или данные вложения и функцию обратного вызова. Ниже представлено определение этого объекта.</span><span class="sxs-lookup"><span data-stu-id="ce20c-p140">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="ce20c-732">String</span><span class="sxs-lookup"><span data-stu-id="ce20c-732">String</span></span> | <span data-ttu-id="ce20c-733">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="ce20c-733">&lt;optional&gt;</span></span> | <span data-ttu-id="ce20c-p141">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="ce20c-p141">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="ce20c-736">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="ce20c-736">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="ce20c-737">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="ce20c-737">&lt;optional&gt;</span></span> | <span data-ttu-id="ce20c-738">Массив объектов JSON, представляющих собой вложенные файлы или элементы.</span><span class="sxs-lookup"><span data-stu-id="ce20c-738">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="ce20c-739">String</span><span class="sxs-lookup"><span data-stu-id="ce20c-739">String</span></span> | | <span data-ttu-id="ce20c-p142">Указывает тип вложения. Допустимые значения: `file` для вложенного файла и `item` для вложенного элемента.</span><span class="sxs-lookup"><span data-stu-id="ce20c-p142">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="ce20c-742">String</span><span class="sxs-lookup"><span data-stu-id="ce20c-742">String</span></span> | | <span data-ttu-id="ce20c-743">Строка, содержащая имя вложения, длиной до 255 символов.</span><span class="sxs-lookup"><span data-stu-id="ce20c-743">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="ce20c-744">String</span><span class="sxs-lookup"><span data-stu-id="ce20c-744">String</span></span> | | <span data-ttu-id="ce20c-p143">Используется, только если свойству `type` задано значение `file`. URI расположения файла.</span><span class="sxs-lookup"><span data-stu-id="ce20c-p143">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.isInline` | <span data-ttu-id="ce20c-747">Логический</span><span class="sxs-lookup"><span data-stu-id="ce20c-747">Boolean</span></span> | | <span data-ttu-id="ce20c-p144">Используется, только если свойству `type` задано значение `file`. Значение `true` указывает на то, что вложение будет встроено в текст сообщения и не должно отображаться в списке вложений.</span><span class="sxs-lookup"><span data-stu-id="ce20c-p144">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="ce20c-750">String</span><span class="sxs-lookup"><span data-stu-id="ce20c-750">String</span></span> | | <span data-ttu-id="ce20c-p145">Используется, только если свойству `type` задано значение `item`. Идентификатор вложения EWS. Это строка длиной до 100 символов.</span><span class="sxs-lookup"><span data-stu-id="ce20c-p145">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="ce20c-754">function</span><span class="sxs-lookup"><span data-stu-id="ce20c-754">function</span></span> | <span data-ttu-id="ce20c-755">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="ce20c-755">&lt;optional&gt;</span></span> | <span data-ttu-id="ce20c-756">По завершении работы метода функция, переданная параметру `callback`, вызывается с помощью одного параметра `asyncResult`, представляющего собой объект [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="ce20c-756">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="ce20c-757">Требования</span><span class="sxs-lookup"><span data-stu-id="ce20c-757">Requirements</span></span>

|<span data-ttu-id="ce20c-758">Требование</span><span class="sxs-lookup"><span data-stu-id="ce20c-758">Requirement</span></span>| <span data-ttu-id="ce20c-759">Значение</span><span class="sxs-lookup"><span data-stu-id="ce20c-759">Value</span></span>|
|---|---|
|[<span data-ttu-id="ce20c-760">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="ce20c-760">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ce20c-761">1.0</span><span class="sxs-lookup"><span data-stu-id="ce20c-761">1.0</span></span>|
|[<span data-ttu-id="ce20c-762">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ce20c-762">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ce20c-763">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ce20c-763">ReadItem</span></span>|
|[<span data-ttu-id="ce20c-764">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ce20c-764">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ce20c-765">Чтение</span><span class="sxs-lookup"><span data-stu-id="ce20c-765">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="ce20c-766">Примеры</span><span class="sxs-lookup"><span data-stu-id="ce20c-766">Examples</span></span>

<span data-ttu-id="ce20c-767">Приведенный ниже код передает строку в функцию `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="ce20c-767">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="ce20c-768">Ответ с пустым текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="ce20c-768">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="ce20c-769">Ответ только с текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="ce20c-769">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="ce20c-770">Ответ с текстом сообщения и вложенным файлом.</span><span class="sxs-lookup"><span data-stu-id="ce20c-770">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="ce20c-771">Ответ с текстом сообщения и вложенным элементом.</span><span class="sxs-lookup"><span data-stu-id="ce20c-771">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="ce20c-772">Ответ с текстом сообщения, вложенным файлом, вложенным элементом и обратным вызовом.</span><span class="sxs-lookup"><span data-stu-id="ce20c-772">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="ce20c-773">displayReplyForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="ce20c-773">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="ce20c-774">Отображает форму ответа, включающую только отправителя выбранного сообщения или организатора выбранной встречи.</span><span class="sxs-lookup"><span data-stu-id="ce20c-774">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="ce20c-775">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="ce20c-775">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="ce20c-776">В Outlook в Интернете форма ответа отображается в виде всплывающей формы в представлении с 3 либо 1 или 2 столбцами.</span><span class="sxs-lookup"><span data-stu-id="ce20c-776">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="ce20c-777">Если любой строковый параметр превышает указанные для него ограничения, `displayReplyForm` возвращает исключение.</span><span class="sxs-lookup"><span data-stu-id="ce20c-777">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="ce20c-p146">Если в параметре `formData.attachments` указаны вложения, Outlook в Интернете и классические клиенты пытаются скачать их и вложить в форму ответа. Если какие-либо вложения добавить не удается, в форме отображается сообщение об ошибке. Если сообщения об ошибках не предусмотрены, то они не отображаются.</span><span class="sxs-lookup"><span data-stu-id="ce20c-p146">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ce20c-781">Параметры</span><span class="sxs-lookup"><span data-stu-id="ce20c-781">Parameters</span></span>

| <span data-ttu-id="ce20c-782">Имя</span><span class="sxs-lookup"><span data-stu-id="ce20c-782">Name</span></span> | <span data-ttu-id="ce20c-783">Тип</span><span class="sxs-lookup"><span data-stu-id="ce20c-783">Type</span></span> | <span data-ttu-id="ce20c-784">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="ce20c-784">Attributes</span></span> | <span data-ttu-id="ce20c-785">Описание</span><span class="sxs-lookup"><span data-stu-id="ce20c-785">Description</span></span> |
|---|---|---|---|
|`formData`| <span data-ttu-id="ce20c-786">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="ce20c-786">String &#124; Object</span></span>| | <span data-ttu-id="ce20c-p147">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="ce20c-p147">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="ce20c-789">**ИЛИ**</span><span class="sxs-lookup"><span data-stu-id="ce20c-789">**OR**</span></span><br/><span data-ttu-id="ce20c-p148">Объект, который содержит текст или данные вложения и функцию обратного вызова. Ниже представлено определение этого объекта.</span><span class="sxs-lookup"><span data-stu-id="ce20c-p148">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="ce20c-792">String</span><span class="sxs-lookup"><span data-stu-id="ce20c-792">String</span></span> | <span data-ttu-id="ce20c-793">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="ce20c-793">&lt;optional&gt;</span></span> | <span data-ttu-id="ce20c-p149">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="ce20c-p149">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="ce20c-796">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="ce20c-796">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="ce20c-797">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="ce20c-797">&lt;optional&gt;</span></span> | <span data-ttu-id="ce20c-798">Массив объектов JSON, представляющих собой вложенные файлы или элементы.</span><span class="sxs-lookup"><span data-stu-id="ce20c-798">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="ce20c-799">String</span><span class="sxs-lookup"><span data-stu-id="ce20c-799">String</span></span> | | <span data-ttu-id="ce20c-p150">Указывает тип вложения. Допустимые значения: `file` для вложенного файла и `item` для вложенного элемента.</span><span class="sxs-lookup"><span data-stu-id="ce20c-p150">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="ce20c-802">String</span><span class="sxs-lookup"><span data-stu-id="ce20c-802">String</span></span> | | <span data-ttu-id="ce20c-803">Строка, содержащая имя вложения, длиной до 255 символов.</span><span class="sxs-lookup"><span data-stu-id="ce20c-803">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="ce20c-804">String</span><span class="sxs-lookup"><span data-stu-id="ce20c-804">String</span></span> | | <span data-ttu-id="ce20c-p151">Используется, только если свойству `type` задано значение `file`. URI расположения файла.</span><span class="sxs-lookup"><span data-stu-id="ce20c-p151">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.isInline` | <span data-ttu-id="ce20c-807">Логический</span><span class="sxs-lookup"><span data-stu-id="ce20c-807">Boolean</span></span> | | <span data-ttu-id="ce20c-p152">Используется, только если свойству `type` задано значение `file`. Значение `true` указывает на то, что вложение будет встроено в текст сообщения и не должно отображаться в списке вложений.</span><span class="sxs-lookup"><span data-stu-id="ce20c-p152">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="ce20c-810">String</span><span class="sxs-lookup"><span data-stu-id="ce20c-810">String</span></span> | | <span data-ttu-id="ce20c-p153">Используется, только если свойству `type` задано значение `item`. Идентификатор вложения EWS. Это строка длиной до 100 символов.</span><span class="sxs-lookup"><span data-stu-id="ce20c-p153">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="ce20c-814">function</span><span class="sxs-lookup"><span data-stu-id="ce20c-814">function</span></span> | <span data-ttu-id="ce20c-815">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="ce20c-815">&lt;optional&gt;</span></span> | <span data-ttu-id="ce20c-816">По завершении работы метода функция, переданная параметру `callback`, вызывается с помощью одного параметра `asyncResult`, представляющего собой объект [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="ce20c-816">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="ce20c-817">Требования</span><span class="sxs-lookup"><span data-stu-id="ce20c-817">Requirements</span></span>

|<span data-ttu-id="ce20c-818">Требование</span><span class="sxs-lookup"><span data-stu-id="ce20c-818">Requirement</span></span>| <span data-ttu-id="ce20c-819">Значение</span><span class="sxs-lookup"><span data-stu-id="ce20c-819">Value</span></span>|
|---|---|
|[<span data-ttu-id="ce20c-820">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="ce20c-820">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ce20c-821">1.0</span><span class="sxs-lookup"><span data-stu-id="ce20c-821">1.0</span></span>|
|[<span data-ttu-id="ce20c-822">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ce20c-822">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ce20c-823">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ce20c-823">ReadItem</span></span>|
|[<span data-ttu-id="ce20c-824">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ce20c-824">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ce20c-825">Чтение</span><span class="sxs-lookup"><span data-stu-id="ce20c-825">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="ce20c-826">Примеры</span><span class="sxs-lookup"><span data-stu-id="ce20c-826">Examples</span></span>

<span data-ttu-id="ce20c-827">Приведенный ниже код передает строку в функцию `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="ce20c-827">The following code passes a string to the `displayReplyForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="ce20c-828">Ответ с пустым текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="ce20c-828">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="ce20c-829">Ответ только с текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="ce20c-829">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="ce20c-830">Ответ с текстом сообщения и вложенным файлом.</span><span class="sxs-lookup"><span data-stu-id="ce20c-830">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="ce20c-831">Ответ с текстом сообщения и вложенным элементом.</span><span class="sxs-lookup"><span data-stu-id="ce20c-831">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="ce20c-832">Ответ с текстом сообщения, вложенным файлом, вложенным элементом и обратным вызовом.</span><span class="sxs-lookup"><span data-stu-id="ce20c-832">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlook15officeentities"></a><span data-ttu-id="ce20c-833">getEntities() → {[Entities](/javascript/api/outlook_1_5/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="ce20c-833">getEntities() → {[Entities](/javascript/api/outlook_1_5/office.entities)}</span></span>

<span data-ttu-id="ce20c-834">Получает сущности, обнаруженные в теле выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="ce20c-834">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="ce20c-835">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="ce20c-835">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="ce20c-836">Требования</span><span class="sxs-lookup"><span data-stu-id="ce20c-836">Requirements</span></span>

|<span data-ttu-id="ce20c-837">Требование</span><span class="sxs-lookup"><span data-stu-id="ce20c-837">Requirement</span></span>| <span data-ttu-id="ce20c-838">Значение</span><span class="sxs-lookup"><span data-stu-id="ce20c-838">Value</span></span>|
|---|---|
|[<span data-ttu-id="ce20c-839">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="ce20c-839">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ce20c-840">1.0</span><span class="sxs-lookup"><span data-stu-id="ce20c-840">1.0</span></span>|
|[<span data-ttu-id="ce20c-841">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ce20c-841">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ce20c-842">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ce20c-842">ReadItem</span></span>|
|[<span data-ttu-id="ce20c-843">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ce20c-843">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ce20c-844">Чтение</span><span class="sxs-lookup"><span data-stu-id="ce20c-844">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="ce20c-845">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="ce20c-845">Returns:</span></span>

<span data-ttu-id="ce20c-846">Тип: [Entities](/javascript/api/outlook_1_5/office.entities)</span><span class="sxs-lookup"><span data-stu-id="ce20c-846">Type: [Entities](/javascript/api/outlook_1_5/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="ce20c-847">Пример</span><span class="sxs-lookup"><span data-stu-id="ce20c-847">Example</span></span>

<span data-ttu-id="ce20c-848">Ниже приведен пример получения доступа к сущностям контактов в теле текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="ce20c-848">The following example accesses the contacts entities in the current item's body.</span></span>

```javascript
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook15officecontactmeetingsuggestionjavascriptapioutlook15officemeetingsuggestionphonenumberjavascriptapioutlook15officephonenumbertasksuggestionjavascriptapioutlook15officetasksuggestion"></a><span data-ttu-id="ce20c-849">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="ce20c-849">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))>}</span></span>

<span data-ttu-id="ce20c-850">Получает массив всех сущностей указанного типа, обнаруженных в теле выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="ce20c-850">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="ce20c-851">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="ce20c-851">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ce20c-852">Параметры</span><span class="sxs-lookup"><span data-stu-id="ce20c-852">Parameters</span></span>

|<span data-ttu-id="ce20c-853">Имя</span><span class="sxs-lookup"><span data-stu-id="ce20c-853">Name</span></span>| <span data-ttu-id="ce20c-854">Тип</span><span class="sxs-lookup"><span data-stu-id="ce20c-854">Type</span></span>| <span data-ttu-id="ce20c-855">Описание</span><span class="sxs-lookup"><span data-stu-id="ce20c-855">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="ce20c-856">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="ce20c-856">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook_1_5/office.mailboxenums.entitytype)|<span data-ttu-id="ce20c-857">Одно из значений перечисления EntityType.</span><span class="sxs-lookup"><span data-stu-id="ce20c-857">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ce20c-858">Требования</span><span class="sxs-lookup"><span data-stu-id="ce20c-858">Requirements</span></span>

|<span data-ttu-id="ce20c-859">Требование</span><span class="sxs-lookup"><span data-stu-id="ce20c-859">Requirement</span></span>| <span data-ttu-id="ce20c-860">Значение</span><span class="sxs-lookup"><span data-stu-id="ce20c-860">Value</span></span>|
|---|---|
|[<span data-ttu-id="ce20c-861">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="ce20c-861">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ce20c-862">1.0</span><span class="sxs-lookup"><span data-stu-id="ce20c-862">1.0</span></span>|
|[<span data-ttu-id="ce20c-863">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ce20c-863">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ce20c-864">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="ce20c-864">Restricted</span></span>|
|[<span data-ttu-id="ce20c-865">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ce20c-865">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ce20c-866">Чтение</span><span class="sxs-lookup"><span data-stu-id="ce20c-866">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="ce20c-867">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="ce20c-867">Returns:</span></span>

<span data-ttu-id="ce20c-868">Если значение, переданное в `entityType`, не является допустимым членом перечисления `EntityType`, метод возвращает значение NULL.</span><span class="sxs-lookup"><span data-stu-id="ce20c-868">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="ce20c-869">Если в теле элемента отсутствуют сущности указанного типа, метод возвращает пустой массив.</span><span class="sxs-lookup"><span data-stu-id="ce20c-869">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="ce20c-870">В противном случае тип объектов в возвращаемом массиве зависит от типа сущности, запрошенной в параметре `entityType`.</span><span class="sxs-lookup"><span data-stu-id="ce20c-870">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="ce20c-871">Хотя минимальный уровень разрешений для использования этого метода — **Restricted**, для некоторых типов сущностей требуется доступ на уровне **ReadItem**, как указано в приведенной ниже таблице.</span><span class="sxs-lookup"><span data-stu-id="ce20c-871">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="ce20c-872">Значение параметра `entityType`</span><span class="sxs-lookup"><span data-stu-id="ce20c-872">Value of `entityType`</span></span> | <span data-ttu-id="ce20c-873">Тип объектов в возвращаемом массиве</span><span class="sxs-lookup"><span data-stu-id="ce20c-873">Type of objects in returned array</span></span> | <span data-ttu-id="ce20c-874">Необходимый уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ce20c-874">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="ce20c-875">String</span><span class="sxs-lookup"><span data-stu-id="ce20c-875">String</span></span> | <span data-ttu-id="ce20c-876">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="ce20c-876">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="ce20c-877">Contact</span><span class="sxs-lookup"><span data-stu-id="ce20c-877">Contact</span></span> | <span data-ttu-id="ce20c-878">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="ce20c-878">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="ce20c-879">String</span><span class="sxs-lookup"><span data-stu-id="ce20c-879">String</span></span> | <span data-ttu-id="ce20c-880">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="ce20c-880">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="ce20c-881">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="ce20c-881">MeetingSuggestion</span></span> | <span data-ttu-id="ce20c-882">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="ce20c-882">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="ce20c-883">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="ce20c-883">PhoneNumber</span></span> | <span data-ttu-id="ce20c-884">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="ce20c-884">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="ce20c-885">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="ce20c-885">TaskSuggestion</span></span> | <span data-ttu-id="ce20c-886">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="ce20c-886">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="ce20c-887">String</span><span class="sxs-lookup"><span data-stu-id="ce20c-887">String</span></span> | <span data-ttu-id="ce20c-888">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="ce20c-888">**Restricted**</span></span> |

<span data-ttu-id="ce20c-889">Тип: Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="ce20c-889">Type: Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="ce20c-890">Пример</span><span class="sxs-lookup"><span data-stu-id="ce20c-890">Example</span></span>

<span data-ttu-id="ce20c-891">В примере ниже показано, как получить доступ к массиву строк, которые представляют собой почтовые адреса в теле текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="ce20c-891">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook15officecontactmeetingsuggestionjavascriptapioutlook15officemeetingsuggestionphonenumberjavascriptapioutlook15officephonenumbertasksuggestionjavascriptapioutlook15officetasksuggestion"></a><span data-ttu-id="ce20c-892">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="ce20c-892">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))>}</span></span>

<span data-ttu-id="ce20c-893">Возвращает известные сущности в выбранном элементе, которые проходят через именованный фильтр, определяемый в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="ce20c-893">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="ce20c-894">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="ce20c-894">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="ce20c-895">Метод `getFilteredEntitiesByName` возвращает сущности, соответствующие регулярному выражению, которое определяется в элементе правила [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) в XML-файле манифеста, с использованием указанного значения элемента `FilterName`.</span><span class="sxs-lookup"><span data-stu-id="ce20c-895">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ce20c-896">Параметры</span><span class="sxs-lookup"><span data-stu-id="ce20c-896">Parameters</span></span>

|<span data-ttu-id="ce20c-897">Имя</span><span class="sxs-lookup"><span data-stu-id="ce20c-897">Name</span></span>| <span data-ttu-id="ce20c-898">Тип</span><span class="sxs-lookup"><span data-stu-id="ce20c-898">Type</span></span>| <span data-ttu-id="ce20c-899">Описание</span><span class="sxs-lookup"><span data-stu-id="ce20c-899">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="ce20c-900">String</span><span class="sxs-lookup"><span data-stu-id="ce20c-900">String</span></span>|<span data-ttu-id="ce20c-901">Имя элемента правила `ItemHasKnownEntity`, определяющее соответствующий фильтр.</span><span class="sxs-lookup"><span data-stu-id="ce20c-901">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ce20c-902">Требования</span><span class="sxs-lookup"><span data-stu-id="ce20c-902">Requirements</span></span>

|<span data-ttu-id="ce20c-903">Требование</span><span class="sxs-lookup"><span data-stu-id="ce20c-903">Requirement</span></span>| <span data-ttu-id="ce20c-904">Значение</span><span class="sxs-lookup"><span data-stu-id="ce20c-904">Value</span></span>|
|---|---|
|[<span data-ttu-id="ce20c-905">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="ce20c-905">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ce20c-906">1.0</span><span class="sxs-lookup"><span data-stu-id="ce20c-906">1.0</span></span>|
|[<span data-ttu-id="ce20c-907">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ce20c-907">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ce20c-908">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ce20c-908">ReadItem</span></span>|
|[<span data-ttu-id="ce20c-909">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ce20c-909">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ce20c-910">Чтение</span><span class="sxs-lookup"><span data-stu-id="ce20c-910">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="ce20c-911">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="ce20c-911">Returns:</span></span>

<span data-ttu-id="ce20c-p155">Если в манифесте нет элемента `ItemHasKnownEntity` со значением `FilterName`, соответствующим параметру `name`, метод возвращает `null`. Если параметр `name` соответствует элементу `ItemHasKnownEntity` в манифесте, но при этом в текущем элементе нет соответствующих сущностей, метод возвращает пустой массив.</span><span class="sxs-lookup"><span data-stu-id="ce20c-p155">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="ce20c-914">Тип: Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="ce20c-914">Type: Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))></span></span>

#### <a name="getregexmatches--object"></a><span data-ttu-id="ce20c-915">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="ce20c-915">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="ce20c-916">Возвращает строковые значения в выбранном элементе, которые соответствуют регулярным выражениям, определенным в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="ce20c-916">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="ce20c-917">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="ce20c-917">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="ce20c-p156">Метод `getRegExMatches` возвращает строки, соответствующие регулярному выражению, которое определяется в каждом элементе правила `ItemHasRegularExpressionMatch` или `ItemHasKnownEntity` в XML-файле манифеста. Для правила `ItemHasRegularExpressionMatch` соответствующую строку должно содержать свойство элемента, указанного этим правилом. Простой тип `PropertyName` определяет поддерживаемые свойства.</span><span class="sxs-lookup"><span data-stu-id="ce20c-p156">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="ce20c-921">Например, рассмотрим манифест надстройки, который содержит указанный ниже элемент `Rule`.</span><span class="sxs-lookup"><span data-stu-id="ce20c-921">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="ce20c-922">Объект, возвращаемый методом `getRegExMatches`, будет содержать два свойства: `fruits` и `veggies`.</span><span class="sxs-lookup"><span data-stu-id="ce20c-922">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="ce20c-p157">Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты. Лучше используйте метод [`Body.getAsync`](/javascript/api/outlook_1_5/office.body#getasync-coerciontype--options--callback-) для этого.</span><span class="sxs-lookup"><span data-stu-id="ce20c-p157">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook_1_5/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="ce20c-926">Requirements</span><span class="sxs-lookup"><span data-stu-id="ce20c-926">Requirements</span></span>

|<span data-ttu-id="ce20c-927">Требование</span><span class="sxs-lookup"><span data-stu-id="ce20c-927">Requirement</span></span>| <span data-ttu-id="ce20c-928">Значение</span><span class="sxs-lookup"><span data-stu-id="ce20c-928">Value</span></span>|
|---|---|
|[<span data-ttu-id="ce20c-929">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="ce20c-929">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ce20c-930">1.0</span><span class="sxs-lookup"><span data-stu-id="ce20c-930">1.0</span></span>|
|[<span data-ttu-id="ce20c-931">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ce20c-931">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ce20c-932">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ce20c-932">ReadItem</span></span>|
|[<span data-ttu-id="ce20c-933">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ce20c-933">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ce20c-934">Чтение</span><span class="sxs-lookup"><span data-stu-id="ce20c-934">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="ce20c-935">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="ce20c-935">Returns:</span></span>

<span data-ttu-id="ce20c-p158">Объект, содержащий массив строк, которые соответствуют регулярным выражениям, определяемым в XML-файле манифеста. Имя каждого массива равно соответствующему значению атрибута `RegExName` подходящего правила `ItemHasRegularExpressionMatch` или атрибута `FilterName` соответствующего правила `ItemHasKnownEntity`.</span><span class="sxs-lookup"><span data-stu-id="ce20c-p158">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="ce20c-938">

<dt>Тип</dt>

</span><span class="sxs-lookup"><span data-stu-id="ce20c-938">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="ce20c-939">Object</span><span class="sxs-lookup"><span data-stu-id="ce20c-939">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="ce20c-940">Пример</span><span class="sxs-lookup"><span data-stu-id="ce20c-940">Example</span></span>

<span data-ttu-id="ce20c-941">В примере ниже показано, как получить доступ к массиву совпадений для <rule>элементов регулярного выражения `fruits` и `veggies`, которые указаны в манифесте</rule>.</span><span class="sxs-lookup"><span data-stu-id="ce20c-941">The following example shows how to access the array of matches for the regular expression <rule>elements `fruits` and `veggies`, which are specified in the manifest.</rule></span></span>

```javascript
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="ce20c-942">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="ce20c-942">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="ce20c-943">Возвращает строковые значения в выбранном элементе, которые соответствуют именованному регулярному выражению, определенному в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="ce20c-943">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="ce20c-944">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="ce20c-944">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="ce20c-945">Метод `getRegExMatchesByName` возвращает строки, соответствующие регулярному выражению, которое определяется в элементе правила `ItemHasRegularExpressionMatch` в XML-файле манифеста, с использованием указанного значения элемента `RegExName`.</span><span class="sxs-lookup"><span data-stu-id="ce20c-945">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="ce20c-p159">Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты.</span><span class="sxs-lookup"><span data-stu-id="ce20c-p159">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ce20c-948">Параметры</span><span class="sxs-lookup"><span data-stu-id="ce20c-948">Parameters</span></span>

|<span data-ttu-id="ce20c-949">Имя</span><span class="sxs-lookup"><span data-stu-id="ce20c-949">Name</span></span>| <span data-ttu-id="ce20c-950">Тип</span><span class="sxs-lookup"><span data-stu-id="ce20c-950">Type</span></span>| <span data-ttu-id="ce20c-951">Описание</span><span class="sxs-lookup"><span data-stu-id="ce20c-951">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="ce20c-952">String</span><span class="sxs-lookup"><span data-stu-id="ce20c-952">String</span></span>|<span data-ttu-id="ce20c-953">Имя элемента правила `ItemHasRegularExpressionMatch`, определяющее соответствующий фильтр.</span><span class="sxs-lookup"><span data-stu-id="ce20c-953">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ce20c-954">Требования</span><span class="sxs-lookup"><span data-stu-id="ce20c-954">Requirements</span></span>

|<span data-ttu-id="ce20c-955">Требование</span><span class="sxs-lookup"><span data-stu-id="ce20c-955">Requirement</span></span>| <span data-ttu-id="ce20c-956">Значение</span><span class="sxs-lookup"><span data-stu-id="ce20c-956">Value</span></span>|
|---|---|
|[<span data-ttu-id="ce20c-957">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="ce20c-957">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ce20c-958">1.0</span><span class="sxs-lookup"><span data-stu-id="ce20c-958">1.0</span></span>|
|[<span data-ttu-id="ce20c-959">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ce20c-959">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ce20c-960">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ce20c-960">ReadItem</span></span>|
|[<span data-ttu-id="ce20c-961">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ce20c-961">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ce20c-962">Чтение</span><span class="sxs-lookup"><span data-stu-id="ce20c-962">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="ce20c-963">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="ce20c-963">Returns:</span></span>

<span data-ttu-id="ce20c-964">Массив строк, соответствующих регулярному выражению, определяемому в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="ce20c-964">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="ce20c-965">

<dt>Тип</dt>

</span><span class="sxs-lookup"><span data-stu-id="ce20c-965">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="ce20c-966">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="ce20c-966">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="ce20c-967">Пример</span><span class="sxs-lookup"><span data-stu-id="ce20c-967">Example</span></span>

```javascript
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

#### <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="ce20c-968">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="ce20c-968">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="ce20c-969">Асинхронно возвращает данные, выбранные в теме или тексте сообщения.</span><span class="sxs-lookup"><span data-stu-id="ce20c-969">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="ce20c-p160">Если выделенный фрагмент отсутствует, но курсор находится в тексте или теме, метод возвращает значение NULL для выбранных данных. Если выбраны не текст и не тема, метод возвращает ошибку `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="ce20c-p160">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ce20c-972">Параметры</span><span class="sxs-lookup"><span data-stu-id="ce20c-972">Parameters</span></span>

|<span data-ttu-id="ce20c-973">Имя</span><span class="sxs-lookup"><span data-stu-id="ce20c-973">Name</span></span>| <span data-ttu-id="ce20c-974">Тип</span><span class="sxs-lookup"><span data-stu-id="ce20c-974">Type</span></span>| <span data-ttu-id="ce20c-975">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="ce20c-975">Attributes</span></span>| <span data-ttu-id="ce20c-976">Описание</span><span class="sxs-lookup"><span data-stu-id="ce20c-976">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="ce20c-977">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="ce20c-977">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="ce20c-p161">Запрашивает формат данных. Если задано значение Text, метод возвращает обычный текст как строку, удаляя все имеющиеся HTML-теги. Если задано значение HTML, метод возвращает выделенный текст (обычный текст или HTML).</span><span class="sxs-lookup"><span data-stu-id="ce20c-p161">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="ce20c-981">Object</span><span class="sxs-lookup"><span data-stu-id="ce20c-981">Object</span></span>| <span data-ttu-id="ce20c-982">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="ce20c-982">&lt;optional&gt;</span></span>|<span data-ttu-id="ce20c-983">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="ce20c-983">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="ce20c-984">Object</span><span class="sxs-lookup"><span data-stu-id="ce20c-984">Object</span></span>| <span data-ttu-id="ce20c-985">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="ce20c-985">&lt;optional&gt;</span></span>|<span data-ttu-id="ce20c-986">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="ce20c-986">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="ce20c-987">функция</span><span class="sxs-lookup"><span data-stu-id="ce20c-987">function</span></span>||<span data-ttu-id="ce20c-988">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="ce20c-988">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="ce20c-989">Чтобы получить доступ к выбранным данным из метода обратного вызова, вызовите `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="ce20c-989">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="ce20c-990">Для доступа к исходному свойству, представляющему собой источник выбранных данных, вызовите параметр `asyncResult.value.sourceProperty`, который может иметь значение `body` или `subject`.</span><span class="sxs-lookup"><span data-stu-id="ce20c-990">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ce20c-991">Требования</span><span class="sxs-lookup"><span data-stu-id="ce20c-991">Requirements</span></span>

|<span data-ttu-id="ce20c-992">Требование</span><span class="sxs-lookup"><span data-stu-id="ce20c-992">Requirement</span></span>| <span data-ttu-id="ce20c-993">Значение</span><span class="sxs-lookup"><span data-stu-id="ce20c-993">Value</span></span>|
|---|---|
|[<span data-ttu-id="ce20c-994">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="ce20c-994">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ce20c-995">1.2</span><span class="sxs-lookup"><span data-stu-id="ce20c-995">1.2</span></span>|
|[<span data-ttu-id="ce20c-996">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ce20c-996">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ce20c-997">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="ce20c-997">ReadWriteItem</span></span>|
|[<span data-ttu-id="ce20c-998">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ce20c-998">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ce20c-999">Создание</span><span class="sxs-lookup"><span data-stu-id="ce20c-999">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="ce20c-1000">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="ce20c-1000">Returns:</span></span>

<span data-ttu-id="ce20c-1001">Выбранные данные в виде строки с форматом, определенным в параметре `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="ce20c-1001">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="ce20c-1002">

<dt>Тип</dt>

</span><span class="sxs-lookup"><span data-stu-id="ce20c-1002">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="ce20c-1003">String</span><span class="sxs-lookup"><span data-stu-id="ce20c-1003">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="ce20c-1004">Пример</span><span class="sxs-lookup"><span data-stu-id="ce20c-1004">Example</span></span>

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

#### <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="ce20c-1005">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="ce20c-1005">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="ce20c-1006">Асинхронно загружает настраиваемые свойства для надстройки для выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="ce20c-1006">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="ce20c-p163">Настраиваемые свойства сохраняются в виде пар "ключ-значение" для каждого приложения и каждого элемента. Этот метод возвращает объект `CustomProperties` при обратном вызове, который предоставляет методы для доступа к настраиваемым свойствам, характерным для текущего элемента и текущей надстройки. Настраиваемые свойства не шифруются для элемента, поэтому этот способ хранения не является безопасным.</span><span class="sxs-lookup"><span data-stu-id="ce20c-p163">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ce20c-1010">Параметры</span><span class="sxs-lookup"><span data-stu-id="ce20c-1010">Parameters</span></span>

|<span data-ttu-id="ce20c-1011">Имя</span><span class="sxs-lookup"><span data-stu-id="ce20c-1011">Name</span></span>| <span data-ttu-id="ce20c-1012">Тип</span><span class="sxs-lookup"><span data-stu-id="ce20c-1012">Type</span></span>| <span data-ttu-id="ce20c-1013">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="ce20c-1013">Attributes</span></span>| <span data-ttu-id="ce20c-1014">Описание</span><span class="sxs-lookup"><span data-stu-id="ce20c-1014">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="ce20c-1015">function</span><span class="sxs-lookup"><span data-stu-id="ce20c-1015">function</span></span>||<span data-ttu-id="ce20c-1016">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="ce20c-1016">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="ce20c-1017">Настраиваемые свойства предоставляются в виде объекта [`CustomProperties`](/javascript/api/outlook_1_5/office.customproperties) в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="ce20c-1017">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook_1_5/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="ce20c-1018">Этот объект позволяет получить, задать и удалить настраиваемые свойства для элемента, а также сохранить изменения, внесенные в набор настраиваемых свойств, на сервере.</span><span class="sxs-lookup"><span data-stu-id="ce20c-1018">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="ce20c-1019">Объект</span><span class="sxs-lookup"><span data-stu-id="ce20c-1019">Object</span></span>| <span data-ttu-id="ce20c-1020">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="ce20c-1020">&lt;optional&gt;</span></span>|<span data-ttu-id="ce20c-1021">Разработчики могут указать любой объект, к которому необходимо получить доступ, в функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="ce20c-1021">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="ce20c-1022">Доступ к этому объекту можно получить с помощью свойства `asyncResult.asyncContext` в функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="ce20c-1022">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ce20c-1023">Требования</span><span class="sxs-lookup"><span data-stu-id="ce20c-1023">Requirements</span></span>

|<span data-ttu-id="ce20c-1024">Требование</span><span class="sxs-lookup"><span data-stu-id="ce20c-1024">Requirement</span></span>| <span data-ttu-id="ce20c-1025">Значение</span><span class="sxs-lookup"><span data-stu-id="ce20c-1025">Value</span></span>|
|---|---|
|[<span data-ttu-id="ce20c-1026">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="ce20c-1026">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ce20c-1027">1.0</span><span class="sxs-lookup"><span data-stu-id="ce20c-1027">1.0</span></span>|
|[<span data-ttu-id="ce20c-1028">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ce20c-1028">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ce20c-1029">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ce20c-1029">ReadItem</span></span>|
|[<span data-ttu-id="ce20c-1030">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ce20c-1030">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ce20c-1031">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="ce20c-1031">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ce20c-1032">Пример</span><span class="sxs-lookup"><span data-stu-id="ce20c-1032">Example</span></span>

<span data-ttu-id="ce20c-p166">Приведенный ниже пример кода показывает, как асинхронно загружать настраиваемые свойства, характерные для текущего элемента, с помощью метода `loadCustomPropertiesAsync`. Этот пример также показывает, как сохранять эти свойства на сервере с помощью метода `CustomProperties.saveAsync`. После загрузки настраиваемых свойств в этом примере кода метод `CustomProperties.get` используется для считывания настраиваемого свойства `myProp`, метод `CustomProperties.set` — для записи настраиваемого свойства `otherProp`, а метод `saveAsync` — для сохранения настраиваемых свойств.</span><span class="sxs-lookup"><span data-stu-id="ce20c-p166">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

#### <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="ce20c-1036">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="ce20c-1036">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="ce20c-1037">Удаляет вложение из сообщения или встречи.</span><span class="sxs-lookup"><span data-stu-id="ce20c-1037">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="ce20c-p167">Метод `removeAttachmentAsync` удаляет из элемента вложение с указанным идентификатором. Идентификатор вложения рекомендуется использовать для удаления вложения, только если оно добавлено тем же почтовым приложением в ходе текущего сеанса. В Outlook в Интернете и на мобильных устройствах идентификатор вложения действителен только в рамках одного сеанса. Сеанс завершается, когда пользователь закрывает приложение или начинает создавать элемент во встроенной форме, а затем переходит из формы в отдельное окно.</span><span class="sxs-lookup"><span data-stu-id="ce20c-p167">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ce20c-1042">Параметры</span><span class="sxs-lookup"><span data-stu-id="ce20c-1042">Parameters</span></span>

|<span data-ttu-id="ce20c-1043">Имя</span><span class="sxs-lookup"><span data-stu-id="ce20c-1043">Name</span></span>| <span data-ttu-id="ce20c-1044">Тип</span><span class="sxs-lookup"><span data-stu-id="ce20c-1044">Type</span></span>| <span data-ttu-id="ce20c-1045">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="ce20c-1045">Attributes</span></span>| <span data-ttu-id="ce20c-1046">Описание</span><span class="sxs-lookup"><span data-stu-id="ce20c-1046">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="ce20c-1047">String</span><span class="sxs-lookup"><span data-stu-id="ce20c-1047">String</span></span>||<span data-ttu-id="ce20c-1048">Идентификатор удаляемого вложения.</span><span class="sxs-lookup"><span data-stu-id="ce20c-1048">The identifier of the attachment to remove.</span></span>|
|`options`| <span data-ttu-id="ce20c-1049">Object</span><span class="sxs-lookup"><span data-stu-id="ce20c-1049">Object</span></span>| <span data-ttu-id="ce20c-1050">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="ce20c-1050">&lt;optional&gt;</span></span>|<span data-ttu-id="ce20c-1051">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="ce20c-1051">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="ce20c-1052">Object</span><span class="sxs-lookup"><span data-stu-id="ce20c-1052">Object</span></span>| <span data-ttu-id="ce20c-1053">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="ce20c-1053">&lt;optional&gt;</span></span>|<span data-ttu-id="ce20c-1054">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="ce20c-1054">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="ce20c-1055">функция</span><span class="sxs-lookup"><span data-stu-id="ce20c-1055">function</span></span>| <span data-ttu-id="ce20c-1056">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="ce20c-1056">&lt;optional&gt;</span></span>|<span data-ttu-id="ce20c-1057">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="ce20c-1057">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="ce20c-1058">Если удалить вложение не удается, свойство `asyncResult.error` содержит код ошибки с указанием ее причины.</span><span class="sxs-lookup"><span data-stu-id="ce20c-1058">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="ce20c-1059">Ошибки</span><span class="sxs-lookup"><span data-stu-id="ce20c-1059">Errors</span></span>

| <span data-ttu-id="ce20c-1060">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="ce20c-1060">Error code</span></span> | <span data-ttu-id="ce20c-1061">Описание</span><span class="sxs-lookup"><span data-stu-id="ce20c-1061">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="ce20c-1062">Идентификатор вложения не существует.</span><span class="sxs-lookup"><span data-stu-id="ce20c-1062">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="ce20c-1063">Требования</span><span class="sxs-lookup"><span data-stu-id="ce20c-1063">Requirements</span></span>

|<span data-ttu-id="ce20c-1064">Требование</span><span class="sxs-lookup"><span data-stu-id="ce20c-1064">Requirement</span></span>| <span data-ttu-id="ce20c-1065">Значение</span><span class="sxs-lookup"><span data-stu-id="ce20c-1065">Value</span></span>|
|---|---|
|[<span data-ttu-id="ce20c-1066">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="ce20c-1066">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ce20c-1067">1.1</span><span class="sxs-lookup"><span data-stu-id="ce20c-1067">1.1</span></span>|
|[<span data-ttu-id="ce20c-1068">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ce20c-1068">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ce20c-1069">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="ce20c-1069">ReadWriteItem</span></span>|
|[<span data-ttu-id="ce20c-1070">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ce20c-1070">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ce20c-1071">Создание</span><span class="sxs-lookup"><span data-stu-id="ce20c-1071">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="ce20c-1072">Пример</span><span class="sxs-lookup"><span data-stu-id="ce20c-1072">Example</span></span>

<span data-ttu-id="ce20c-1073">Указанный ниже код удаляет вложение с идентификатором "0".</span><span class="sxs-lookup"><span data-stu-id="ce20c-1073">The following code removes an attachment with an identifier of '0'.</span></span>

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

#### <a name="saveasyncoptions-callback"></a><span data-ttu-id="ce20c-1074">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="ce20c-1074">saveAsync([options], callback)</span></span>

<span data-ttu-id="ce20c-1075">Асинхронно сохраняет элемент.</span><span class="sxs-lookup"><span data-stu-id="ce20c-1075">Asynchronously saves an item.</span></span>

<span data-ttu-id="ce20c-p168">При вызове этот метод сохраняет текущее сообщение в виде черновика и возвращает идентификатор элемента с помощью метода обратного вызова. В Outlook в Интернете или интерактивном режиме Outlook этот элемент сохраняется на сервере. В Outlook в режиме кэширования этот элемент сохраняется в локальном кэше.</span><span class="sxs-lookup"><span data-stu-id="ce20c-p168">When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="ce20c-1079">Если в режиме создания надстройка вызывает для элемента метод `saveAsync`, чтобы получить параметр `itemId` для использования с EWS или REST API, необходимо помнить, что синхронизация элемента с сервером может занять много времени, если Outlook работает в режиме кэширования данных.</span><span class="sxs-lookup"><span data-stu-id="ce20c-1079">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="ce20c-1080">До окончания синхронизации элемента применение параметра `itemId` будет приводить к ошибке.</span><span class="sxs-lookup"><span data-stu-id="ce20c-1080">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="ce20c-p170">Если метод `saveAsync` вызывается для встречи в режиме создания, она сохраняется как обычная встреча в календаре пользователя, а не как черновик. При сохранении новой встречи приглашения не отправляются. При сохранении существующей встречи уведомления отправляются добавленным или удаленным участникам.</span><span class="sxs-lookup"><span data-stu-id="ce20c-p170">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="ce20c-1084">Следующие клиенты отличаются другим поведением `saveAsync` в отношении встреч в режиме создания:</span><span class="sxs-lookup"><span data-stu-id="ce20c-1084">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="ce20c-1085">Outlook для Mac не поддерживает сохранение собрания.</span><span class="sxs-lookup"><span data-stu-id="ce20c-1085">Outlook on Mac does not support saving a meeting.</span></span> <span data-ttu-id="ce20c-1086">Метод `saveAsync` не работает при вызове из собрания в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="ce20c-1086">The `saveAsync` method fails when called from a meeting in compose mode.</span></span> <span data-ttu-id="ce20c-1087">Временное решение представлено в статье [Не удается сохранить встречу как черновик в Outlook для Mac с помощью API JS для Office](https://support.microsoft.com/help/4505745).</span><span class="sxs-lookup"><span data-stu-id="ce20c-1087">See [Cannot save a meeting as a draft in Outlook for Mac by using Office JS API](https://support.microsoft.com/help/4505745) for a workaround.</span></span>
> - <span data-ttu-id="ce20c-1088">Outlook в Интернете всегда отправляет приглашение или обновление при вызове `saveAsync` для встречи в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="ce20c-1088">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ce20c-1089">Параметры</span><span class="sxs-lookup"><span data-stu-id="ce20c-1089">Parameters</span></span>

|<span data-ttu-id="ce20c-1090">Имя</span><span class="sxs-lookup"><span data-stu-id="ce20c-1090">Name</span></span>| <span data-ttu-id="ce20c-1091">Тип</span><span class="sxs-lookup"><span data-stu-id="ce20c-1091">Type</span></span>| <span data-ttu-id="ce20c-1092">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="ce20c-1092">Attributes</span></span>| <span data-ttu-id="ce20c-1093">Описание</span><span class="sxs-lookup"><span data-stu-id="ce20c-1093">Description</span></span>|
|---|---|---|---|
|`options`| <span data-ttu-id="ce20c-1094">Object</span><span class="sxs-lookup"><span data-stu-id="ce20c-1094">Object</span></span>| <span data-ttu-id="ce20c-1095">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="ce20c-1095">&lt;optional&gt;</span></span>|<span data-ttu-id="ce20c-1096">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="ce20c-1096">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="ce20c-1097">Object</span><span class="sxs-lookup"><span data-stu-id="ce20c-1097">Object</span></span>| <span data-ttu-id="ce20c-1098">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="ce20c-1098">&lt;optional&gt;</span></span>|<span data-ttu-id="ce20c-1099">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="ce20c-1099">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="ce20c-1100">функция</span><span class="sxs-lookup"><span data-stu-id="ce20c-1100">function</span></span>||<span data-ttu-id="ce20c-1101">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="ce20c-1101">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="ce20c-1102">После успешного выполнения идентификатор элемента будет предоставлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="ce20c-1102">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ce20c-1103">Требования</span><span class="sxs-lookup"><span data-stu-id="ce20c-1103">Requirements</span></span>

|<span data-ttu-id="ce20c-1104">Требование</span><span class="sxs-lookup"><span data-stu-id="ce20c-1104">Requirement</span></span>| <span data-ttu-id="ce20c-1105">Значение</span><span class="sxs-lookup"><span data-stu-id="ce20c-1105">Value</span></span>|
|---|---|
|[<span data-ttu-id="ce20c-1106">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="ce20c-1106">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ce20c-1107">1.3</span><span class="sxs-lookup"><span data-stu-id="ce20c-1107">1.3</span></span>|
|[<span data-ttu-id="ce20c-1108">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ce20c-1108">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ce20c-1109">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="ce20c-1109">ReadWriteItem</span></span>|
|[<span data-ttu-id="ce20c-1110">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ce20c-1110">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ce20c-1111">Создание</span><span class="sxs-lookup"><span data-stu-id="ce20c-1111">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="ce20c-1112">Примеры</span><span class="sxs-lookup"><span data-stu-id="ce20c-1112">Examples</span></span>

```javascript
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result.
  });
```

<span data-ttu-id="ce20c-p172">Ниже приведен пример параметра `result`, переданного функции обратного вызова. Свойство `value` содержит идентификатор элемента.</span><span class="sxs-lookup"><span data-stu-id="ce20c-p172">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

#### <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="ce20c-1115">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="ce20c-1115">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="ce20c-1116">Асинхронно вставляет данные в текст или тему сообщения.</span><span class="sxs-lookup"><span data-stu-id="ce20c-1116">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="ce20c-p173">Метод `setSelectedDataAsync` вставляет указанную строку в местоположение курсора в теме или тексте элемента либо, если текст выделен в редакторе, он заменяет выделенный текст. Если курсор находится вне текста или темы элемента, возвращается ошибка. После вставки курсор помещается в конец вставленного содержимого.</span><span class="sxs-lookup"><span data-stu-id="ce20c-p173">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ce20c-1120">Параметры</span><span class="sxs-lookup"><span data-stu-id="ce20c-1120">Parameters</span></span>

|<span data-ttu-id="ce20c-1121">Имя</span><span class="sxs-lookup"><span data-stu-id="ce20c-1121">Name</span></span>| <span data-ttu-id="ce20c-1122">Тип</span><span class="sxs-lookup"><span data-stu-id="ce20c-1122">Type</span></span>| <span data-ttu-id="ce20c-1123">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="ce20c-1123">Attributes</span></span>| <span data-ttu-id="ce20c-1124">Описание</span><span class="sxs-lookup"><span data-stu-id="ce20c-1124">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="ce20c-1125">String</span><span class="sxs-lookup"><span data-stu-id="ce20c-1125">String</span></span>||<span data-ttu-id="ce20c-p174">Вставляемые данные. Объем данных не должен превышать 1 000 000 символов. Если передано больше 1 000 000 символов, возвращается исключение `ArgumentOutOfRange`.</span><span class="sxs-lookup"><span data-stu-id="ce20c-p174">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="ce20c-1129">Object</span><span class="sxs-lookup"><span data-stu-id="ce20c-1129">Object</span></span>| <span data-ttu-id="ce20c-1130">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="ce20c-1130">&lt;optional&gt;</span></span>|<span data-ttu-id="ce20c-1131">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="ce20c-1131">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="ce20c-1132">Object</span><span class="sxs-lookup"><span data-stu-id="ce20c-1132">Object</span></span>| <span data-ttu-id="ce20c-1133">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="ce20c-1133">&lt;optional&gt;</span></span>|<span data-ttu-id="ce20c-1134">В методе обратного вызова разработчики могут указать любой объект, к которому необходимо получить доступ.</span><span class="sxs-lookup"><span data-stu-id="ce20c-1134">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="ce20c-1135">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="ce20c-1135">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="ce20c-1136">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="ce20c-1136">&lt;optional&gt;</span></span>|<span data-ttu-id="ce20c-1137">Если задано значение `text`, текущий стиль применяется в Outlook в Интернете и классических клиентах.</span><span class="sxs-lookup"><span data-stu-id="ce20c-1137">If `text`, the current style is applied in Outlook on the web and desktop clients.</span></span> <span data-ttu-id="ce20c-1138">Если поле представляет собой редактор HTML, вставляются только текстовые данные, даже если они имеют формат HTML.</span><span class="sxs-lookup"><span data-stu-id="ce20c-1138">If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="ce20c-1139">Если задано значение `html` и поле (не тема) поддерживает HTML, в Outlook в Интернете применяется текущий стиль, а в классических клиентах Outlook — стиль по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="ce20c-1139">If html and the field supports HTML (the subject doesn't), the current style is applied in olwebshort and the default style is applied in Outlook.</span></span> <span data-ttu-id="ce20c-1140">Если поле является текстовым, возвращается ошибка `InvalidDataFormat`.</span><span class="sxs-lookup"><span data-stu-id="ce20c-1140">If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="ce20c-1141">Если свойство `coercionType` не задано, результат зависит от поля: если поле имеет формат HTML, используется текст в формате HTML, а если поле текстовое, применяется обычный текст.</span><span class="sxs-lookup"><span data-stu-id="ce20c-1141">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="ce20c-1142">функция</span><span class="sxs-lookup"><span data-stu-id="ce20c-1142">function</span></span>||<span data-ttu-id="ce20c-1143">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="ce20c-1143">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="ce20c-1144">Требования</span><span class="sxs-lookup"><span data-stu-id="ce20c-1144">Requirements</span></span>

|<span data-ttu-id="ce20c-1145">Требование</span><span class="sxs-lookup"><span data-stu-id="ce20c-1145">Requirement</span></span>| <span data-ttu-id="ce20c-1146">Значение</span><span class="sxs-lookup"><span data-stu-id="ce20c-1146">Value</span></span>|
|---|---|
|[<span data-ttu-id="ce20c-1147">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="ce20c-1147">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ce20c-1148">1.2</span><span class="sxs-lookup"><span data-stu-id="ce20c-1148">1.2</span></span>|
|[<span data-ttu-id="ce20c-1149">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ce20c-1149">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ce20c-1150">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="ce20c-1150">ReadWriteItem</span></span>|
|[<span data-ttu-id="ce20c-1151">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ce20c-1151">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ce20c-1152">Создание</span><span class="sxs-lookup"><span data-stu-id="ce20c-1152">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="ce20c-1153">Пример</span><span class="sxs-lookup"><span data-stu-id="ce20c-1153">Example</span></span>

```javascript
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
