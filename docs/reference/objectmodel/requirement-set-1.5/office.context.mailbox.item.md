---
title: Office.context.mailbox.item — набор обязательных элементов 1.5
description: ''
ms.date: 11/25/2019
localization_priority: Priority
ms.openlocfilehash: f52f6bfa510d5949da5b1b542ca2a0f6e6f22750
ms.sourcegitcommit: 05a883a7fd89136301ce35aabc57638e9f563288
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 11/27/2019
ms.locfileid: "39629709"
---
# <a name="item"></a><span data-ttu-id="dd864-102">item</span><span class="sxs-lookup"><span data-stu-id="dd864-102">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="dd864-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="dd864-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="dd864-p101">Пространство имен `item` используется для доступа к выбранному в данный момент сообщению, приглашению на собрание или описанию встречи. Вы можете определить тип пространства имен `item` с помощью свойства [itemType](#itemtype-officemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="dd864-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="dd864-106">Requirements</span><span class="sxs-lookup"><span data-stu-id="dd864-106">Requirements</span></span>

|<span data-ttu-id="dd864-107">Требование</span><span class="sxs-lookup"><span data-stu-id="dd864-107">Requirement</span></span>| <span data-ttu-id="dd864-108">Значение</span><span class="sxs-lookup"><span data-stu-id="dd864-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="dd864-109">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="dd864-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dd864-110">1.0</span><span class="sxs-lookup"><span data-stu-id="dd864-110">1.0</span></span>|
|[<span data-ttu-id="dd864-111">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="dd864-111">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dd864-112">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="dd864-112">Restricted</span></span>|
|[<span data-ttu-id="dd864-113">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="dd864-113">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="dd864-114">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="dd864-114">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="dd864-115">Элементы и методы</span><span class="sxs-lookup"><span data-stu-id="dd864-115">Members and methods</span></span>

| <span data-ttu-id="dd864-116">Элемент</span><span class="sxs-lookup"><span data-stu-id="dd864-116">Member</span></span> | <span data-ttu-id="dd864-117">Тип</span><span class="sxs-lookup"><span data-stu-id="dd864-117">Type</span></span> |
|--------|------|
| [<span data-ttu-id="dd864-118">attachments</span><span class="sxs-lookup"><span data-stu-id="dd864-118">attachments</span></span>](#attachments-arrayattachmentdetails) | <span data-ttu-id="dd864-119">Элемент</span><span class="sxs-lookup"><span data-stu-id="dd864-119">Member</span></span> |
| [<span data-ttu-id="dd864-120">bcc</span><span class="sxs-lookup"><span data-stu-id="dd864-120">bcc</span></span>](#bcc-recipients) | <span data-ttu-id="dd864-121">Элемент</span><span class="sxs-lookup"><span data-stu-id="dd864-121">Member</span></span> |
| [<span data-ttu-id="dd864-122">body</span><span class="sxs-lookup"><span data-stu-id="dd864-122">body</span></span>](#body-body) | <span data-ttu-id="dd864-123">Элемент</span><span class="sxs-lookup"><span data-stu-id="dd864-123">Member</span></span> |
| [<span data-ttu-id="dd864-124">cc</span><span class="sxs-lookup"><span data-stu-id="dd864-124">cc</span></span>](#cc-arrayemailaddressdetailsrecipients) | <span data-ttu-id="dd864-125">Элемент</span><span class="sxs-lookup"><span data-stu-id="dd864-125">Member</span></span> |
| [<span data-ttu-id="dd864-126">conversationId</span><span class="sxs-lookup"><span data-stu-id="dd864-126">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="dd864-127">Элемент</span><span class="sxs-lookup"><span data-stu-id="dd864-127">Member</span></span> |
| [<span data-ttu-id="dd864-128">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="dd864-128">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="dd864-129">Элемент</span><span class="sxs-lookup"><span data-stu-id="dd864-129">Member</span></span> |
| [<span data-ttu-id="dd864-130">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="dd864-130">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="dd864-131">Элемент</span><span class="sxs-lookup"><span data-stu-id="dd864-131">Member</span></span> |
| [<span data-ttu-id="dd864-132">end</span><span class="sxs-lookup"><span data-stu-id="dd864-132">end</span></span>](#end-datetime) | <span data-ttu-id="dd864-133">Элемент</span><span class="sxs-lookup"><span data-stu-id="dd864-133">Member</span></span> |
| [<span data-ttu-id="dd864-134">from</span><span class="sxs-lookup"><span data-stu-id="dd864-134">from</span></span>](#from-emailaddressdetails) | <span data-ttu-id="dd864-135">Элемент</span><span class="sxs-lookup"><span data-stu-id="dd864-135">Member</span></span> |
| [<span data-ttu-id="dd864-136">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="dd864-136">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="dd864-137">Элемент</span><span class="sxs-lookup"><span data-stu-id="dd864-137">Member</span></span> |
| [<span data-ttu-id="dd864-138">itemClass</span><span class="sxs-lookup"><span data-stu-id="dd864-138">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="dd864-139">Элемент</span><span class="sxs-lookup"><span data-stu-id="dd864-139">Member</span></span> |
| [<span data-ttu-id="dd864-140">itemId</span><span class="sxs-lookup"><span data-stu-id="dd864-140">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="dd864-141">Элемент</span><span class="sxs-lookup"><span data-stu-id="dd864-141">Member</span></span> |
| [<span data-ttu-id="dd864-142">itemType</span><span class="sxs-lookup"><span data-stu-id="dd864-142">itemType</span></span>](#itemtype-officemailboxenumsitemtype) | <span data-ttu-id="dd864-143">Элемент</span><span class="sxs-lookup"><span data-stu-id="dd864-143">Member</span></span> |
| [<span data-ttu-id="dd864-144">location</span><span class="sxs-lookup"><span data-stu-id="dd864-144">location</span></span>](#location-stringlocation) | <span data-ttu-id="dd864-145">Элемент</span><span class="sxs-lookup"><span data-stu-id="dd864-145">Member</span></span> |
| [<span data-ttu-id="dd864-146">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="dd864-146">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="dd864-147">Элемент</span><span class="sxs-lookup"><span data-stu-id="dd864-147">Member</span></span> |
| [<span data-ttu-id="dd864-148">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="dd864-148">notificationMessages</span></span>](#notificationmessages-notificationmessages) | <span data-ttu-id="dd864-149">Элемент</span><span class="sxs-lookup"><span data-stu-id="dd864-149">Member</span></span> |
| [<span data-ttu-id="dd864-150">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="dd864-150">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="dd864-151">Элемент</span><span class="sxs-lookup"><span data-stu-id="dd864-151">Member</span></span> |
| [<span data-ttu-id="dd864-152">organizer</span><span class="sxs-lookup"><span data-stu-id="dd864-152">organizer</span></span>](#organizer-emailaddressdetails) | <span data-ttu-id="dd864-153">Элемент</span><span class="sxs-lookup"><span data-stu-id="dd864-153">Member</span></span> |
| [<span data-ttu-id="dd864-154">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="dd864-154">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="dd864-155">Member</span><span class="sxs-lookup"><span data-stu-id="dd864-155">Member</span></span> |
| [<span data-ttu-id="dd864-156">sender</span><span class="sxs-lookup"><span data-stu-id="dd864-156">sender</span></span>](#sender-emailaddressdetails) | <span data-ttu-id="dd864-157">Элемент</span><span class="sxs-lookup"><span data-stu-id="dd864-157">Member</span></span> |
| [<span data-ttu-id="dd864-158">start</span><span class="sxs-lookup"><span data-stu-id="dd864-158">start</span></span>](#start-datetime) | <span data-ttu-id="dd864-159">Элемент</span><span class="sxs-lookup"><span data-stu-id="dd864-159">Member</span></span> |
| [<span data-ttu-id="dd864-160">subject</span><span class="sxs-lookup"><span data-stu-id="dd864-160">subject</span></span>](#subject-stringsubject) | <span data-ttu-id="dd864-161">Элемент</span><span class="sxs-lookup"><span data-stu-id="dd864-161">Member</span></span> |
| [<span data-ttu-id="dd864-162">to</span><span class="sxs-lookup"><span data-stu-id="dd864-162">to</span></span>](#to-arrayemailaddressdetailsrecipients) | <span data-ttu-id="dd864-163">Элемент</span><span class="sxs-lookup"><span data-stu-id="dd864-163">Member</span></span> |
| [<span data-ttu-id="dd864-164">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="dd864-164">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="dd864-165">Метод</span><span class="sxs-lookup"><span data-stu-id="dd864-165">Method</span></span> |
| [<span data-ttu-id="dd864-166">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="dd864-166">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="dd864-167">Метод</span><span class="sxs-lookup"><span data-stu-id="dd864-167">Method</span></span> |
| [<span data-ttu-id="dd864-168">close</span><span class="sxs-lookup"><span data-stu-id="dd864-168">close</span></span>](#close) | <span data-ttu-id="dd864-169">Метод</span><span class="sxs-lookup"><span data-stu-id="dd864-169">Method</span></span> |
| [<span data-ttu-id="dd864-170">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="dd864-170">displayReplyAllForm</span></span>](#displayreplyallformformdata-callback) | <span data-ttu-id="dd864-171">Метод</span><span class="sxs-lookup"><span data-stu-id="dd864-171">Method</span></span> |
| [<span data-ttu-id="dd864-172">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="dd864-172">displayReplyForm</span></span>](#displayreplyformformdata-callback) | <span data-ttu-id="dd864-173">Метод</span><span class="sxs-lookup"><span data-stu-id="dd864-173">Method</span></span> |
| [<span data-ttu-id="dd864-174">getEntities</span><span class="sxs-lookup"><span data-stu-id="dd864-174">getEntities</span></span>](#getentities--entities) | <span data-ttu-id="dd864-175">Метод</span><span class="sxs-lookup"><span data-stu-id="dd864-175">Method</span></span> |
| [<span data-ttu-id="dd864-176">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="dd864-176">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="dd864-177">Метод</span><span class="sxs-lookup"><span data-stu-id="dd864-177">Method</span></span> |
| [<span data-ttu-id="dd864-178">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="dd864-178">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="dd864-179">Метод</span><span class="sxs-lookup"><span data-stu-id="dd864-179">Method</span></span> |
| [<span data-ttu-id="dd864-180">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="dd864-180">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="dd864-181">Метод</span><span class="sxs-lookup"><span data-stu-id="dd864-181">Method</span></span> |
| [<span data-ttu-id="dd864-182">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="dd864-182">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="dd864-183">Метод</span><span class="sxs-lookup"><span data-stu-id="dd864-183">Method</span></span> |
| [<span data-ttu-id="dd864-184">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="dd864-184">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="dd864-185">Метод</span><span class="sxs-lookup"><span data-stu-id="dd864-185">Method</span></span> |
| [<span data-ttu-id="dd864-186">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="dd864-186">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="dd864-187">Метод</span><span class="sxs-lookup"><span data-stu-id="dd864-187">Method</span></span> |
| [<span data-ttu-id="dd864-188">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="dd864-188">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="dd864-189">Метод</span><span class="sxs-lookup"><span data-stu-id="dd864-189">Method</span></span> |
| [<span data-ttu-id="dd864-190">saveAsync</span><span class="sxs-lookup"><span data-stu-id="dd864-190">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="dd864-191">Метод</span><span class="sxs-lookup"><span data-stu-id="dd864-191">Method</span></span> |
| [<span data-ttu-id="dd864-192">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="dd864-192">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="dd864-193">Метод</span><span class="sxs-lookup"><span data-stu-id="dd864-193">Method</span></span> |

### <a name="example"></a><span data-ttu-id="dd864-194">Пример</span><span class="sxs-lookup"><span data-stu-id="dd864-194">Example</span></span>

<span data-ttu-id="dd864-195">В примере кода JavaScript, приведенном ниже, показано, как получить доступ к свойству `subject` текущего элемента в Outlook.</span><span class="sxs-lookup"><span data-stu-id="dd864-195">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="dd864-196">Members</span><span class="sxs-lookup"><span data-stu-id="dd864-196">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetailsviewoutlook-js-15"></a><span data-ttu-id="dd864-197">attachments: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.5)></span><span class="sxs-lookup"><span data-stu-id="dd864-197">attachments: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.5)></span></span>

<span data-ttu-id="dd864-p102">Получает массив вложений для элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="dd864-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="dd864-200">Outlook блокирует определенные типы файлов из-за потенциальных проблем с безопасностью, поэтому они не возвращаются.</span><span class="sxs-lookup"><span data-stu-id="dd864-200">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="dd864-201">Дополнительные сведения см. в статье [Блокировка вложений в Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="dd864-201">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="dd864-202">Тип</span><span class="sxs-lookup"><span data-stu-id="dd864-202">Type</span></span>

*   <span data-ttu-id="dd864-203">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.5)></span><span class="sxs-lookup"><span data-stu-id="dd864-203">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.5)></span></span>

##### <a name="requirements"></a><span data-ttu-id="dd864-204">Требования</span><span class="sxs-lookup"><span data-stu-id="dd864-204">Requirements</span></span>

|<span data-ttu-id="dd864-205">Требование</span><span class="sxs-lookup"><span data-stu-id="dd864-205">Requirement</span></span>| <span data-ttu-id="dd864-206">Значение</span><span class="sxs-lookup"><span data-stu-id="dd864-206">Value</span></span>|
|---|---|
|[<span data-ttu-id="dd864-207">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="dd864-207">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dd864-208">1.0</span><span class="sxs-lookup"><span data-stu-id="dd864-208">1.0</span></span>|
|[<span data-ttu-id="dd864-209">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="dd864-209">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dd864-210">ReadItem</span><span class="sxs-lookup"><span data-stu-id="dd864-210">ReadItem</span></span>|
|[<span data-ttu-id="dd864-211">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="dd864-211">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="dd864-212">Чтение</span><span class="sxs-lookup"><span data-stu-id="dd864-212">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="dd864-213">Пример</span><span class="sxs-lookup"><span data-stu-id="dd864-213">Example</span></span>

<span data-ttu-id="dd864-214">С помощью приведенного ниже кода можно создать HTML-строку с подробными сведениями обо всех вложениях для текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="dd864-214">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

#### <a name="bcc-recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-15"></a><span data-ttu-id="dd864-215">bcc: [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="dd864-215">bcc: [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.5)</span></span>

<span data-ttu-id="dd864-216">Получает объект, который предоставляет методы для получения или обновления получателей скрытой копии сообщения.</span><span class="sxs-lookup"><span data-stu-id="dd864-216">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="dd864-217">Только в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="dd864-217">Compose mode only.</span></span>

<span data-ttu-id="dd864-218">Коллекция может включать не более 100 элементов по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="dd864-218">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="dd864-219">Однако для компьютеров с Windows и Mac OS действуют указанные ниже ограничения.</span><span class="sxs-lookup"><span data-stu-id="dd864-219">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="dd864-220">Максимальное количество элементов — 500.</span><span class="sxs-lookup"><span data-stu-id="dd864-220">Get 500 members maximum.</span></span>
- <span data-ttu-id="dd864-221">Установить ограничение количества элементов на вызов — не более 100, общего количества — не более 500.</span><span class="sxs-lookup"><span data-stu-id="dd864-221">Set a maximum of 100 members per call, up to 500 members total.</span></span>

##### <a name="type"></a><span data-ttu-id="dd864-222">Тип</span><span class="sxs-lookup"><span data-stu-id="dd864-222">Type</span></span>

*   [<span data-ttu-id="dd864-223">Получатели</span><span class="sxs-lookup"><span data-stu-id="dd864-223">Recipients</span></span>](/javascript/api/outlook/office.recipients?view=outlook-js-1.5)

##### <a name="requirements"></a><span data-ttu-id="dd864-224">Требования</span><span class="sxs-lookup"><span data-stu-id="dd864-224">Requirements</span></span>

|<span data-ttu-id="dd864-225">Требование</span><span class="sxs-lookup"><span data-stu-id="dd864-225">Requirement</span></span>| <span data-ttu-id="dd864-226">Значение</span><span class="sxs-lookup"><span data-stu-id="dd864-226">Value</span></span>|
|---|---|
|[<span data-ttu-id="dd864-227">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="dd864-227">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dd864-228">1.1</span><span class="sxs-lookup"><span data-stu-id="dd864-228">1.1</span></span>|
|[<span data-ttu-id="dd864-229">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="dd864-229">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dd864-230">ReadItem</span><span class="sxs-lookup"><span data-stu-id="dd864-230">ReadItem</span></span>|
|[<span data-ttu-id="dd864-231">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="dd864-231">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="dd864-232">Создание</span><span class="sxs-lookup"><span data-stu-id="dd864-232">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="dd864-233">Пример</span><span class="sxs-lookup"><span data-stu-id="dd864-233">Example</span></span>

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

#### <a name="body-bodyjavascriptapioutlookofficebodyviewoutlook-js-15"></a><span data-ttu-id="dd864-234">body: [Body](/javascript/api/outlook/office.body?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="dd864-234">body: [Body](/javascript/api/outlook/office.body?view=outlook-js-1.5)</span></span>

<span data-ttu-id="dd864-235">Получает объект, предоставляющий методы для работы с основным текстом элемента.</span><span class="sxs-lookup"><span data-stu-id="dd864-235">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="dd864-236">Тип</span><span class="sxs-lookup"><span data-stu-id="dd864-236">Type</span></span>

*   [<span data-ttu-id="dd864-237">Body</span><span class="sxs-lookup"><span data-stu-id="dd864-237">Body</span></span>](/javascript/api/outlook/office.body?view=outlook-js-1.5)

##### <a name="requirements"></a><span data-ttu-id="dd864-238">Требования</span><span class="sxs-lookup"><span data-stu-id="dd864-238">Requirements</span></span>

|<span data-ttu-id="dd864-239">Требование</span><span class="sxs-lookup"><span data-stu-id="dd864-239">Requirement</span></span>| <span data-ttu-id="dd864-240">Значение</span><span class="sxs-lookup"><span data-stu-id="dd864-240">Value</span></span>|
|---|---|
|[<span data-ttu-id="dd864-241">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="dd864-241">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dd864-242">1.1</span><span class="sxs-lookup"><span data-stu-id="dd864-242">1.1</span></span>|
|[<span data-ttu-id="dd864-243">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="dd864-243">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dd864-244">ReadItem</span><span class="sxs-lookup"><span data-stu-id="dd864-244">ReadItem</span></span>|
|[<span data-ttu-id="dd864-245">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="dd864-245">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="dd864-246">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="dd864-246">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="dd864-247">Пример</span><span class="sxs-lookup"><span data-stu-id="dd864-247">Example</span></span>

<span data-ttu-id="dd864-248">В этом примере возвращается текст сообщения в формате обычного текста.</span><span class="sxs-lookup"><span data-stu-id="dd864-248">This example gets the body of the message in plain text.</span></span>

```js
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="dd864-249">Ниже приведен пример итогового параметра, переданного функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="dd864-249">The following is an example of the result parameter passed to the callback function.</span></span>

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

#### <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-15recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-15"></a><span data-ttu-id="dd864-250">cc: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="dd864-250">cc: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.5)</span></span>

<span data-ttu-id="dd864-251">Предоставляет доступ к получателям копии сообщения.</span><span class="sxs-lookup"><span data-stu-id="dd864-251">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="dd864-252">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="dd864-252">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="dd864-253">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="dd864-253">Read mode</span></span>

<span data-ttu-id="dd864-254">Свойство `cc` возвращает массив, который содержит объект `EmailAddressDetails` для каждого получателя, указанного в строке **Копия** сообщения.</span><span class="sxs-lookup"><span data-stu-id="dd864-254">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message.</span></span> <span data-ttu-id="dd864-255">Коллекция может включать не более 100 элементов по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="dd864-255">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="dd864-256">Однако на компьютерах с Windows и Mac OS может быть до 500 элементов.</span><span class="sxs-lookup"><span data-stu-id="dd864-256">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="dd864-257">Режим создания</span><span class="sxs-lookup"><span data-stu-id="dd864-257">Compose mode</span></span>

<span data-ttu-id="dd864-258">Свойство `cc` возвращает объект `Recipients`, предоставляющий методы для получения или обновления получателей, которые указаны в строке **Копия** сообщения.</span><span class="sxs-lookup"><span data-stu-id="dd864-258">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span> <span data-ttu-id="dd864-259">Коллекция может включать не более 100 элементов по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="dd864-259">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="dd864-260">Однако для компьютеров с Windows и Mac OS действуют указанные ниже ограничения.</span><span class="sxs-lookup"><span data-stu-id="dd864-260">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="dd864-261">Максимальное количество элементов — 500.</span><span class="sxs-lookup"><span data-stu-id="dd864-261">Get 500 members maximum.</span></span>
- <span data-ttu-id="dd864-262">Установить ограничение количества элементов на вызов — не более 100, общего количества — не более 500.</span><span class="sxs-lookup"><span data-stu-id="dd864-262">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="dd864-263">Тип</span><span class="sxs-lookup"><span data-stu-id="dd864-263">Type</span></span>

*   <span data-ttu-id="dd864-264">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="dd864-264">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.5)</span></span>

##### <a name="requirements"></a><span data-ttu-id="dd864-265">Требования</span><span class="sxs-lookup"><span data-stu-id="dd864-265">Requirements</span></span>

|<span data-ttu-id="dd864-266">Требование</span><span class="sxs-lookup"><span data-stu-id="dd864-266">Requirement</span></span>| <span data-ttu-id="dd864-267">Значение</span><span class="sxs-lookup"><span data-stu-id="dd864-267">Value</span></span>|
|---|---|
|[<span data-ttu-id="dd864-268">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="dd864-268">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dd864-269">1.0</span><span class="sxs-lookup"><span data-stu-id="dd864-269">1.0</span></span>|
|[<span data-ttu-id="dd864-270">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="dd864-270">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dd864-271">ReadItem</span><span class="sxs-lookup"><span data-stu-id="dd864-271">ReadItem</span></span>|
|[<span data-ttu-id="dd864-272">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="dd864-272">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="dd864-273">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="dd864-273">Compose or Read</span></span>|

<br>

---
---

#### <a name="nullable-conversationid-string"></a><span data-ttu-id="dd864-274">(nullable) conversationId: String</span><span class="sxs-lookup"><span data-stu-id="dd864-274">(nullable) conversationId: String</span></span>

<span data-ttu-id="dd864-275">Получает идентификатор разговора по электронной почте, содержащего конкретное сообщение.</span><span class="sxs-lookup"><span data-stu-id="dd864-275">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="dd864-p109">Вы можете получить целочисленное значение этого свойства, если ваше почтовое приложение активируется в формах просмотра или формах создания ответов. Если пользователь изменит тему ответа, после его отправки идентификатор беседы будет изменен, и полученное ранее значение будет недействительным.</span><span class="sxs-lookup"><span data-stu-id="dd864-p109">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="dd864-p110">Это свойство имеет значение NULL для нового элемента в форме создания. Свойство `conversationId` вернет значение, если пользователь задаст тему и сохранит элемент.</span><span class="sxs-lookup"><span data-stu-id="dd864-p110">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="dd864-280">Тип</span><span class="sxs-lookup"><span data-stu-id="dd864-280">Type</span></span>

*   <span data-ttu-id="dd864-281">String</span><span class="sxs-lookup"><span data-stu-id="dd864-281">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="dd864-282">Требования</span><span class="sxs-lookup"><span data-stu-id="dd864-282">Requirements</span></span>

|<span data-ttu-id="dd864-283">Требование</span><span class="sxs-lookup"><span data-stu-id="dd864-283">Requirement</span></span>| <span data-ttu-id="dd864-284">Значение</span><span class="sxs-lookup"><span data-stu-id="dd864-284">Value</span></span>|
|---|---|
|[<span data-ttu-id="dd864-285">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="dd864-285">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dd864-286">1.0</span><span class="sxs-lookup"><span data-stu-id="dd864-286">1.0</span></span>|
|[<span data-ttu-id="dd864-287">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="dd864-287">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dd864-288">ReadItem</span><span class="sxs-lookup"><span data-stu-id="dd864-288">ReadItem</span></span>|
|[<span data-ttu-id="dd864-289">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="dd864-289">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="dd864-290">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="dd864-290">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="dd864-291">Пример</span><span class="sxs-lookup"><span data-stu-id="dd864-291">Example</span></span>

```js
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

<br>

---
---

#### <a name="datetimecreated-date"></a><span data-ttu-id="dd864-292">dateTimeCreated: Date</span><span class="sxs-lookup"><span data-stu-id="dd864-292">dateTimeCreated: Date</span></span>

<span data-ttu-id="dd864-p111">Получает дату и время создания элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="dd864-p111">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="dd864-295">Тип</span><span class="sxs-lookup"><span data-stu-id="dd864-295">Type</span></span>

*   <span data-ttu-id="dd864-296">Дата</span><span class="sxs-lookup"><span data-stu-id="dd864-296">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="dd864-297">Требования</span><span class="sxs-lookup"><span data-stu-id="dd864-297">Requirements</span></span>

|<span data-ttu-id="dd864-298">Требование</span><span class="sxs-lookup"><span data-stu-id="dd864-298">Requirement</span></span>| <span data-ttu-id="dd864-299">Значение</span><span class="sxs-lookup"><span data-stu-id="dd864-299">Value</span></span>|
|---|---|
|[<span data-ttu-id="dd864-300">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="dd864-300">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dd864-301">1.0</span><span class="sxs-lookup"><span data-stu-id="dd864-301">1.0</span></span>|
|[<span data-ttu-id="dd864-302">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="dd864-302">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dd864-303">ReadItem</span><span class="sxs-lookup"><span data-stu-id="dd864-303">ReadItem</span></span>|
|[<span data-ttu-id="dd864-304">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="dd864-304">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="dd864-305">Чтение</span><span class="sxs-lookup"><span data-stu-id="dd864-305">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="dd864-306">Пример</span><span class="sxs-lookup"><span data-stu-id="dd864-306">Example</span></span>

```js
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

<br>

---
---

#### <a name="datetimemodified-date"></a><span data-ttu-id="dd864-307">dateTimeModified: Date</span><span class="sxs-lookup"><span data-stu-id="dd864-307">dateTimeModified: Date</span></span>

<span data-ttu-id="dd864-p112">Получает дату и время последнего изменения элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="dd864-p112">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="dd864-310">Этот элемент не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="dd864-310">This member is not supported in Outlook on iOS or Android.</span></span>

##### <a name="type"></a><span data-ttu-id="dd864-311">Тип</span><span class="sxs-lookup"><span data-stu-id="dd864-311">Type</span></span>

*   <span data-ttu-id="dd864-312">Дата</span><span class="sxs-lookup"><span data-stu-id="dd864-312">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="dd864-313">Требования</span><span class="sxs-lookup"><span data-stu-id="dd864-313">Requirements</span></span>

|<span data-ttu-id="dd864-314">Требование</span><span class="sxs-lookup"><span data-stu-id="dd864-314">Requirement</span></span>| <span data-ttu-id="dd864-315">Значение</span><span class="sxs-lookup"><span data-stu-id="dd864-315">Value</span></span>|
|---|---|
|[<span data-ttu-id="dd864-316">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="dd864-316">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dd864-317">1.0</span><span class="sxs-lookup"><span data-stu-id="dd864-317">1.0</span></span>|
|[<span data-ttu-id="dd864-318">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="dd864-318">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dd864-319">ReadItem</span><span class="sxs-lookup"><span data-stu-id="dd864-319">ReadItem</span></span>|
|[<span data-ttu-id="dd864-320">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="dd864-320">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="dd864-321">Чтение</span><span class="sxs-lookup"><span data-stu-id="dd864-321">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="dd864-322">Пример</span><span class="sxs-lookup"><span data-stu-id="dd864-322">Example</span></span>

```js
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

<br>

---
---

#### <a name="end-datetimejavascriptapioutlookofficetimeviewoutlook-js-15"></a><span data-ttu-id="dd864-323">end: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="dd864-323">end: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.5)</span></span>

<span data-ttu-id="dd864-324">Получает или задает дату и время окончания встречи.</span><span class="sxs-lookup"><span data-stu-id="dd864-324">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="dd864-p113">Свойство `end` представлено в виде значения даты и времени в формате UTC. Преобразовать значение свойства end в местные значения даты и времени клиента можно с помощью метода [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime).</span><span class="sxs-lookup"><span data-stu-id="dd864-p113">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="dd864-327">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="dd864-327">Read mode</span></span>

<span data-ttu-id="dd864-328">Свойство `end` возвращает объект `Date`.</span><span class="sxs-lookup"><span data-stu-id="dd864-328">The `end` property returns a `Date` object.</span></span>

```js
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="dd864-329">Режим создания</span><span class="sxs-lookup"><span data-stu-id="dd864-329">Compose mode</span></span>

<span data-ttu-id="dd864-330">Свойство `end` возвращает объект `Time`.</span><span class="sxs-lookup"><span data-stu-id="dd864-330">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="dd864-331">Если вы задаете время окончания с помощью метода [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.5#setasync-datetime--options--callback-), необходимо использовать метод [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) для преобразования местного времени на клиенте в формат UTC для сервера.</span><span class="sxs-lookup"><span data-stu-id="dd864-331">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.5#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="dd864-332">В примере ниже показано, как с помощью метода [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.5#setasync-datetime--options--callback-) объекта `Time` задать время окончания встречи.</span><span class="sxs-lookup"><span data-stu-id="dd864-332">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.5#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="dd864-333">Тип</span><span class="sxs-lookup"><span data-stu-id="dd864-333">Type</span></span>

*   <span data-ttu-id="dd864-334">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="dd864-334">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.5)</span></span>

##### <a name="requirements"></a><span data-ttu-id="dd864-335">Требования</span><span class="sxs-lookup"><span data-stu-id="dd864-335">Requirements</span></span>

|<span data-ttu-id="dd864-336">Требование</span><span class="sxs-lookup"><span data-stu-id="dd864-336">Requirement</span></span>| <span data-ttu-id="dd864-337">Значение</span><span class="sxs-lookup"><span data-stu-id="dd864-337">Value</span></span>|
|---|---|
|[<span data-ttu-id="dd864-338">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="dd864-338">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dd864-339">1.0</span><span class="sxs-lookup"><span data-stu-id="dd864-339">1.0</span></span>|
|[<span data-ttu-id="dd864-340">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="dd864-340">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dd864-341">ReadItem</span><span class="sxs-lookup"><span data-stu-id="dd864-341">ReadItem</span></span>|
|[<span data-ttu-id="dd864-342">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="dd864-342">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="dd864-343">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="dd864-343">Compose or Read</span></span>|

<br>

---
---

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-15"></a><span data-ttu-id="dd864-344">from: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="dd864-344">from: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)</span></span>

<span data-ttu-id="dd864-p114">Получает электронный адрес отправителя сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="dd864-p114">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="dd864-p115">Свойства `from` и [`sender`](#sender-emailaddressdetails) представляют одно лицо, если сообщение не отправлено представителем. В противном случае свойство `from` представляет лицо, делегировавшее полномочия, а свойство sender — представителя.</span><span class="sxs-lookup"><span data-stu-id="dd864-p115">The `from` and [`sender`](#sender-emailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="dd864-349">Свойству `recipientType`, принадлежащему объекту `EmailAddressDetails` в свойстве `from`, задано значение `undefined`.</span><span class="sxs-lookup"><span data-stu-id="dd864-349">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="dd864-350">Тип</span><span class="sxs-lookup"><span data-stu-id="dd864-350">Type</span></span>

*   [<span data-ttu-id="dd864-351">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="dd864-351">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)

##### <a name="requirements"></a><span data-ttu-id="dd864-352">Требования</span><span class="sxs-lookup"><span data-stu-id="dd864-352">Requirements</span></span>

|<span data-ttu-id="dd864-353">Требование</span><span class="sxs-lookup"><span data-stu-id="dd864-353">Requirement</span></span>| <span data-ttu-id="dd864-354">Значение</span><span class="sxs-lookup"><span data-stu-id="dd864-354">Value</span></span>|
|---|---|
|[<span data-ttu-id="dd864-355">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="dd864-355">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dd864-356">1.0</span><span class="sxs-lookup"><span data-stu-id="dd864-356">1.0</span></span>|
|[<span data-ttu-id="dd864-357">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="dd864-357">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dd864-358">ReadItem</span><span class="sxs-lookup"><span data-stu-id="dd864-358">ReadItem</span></span>|
|[<span data-ttu-id="dd864-359">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="dd864-359">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="dd864-360">Чтение</span><span class="sxs-lookup"><span data-stu-id="dd864-360">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="dd864-361">Пример</span><span class="sxs-lookup"><span data-stu-id="dd864-361">Example</span></span>

```js
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

<br>

---
---

#### <a name="internetmessageid-string"></a><span data-ttu-id="dd864-362">internetMessageId: String</span><span class="sxs-lookup"><span data-stu-id="dd864-362">internetMessageId: String</span></span>

<span data-ttu-id="dd864-p116">Получает идентификатор интернет-сообщения для электронного сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="dd864-p116">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="dd864-365">Тип</span><span class="sxs-lookup"><span data-stu-id="dd864-365">Type</span></span>

*   <span data-ttu-id="dd864-366">String</span><span class="sxs-lookup"><span data-stu-id="dd864-366">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="dd864-367">Требования</span><span class="sxs-lookup"><span data-stu-id="dd864-367">Requirements</span></span>

|<span data-ttu-id="dd864-368">Требование</span><span class="sxs-lookup"><span data-stu-id="dd864-368">Requirement</span></span>| <span data-ttu-id="dd864-369">Значение</span><span class="sxs-lookup"><span data-stu-id="dd864-369">Value</span></span>|
|---|---|
|[<span data-ttu-id="dd864-370">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="dd864-370">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dd864-371">1.0</span><span class="sxs-lookup"><span data-stu-id="dd864-371">1.0</span></span>|
|[<span data-ttu-id="dd864-372">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="dd864-372">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dd864-373">ReadItem</span><span class="sxs-lookup"><span data-stu-id="dd864-373">ReadItem</span></span>|
|[<span data-ttu-id="dd864-374">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="dd864-374">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="dd864-375">Чтение</span><span class="sxs-lookup"><span data-stu-id="dd864-375">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="dd864-376">Пример</span><span class="sxs-lookup"><span data-stu-id="dd864-376">Example</span></span>

```js
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

<br>

---
---

#### <a name="itemclass-string"></a><span data-ttu-id="dd864-377">itemClass: String</span><span class="sxs-lookup"><span data-stu-id="dd864-377">itemClass: String</span></span>

<span data-ttu-id="dd864-p117">Получает класс элемента веб-служб Exchange для выбранного элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="dd864-p117">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="dd864-p118">Свойство `itemClass` указывает класс сообщения выбранного элемента. Ниже приводятся классы сообщения по умолчанию для элемента сообщения или встречи.</span><span class="sxs-lookup"><span data-stu-id="dd864-p118">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="dd864-382">Тип</span><span class="sxs-lookup"><span data-stu-id="dd864-382">Type</span></span> | <span data-ttu-id="dd864-383">Описание</span><span class="sxs-lookup"><span data-stu-id="dd864-383">Description</span></span> | <span data-ttu-id="dd864-384">Класс элемента</span><span class="sxs-lookup"><span data-stu-id="dd864-384">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="dd864-385">Элементы встречи</span><span class="sxs-lookup"><span data-stu-id="dd864-385">Appointment items</span></span> | <span data-ttu-id="dd864-386">Это элементы календаря для класса элемента `IPM.Appointment` или `IPM.Appointment.Occurrence`.</span><span class="sxs-lookup"><span data-stu-id="dd864-386">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurrence` |
| <span data-ttu-id="dd864-387">Элементы сообщения</span><span class="sxs-lookup"><span data-stu-id="dd864-387">Message items</span></span> | <span data-ttu-id="dd864-388">Сюда входят электронные сообщения, для которых по умолчанию задан класс сообщения `IPM.Note`, а также приглашения на собрания, ответы на них и уведомления об их отмене, использующие `IPM.Schedule.Meeting` в качестве базового класса сообщения.</span><span class="sxs-lookup"><span data-stu-id="dd864-388">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="dd864-389">Можно создавать настраиваемые классы сообщения, расширяющие классы сообщения по умолчанию, например настраиваемый класс сообщения о встрече `IPM.Appointment.Contoso`.</span><span class="sxs-lookup"><span data-stu-id="dd864-389">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="dd864-390">Тип</span><span class="sxs-lookup"><span data-stu-id="dd864-390">Type</span></span>

*   <span data-ttu-id="dd864-391">String</span><span class="sxs-lookup"><span data-stu-id="dd864-391">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="dd864-392">Требования</span><span class="sxs-lookup"><span data-stu-id="dd864-392">Requirements</span></span>

|<span data-ttu-id="dd864-393">Требование</span><span class="sxs-lookup"><span data-stu-id="dd864-393">Requirement</span></span>| <span data-ttu-id="dd864-394">Значение</span><span class="sxs-lookup"><span data-stu-id="dd864-394">Value</span></span>|
|---|---|
|[<span data-ttu-id="dd864-395">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="dd864-395">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dd864-396">1.0</span><span class="sxs-lookup"><span data-stu-id="dd864-396">1.0</span></span>|
|[<span data-ttu-id="dd864-397">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="dd864-397">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dd864-398">ReadItem</span><span class="sxs-lookup"><span data-stu-id="dd864-398">ReadItem</span></span>|
|[<span data-ttu-id="dd864-399">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="dd864-399">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="dd864-400">Чтение</span><span class="sxs-lookup"><span data-stu-id="dd864-400">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="dd864-401">Пример</span><span class="sxs-lookup"><span data-stu-id="dd864-401">Example</span></span>

```js
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

<br>

---
---

#### <a name="nullable-itemid-string"></a><span data-ttu-id="dd864-402">(nullable) itemId: String</span><span class="sxs-lookup"><span data-stu-id="dd864-402">(nullable) itemId: String</span></span>

<span data-ttu-id="dd864-p119">Получает [идентификатор элемента веб-служб Exchange](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange) для текущего элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="dd864-p119">Gets the [Exchange Web Services item identifier](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange) for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="dd864-405">Идентификатор, возвращаемый свойством `itemId`, совпадает с [идентификатором элемента веб-служб Exchange](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange).</span><span class="sxs-lookup"><span data-stu-id="dd864-405">The identifier returned by the `itemId` property is the same as the [Exchange Web Services item identifier](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange).</span></span> <span data-ttu-id="dd864-406">Свойство `itemId` не совпадает с идентификатором записи Outlook, а также идентификатором, который используется REST API Outlook.</span><span class="sxs-lookup"><span data-stu-id="dd864-406">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="dd864-407">Прежде чем совершать вызовы REST API, используя это значение, его необходимо преобразовать с помощью [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="dd864-407">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="dd864-408">Дополнительные сведения см. в статье [Использование REST API Outlook из надстройки Outlook](/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="dd864-408">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="dd864-p121">Свойство `itemId` недоступно в режиме создания. Если требуется идентификатор элемента, с помощью метода [`saveAsync`](#saveasyncoptions-callback) можно сохранить элемент в хранилище. При этом в параметре [`AsyncResult.value`](/javascript/api/office/office.asyncresult) функции обратного вызова возвращается идентификатор элемента.</span><span class="sxs-lookup"><span data-stu-id="dd864-p121">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="dd864-411">Тип</span><span class="sxs-lookup"><span data-stu-id="dd864-411">Type</span></span>

*   <span data-ttu-id="dd864-412">String</span><span class="sxs-lookup"><span data-stu-id="dd864-412">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="dd864-413">Требования</span><span class="sxs-lookup"><span data-stu-id="dd864-413">Requirements</span></span>

|<span data-ttu-id="dd864-414">Требование</span><span class="sxs-lookup"><span data-stu-id="dd864-414">Requirement</span></span>| <span data-ttu-id="dd864-415">Значение</span><span class="sxs-lookup"><span data-stu-id="dd864-415">Value</span></span>|
|---|---|
|[<span data-ttu-id="dd864-416">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="dd864-416">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dd864-417">1.0</span><span class="sxs-lookup"><span data-stu-id="dd864-417">1.0</span></span>|
|[<span data-ttu-id="dd864-418">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="dd864-418">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dd864-419">ReadItem</span><span class="sxs-lookup"><span data-stu-id="dd864-419">ReadItem</span></span>|
|[<span data-ttu-id="dd864-420">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="dd864-420">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="dd864-421">Чтение</span><span class="sxs-lookup"><span data-stu-id="dd864-421">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="dd864-422">Пример</span><span class="sxs-lookup"><span data-stu-id="dd864-422">Example</span></span>

<span data-ttu-id="dd864-p122">Указанный ниже код проверяет наличие идентификатора элемента. Если свойство `itemId` возвращает значение `null` или `undefined`, элемент будет сохранен в хранилище, а из асинхронного результата будет получен идентификатор элемента.</span><span class="sxs-lookup"><span data-stu-id="dd864-p122">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

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

#### <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtypeviewoutlook-js-15"></a><span data-ttu-id="dd864-425">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="dd864-425">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.5)</span></span>

<span data-ttu-id="dd864-426">Получает тип элемента, который представляет экземпляр.</span><span class="sxs-lookup"><span data-stu-id="dd864-426">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="dd864-427">Свойство `itemType` возвращает одно из значений перечисления `ItemType`, которое указывает, является ли экземпляр объекта `item` сообщением или встречей.</span><span class="sxs-lookup"><span data-stu-id="dd864-427">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="dd864-428">Тип</span><span class="sxs-lookup"><span data-stu-id="dd864-428">Type</span></span>

*   [<span data-ttu-id="dd864-429">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="dd864-429">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.5)

##### <a name="requirements"></a><span data-ttu-id="dd864-430">Требования</span><span class="sxs-lookup"><span data-stu-id="dd864-430">Requirements</span></span>

|<span data-ttu-id="dd864-431">Требование</span><span class="sxs-lookup"><span data-stu-id="dd864-431">Requirement</span></span>| <span data-ttu-id="dd864-432">Значение</span><span class="sxs-lookup"><span data-stu-id="dd864-432">Value</span></span>|
|---|---|
|[<span data-ttu-id="dd864-433">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="dd864-433">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dd864-434">1.0</span><span class="sxs-lookup"><span data-stu-id="dd864-434">1.0</span></span>|
|[<span data-ttu-id="dd864-435">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="dd864-435">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dd864-436">ReadItem</span><span class="sxs-lookup"><span data-stu-id="dd864-436">ReadItem</span></span>|
|[<span data-ttu-id="dd864-437">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="dd864-437">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="dd864-438">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="dd864-438">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="dd864-439">Пример</span><span class="sxs-lookup"><span data-stu-id="dd864-439">Example</span></span>

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

#### <a name="location-stringlocationjavascriptapioutlookofficelocationviewoutlook-js-15"></a><span data-ttu-id="dd864-440">location: String|[Location](/javascript/api/outlook/office.location?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="dd864-440">location: String|[Location](/javascript/api/outlook/office.location?view=outlook-js-1.5)</span></span>

<span data-ttu-id="dd864-441">Получает или задает место встречи.</span><span class="sxs-lookup"><span data-stu-id="dd864-441">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="dd864-442">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="dd864-442">Read mode</span></span>

<span data-ttu-id="dd864-443">Свойство `location` возвращает строку, содержащую сведения о месте встречи.</span><span class="sxs-lookup"><span data-stu-id="dd864-443">The `location` property returns a string that contains the location of the appointment.</span></span>

```js
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="dd864-444">Режим создания</span><span class="sxs-lookup"><span data-stu-id="dd864-444">Compose mode</span></span>

<span data-ttu-id="dd864-445">Свойство `location` возвращает объект `Location`, предоставляющий методы, которые используются для получения и задания места встречи.</span><span class="sxs-lookup"><span data-stu-id="dd864-445">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```js
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="dd864-446">Тип</span><span class="sxs-lookup"><span data-stu-id="dd864-446">Type</span></span>

*   <span data-ttu-id="dd864-447">String | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="dd864-447">String | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.5)</span></span>

##### <a name="requirements"></a><span data-ttu-id="dd864-448">Требования</span><span class="sxs-lookup"><span data-stu-id="dd864-448">Requirements</span></span>

|<span data-ttu-id="dd864-449">Требование</span><span class="sxs-lookup"><span data-stu-id="dd864-449">Requirement</span></span>| <span data-ttu-id="dd864-450">Значение</span><span class="sxs-lookup"><span data-stu-id="dd864-450">Value</span></span>|
|---|---|
|[<span data-ttu-id="dd864-451">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="dd864-451">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dd864-452">1.0</span><span class="sxs-lookup"><span data-stu-id="dd864-452">1.0</span></span>|
|[<span data-ttu-id="dd864-453">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="dd864-453">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dd864-454">ReadItem</span><span class="sxs-lookup"><span data-stu-id="dd864-454">ReadItem</span></span>|
|[<span data-ttu-id="dd864-455">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="dd864-455">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="dd864-456">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="dd864-456">Compose or Read</span></span>|

<br>

---
---

#### <a name="normalizedsubject-string"></a><span data-ttu-id="dd864-457">normalizedSubject: String</span><span class="sxs-lookup"><span data-stu-id="dd864-457">normalizedSubject: String</span></span>

<span data-ttu-id="dd864-p123">Получает тему элемента со всеми удаленными префиксами (включая `RE:` и `FWD:`). Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="dd864-p123">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="dd864-p124">Свойство normalizedSubject получает тему элемента со стандартными префиксами (такими как `RE:` и `FW:`), добавляемыми почтовыми программами. Для получения темы элемента с неизмененными префиксами используйте свойство [`subject`](#subject-stringsubject).</span><span class="sxs-lookup"><span data-stu-id="dd864-p124">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="dd864-462">Тип</span><span class="sxs-lookup"><span data-stu-id="dd864-462">Type</span></span>

*   <span data-ttu-id="dd864-463">String</span><span class="sxs-lookup"><span data-stu-id="dd864-463">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="dd864-464">Требования</span><span class="sxs-lookup"><span data-stu-id="dd864-464">Requirements</span></span>

|<span data-ttu-id="dd864-465">Требование</span><span class="sxs-lookup"><span data-stu-id="dd864-465">Requirement</span></span>| <span data-ttu-id="dd864-466">Значение</span><span class="sxs-lookup"><span data-stu-id="dd864-466">Value</span></span>|
|---|---|
|[<span data-ttu-id="dd864-467">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="dd864-467">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dd864-468">1.0</span><span class="sxs-lookup"><span data-stu-id="dd864-468">1.0</span></span>|
|[<span data-ttu-id="dd864-469">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="dd864-469">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dd864-470">ReadItem</span><span class="sxs-lookup"><span data-stu-id="dd864-470">ReadItem</span></span>|
|[<span data-ttu-id="dd864-471">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="dd864-471">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="dd864-472">Чтение</span><span class="sxs-lookup"><span data-stu-id="dd864-472">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="dd864-473">Пример</span><span class="sxs-lookup"><span data-stu-id="dd864-473">Example</span></span>

```js
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

<br>

---
---

#### <a name="notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessagesviewoutlook-js-15"></a><span data-ttu-id="dd864-474">notificationMessages: [NotificationMessages](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="dd864-474">notificationMessages: [NotificationMessages](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.5)</span></span>

<span data-ttu-id="dd864-475">Получает сообщения уведомления для элемента.</span><span class="sxs-lookup"><span data-stu-id="dd864-475">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="dd864-476">Тип</span><span class="sxs-lookup"><span data-stu-id="dd864-476">Type</span></span>

*   [<span data-ttu-id="dd864-477">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="dd864-477">NotificationMessages</span></span>](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.5)

##### <a name="requirements"></a><span data-ttu-id="dd864-478">Требования</span><span class="sxs-lookup"><span data-stu-id="dd864-478">Requirements</span></span>

|<span data-ttu-id="dd864-479">Требование</span><span class="sxs-lookup"><span data-stu-id="dd864-479">Requirement</span></span>| <span data-ttu-id="dd864-480">Значение</span><span class="sxs-lookup"><span data-stu-id="dd864-480">Value</span></span>|
|---|---|
|[<span data-ttu-id="dd864-481">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="dd864-481">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dd864-482">1.3</span><span class="sxs-lookup"><span data-stu-id="dd864-482">1.3</span></span>|
|[<span data-ttu-id="dd864-483">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="dd864-483">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dd864-484">ReadItem</span><span class="sxs-lookup"><span data-stu-id="dd864-484">ReadItem</span></span>|
|[<span data-ttu-id="dd864-485">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="dd864-485">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="dd864-486">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="dd864-486">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="dd864-487">Пример</span><span class="sxs-lookup"><span data-stu-id="dd864-487">Example</span></span>

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

#### <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-15recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-15"></a><span data-ttu-id="dd864-488">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="dd864-488">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.5)</span></span>

<span data-ttu-id="dd864-489">Предоставляет доступ к необязательным участникам события.</span><span class="sxs-lookup"><span data-stu-id="dd864-489">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="dd864-490">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="dd864-490">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="dd864-491">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="dd864-491">Read mode</span></span>

<span data-ttu-id="dd864-492">Свойство `optionalAttendees` возвращает массив, содержащий объект `EmailAddressDetails` для каждого необязательного участника собрания.</span><span class="sxs-lookup"><span data-stu-id="dd864-492">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span> <span data-ttu-id="dd864-493">Коллекция может включать не более 100 элементов по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="dd864-493">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="dd864-494">Однако на компьютерах с Windows и Mac OS может быть до 500 элементов.</span><span class="sxs-lookup"><span data-stu-id="dd864-494">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="dd864-495">Режим создания</span><span class="sxs-lookup"><span data-stu-id="dd864-495">Compose mode</span></span>

<span data-ttu-id="dd864-496">Свойство `optionalAttendees` возвращает объект `Recipients`, который предоставляет методы для получения или обновления необязательных участников собрания.</span><span class="sxs-lookup"><span data-stu-id="dd864-496">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span> <span data-ttu-id="dd864-497">Коллекция может включать не более 100 элементов по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="dd864-497">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="dd864-498">Однако для компьютеров с Windows и Mac OS действуют указанные ниже ограничения.</span><span class="sxs-lookup"><span data-stu-id="dd864-498">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="dd864-499">Максимальное количество элементов — 500.</span><span class="sxs-lookup"><span data-stu-id="dd864-499">Get 500 members maximum.</span></span>
- <span data-ttu-id="dd864-500">Установить ограничение количества элементов на вызов — не более 100, общего количества — не более 500.</span><span class="sxs-lookup"><span data-stu-id="dd864-500">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="dd864-501">Тип</span><span class="sxs-lookup"><span data-stu-id="dd864-501">Type</span></span>

*   <span data-ttu-id="dd864-502">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="dd864-502">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.5)</span></span>

##### <a name="requirements"></a><span data-ttu-id="dd864-503">Требования</span><span class="sxs-lookup"><span data-stu-id="dd864-503">Requirements</span></span>

|<span data-ttu-id="dd864-504">Требование</span><span class="sxs-lookup"><span data-stu-id="dd864-504">Requirement</span></span>| <span data-ttu-id="dd864-505">Значение</span><span class="sxs-lookup"><span data-stu-id="dd864-505">Value</span></span>|
|---|---|
|[<span data-ttu-id="dd864-506">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="dd864-506">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dd864-507">1.0</span><span class="sxs-lookup"><span data-stu-id="dd864-507">1.0</span></span>|
|[<span data-ttu-id="dd864-508">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="dd864-508">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dd864-509">ReadItem</span><span class="sxs-lookup"><span data-stu-id="dd864-509">ReadItem</span></span>|
|[<span data-ttu-id="dd864-510">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="dd864-510">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="dd864-511">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="dd864-511">Compose or Read</span></span>|

<br>

---
---

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-15"></a><span data-ttu-id="dd864-512">organizer: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="dd864-512">organizer: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)</span></span>

<span data-ttu-id="dd864-p128">Получает электронный адрес организатора указанного собрания. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="dd864-p128">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="dd864-515">Тип</span><span class="sxs-lookup"><span data-stu-id="dd864-515">Type</span></span>

*   [<span data-ttu-id="dd864-516">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="dd864-516">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)

##### <a name="requirements"></a><span data-ttu-id="dd864-517">Требования</span><span class="sxs-lookup"><span data-stu-id="dd864-517">Requirements</span></span>

|<span data-ttu-id="dd864-518">Требование</span><span class="sxs-lookup"><span data-stu-id="dd864-518">Requirement</span></span>| <span data-ttu-id="dd864-519">Значение</span><span class="sxs-lookup"><span data-stu-id="dd864-519">Value</span></span>|
|---|---|
|[<span data-ttu-id="dd864-520">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="dd864-520">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dd864-521">1.0</span><span class="sxs-lookup"><span data-stu-id="dd864-521">1.0</span></span>|
|[<span data-ttu-id="dd864-522">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="dd864-522">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dd864-523">ReadItem</span><span class="sxs-lookup"><span data-stu-id="dd864-523">ReadItem</span></span>|
|[<span data-ttu-id="dd864-524">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="dd864-524">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="dd864-525">Чтение</span><span class="sxs-lookup"><span data-stu-id="dd864-525">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="dd864-526">Пример</span><span class="sxs-lookup"><span data-stu-id="dd864-526">Example</span></span>

```js
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

<br>

---
---

#### <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-15recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-15"></a><span data-ttu-id="dd864-527">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="dd864-527">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.5)</span></span>

<span data-ttu-id="dd864-528">Предоставляет доступ к обязательным участникам события.</span><span class="sxs-lookup"><span data-stu-id="dd864-528">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="dd864-529">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="dd864-529">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="dd864-530">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="dd864-530">Read mode</span></span>

<span data-ttu-id="dd864-531">Свойство `requiredAttendees` возвращает массив, содержащий объект `EmailAddressDetails` для каждого обязательного участника собрания.</span><span class="sxs-lookup"><span data-stu-id="dd864-531">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span> <span data-ttu-id="dd864-532">Коллекция может включать не более 100 элементов по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="dd864-532">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="dd864-533">Однако на компьютерах с Windows и Mac OS может быть до 500 элементов.</span><span class="sxs-lookup"><span data-stu-id="dd864-533">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="dd864-534">Режим создания</span><span class="sxs-lookup"><span data-stu-id="dd864-534">Compose mode</span></span>

<span data-ttu-id="dd864-535">Свойство `requiredAttendees` возвращает объект `Recipients`, предоставляющий методы, с помощью которых можно получить или обновить сведения об обязательных участниках собрания.</span><span class="sxs-lookup"><span data-stu-id="dd864-535">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span> <span data-ttu-id="dd864-536">Коллекция может включать не более 100 элементов по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="dd864-536">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="dd864-537">Однако для компьютеров с Windows и Mac OS действуют указанные ниже ограничения.</span><span class="sxs-lookup"><span data-stu-id="dd864-537">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="dd864-538">Максимальное количество элементов — 500.</span><span class="sxs-lookup"><span data-stu-id="dd864-538">Get 500 members maximum.</span></span>
- <span data-ttu-id="dd864-539">Установить ограничение количества элементов на вызов — не более 100, общего количества — не более 500.</span><span class="sxs-lookup"><span data-stu-id="dd864-539">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="dd864-540">Тип</span><span class="sxs-lookup"><span data-stu-id="dd864-540">Type</span></span>

*   <span data-ttu-id="dd864-541">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="dd864-541">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.5)</span></span>

##### <a name="requirements"></a><span data-ttu-id="dd864-542">Требования</span><span class="sxs-lookup"><span data-stu-id="dd864-542">Requirements</span></span>

|<span data-ttu-id="dd864-543">Требование</span><span class="sxs-lookup"><span data-stu-id="dd864-543">Requirement</span></span>| <span data-ttu-id="dd864-544">Значение</span><span class="sxs-lookup"><span data-stu-id="dd864-544">Value</span></span>|
|---|---|
|[<span data-ttu-id="dd864-545">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="dd864-545">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dd864-546">1.0</span><span class="sxs-lookup"><span data-stu-id="dd864-546">1.0</span></span>|
|[<span data-ttu-id="dd864-547">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="dd864-547">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dd864-548">ReadItem</span><span class="sxs-lookup"><span data-stu-id="dd864-548">ReadItem</span></span>|
|[<span data-ttu-id="dd864-549">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="dd864-549">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="dd864-550">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="dd864-550">Compose or Read</span></span>|

<br>

---
---

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-15"></a><span data-ttu-id="dd864-551">sender: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="dd864-551">sender: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)</span></span>

<span data-ttu-id="dd864-p132">Получает электронный адрес отправителя электронного сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="dd864-p132">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="dd864-p133">Свойства [`from`](#from-emailaddressdetails) и `sender` представляют одно лицо, если сообщение не отправлено представителем. В противном случае свойство `from` представляет лицо, делегировавшее полномочия, а свойство sender — представителя.</span><span class="sxs-lookup"><span data-stu-id="dd864-p133">The [`from`](#from-emailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="dd864-556">Свойству `recipientType`, принадлежащему объекту `EmailAddressDetails` в свойстве `sender`, задано значение `undefined`.</span><span class="sxs-lookup"><span data-stu-id="dd864-556">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="dd864-557">Тип</span><span class="sxs-lookup"><span data-stu-id="dd864-557">Type</span></span>

*   [<span data-ttu-id="dd864-558">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="dd864-558">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)

##### <a name="requirements"></a><span data-ttu-id="dd864-559">Требования</span><span class="sxs-lookup"><span data-stu-id="dd864-559">Requirements</span></span>

|<span data-ttu-id="dd864-560">Требование</span><span class="sxs-lookup"><span data-stu-id="dd864-560">Requirement</span></span>| <span data-ttu-id="dd864-561">Значение</span><span class="sxs-lookup"><span data-stu-id="dd864-561">Value</span></span>|
|---|---|
|[<span data-ttu-id="dd864-562">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="dd864-562">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dd864-563">1.0</span><span class="sxs-lookup"><span data-stu-id="dd864-563">1.0</span></span>|
|[<span data-ttu-id="dd864-564">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="dd864-564">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dd864-565">ReadItem</span><span class="sxs-lookup"><span data-stu-id="dd864-565">ReadItem</span></span>|
|[<span data-ttu-id="dd864-566">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="dd864-566">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="dd864-567">Чтение</span><span class="sxs-lookup"><span data-stu-id="dd864-567">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="dd864-568">Пример</span><span class="sxs-lookup"><span data-stu-id="dd864-568">Example</span></span>

```js
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

<br>

---
---

#### <a name="start-datetimejavascriptapioutlookofficetimeviewoutlook-js-15"></a><span data-ttu-id="dd864-569">start: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="dd864-569">start: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.5)</span></span>

<span data-ttu-id="dd864-570">Получает или задает дату и время начала встречи.</span><span class="sxs-lookup"><span data-stu-id="dd864-570">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="dd864-p134">Свойство `start` представлено в виде значения даты и времени в формате UTC. Это значение можно преобразовать в местные значения даты и времени клиента с помощью метода [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime).</span><span class="sxs-lookup"><span data-stu-id="dd864-p134">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="dd864-573">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="dd864-573">Read mode</span></span>

<span data-ttu-id="dd864-574">Свойство `start` возвращает объект `Date`.</span><span class="sxs-lookup"><span data-stu-id="dd864-574">The `start` property returns a `Date` object.</span></span>

```js
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="dd864-575">Режим создания</span><span class="sxs-lookup"><span data-stu-id="dd864-575">Compose mode</span></span>

<span data-ttu-id="dd864-576">Свойство `start` возвращает объект `Time`.</span><span class="sxs-lookup"><span data-stu-id="dd864-576">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="dd864-577">Если вы задаете время начала с помощью метода [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.5#setasync-datetime--options--callback-), необходимо использовать метод [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) для преобразования местного времени на клиенте в формат UTC для сервера.</span><span class="sxs-lookup"><span data-stu-id="dd864-577">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.5#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="dd864-578">В примере ниже с помощью метода [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.5#setasync-datetime--options--callback-) объекта `Time` задается время начала встречи в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="dd864-578">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.5#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="dd864-579">Тип</span><span class="sxs-lookup"><span data-stu-id="dd864-579">Type</span></span>

*   <span data-ttu-id="dd864-580">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="dd864-580">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.5)</span></span>

##### <a name="requirements"></a><span data-ttu-id="dd864-581">Требования</span><span class="sxs-lookup"><span data-stu-id="dd864-581">Requirements</span></span>

|<span data-ttu-id="dd864-582">Требование</span><span class="sxs-lookup"><span data-stu-id="dd864-582">Requirement</span></span>| <span data-ttu-id="dd864-583">Значение</span><span class="sxs-lookup"><span data-stu-id="dd864-583">Value</span></span>|
|---|---|
|[<span data-ttu-id="dd864-584">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="dd864-584">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dd864-585">1.0</span><span class="sxs-lookup"><span data-stu-id="dd864-585">1.0</span></span>|
|[<span data-ttu-id="dd864-586">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="dd864-586">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dd864-587">ReadItem</span><span class="sxs-lookup"><span data-stu-id="dd864-587">ReadItem</span></span>|
|[<span data-ttu-id="dd864-588">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="dd864-588">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="dd864-589">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="dd864-589">Compose or Read</span></span>|

<br>

---
---

#### <a name="subject-stringsubjectjavascriptapioutlookofficesubjectviewoutlook-js-15"></a><span data-ttu-id="dd864-590">subject: String|[Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="dd864-590">subject: String|[Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.5)</span></span>

<span data-ttu-id="dd864-591">Получает или задает описание, которое отображается в поле темы элемента.</span><span class="sxs-lookup"><span data-stu-id="dd864-591">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="dd864-592">Свойство `subject` получает или задает всю тему элемента для отправки с почтового сервера.</span><span class="sxs-lookup"><span data-stu-id="dd864-592">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="dd864-593">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="dd864-593">Read mode</span></span>

<span data-ttu-id="dd864-p135">Свойство `subject` возвращает строку. С помощью свойства [`normalizedSubject`](#normalizedsubject-string) можно получить тему без начальных префиксов, таких как `RE:` и `FW:`.</span><span class="sxs-lookup"><span data-stu-id="dd864-p135">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```js
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="dd864-596">Режим создания</span><span class="sxs-lookup"><span data-stu-id="dd864-596">Compose mode</span></span>

<span data-ttu-id="dd864-597">Свойство `subject` возвращает объект `Subject`, который предоставляет методы для получения и задания темы.</span><span class="sxs-lookup"><span data-stu-id="dd864-597">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```js
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="dd864-598">Тип</span><span class="sxs-lookup"><span data-stu-id="dd864-598">Type</span></span>

*   <span data-ttu-id="dd864-599">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="dd864-599">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.5)</span></span>

##### <a name="requirements"></a><span data-ttu-id="dd864-600">Требования</span><span class="sxs-lookup"><span data-stu-id="dd864-600">Requirements</span></span>

|<span data-ttu-id="dd864-601">Требование</span><span class="sxs-lookup"><span data-stu-id="dd864-601">Requirement</span></span>| <span data-ttu-id="dd864-602">Значение</span><span class="sxs-lookup"><span data-stu-id="dd864-602">Value</span></span>|
|---|---|
|[<span data-ttu-id="dd864-603">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="dd864-603">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dd864-604">1.0</span><span class="sxs-lookup"><span data-stu-id="dd864-604">1.0</span></span>|
|[<span data-ttu-id="dd864-605">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="dd864-605">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dd864-606">ReadItem</span><span class="sxs-lookup"><span data-stu-id="dd864-606">ReadItem</span></span>|
|[<span data-ttu-id="dd864-607">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="dd864-607">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="dd864-608">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="dd864-608">Compose or Read</span></span>|

<br>

---
---

#### <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-15recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-15"></a><span data-ttu-id="dd864-609">to: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="dd864-609">to: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.5)</span></span>

<span data-ttu-id="dd864-610">Предоставляет доступ к получателям, указанным в строке **Кому** сообщения.</span><span class="sxs-lookup"><span data-stu-id="dd864-610">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="dd864-611">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="dd864-611">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="dd864-612">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="dd864-612">Read mode</span></span>

<span data-ttu-id="dd864-613">Свойство `to` возвращает массив, содержащий объект `EmailAddressDetails` для каждого получателя в строке **Кому** сообщения.</span><span class="sxs-lookup"><span data-stu-id="dd864-613">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message.</span></span> <span data-ttu-id="dd864-614">Коллекция может включать не более 100 элементов по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="dd864-614">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="dd864-615">Однако на компьютерах с Windows и Mac OS может быть до 500 элементов.</span><span class="sxs-lookup"><span data-stu-id="dd864-615">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="dd864-616">Режим создания</span><span class="sxs-lookup"><span data-stu-id="dd864-616">Compose mode</span></span>

<span data-ttu-id="dd864-617">Свойство `to` возвращает объект `Recipients`, предоставляющий методы для получения или обновления получателей, которые указаны в строке **Кому** сообщения.</span><span class="sxs-lookup"><span data-stu-id="dd864-617">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span> <span data-ttu-id="dd864-618">Коллекция может включать не более 100 элементов по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="dd864-618">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="dd864-619">Однако для компьютеров с Windows и Mac OS действуют указанные ниже ограничения.</span><span class="sxs-lookup"><span data-stu-id="dd864-619">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="dd864-620">Максимальное количество элементов — 500.</span><span class="sxs-lookup"><span data-stu-id="dd864-620">Get 500 members maximum.</span></span>
- <span data-ttu-id="dd864-621">Установить ограничение количества элементов на вызов — не более 100, общего количества — не более 500.</span><span class="sxs-lookup"><span data-stu-id="dd864-621">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="dd864-622">Тип</span><span class="sxs-lookup"><span data-stu-id="dd864-622">Type</span></span>

*   <span data-ttu-id="dd864-623">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="dd864-623">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.5)</span></span>

##### <a name="requirements"></a><span data-ttu-id="dd864-624">Требования</span><span class="sxs-lookup"><span data-stu-id="dd864-624">Requirements</span></span>

|<span data-ttu-id="dd864-625">Требование</span><span class="sxs-lookup"><span data-stu-id="dd864-625">Requirement</span></span>| <span data-ttu-id="dd864-626">Значение</span><span class="sxs-lookup"><span data-stu-id="dd864-626">Value</span></span>|
|---|---|
|[<span data-ttu-id="dd864-627">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="dd864-627">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dd864-628">1.0</span><span class="sxs-lookup"><span data-stu-id="dd864-628">1.0</span></span>|
|[<span data-ttu-id="dd864-629">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="dd864-629">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dd864-630">ReadItem</span><span class="sxs-lookup"><span data-stu-id="dd864-630">ReadItem</span></span>|
|[<span data-ttu-id="dd864-631">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="dd864-631">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="dd864-632">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="dd864-632">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="dd864-633">Методы</span><span class="sxs-lookup"><span data-stu-id="dd864-633">Methods</span></span>

#### <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="dd864-634">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="dd864-634">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="dd864-635">Добавляет файл в сообщение или встречу в качестве вложения.</span><span class="sxs-lookup"><span data-stu-id="dd864-635">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="dd864-636">Метод `addFileAttachmentAsync` передает файл по указанному универсальному коду ресурса (URI) и вкладывает его в элемент в форме создания.</span><span class="sxs-lookup"><span data-stu-id="dd864-636">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="dd864-637">Идентификатор можно использовать с методом [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback), чтобы удалить вложение, добавленное во время текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="dd864-637">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="dd864-638">Параметры</span><span class="sxs-lookup"><span data-stu-id="dd864-638">Parameters</span></span>

|<span data-ttu-id="dd864-639">Имя</span><span class="sxs-lookup"><span data-stu-id="dd864-639">Name</span></span>| <span data-ttu-id="dd864-640">Тип</span><span class="sxs-lookup"><span data-stu-id="dd864-640">Type</span></span>| <span data-ttu-id="dd864-641">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="dd864-641">Attributes</span></span>| <span data-ttu-id="dd864-642">Описание</span><span class="sxs-lookup"><span data-stu-id="dd864-642">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="dd864-643">String</span><span class="sxs-lookup"><span data-stu-id="dd864-643">String</span></span>||<span data-ttu-id="dd864-p139">Универсальный код ресурса (URI), представляющий расположение файла, который нужно вложить в сообщение или встречу. Максимальная длина — 2048 символов.</span><span class="sxs-lookup"><span data-stu-id="dd864-p139">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="dd864-646">String</span><span class="sxs-lookup"><span data-stu-id="dd864-646">String</span></span>||<span data-ttu-id="dd864-p140">Имя вложения, которое отображается при передаче вложения. Максимальная длина — 255 символов.</span><span class="sxs-lookup"><span data-stu-id="dd864-p140">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="dd864-649">Object</span><span class="sxs-lookup"><span data-stu-id="dd864-649">Object</span></span>| <span data-ttu-id="dd864-650">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="dd864-650">&lt;optional&gt;</span></span>|<span data-ttu-id="dd864-651">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="dd864-651">An object literal that contains one or more of the following properties.</span></span>|
| `options.asyncContext` | <span data-ttu-id="dd864-652">Object</span><span class="sxs-lookup"><span data-stu-id="dd864-652">Object</span></span> | <span data-ttu-id="dd864-653">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="dd864-653">&lt;optional&gt;</span></span> | <span data-ttu-id="dd864-654">В методе обратного вызова разработчики могут указать любой объект, к которому необходимо получить доступ.</span><span class="sxs-lookup"><span data-stu-id="dd864-654">Developers can provide any object they wish to access in the callback method.</span></span> |
| `options.isInline` | <span data-ttu-id="dd864-655">Boolean</span><span class="sxs-lookup"><span data-stu-id="dd864-655">Boolean</span></span> | <span data-ttu-id="dd864-656">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="dd864-656">&lt;optional&gt;</span></span> | <span data-ttu-id="dd864-657">Значение `true` указывает на то, что вложение будет встроено в текст сообщения и не должно отображаться в списке вложений.</span><span class="sxs-lookup"><span data-stu-id="dd864-657">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
|`callback`| <span data-ttu-id="dd864-658">function</span><span class="sxs-lookup"><span data-stu-id="dd864-658">function</span></span>| <span data-ttu-id="dd864-659">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="dd864-659">&lt;optional&gt;</span></span>|<span data-ttu-id="dd864-660">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="dd864-660">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="dd864-661">После успешного выполнения идентификатор вложения будет представлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="dd864-661">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="dd864-662">Если добавить вложение не удастся, объект `asyncResult` будет содержать объект `Error` с описанием ошибки.</span><span class="sxs-lookup"><span data-stu-id="dd864-662">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="dd864-663">Ошибки</span><span class="sxs-lookup"><span data-stu-id="dd864-663">Errors</span></span>

| <span data-ttu-id="dd864-664">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="dd864-664">Error code</span></span> | <span data-ttu-id="dd864-665">Описание</span><span class="sxs-lookup"><span data-stu-id="dd864-665">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="dd864-666">Вложение превышает максимальный размер.</span><span class="sxs-lookup"><span data-stu-id="dd864-666">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="dd864-667">Расширение вложения не поддерживается.</span><span class="sxs-lookup"><span data-stu-id="dd864-667">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="dd864-668">Сообщение или встреча содержат слишком много вложений.</span><span class="sxs-lookup"><span data-stu-id="dd864-668">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="dd864-669">Требования</span><span class="sxs-lookup"><span data-stu-id="dd864-669">Requirements</span></span>

|<span data-ttu-id="dd864-670">Требование</span><span class="sxs-lookup"><span data-stu-id="dd864-670">Requirement</span></span>| <span data-ttu-id="dd864-671">Значение</span><span class="sxs-lookup"><span data-stu-id="dd864-671">Value</span></span>|
|---|---|
|[<span data-ttu-id="dd864-672">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="dd864-672">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dd864-673">1.1</span><span class="sxs-lookup"><span data-stu-id="dd864-673">1.1</span></span>|
|[<span data-ttu-id="dd864-674">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="dd864-674">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dd864-675">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="dd864-675">ReadWriteItem</span></span>|
|[<span data-ttu-id="dd864-676">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="dd864-676">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="dd864-677">Создание</span><span class="sxs-lookup"><span data-stu-id="dd864-677">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="dd864-678">Примеры</span><span class="sxs-lookup"><span data-stu-id="dd864-678">Examples</span></span>

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

<span data-ttu-id="dd864-679">В приведенном ниже примере файл изображения добавляется в качестве встроенного вложения, а ссылка на вложение добавляется в текст сообщения.</span><span class="sxs-lookup"><span data-stu-id="dd864-679">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

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

#### <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="dd864-680">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="dd864-680">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="dd864-681">Добавляет к сообщению элемент Exchange, например сообщение, в виде вложения.</span><span class="sxs-lookup"><span data-stu-id="dd864-681">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="dd864-p141">С помощью метода `addItemAttachmentAsync` можно в элемент формы создания вложить элемент с указанным идентификатором Exchange. Если указать метод обратного вызова, то этот метод вызывается с помощью параметра `asyncResult`, который содержит идентификатор вложения или код, указывающий на ошибки, которые произошли при вложении элемента. При необходимости можно использовать параметр `options` для передачи сведений о состоянии методу обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="dd864-p141">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="dd864-685">Идентификатор можно использовать с методом [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback), чтобы удалить вложение, добавленное во время текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="dd864-685">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="dd864-686">Если ваша надстройка Office выполняется в Outlook в Интернете, метод `addItemAttachmentAsync` обеспечивает вложение элементов в элементы, отличные от редактируемого. Однако это действие не рекомендуется выполнять, так как оно не поддерживается.</span><span class="sxs-lookup"><span data-stu-id="dd864-686">If your Office Add-in is running in Outlook on the web, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="dd864-687">Параметры</span><span class="sxs-lookup"><span data-stu-id="dd864-687">Parameters</span></span>

|<span data-ttu-id="dd864-688">Имя</span><span class="sxs-lookup"><span data-stu-id="dd864-688">Name</span></span>| <span data-ttu-id="dd864-689">Тип</span><span class="sxs-lookup"><span data-stu-id="dd864-689">Type</span></span>| <span data-ttu-id="dd864-690">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="dd864-690">Attributes</span></span>| <span data-ttu-id="dd864-691">Описание</span><span class="sxs-lookup"><span data-stu-id="dd864-691">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="dd864-692">String</span><span class="sxs-lookup"><span data-stu-id="dd864-692">String</span></span>||<span data-ttu-id="dd864-p142">Идентификатор Exchange для вкладываемого элемента. Максимальная длина — 100 символов.</span><span class="sxs-lookup"><span data-stu-id="dd864-p142">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="dd864-695">String</span><span class="sxs-lookup"><span data-stu-id="dd864-695">String</span></span>||<span data-ttu-id="dd864-696">Тема вкладываемого элемента.</span><span class="sxs-lookup"><span data-stu-id="dd864-696">The subject of the item to be attached.</span></span> <span data-ttu-id="dd864-697">Максимальная длина: 255 символов.</span><span class="sxs-lookup"><span data-stu-id="dd864-697">The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="dd864-698">Object</span><span class="sxs-lookup"><span data-stu-id="dd864-698">Object</span></span>| <span data-ttu-id="dd864-699">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="dd864-699">&lt;optional&gt;</span></span>|<span data-ttu-id="dd864-700">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="dd864-700">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="dd864-701">Object</span><span class="sxs-lookup"><span data-stu-id="dd864-701">Object</span></span>| <span data-ttu-id="dd864-702">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="dd864-702">&lt;optional&gt;</span></span>|<span data-ttu-id="dd864-703">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="dd864-703">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="dd864-704">функция</span><span class="sxs-lookup"><span data-stu-id="dd864-704">function</span></span>| <span data-ttu-id="dd864-705">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="dd864-705">&lt;optional&gt;</span></span>|<span data-ttu-id="dd864-706">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="dd864-706">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="dd864-707">После успешного выполнения идентификатор вложения будет представлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="dd864-707">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="dd864-708">Если добавить вложение не удастся, объект `asyncResult` будет содержать объект `Error` с описанием ошибки.</span><span class="sxs-lookup"><span data-stu-id="dd864-708">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="dd864-709">Ошибки</span><span class="sxs-lookup"><span data-stu-id="dd864-709">Errors</span></span>

| <span data-ttu-id="dd864-710">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="dd864-710">Error code</span></span> | <span data-ttu-id="dd864-711">Описание</span><span class="sxs-lookup"><span data-stu-id="dd864-711">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="dd864-712">Сообщение или встреча содержат слишком много вложений.</span><span class="sxs-lookup"><span data-stu-id="dd864-712">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="dd864-713">Требования</span><span class="sxs-lookup"><span data-stu-id="dd864-713">Requirements</span></span>

|<span data-ttu-id="dd864-714">Требование</span><span class="sxs-lookup"><span data-stu-id="dd864-714">Requirement</span></span>| <span data-ttu-id="dd864-715">Значение</span><span class="sxs-lookup"><span data-stu-id="dd864-715">Value</span></span>|
|---|---|
|[<span data-ttu-id="dd864-716">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="dd864-716">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dd864-717">1.1</span><span class="sxs-lookup"><span data-stu-id="dd864-717">1.1</span></span>|
|[<span data-ttu-id="dd864-718">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="dd864-718">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dd864-719">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="dd864-719">ReadWriteItem</span></span>|
|[<span data-ttu-id="dd864-720">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="dd864-720">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="dd864-721">Создание</span><span class="sxs-lookup"><span data-stu-id="dd864-721">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="dd864-722">Пример</span><span class="sxs-lookup"><span data-stu-id="dd864-722">Example</span></span>

<span data-ttu-id="dd864-723">В следующем примере существующий элемент Outlook добавляется в виде вложения с именем `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="dd864-723">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

#### <a name="close"></a><span data-ttu-id="dd864-724">close()</span><span class="sxs-lookup"><span data-stu-id="dd864-724">close()</span></span>

<span data-ttu-id="dd864-725">Закрывает текущий создаваемый элемент.</span><span class="sxs-lookup"><span data-stu-id="dd864-725">Closes the current item that is being composed.</span></span>

<span data-ttu-id="dd864-p144">Работа метода `close` зависит от текущего состояния создаваемого элемента. Если элемент содержит несохраненные изменения, клиент предложит пользователю сохранить или отклонить их либо отменить действие закрытия.</span><span class="sxs-lookup"><span data-stu-id="dd864-p144">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="dd864-728">Если в Outlook в Интернете элемент представляет собой встречу, ранее сохраненную с помощью метода `saveAsync`, пользователю предлагается сохранить, отклонить или отменить действие, даже если с момента последнего сохранения элемента не вносились какие-либо изменения.</span><span class="sxs-lookup"><span data-stu-id="dd864-728">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="dd864-729">Если в клиенте Outlook для настольных ПК сообщение представляет собой ответ в тексте, метод `close` не работает.</span><span class="sxs-lookup"><span data-stu-id="dd864-729">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="dd864-730">Требования</span><span class="sxs-lookup"><span data-stu-id="dd864-730">Requirements</span></span>

|<span data-ttu-id="dd864-731">Требование</span><span class="sxs-lookup"><span data-stu-id="dd864-731">Requirement</span></span>| <span data-ttu-id="dd864-732">Значение</span><span class="sxs-lookup"><span data-stu-id="dd864-732">Value</span></span>|
|---|---|
|[<span data-ttu-id="dd864-733">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="dd864-733">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dd864-734">1.3</span><span class="sxs-lookup"><span data-stu-id="dd864-734">1.3</span></span>|
|[<span data-ttu-id="dd864-735">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="dd864-735">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dd864-736">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="dd864-736">Restricted</span></span>|
|[<span data-ttu-id="dd864-737">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="dd864-737">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="dd864-738">Создание</span><span class="sxs-lookup"><span data-stu-id="dd864-738">Compose</span></span>|

<br>

---
---

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="dd864-739">displayReplyAllForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="dd864-739">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="dd864-740">Отображает форму ответа, включающую отправителя и всех получателей выбранного сообщения или организатора и всех участников выбранной встречи.</span><span class="sxs-lookup"><span data-stu-id="dd864-740">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="dd864-741">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="dd864-741">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="dd864-742">В Outlook в Интернете форма ответа отображается в виде всплывающей формы в представлении с 3 либо 1 или 2 столбцами.</span><span class="sxs-lookup"><span data-stu-id="dd864-742">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="dd864-743">Если любой строковый параметр превышает указанные для него ограничения, `displayReplyAllForm` возвращает исключение.</span><span class="sxs-lookup"><span data-stu-id="dd864-743">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="dd864-p145">Если в параметре `formData.attachments` указаны вложения, Outlook в Интернете и классические клиенты пытаются скачать их и вложить в форму ответа. Если какие-либо вложения добавить не удается, в форме отображается сообщение об ошибке. Если сообщения об ошибках не предусмотрены, то они не отображаются.</span><span class="sxs-lookup"><span data-stu-id="dd864-p145">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="dd864-747">Параметры</span><span class="sxs-lookup"><span data-stu-id="dd864-747">Parameters</span></span>

| <span data-ttu-id="dd864-748">Имя</span><span class="sxs-lookup"><span data-stu-id="dd864-748">Name</span></span> | <span data-ttu-id="dd864-749">Тип</span><span class="sxs-lookup"><span data-stu-id="dd864-749">Type</span></span> | <span data-ttu-id="dd864-750">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="dd864-750">Attributes</span></span> | <span data-ttu-id="dd864-751">Описание</span><span class="sxs-lookup"><span data-stu-id="dd864-751">Description</span></span> |
|---|---|---|---|
|`formData`| <span data-ttu-id="dd864-752">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="dd864-752">String &#124; Object</span></span>| |<span data-ttu-id="dd864-p146">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="dd864-p146">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="dd864-755">**ИЛИ**</span><span class="sxs-lookup"><span data-stu-id="dd864-755">**OR**</span></span><br/><span data-ttu-id="dd864-p147">Объект, который содержит текст или данные вложения и функцию обратного вызова. Ниже представлено определение этого объекта.</span><span class="sxs-lookup"><span data-stu-id="dd864-p147">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="dd864-758">String</span><span class="sxs-lookup"><span data-stu-id="dd864-758">String</span></span> | <span data-ttu-id="dd864-759">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="dd864-759">&lt;optional&gt;</span></span> | <span data-ttu-id="dd864-p148">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="dd864-p148">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="dd864-762">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="dd864-762">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="dd864-763">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="dd864-763">&lt;optional&gt;</span></span> | <span data-ttu-id="dd864-764">Массив объектов JSON, представляющих собой вложенные файлы или элементы.</span><span class="sxs-lookup"><span data-stu-id="dd864-764">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="dd864-765">String</span><span class="sxs-lookup"><span data-stu-id="dd864-765">String</span></span> | | <span data-ttu-id="dd864-p149">Указывает тип вложения. Допустимые значения: `file` для вложенного файла и `item` для вложенного элемента.</span><span class="sxs-lookup"><span data-stu-id="dd864-p149">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="dd864-768">String</span><span class="sxs-lookup"><span data-stu-id="dd864-768">String</span></span> | | <span data-ttu-id="dd864-769">Строка, содержащая имя вложения, длиной до 255 символов.</span><span class="sxs-lookup"><span data-stu-id="dd864-769">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="dd864-770">String</span><span class="sxs-lookup"><span data-stu-id="dd864-770">String</span></span> | | <span data-ttu-id="dd864-p150">Используется, только если свойству `type` задано значение `file`. URI расположения файла.</span><span class="sxs-lookup"><span data-stu-id="dd864-p150">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.isInline` | <span data-ttu-id="dd864-773">Логический</span><span class="sxs-lookup"><span data-stu-id="dd864-773">Boolean</span></span> | | <span data-ttu-id="dd864-p151">Используется, только если свойству `type` задано значение `file`. Значение `true` указывает на то, что вложение будет встроено в текст сообщения и не должно отображаться в списке вложений.</span><span class="sxs-lookup"><span data-stu-id="dd864-p151">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="dd864-776">String</span><span class="sxs-lookup"><span data-stu-id="dd864-776">String</span></span> | | <span data-ttu-id="dd864-p152">Используется, только если свойству `type` задано значение `item`. Идентификатор вложения EWS. Это строка длиной до 100 символов.</span><span class="sxs-lookup"><span data-stu-id="dd864-p152">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="dd864-780">function</span><span class="sxs-lookup"><span data-stu-id="dd864-780">function</span></span> | <span data-ttu-id="dd864-781">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="dd864-781">&lt;optional&gt;</span></span> | <span data-ttu-id="dd864-782">По завершении работы метода функция, переданная параметру `callback`, вызывается с помощью одного параметра `asyncResult`, представляющего собой объект [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="dd864-782">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="dd864-783">Требования</span><span class="sxs-lookup"><span data-stu-id="dd864-783">Requirements</span></span>

|<span data-ttu-id="dd864-784">Требование</span><span class="sxs-lookup"><span data-stu-id="dd864-784">Requirement</span></span>| <span data-ttu-id="dd864-785">Значение</span><span class="sxs-lookup"><span data-stu-id="dd864-785">Value</span></span>|
|---|---|
|[<span data-ttu-id="dd864-786">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="dd864-786">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dd864-787">1.0</span><span class="sxs-lookup"><span data-stu-id="dd864-787">1.0</span></span>|
|[<span data-ttu-id="dd864-788">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="dd864-788">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dd864-789">ReadItem</span><span class="sxs-lookup"><span data-stu-id="dd864-789">ReadItem</span></span>|
|[<span data-ttu-id="dd864-790">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="dd864-790">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="dd864-791">Чтение</span><span class="sxs-lookup"><span data-stu-id="dd864-791">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="dd864-792">Примеры</span><span class="sxs-lookup"><span data-stu-id="dd864-792">Examples</span></span>

<span data-ttu-id="dd864-793">Приведенный ниже код передает строку в функцию `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="dd864-793">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="dd864-794">Ответ с пустым текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="dd864-794">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="dd864-795">Ответ только с текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="dd864-795">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="dd864-796">Ответ с текстом сообщения и вложенным файлом.</span><span class="sxs-lookup"><span data-stu-id="dd864-796">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="dd864-797">Ответ с текстом сообщения и вложенным элементом.</span><span class="sxs-lookup"><span data-stu-id="dd864-797">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="dd864-798">Ответ с текстом сообщения, вложенным файлом, вложенным элементом и обратным вызовом.</span><span class="sxs-lookup"><span data-stu-id="dd864-798">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="dd864-799">displayReplyForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="dd864-799">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="dd864-800">Отображает форму ответа, включающую только отправителя выбранного сообщения или организатора выбранной встречи.</span><span class="sxs-lookup"><span data-stu-id="dd864-800">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="dd864-801">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="dd864-801">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="dd864-802">В Outlook в Интернете форма ответа отображается в виде всплывающей формы в представлении с 3 либо 1 или 2 столбцами.</span><span class="sxs-lookup"><span data-stu-id="dd864-802">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="dd864-803">Если любой строковый параметр превышает указанные для него ограничения, `displayReplyForm` возвращает исключение.</span><span class="sxs-lookup"><span data-stu-id="dd864-803">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="dd864-p153">Если в параметре `formData.attachments` указаны вложения, Outlook в Интернете и классические клиенты пытаются скачать их и вложить в форму ответа. Если какие-либо вложения добавить не удается, в форме отображается сообщение об ошибке. Если сообщения об ошибках не предусмотрены, то они не отображаются.</span><span class="sxs-lookup"><span data-stu-id="dd864-p153">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="dd864-807">Параметры</span><span class="sxs-lookup"><span data-stu-id="dd864-807">Parameters</span></span>

| <span data-ttu-id="dd864-808">Имя</span><span class="sxs-lookup"><span data-stu-id="dd864-808">Name</span></span> | <span data-ttu-id="dd864-809">Тип</span><span class="sxs-lookup"><span data-stu-id="dd864-809">Type</span></span> | <span data-ttu-id="dd864-810">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="dd864-810">Attributes</span></span> | <span data-ttu-id="dd864-811">Описание</span><span class="sxs-lookup"><span data-stu-id="dd864-811">Description</span></span> |
|---|---|---|---|
|`formData`| <span data-ttu-id="dd864-812">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="dd864-812">String &#124; Object</span></span>| | <span data-ttu-id="dd864-p154">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="dd864-p154">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="dd864-815">**ИЛИ**</span><span class="sxs-lookup"><span data-stu-id="dd864-815">**OR**</span></span><br/><span data-ttu-id="dd864-p155">Объект, который содержит текст или данные вложения и функцию обратного вызова. Ниже представлено определение этого объекта.</span><span class="sxs-lookup"><span data-stu-id="dd864-p155">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="dd864-818">String</span><span class="sxs-lookup"><span data-stu-id="dd864-818">String</span></span> | <span data-ttu-id="dd864-819">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="dd864-819">&lt;optional&gt;</span></span> | <span data-ttu-id="dd864-p156">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="dd864-p156">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="dd864-822">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="dd864-822">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="dd864-823">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="dd864-823">&lt;optional&gt;</span></span> | <span data-ttu-id="dd864-824">Массив объектов JSON, представляющих собой вложенные файлы или элементы.</span><span class="sxs-lookup"><span data-stu-id="dd864-824">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="dd864-825">String</span><span class="sxs-lookup"><span data-stu-id="dd864-825">String</span></span> | | <span data-ttu-id="dd864-p157">Указывает тип вложения. Допустимые значения: `file` для вложенного файла и `item` для вложенного элемента.</span><span class="sxs-lookup"><span data-stu-id="dd864-p157">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="dd864-828">String</span><span class="sxs-lookup"><span data-stu-id="dd864-828">String</span></span> | | <span data-ttu-id="dd864-829">Строка, содержащая имя вложения, длиной до 255 символов.</span><span class="sxs-lookup"><span data-stu-id="dd864-829">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="dd864-830">String</span><span class="sxs-lookup"><span data-stu-id="dd864-830">String</span></span> | | <span data-ttu-id="dd864-p158">Используется, только если свойству `type` задано значение `file`. URI расположения файла.</span><span class="sxs-lookup"><span data-stu-id="dd864-p158">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.isInline` | <span data-ttu-id="dd864-833">Логический</span><span class="sxs-lookup"><span data-stu-id="dd864-833">Boolean</span></span> | | <span data-ttu-id="dd864-p159">Используется, только если свойству `type` задано значение `file`. Значение `true` указывает на то, что вложение будет встроено в текст сообщения и не должно отображаться в списке вложений.</span><span class="sxs-lookup"><span data-stu-id="dd864-p159">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="dd864-836">String</span><span class="sxs-lookup"><span data-stu-id="dd864-836">String</span></span> | | <span data-ttu-id="dd864-p160">Используется, только если свойству `type` задано значение `item`. Идентификатор вложения EWS. Это строка длиной до 100 символов.</span><span class="sxs-lookup"><span data-stu-id="dd864-p160">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="dd864-840">function</span><span class="sxs-lookup"><span data-stu-id="dd864-840">function</span></span> | <span data-ttu-id="dd864-841">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="dd864-841">&lt;optional&gt;</span></span> | <span data-ttu-id="dd864-842">По завершении работы метода функция, переданная параметру `callback`, вызывается с помощью одного параметра `asyncResult`, представляющего собой объект [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="dd864-842">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="dd864-843">Требования</span><span class="sxs-lookup"><span data-stu-id="dd864-843">Requirements</span></span>

|<span data-ttu-id="dd864-844">Требование</span><span class="sxs-lookup"><span data-stu-id="dd864-844">Requirement</span></span>| <span data-ttu-id="dd864-845">Значение</span><span class="sxs-lookup"><span data-stu-id="dd864-845">Value</span></span>|
|---|---|
|[<span data-ttu-id="dd864-846">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="dd864-846">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dd864-847">1.0</span><span class="sxs-lookup"><span data-stu-id="dd864-847">1.0</span></span>|
|[<span data-ttu-id="dd864-848">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="dd864-848">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dd864-849">ReadItem</span><span class="sxs-lookup"><span data-stu-id="dd864-849">ReadItem</span></span>|
|[<span data-ttu-id="dd864-850">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="dd864-850">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="dd864-851">Чтение</span><span class="sxs-lookup"><span data-stu-id="dd864-851">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="dd864-852">Примеры</span><span class="sxs-lookup"><span data-stu-id="dd864-852">Examples</span></span>

<span data-ttu-id="dd864-853">Приведенный ниже код передает строку в функцию `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="dd864-853">The following code passes a string to the `displayReplyForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="dd864-854">Ответ с пустым текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="dd864-854">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="dd864-855">Ответ только с текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="dd864-855">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="dd864-856">Ответ с текстом сообщения и вложенным файлом.</span><span class="sxs-lookup"><span data-stu-id="dd864-856">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="dd864-857">Ответ с текстом сообщения и вложенным элементом.</span><span class="sxs-lookup"><span data-stu-id="dd864-857">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="dd864-858">Ответ с текстом сообщения, вложенным файлом, вложенным элементом и обратным вызовом.</span><span class="sxs-lookup"><span data-stu-id="dd864-858">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlookofficeentitiesviewoutlook-js-15"></a><span data-ttu-id="dd864-859">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.5)}</span><span class="sxs-lookup"><span data-stu-id="dd864-859">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.5)}</span></span>

<span data-ttu-id="dd864-860">Получает сущности, обнаруженные в теле выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="dd864-860">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="dd864-861">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="dd864-861">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="dd864-862">Требования</span><span class="sxs-lookup"><span data-stu-id="dd864-862">Requirements</span></span>

|<span data-ttu-id="dd864-863">Требование</span><span class="sxs-lookup"><span data-stu-id="dd864-863">Requirement</span></span>| <span data-ttu-id="dd864-864">Значение</span><span class="sxs-lookup"><span data-stu-id="dd864-864">Value</span></span>|
|---|---|
|[<span data-ttu-id="dd864-865">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="dd864-865">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dd864-866">1.0</span><span class="sxs-lookup"><span data-stu-id="dd864-866">1.0</span></span>|
|[<span data-ttu-id="dd864-867">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="dd864-867">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dd864-868">ReadItem</span><span class="sxs-lookup"><span data-stu-id="dd864-868">ReadItem</span></span>|
|[<span data-ttu-id="dd864-869">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="dd864-869">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="dd864-870">Чтение</span><span class="sxs-lookup"><span data-stu-id="dd864-870">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="dd864-871">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="dd864-871">Returns:</span></span>

<span data-ttu-id="dd864-872">Тип: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="dd864-872">Type: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.5)</span></span>

##### <a name="example"></a><span data-ttu-id="dd864-873">Пример</span><span class="sxs-lookup"><span data-stu-id="dd864-873">Example</span></span>

<span data-ttu-id="dd864-874">Ниже приведен пример получения доступа к сущностям контактов в теле текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="dd864-874">The following example accesses the contacts entities in the current item's body.</span></span>

```js
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

<br>

---
---

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-15meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-15phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-15tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-15"></a><span data-ttu-id="dd864-875">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.5)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.5)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.5)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.5))>}</span><span class="sxs-lookup"><span data-stu-id="dd864-875">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.5)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.5)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.5)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.5))>}</span></span>

<span data-ttu-id="dd864-876">Получает массив всех сущностей указанного типа, обнаруженных в теле выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="dd864-876">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="dd864-877">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="dd864-877">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="dd864-878">Параметры</span><span class="sxs-lookup"><span data-stu-id="dd864-878">Parameters</span></span>

|<span data-ttu-id="dd864-879">Имя</span><span class="sxs-lookup"><span data-stu-id="dd864-879">Name</span></span>| <span data-ttu-id="dd864-880">Тип</span><span class="sxs-lookup"><span data-stu-id="dd864-880">Type</span></span>| <span data-ttu-id="dd864-881">Описание</span><span class="sxs-lookup"><span data-stu-id="dd864-881">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="dd864-882">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="dd864-882">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook/office.mailboxenums.entitytype?view=outlook-js-1.5)|<span data-ttu-id="dd864-883">Одно из значений перечисления EntityType.</span><span class="sxs-lookup"><span data-stu-id="dd864-883">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="dd864-884">Требования</span><span class="sxs-lookup"><span data-stu-id="dd864-884">Requirements</span></span>

|<span data-ttu-id="dd864-885">Требование</span><span class="sxs-lookup"><span data-stu-id="dd864-885">Requirement</span></span>| <span data-ttu-id="dd864-886">Значение</span><span class="sxs-lookup"><span data-stu-id="dd864-886">Value</span></span>|
|---|---|
|[<span data-ttu-id="dd864-887">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="dd864-887">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dd864-888">1.0</span><span class="sxs-lookup"><span data-stu-id="dd864-888">1.0</span></span>|
|[<span data-ttu-id="dd864-889">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="dd864-889">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dd864-890">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="dd864-890">Restricted</span></span>|
|[<span data-ttu-id="dd864-891">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="dd864-891">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="dd864-892">Чтение</span><span class="sxs-lookup"><span data-stu-id="dd864-892">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="dd864-893">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="dd864-893">Returns:</span></span>

<span data-ttu-id="dd864-894">Если значение, переданное в `entityType`, не является допустимым членом перечисления `EntityType`, метод возвращает значение NULL.</span><span class="sxs-lookup"><span data-stu-id="dd864-894">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="dd864-895">Если в теле элемента отсутствуют сущности указанного типа, метод возвращает пустой массив.</span><span class="sxs-lookup"><span data-stu-id="dd864-895">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="dd864-896">В противном случае тип объектов в возвращаемом массиве зависит от типа сущности, запрошенной в параметре `entityType`.</span><span class="sxs-lookup"><span data-stu-id="dd864-896">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="dd864-897">Хотя минимальный уровень разрешений для использования этого метода — **Restricted**, для некоторых типов сущностей требуется доступ на уровне **ReadItem**, как указано в приведенной ниже таблице.</span><span class="sxs-lookup"><span data-stu-id="dd864-897">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="dd864-898">Значение параметра `entityType`</span><span class="sxs-lookup"><span data-stu-id="dd864-898">Value of `entityType`</span></span> | <span data-ttu-id="dd864-899">Тип объектов в возвращаемом массиве</span><span class="sxs-lookup"><span data-stu-id="dd864-899">Type of objects in returned array</span></span> | <span data-ttu-id="dd864-900">Необходимый уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="dd864-900">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="dd864-901">String</span><span class="sxs-lookup"><span data-stu-id="dd864-901">String</span></span> | <span data-ttu-id="dd864-902">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="dd864-902">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="dd864-903">Contact</span><span class="sxs-lookup"><span data-stu-id="dd864-903">Contact</span></span> | <span data-ttu-id="dd864-904">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="dd864-904">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="dd864-905">String</span><span class="sxs-lookup"><span data-stu-id="dd864-905">String</span></span> | <span data-ttu-id="dd864-906">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="dd864-906">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="dd864-907">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="dd864-907">MeetingSuggestion</span></span> | <span data-ttu-id="dd864-908">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="dd864-908">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="dd864-909">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="dd864-909">PhoneNumber</span></span> | <span data-ttu-id="dd864-910">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="dd864-910">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="dd864-911">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="dd864-911">TaskSuggestion</span></span> | <span data-ttu-id="dd864-912">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="dd864-912">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="dd864-913">String</span><span class="sxs-lookup"><span data-stu-id="dd864-913">String</span></span> | <span data-ttu-id="dd864-914">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="dd864-914">**Restricted**</span></span> |

<span data-ttu-id="dd864-915">Тип: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.5)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.5)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.5)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.5))></span><span class="sxs-lookup"><span data-stu-id="dd864-915">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.5)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.5)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.5)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.5))></span></span>

##### <a name="example"></a><span data-ttu-id="dd864-916">Пример</span><span class="sxs-lookup"><span data-stu-id="dd864-916">Example</span></span>

<span data-ttu-id="dd864-917">В примере ниже показано, как получить доступ к массиву строк, которые представляют собой почтовые адреса в теле текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="dd864-917">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-15meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-15phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-15tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-15"></a><span data-ttu-id="dd864-918">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.5)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.5)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.5)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.5))>}</span><span class="sxs-lookup"><span data-stu-id="dd864-918">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.5)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.5)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.5)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.5))>}</span></span>

<span data-ttu-id="dd864-919">Возвращает известные сущности в выбранном элементе, которые проходят через именованный фильтр, определяемый в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="dd864-919">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="dd864-920">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="dd864-920">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="dd864-921">Метод `getFilteredEntitiesByName` возвращает сущности, соответствующие регулярному выражению, которое определяется в элементе правила [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) в XML-файле манифеста, с использованием указанного значения элемента `FilterName`.</span><span class="sxs-lookup"><span data-stu-id="dd864-921">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="dd864-922">Параметры</span><span class="sxs-lookup"><span data-stu-id="dd864-922">Parameters</span></span>

|<span data-ttu-id="dd864-923">Имя</span><span class="sxs-lookup"><span data-stu-id="dd864-923">Name</span></span>| <span data-ttu-id="dd864-924">Тип</span><span class="sxs-lookup"><span data-stu-id="dd864-924">Type</span></span>| <span data-ttu-id="dd864-925">Описание</span><span class="sxs-lookup"><span data-stu-id="dd864-925">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="dd864-926">String</span><span class="sxs-lookup"><span data-stu-id="dd864-926">String</span></span>|<span data-ttu-id="dd864-927">Имя элемента правила `ItemHasKnownEntity`, определяющее соответствующий фильтр.</span><span class="sxs-lookup"><span data-stu-id="dd864-927">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="dd864-928">Требования</span><span class="sxs-lookup"><span data-stu-id="dd864-928">Requirements</span></span>

|<span data-ttu-id="dd864-929">Требование</span><span class="sxs-lookup"><span data-stu-id="dd864-929">Requirement</span></span>| <span data-ttu-id="dd864-930">Значение</span><span class="sxs-lookup"><span data-stu-id="dd864-930">Value</span></span>|
|---|---|
|[<span data-ttu-id="dd864-931">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="dd864-931">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dd864-932">1.0</span><span class="sxs-lookup"><span data-stu-id="dd864-932">1.0</span></span>|
|[<span data-ttu-id="dd864-933">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="dd864-933">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dd864-934">ReadItem</span><span class="sxs-lookup"><span data-stu-id="dd864-934">ReadItem</span></span>|
|[<span data-ttu-id="dd864-935">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="dd864-935">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="dd864-936">Чтение</span><span class="sxs-lookup"><span data-stu-id="dd864-936">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="dd864-937">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="dd864-937">Returns:</span></span>

<span data-ttu-id="dd864-p162">Если в манифесте нет элемента `ItemHasKnownEntity` со значением `FilterName`, соответствующим параметру `name`, метод возвращает `null`. Если параметр `name` соответствует элементу `ItemHasKnownEntity` в манифесте, но при этом в текущем элементе нет соответствующих сущностей, метод возвращает пустой массив.</span><span class="sxs-lookup"><span data-stu-id="dd864-p162">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="dd864-940">Тип: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.5)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.5)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.5)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.5))></span><span class="sxs-lookup"><span data-stu-id="dd864-940">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.5)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.5)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.5)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.5))></span></span>

<br>

---
---

#### <a name="getregexmatches--object"></a><span data-ttu-id="dd864-941">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="dd864-941">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="dd864-942">Возвращает строковые значения в выбранном элементе, которые соответствуют регулярным выражениям, определенным в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="dd864-942">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="dd864-943">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="dd864-943">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="dd864-p163">Метод `getRegExMatches` возвращает строки, соответствующие регулярному выражению, которое определяется в каждом элементе правила `ItemHasRegularExpressionMatch` или `ItemHasKnownEntity` в XML-файле манифеста. Для правила `ItemHasRegularExpressionMatch` соответствующую строку должно содержать свойство элемента, указанного этим правилом. Простой тип `PropertyName` определяет поддерживаемые свойства.</span><span class="sxs-lookup"><span data-stu-id="dd864-p163">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="dd864-947">Например, рассмотрим манифест надстройки, который содержит указанный ниже элемент `Rule`.</span><span class="sxs-lookup"><span data-stu-id="dd864-947">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="dd864-948">Объект, возвращаемый методом `getRegExMatches`, будет содержать два свойства: `fruits` и `veggies`.</span><span class="sxs-lookup"><span data-stu-id="dd864-948">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="dd864-p164">Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты. Лучше используйте метод [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.5#getasync-coerciontype--options--callback-) для этого.</span><span class="sxs-lookup"><span data-stu-id="dd864-p164">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.5#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="dd864-952">Requirements</span><span class="sxs-lookup"><span data-stu-id="dd864-952">Requirements</span></span>

|<span data-ttu-id="dd864-953">Требование</span><span class="sxs-lookup"><span data-stu-id="dd864-953">Requirement</span></span>| <span data-ttu-id="dd864-954">Значение</span><span class="sxs-lookup"><span data-stu-id="dd864-954">Value</span></span>|
|---|---|
|[<span data-ttu-id="dd864-955">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="dd864-955">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dd864-956">1.0</span><span class="sxs-lookup"><span data-stu-id="dd864-956">1.0</span></span>|
|[<span data-ttu-id="dd864-957">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="dd864-957">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dd864-958">ReadItem</span><span class="sxs-lookup"><span data-stu-id="dd864-958">ReadItem</span></span>|
|[<span data-ttu-id="dd864-959">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="dd864-959">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="dd864-960">Чтение</span><span class="sxs-lookup"><span data-stu-id="dd864-960">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="dd864-961">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="dd864-961">Returns:</span></span>

<span data-ttu-id="dd864-p165">Объект, содержащий массив строк, которые соответствуют регулярным выражениям, определяемым в XML-файле манифеста. Имя каждого массива равно соответствующему значению атрибута `RegExName` подходящего правила `ItemHasRegularExpressionMatch` или атрибута `FilterName` соответствующего правила `ItemHasKnownEntity`.</span><span class="sxs-lookup"><span data-stu-id="dd864-p165">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<span data-ttu-id="dd864-964">Тип: Object</span><span class="sxs-lookup"><span data-stu-id="dd864-964">Type: Object</span></span>

##### <a name="example"></a><span data-ttu-id="dd864-965">Пример</span><span class="sxs-lookup"><span data-stu-id="dd864-965">Example</span></span>

<span data-ttu-id="dd864-966">В примере ниже показано, как получить доступ к массиву совпадений для <rule>элементов регулярного выражения `fruits` и `veggies`, которые указаны в манифесте</rule>.</span><span class="sxs-lookup"><span data-stu-id="dd864-966">The following example shows how to access the array of matches for the regular expression <rule>elements `fruits` and `veggies`, which are specified in the manifest.</rule></span></span>

```js
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

<br>

---
---

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="dd864-967">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="dd864-967">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="dd864-968">Возвращает строковые значения в выбранном элементе, которые соответствуют именованному регулярному выражению, определенному в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="dd864-968">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="dd864-969">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="dd864-969">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="dd864-970">Метод `getRegExMatchesByName` возвращает строки, соответствующие регулярному выражению, которое определяется в элементе правила `ItemHasRegularExpressionMatch` в XML-файле манифеста, с использованием указанного значения элемента `RegExName`.</span><span class="sxs-lookup"><span data-stu-id="dd864-970">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="dd864-p166">Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты.</span><span class="sxs-lookup"><span data-stu-id="dd864-p166">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="dd864-973">Параметры</span><span class="sxs-lookup"><span data-stu-id="dd864-973">Parameters</span></span>

|<span data-ttu-id="dd864-974">Имя</span><span class="sxs-lookup"><span data-stu-id="dd864-974">Name</span></span>| <span data-ttu-id="dd864-975">Тип</span><span class="sxs-lookup"><span data-stu-id="dd864-975">Type</span></span>| <span data-ttu-id="dd864-976">Описание</span><span class="sxs-lookup"><span data-stu-id="dd864-976">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="dd864-977">String</span><span class="sxs-lookup"><span data-stu-id="dd864-977">String</span></span>|<span data-ttu-id="dd864-978">Имя элемента правила `ItemHasRegularExpressionMatch`, определяющее соответствующий фильтр.</span><span class="sxs-lookup"><span data-stu-id="dd864-978">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="dd864-979">Требования</span><span class="sxs-lookup"><span data-stu-id="dd864-979">Requirements</span></span>

|<span data-ttu-id="dd864-980">Требование</span><span class="sxs-lookup"><span data-stu-id="dd864-980">Requirement</span></span>| <span data-ttu-id="dd864-981">Значение</span><span class="sxs-lookup"><span data-stu-id="dd864-981">Value</span></span>|
|---|---|
|[<span data-ttu-id="dd864-982">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="dd864-982">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dd864-983">1.0</span><span class="sxs-lookup"><span data-stu-id="dd864-983">1.0</span></span>|
|[<span data-ttu-id="dd864-984">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="dd864-984">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dd864-985">ReadItem</span><span class="sxs-lookup"><span data-stu-id="dd864-985">ReadItem</span></span>|
|[<span data-ttu-id="dd864-986">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="dd864-986">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="dd864-987">Чтение</span><span class="sxs-lookup"><span data-stu-id="dd864-987">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="dd864-988">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="dd864-988">Returns:</span></span>

<span data-ttu-id="dd864-989">Массив строк, соответствующих регулярному выражению, определяемому в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="dd864-989">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<span data-ttu-id="dd864-990">Тип: Array.< String ></span><span class="sxs-lookup"><span data-stu-id="dd864-990">Type: Array.< String ></span></span>

##### <a name="example"></a><span data-ttu-id="dd864-991">Пример</span><span class="sxs-lookup"><span data-stu-id="dd864-991">Example</span></span>

```js
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

<br>

---
---

#### <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="dd864-992">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="dd864-992">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="dd864-993">Асинхронно возвращает данные, выбранные в теме или тексте сообщения.</span><span class="sxs-lookup"><span data-stu-id="dd864-993">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="dd864-p167">Если выделенный фрагмент отсутствует, но курсор находится в тексте или теме, метод возвращает пустую строку для выбранных данных. Если выбраны не текст и не тема, метод возвращает ошибку `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="dd864-p167">If there is no selection but the cursor is in the body or subject, the method returns an empty string for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="dd864-996">Параметры</span><span class="sxs-lookup"><span data-stu-id="dd864-996">Parameters</span></span>

|<span data-ttu-id="dd864-997">Имя</span><span class="sxs-lookup"><span data-stu-id="dd864-997">Name</span></span>| <span data-ttu-id="dd864-998">Тип</span><span class="sxs-lookup"><span data-stu-id="dd864-998">Type</span></span>| <span data-ttu-id="dd864-999">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="dd864-999">Attributes</span></span>| <span data-ttu-id="dd864-1000">Описание</span><span class="sxs-lookup"><span data-stu-id="dd864-1000">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="dd864-1001">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="dd864-1001">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="dd864-p168">Запрашивает формат данных. Если задано значение Text, метод возвращает обычный текст как строку, удаляя все имеющиеся HTML-теги. Если задано значение HTML, метод возвращает выделенный текст (обычный текст или HTML).</span><span class="sxs-lookup"><span data-stu-id="dd864-p168">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="dd864-1005">Object</span><span class="sxs-lookup"><span data-stu-id="dd864-1005">Object</span></span>| <span data-ttu-id="dd864-1006">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="dd864-1006">&lt;optional&gt;</span></span>|<span data-ttu-id="dd864-1007">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="dd864-1007">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="dd864-1008">Object</span><span class="sxs-lookup"><span data-stu-id="dd864-1008">Object</span></span>| <span data-ttu-id="dd864-1009">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="dd864-1009">&lt;optional&gt;</span></span>|<span data-ttu-id="dd864-1010">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="dd864-1010">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="dd864-1011">функция</span><span class="sxs-lookup"><span data-stu-id="dd864-1011">function</span></span>||<span data-ttu-id="dd864-1012">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="dd864-1012">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="dd864-1013">Чтобы получить доступ к выбранным данным из метода обратного вызова, вызовите `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="dd864-1013">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="dd864-1014">Для доступа к исходному свойству, представляющему собой источник выбранных данных, вызовите параметр `asyncResult.value.sourceProperty`, который может иметь значение `body` или `subject`.</span><span class="sxs-lookup"><span data-stu-id="dd864-1014">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="dd864-1015">Требования</span><span class="sxs-lookup"><span data-stu-id="dd864-1015">Requirements</span></span>

|<span data-ttu-id="dd864-1016">Требование</span><span class="sxs-lookup"><span data-stu-id="dd864-1016">Requirement</span></span>| <span data-ttu-id="dd864-1017">Значение</span><span class="sxs-lookup"><span data-stu-id="dd864-1017">Value</span></span>|
|---|---|
|[<span data-ttu-id="dd864-1018">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="dd864-1018">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dd864-1019">1.2</span><span class="sxs-lookup"><span data-stu-id="dd864-1019">1.2</span></span>|
|[<span data-ttu-id="dd864-1020">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="dd864-1020">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dd864-1021">ReadItem</span><span class="sxs-lookup"><span data-stu-id="dd864-1021">ReadItem</span></span>|
|[<span data-ttu-id="dd864-1022">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="dd864-1022">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="dd864-1023">Создание</span><span class="sxs-lookup"><span data-stu-id="dd864-1023">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="dd864-1024">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="dd864-1024">Returns:</span></span>

<span data-ttu-id="dd864-1025">Выбранные данные в виде строки с форматом, определенным в параметре `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="dd864-1025">The selected data as a string with format determined by `coercionType`.</span></span>

<span data-ttu-id="dd864-1026">Тип: строка</span><span class="sxs-lookup"><span data-stu-id="dd864-1026">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="dd864-1027">Пример</span><span class="sxs-lookup"><span data-stu-id="dd864-1027">Example</span></span>

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

#### <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="dd864-1028">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="dd864-1028">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="dd864-1029">Асинхронно загружает настраиваемые свойства для надстройки для выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="dd864-1029">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="dd864-p170">Настраиваемые свойства сохраняются в виде пар "ключ-значение" для каждого приложения и каждого элемента. Этот метод возвращает объект `CustomProperties` при обратном вызове, который предоставляет методы для доступа к настраиваемым свойствам, характерным для текущего элемента и текущей надстройки. Настраиваемые свойства не шифруются для элемента, поэтому этот способ хранения не является безопасным.</span><span class="sxs-lookup"><span data-stu-id="dd864-p170">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="dd864-1033">Параметры</span><span class="sxs-lookup"><span data-stu-id="dd864-1033">Parameters</span></span>

|<span data-ttu-id="dd864-1034">Имя</span><span class="sxs-lookup"><span data-stu-id="dd864-1034">Name</span></span>| <span data-ttu-id="dd864-1035">Тип</span><span class="sxs-lookup"><span data-stu-id="dd864-1035">Type</span></span>| <span data-ttu-id="dd864-1036">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="dd864-1036">Attributes</span></span>| <span data-ttu-id="dd864-1037">Описание</span><span class="sxs-lookup"><span data-stu-id="dd864-1037">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="dd864-1038">function</span><span class="sxs-lookup"><span data-stu-id="dd864-1038">function</span></span>||<span data-ttu-id="dd864-1039">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="dd864-1039">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="dd864-1040">Настраиваемые свойства предоставляются в виде объекта [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.5) в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="dd864-1040">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.5) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="dd864-1041">Этот объект позволяет получить, задать и удалить настраиваемые свойства для элемента, а также сохранить изменения, внесенные в набор настраиваемых свойств, на сервере.</span><span class="sxs-lookup"><span data-stu-id="dd864-1041">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="dd864-1042">Объект</span><span class="sxs-lookup"><span data-stu-id="dd864-1042">Object</span></span>| <span data-ttu-id="dd864-1043">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="dd864-1043">&lt;optional&gt;</span></span>|<span data-ttu-id="dd864-1044">Разработчики могут указать любой объект, к которому необходимо получить доступ, в функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="dd864-1044">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="dd864-1045">Доступ к этому объекту можно получить с помощью свойства `asyncResult.asyncContext` в функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="dd864-1045">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="dd864-1046">Требования</span><span class="sxs-lookup"><span data-stu-id="dd864-1046">Requirements</span></span>

|<span data-ttu-id="dd864-1047">Требование</span><span class="sxs-lookup"><span data-stu-id="dd864-1047">Requirement</span></span>| <span data-ttu-id="dd864-1048">Значение</span><span class="sxs-lookup"><span data-stu-id="dd864-1048">Value</span></span>|
|---|---|
|[<span data-ttu-id="dd864-1049">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="dd864-1049">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dd864-1050">1.0</span><span class="sxs-lookup"><span data-stu-id="dd864-1050">1.0</span></span>|
|[<span data-ttu-id="dd864-1051">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="dd864-1051">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dd864-1052">ReadItem</span><span class="sxs-lookup"><span data-stu-id="dd864-1052">ReadItem</span></span>|
|[<span data-ttu-id="dd864-1053">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="dd864-1053">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="dd864-1054">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="dd864-1054">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="dd864-1055">Пример</span><span class="sxs-lookup"><span data-stu-id="dd864-1055">Example</span></span>

<span data-ttu-id="dd864-p173">Приведенный ниже пример кода показывает, как асинхронно загружать настраиваемые свойства, характерные для текущего элемента, с помощью метода `loadCustomPropertiesAsync`. Этот пример также показывает, как сохранять эти свойства на сервере с помощью метода `CustomProperties.saveAsync`. После загрузки настраиваемых свойств в этом примере кода метод `CustomProperties.get` используется для считывания настраиваемого свойства `myProp`, метод `CustomProperties.set` — для записи настраиваемого свойства `otherProp`, а метод `saveAsync` — для сохранения настраиваемых свойств.</span><span class="sxs-lookup"><span data-stu-id="dd864-p173">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

#### <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="dd864-1059">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="dd864-1059">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="dd864-1060">Удаляет вложение из сообщения или встречи.</span><span class="sxs-lookup"><span data-stu-id="dd864-1060">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="dd864-1061">Метод `removeAttachmentAsync` удаляет из элемента вложение с указанным идентификатором.</span><span class="sxs-lookup"><span data-stu-id="dd864-1061">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="dd864-1062">Идентификатор вложения рекомендуется использовать для удаления вложения, только если оно добавлено тем же почтовым приложением в ходе текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="dd864-1062">As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session.</span></span> <span data-ttu-id="dd864-1063">В Outlook в Интернете и на мобильных устройствах идентификатор вложения действителен только в течение одного сеанса.</span><span class="sxs-lookup"><span data-stu-id="dd864-1063">In Outlook on the web and mobile devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="dd864-1064">Сеанс завершается, когда пользователь закрывает приложение или начинает создавать элемент во встроенной форме, а затем переходит из формы в отдельное окно.</span><span class="sxs-lookup"><span data-stu-id="dd864-1064">A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="dd864-1065">Параметры</span><span class="sxs-lookup"><span data-stu-id="dd864-1065">Parameters</span></span>

|<span data-ttu-id="dd864-1066">Имя</span><span class="sxs-lookup"><span data-stu-id="dd864-1066">Name</span></span>| <span data-ttu-id="dd864-1067">Тип</span><span class="sxs-lookup"><span data-stu-id="dd864-1067">Type</span></span>| <span data-ttu-id="dd864-1068">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="dd864-1068">Attributes</span></span>| <span data-ttu-id="dd864-1069">Описание</span><span class="sxs-lookup"><span data-stu-id="dd864-1069">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="dd864-1070">String</span><span class="sxs-lookup"><span data-stu-id="dd864-1070">String</span></span>||<span data-ttu-id="dd864-1071">Идентификатор удаляемого вложения.</span><span class="sxs-lookup"><span data-stu-id="dd864-1071">The identifier of the attachment to remove.</span></span>|
|`options`| <span data-ttu-id="dd864-1072">Object</span><span class="sxs-lookup"><span data-stu-id="dd864-1072">Object</span></span>| <span data-ttu-id="dd864-1073">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="dd864-1073">&lt;optional&gt;</span></span>|<span data-ttu-id="dd864-1074">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="dd864-1074">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="dd864-1075">Object</span><span class="sxs-lookup"><span data-stu-id="dd864-1075">Object</span></span>| <span data-ttu-id="dd864-1076">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="dd864-1076">&lt;optional&gt;</span></span>|<span data-ttu-id="dd864-1077">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="dd864-1077">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="dd864-1078">функция</span><span class="sxs-lookup"><span data-stu-id="dd864-1078">function</span></span>| <span data-ttu-id="dd864-1079">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="dd864-1079">&lt;optional&gt;</span></span>|<span data-ttu-id="dd864-1080">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="dd864-1080">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="dd864-1081">Если удалить вложение не удается, свойство `asyncResult.error` содержит код ошибки с указанием ее причины.</span><span class="sxs-lookup"><span data-stu-id="dd864-1081">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="dd864-1082">Ошибки</span><span class="sxs-lookup"><span data-stu-id="dd864-1082">Errors</span></span>

| <span data-ttu-id="dd864-1083">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="dd864-1083">Error code</span></span> | <span data-ttu-id="dd864-1084">Описание</span><span class="sxs-lookup"><span data-stu-id="dd864-1084">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="dd864-1085">Идентификатор вложения не существует.</span><span class="sxs-lookup"><span data-stu-id="dd864-1085">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="dd864-1086">Требования</span><span class="sxs-lookup"><span data-stu-id="dd864-1086">Requirements</span></span>

|<span data-ttu-id="dd864-1087">Требование</span><span class="sxs-lookup"><span data-stu-id="dd864-1087">Requirement</span></span>| <span data-ttu-id="dd864-1088">Значение</span><span class="sxs-lookup"><span data-stu-id="dd864-1088">Value</span></span>|
|---|---|
|[<span data-ttu-id="dd864-1089">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="dd864-1089">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dd864-1090">1.1</span><span class="sxs-lookup"><span data-stu-id="dd864-1090">1.1</span></span>|
|[<span data-ttu-id="dd864-1091">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="dd864-1091">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dd864-1092">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="dd864-1092">ReadWriteItem</span></span>|
|[<span data-ttu-id="dd864-1093">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="dd864-1093">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="dd864-1094">Создание</span><span class="sxs-lookup"><span data-stu-id="dd864-1094">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="dd864-1095">Пример</span><span class="sxs-lookup"><span data-stu-id="dd864-1095">Example</span></span>

<span data-ttu-id="dd864-1096">Указанный ниже код удаляет вложение с идентификатором "0".</span><span class="sxs-lookup"><span data-stu-id="dd864-1096">The following code removes an attachment with an identifier of '0'.</span></span>

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

#### <a name="saveasyncoptions-callback"></a><span data-ttu-id="dd864-1097">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="dd864-1097">saveAsync([options], callback)</span></span>

<span data-ttu-id="dd864-1098">Асинхронно сохраняет элемент.</span><span class="sxs-lookup"><span data-stu-id="dd864-1098">Asynchronously saves an item.</span></span>

<span data-ttu-id="dd864-1099">При вызове этот метод сохраняет текущее сообщение в виде черновика и возвращает идентификатор элемента с помощью метода обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="dd864-1099">When invoked, this method saves the current message as a draft and returns the item id via the callback method.</span></span> <span data-ttu-id="dd864-1100">В Outlook в Интернете или интерактивном режиме Outlook этот элемент сохраняется на сервере.</span><span class="sxs-lookup"><span data-stu-id="dd864-1100">In Outlook on the web or Outlook in online mode, the item is saved to the server.</span></span> <span data-ttu-id="dd864-1101">В Outlook в режиме кэширования этот элемент сохраняется в локальном кэше.</span><span class="sxs-lookup"><span data-stu-id="dd864-1101">In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="dd864-1102">Если в режиме создания надстройка вызывает для элемента метод `saveAsync`, чтобы получить параметр `itemId` для использования с EWS или REST API, необходимо помнить, что синхронизация элемента с сервером может занять много времени, если Outlook работает в режиме кэширования данных.</span><span class="sxs-lookup"><span data-stu-id="dd864-1102">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="dd864-1103">До окончания синхронизации элемента применение параметра `itemId` будет приводить к ошибке.</span><span class="sxs-lookup"><span data-stu-id="dd864-1103">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="dd864-p177">Если метод `saveAsync` вызывается для встречи в режиме создания, она сохраняется как обычная встреча в календаре пользователя, а не как черновик. При сохранении новой встречи приглашения не отправляются. При сохранении существующей встречи уведомления отправляются добавленным или удаленным участникам.</span><span class="sxs-lookup"><span data-stu-id="dd864-p177">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="dd864-1107">Следующие клиенты отличаются другим поведением `saveAsync` в отношении встреч в режиме создания:</span><span class="sxs-lookup"><span data-stu-id="dd864-1107">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="dd864-1108">Outlook для Mac не поддерживает сохранение собрания.</span><span class="sxs-lookup"><span data-stu-id="dd864-1108">Outlook on Mac does not support saving a meeting.</span></span> <span data-ttu-id="dd864-1109">Метод `saveAsync` не работает при вызове из собрания в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="dd864-1109">The `saveAsync` method fails when called from a meeting in compose mode.</span></span> <span data-ttu-id="dd864-1110">Временное решение представлено в статье [Не удается сохранить встречу как черновик в Outlook для Mac с помощью API JS для Office](https://support.microsoft.com/help/4505745).</span><span class="sxs-lookup"><span data-stu-id="dd864-1110">See [Cannot save a meeting as a draft in Outlook for Mac by using Office JS API](https://support.microsoft.com/help/4505745) for a workaround.</span></span>
> - <span data-ttu-id="dd864-1111">Outlook в Интернете всегда отправляет приглашение или обновление при вызове `saveAsync` для встречи в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="dd864-1111">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="dd864-1112">Параметры</span><span class="sxs-lookup"><span data-stu-id="dd864-1112">Parameters</span></span>

|<span data-ttu-id="dd864-1113">Имя</span><span class="sxs-lookup"><span data-stu-id="dd864-1113">Name</span></span>| <span data-ttu-id="dd864-1114">Тип</span><span class="sxs-lookup"><span data-stu-id="dd864-1114">Type</span></span>| <span data-ttu-id="dd864-1115">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="dd864-1115">Attributes</span></span>| <span data-ttu-id="dd864-1116">Описание</span><span class="sxs-lookup"><span data-stu-id="dd864-1116">Description</span></span>|
|---|---|---|---|
|`options`| <span data-ttu-id="dd864-1117">Object</span><span class="sxs-lookup"><span data-stu-id="dd864-1117">Object</span></span>| <span data-ttu-id="dd864-1118">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="dd864-1118">&lt;optional&gt;</span></span>|<span data-ttu-id="dd864-1119">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="dd864-1119">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="dd864-1120">Object</span><span class="sxs-lookup"><span data-stu-id="dd864-1120">Object</span></span>| <span data-ttu-id="dd864-1121">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="dd864-1121">&lt;optional&gt;</span></span>|<span data-ttu-id="dd864-1122">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="dd864-1122">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="dd864-1123">функция</span><span class="sxs-lookup"><span data-stu-id="dd864-1123">function</span></span>||<span data-ttu-id="dd864-1124">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="dd864-1124">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="dd864-1125">После успешного выполнения идентификатор элемента будет предоставлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="dd864-1125">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="dd864-1126">Требования</span><span class="sxs-lookup"><span data-stu-id="dd864-1126">Requirements</span></span>

|<span data-ttu-id="dd864-1127">Требование</span><span class="sxs-lookup"><span data-stu-id="dd864-1127">Requirement</span></span>| <span data-ttu-id="dd864-1128">Значение</span><span class="sxs-lookup"><span data-stu-id="dd864-1128">Value</span></span>|
|---|---|
|[<span data-ttu-id="dd864-1129">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="dd864-1129">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dd864-1130">1.3</span><span class="sxs-lookup"><span data-stu-id="dd864-1130">1.3</span></span>|
|[<span data-ttu-id="dd864-1131">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="dd864-1131">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dd864-1132">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="dd864-1132">ReadWriteItem</span></span>|
|[<span data-ttu-id="dd864-1133">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="dd864-1133">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="dd864-1134">Создание</span><span class="sxs-lookup"><span data-stu-id="dd864-1134">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="dd864-1135">Примеры</span><span class="sxs-lookup"><span data-stu-id="dd864-1135">Examples</span></span>

```js
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result.
  });
```

<span data-ttu-id="dd864-p179">Ниже приведен пример параметра `result`, переданного функции обратного вызова. Свойство `value` содержит идентификатор элемента.</span><span class="sxs-lookup"><span data-stu-id="dd864-p179">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

<br>

---
---

#### <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="dd864-1138">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="dd864-1138">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="dd864-1139">Асинхронно вставляет данные в текст или тему сообщения.</span><span class="sxs-lookup"><span data-stu-id="dd864-1139">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="dd864-p180">Метод `setSelectedDataAsync` вставляет указанную строку в местоположение курсора в теме или тексте элемента либо, если текст выделен в редакторе, он заменяет выделенный текст. Если курсор находится вне текста или темы элемента, возвращается ошибка. После вставки курсор помещается в конец вставленного содержимого.</span><span class="sxs-lookup"><span data-stu-id="dd864-p180">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="dd864-1143">Параметры</span><span class="sxs-lookup"><span data-stu-id="dd864-1143">Parameters</span></span>

|<span data-ttu-id="dd864-1144">Имя</span><span class="sxs-lookup"><span data-stu-id="dd864-1144">Name</span></span>| <span data-ttu-id="dd864-1145">Тип</span><span class="sxs-lookup"><span data-stu-id="dd864-1145">Type</span></span>| <span data-ttu-id="dd864-1146">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="dd864-1146">Attributes</span></span>| <span data-ttu-id="dd864-1147">Описание</span><span class="sxs-lookup"><span data-stu-id="dd864-1147">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="dd864-1148">String</span><span class="sxs-lookup"><span data-stu-id="dd864-1148">String</span></span>||<span data-ttu-id="dd864-p181">Вставляемые данные. Объем данных не должен превышать 1 000 000 символов. Если передано больше 1 000 000 символов, возвращается исключение `ArgumentOutOfRange`.</span><span class="sxs-lookup"><span data-stu-id="dd864-p181">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="dd864-1152">Object</span><span class="sxs-lookup"><span data-stu-id="dd864-1152">Object</span></span>| <span data-ttu-id="dd864-1153">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="dd864-1153">&lt;optional&gt;</span></span>|<span data-ttu-id="dd864-1154">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="dd864-1154">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="dd864-1155">Object</span><span class="sxs-lookup"><span data-stu-id="dd864-1155">Object</span></span>| <span data-ttu-id="dd864-1156">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="dd864-1156">&lt;optional&gt;</span></span>|<span data-ttu-id="dd864-1157">В методе обратного вызова разработчики могут указать любой объект, к которому необходимо получить доступ.</span><span class="sxs-lookup"><span data-stu-id="dd864-1157">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="dd864-1158">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="dd864-1158">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="dd864-1159">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="dd864-1159">&lt;optional&gt;</span></span>|<span data-ttu-id="dd864-1160">Если задано значение `text`, текущий стиль применяется в Outlook в Интернете и классических клиентах.</span><span class="sxs-lookup"><span data-stu-id="dd864-1160">If `text`, the current style is applied in Outlook on the web and desktop clients.</span></span> <span data-ttu-id="dd864-1161">Если поле представляет собой редактор HTML, вставляются только текстовые данные, даже если они имеют формат HTML.</span><span class="sxs-lookup"><span data-stu-id="dd864-1161">If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="dd864-1162">Если задано значение `html` и поле (не тема) поддерживает HTML, в Outlook в Интернете применяется текущий стиль, а в классических клиентах Outlook — стиль по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="dd864-1162">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook on the web and the default style is applied in Outlook desktop clients.</span></span> <span data-ttu-id="dd864-1163">Если поле является текстовым, возвращается ошибка `InvalidDataFormat`.</span><span class="sxs-lookup"><span data-stu-id="dd864-1163">If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="dd864-1164">Если свойство `coercionType` не задано, результат зависит от поля: если поле имеет формат HTML, используется текст в формате HTML, а если поле текстовое, применяется обычный текст.</span><span class="sxs-lookup"><span data-stu-id="dd864-1164">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="dd864-1165">функция</span><span class="sxs-lookup"><span data-stu-id="dd864-1165">function</span></span>||<span data-ttu-id="dd864-1166">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="dd864-1166">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="dd864-1167">Требования</span><span class="sxs-lookup"><span data-stu-id="dd864-1167">Requirements</span></span>

|<span data-ttu-id="dd864-1168">Требование</span><span class="sxs-lookup"><span data-stu-id="dd864-1168">Requirement</span></span>| <span data-ttu-id="dd864-1169">Значение</span><span class="sxs-lookup"><span data-stu-id="dd864-1169">Value</span></span>|
|---|---|
|[<span data-ttu-id="dd864-1170">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="dd864-1170">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dd864-1171">1.2</span><span class="sxs-lookup"><span data-stu-id="dd864-1171">1.2</span></span>|
|[<span data-ttu-id="dd864-1172">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="dd864-1172">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dd864-1173">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="dd864-1173">ReadWriteItem</span></span>|
|[<span data-ttu-id="dd864-1174">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="dd864-1174">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="dd864-1175">Создание</span><span class="sxs-lookup"><span data-stu-id="dd864-1175">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="dd864-1176">Пример</span><span class="sxs-lookup"><span data-stu-id="dd864-1176">Example</span></span>

```js
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
