---
title: Office. Context. Mailbox. Item — набор требований 1,2
description: ''
ms.date: 11/25/2019
localization_priority: Normal
ms.openlocfilehash: ab8c55d2f91b250b419c7c9c71fc044b6fa68279
ms.sourcegitcommit: 05a883a7fd89136301ce35aabc57638e9f563288
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 11/27/2019
ms.locfileid: "39629211"
---
# <a name="item"></a><span data-ttu-id="ad140-102">item</span><span class="sxs-lookup"><span data-stu-id="ad140-102">item</span></span>

### <span data-ttu-id="ad140-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md). item</span><span class="sxs-lookup"><span data-stu-id="ad140-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md). item</span></span>

<span data-ttu-id="ad140-p102">Пространство имен `item` используется для доступа к выбранному в данный момент сообщению, приглашению на собрание или описанию встречи. Вы можете определить тип пространства имен `item` с помощью свойства [itemType](#itemtype-officemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="ad140-p102">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="ad140-107">Requirements</span><span class="sxs-lookup"><span data-stu-id="ad140-107">Requirements</span></span>

|<span data-ttu-id="ad140-108">Требование</span><span class="sxs-lookup"><span data-stu-id="ad140-108">Requirement</span></span>| <span data-ttu-id="ad140-109">Значение</span><span class="sxs-lookup"><span data-stu-id="ad140-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="ad140-110">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="ad140-110">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ad140-111">1.0</span><span class="sxs-lookup"><span data-stu-id="ad140-111">1.0</span></span>|
|[<span data-ttu-id="ad140-112">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ad140-112">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ad140-113">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="ad140-113">Restricted</span></span>|
|[<span data-ttu-id="ad140-114">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ad140-114">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ad140-115">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="ad140-115">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="ad140-116">Элементы и методы</span><span class="sxs-lookup"><span data-stu-id="ad140-116">Members and methods</span></span>

| <span data-ttu-id="ad140-117">Элемент	</span><span class="sxs-lookup"><span data-stu-id="ad140-117">Member</span></span> | <span data-ttu-id="ad140-118">Тип</span><span class="sxs-lookup"><span data-stu-id="ad140-118">Type</span></span> |
|--------|------|
| [<span data-ttu-id="ad140-119">attachments</span><span class="sxs-lookup"><span data-stu-id="ad140-119">attachments</span></span>](#attachments-arrayattachmentdetails) | <span data-ttu-id="ad140-120">Member</span><span class="sxs-lookup"><span data-stu-id="ad140-120">Member</span></span> |
| [<span data-ttu-id="ad140-121">bcc</span><span class="sxs-lookup"><span data-stu-id="ad140-121">bcc</span></span>](#bcc-recipients) | <span data-ttu-id="ad140-122">Member</span><span class="sxs-lookup"><span data-stu-id="ad140-122">Member</span></span> |
| [<span data-ttu-id="ad140-123">body</span><span class="sxs-lookup"><span data-stu-id="ad140-123">body</span></span>](#body-body) | <span data-ttu-id="ad140-124">Member</span><span class="sxs-lookup"><span data-stu-id="ad140-124">Member</span></span> |
| [<span data-ttu-id="ad140-125">cc</span><span class="sxs-lookup"><span data-stu-id="ad140-125">cc</span></span>](#cc-arrayemailaddressdetailsrecipients) | <span data-ttu-id="ad140-126">Member</span><span class="sxs-lookup"><span data-stu-id="ad140-126">Member</span></span> |
| [<span data-ttu-id="ad140-127">conversationId</span><span class="sxs-lookup"><span data-stu-id="ad140-127">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="ad140-128">Member</span><span class="sxs-lookup"><span data-stu-id="ad140-128">Member</span></span> |
| [<span data-ttu-id="ad140-129">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="ad140-129">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="ad140-130">Member</span><span class="sxs-lookup"><span data-stu-id="ad140-130">Member</span></span> |
| [<span data-ttu-id="ad140-131">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="ad140-131">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="ad140-132">Member</span><span class="sxs-lookup"><span data-stu-id="ad140-132">Member</span></span> |
| [<span data-ttu-id="ad140-133">end</span><span class="sxs-lookup"><span data-stu-id="ad140-133">end</span></span>](#end-datetime) | <span data-ttu-id="ad140-134">Member</span><span class="sxs-lookup"><span data-stu-id="ad140-134">Member</span></span> |
| [<span data-ttu-id="ad140-135">from</span><span class="sxs-lookup"><span data-stu-id="ad140-135">from</span></span>](#from-emailaddressdetails) | <span data-ttu-id="ad140-136">Member</span><span class="sxs-lookup"><span data-stu-id="ad140-136">Member</span></span> |
| [<span data-ttu-id="ad140-137">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="ad140-137">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="ad140-138">Member</span><span class="sxs-lookup"><span data-stu-id="ad140-138">Member</span></span> |
| [<span data-ttu-id="ad140-139">itemClass</span><span class="sxs-lookup"><span data-stu-id="ad140-139">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="ad140-140">Member</span><span class="sxs-lookup"><span data-stu-id="ad140-140">Member</span></span> |
| [<span data-ttu-id="ad140-141">itemId</span><span class="sxs-lookup"><span data-stu-id="ad140-141">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="ad140-142">Member</span><span class="sxs-lookup"><span data-stu-id="ad140-142">Member</span></span> |
| [<span data-ttu-id="ad140-143">itemType</span><span class="sxs-lookup"><span data-stu-id="ad140-143">itemType</span></span>](#itemtype-officemailboxenumsitemtype) | <span data-ttu-id="ad140-144">Member</span><span class="sxs-lookup"><span data-stu-id="ad140-144">Member</span></span> |
| [<span data-ttu-id="ad140-145">location</span><span class="sxs-lookup"><span data-stu-id="ad140-145">location</span></span>](#location-stringlocation) | <span data-ttu-id="ad140-146">Member</span><span class="sxs-lookup"><span data-stu-id="ad140-146">Member</span></span> |
| [<span data-ttu-id="ad140-147">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="ad140-147">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="ad140-148">Member</span><span class="sxs-lookup"><span data-stu-id="ad140-148">Member</span></span> |
| [<span data-ttu-id="ad140-149">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="ad140-149">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="ad140-150">Member</span><span class="sxs-lookup"><span data-stu-id="ad140-150">Member</span></span> |
| [<span data-ttu-id="ad140-151">organizer</span><span class="sxs-lookup"><span data-stu-id="ad140-151">organizer</span></span>](#organizer-emailaddressdetails) | <span data-ttu-id="ad140-152">Member</span><span class="sxs-lookup"><span data-stu-id="ad140-152">Member</span></span> |
| [<span data-ttu-id="ad140-153">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="ad140-153">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="ad140-154">Member</span><span class="sxs-lookup"><span data-stu-id="ad140-154">Member</span></span> |
| [<span data-ttu-id="ad140-155">sender</span><span class="sxs-lookup"><span data-stu-id="ad140-155">sender</span></span>](#sender-emailaddressdetails) | <span data-ttu-id="ad140-156">Member</span><span class="sxs-lookup"><span data-stu-id="ad140-156">Member</span></span> |
| [<span data-ttu-id="ad140-157">start</span><span class="sxs-lookup"><span data-stu-id="ad140-157">start</span></span>](#start-datetime) | <span data-ttu-id="ad140-158">Member</span><span class="sxs-lookup"><span data-stu-id="ad140-158">Member</span></span> |
| [<span data-ttu-id="ad140-159">subject</span><span class="sxs-lookup"><span data-stu-id="ad140-159">subject</span></span>](#subject-stringsubject) | <span data-ttu-id="ad140-160">Элемент</span><span class="sxs-lookup"><span data-stu-id="ad140-160">Member</span></span> |
| [<span data-ttu-id="ad140-161">to</span><span class="sxs-lookup"><span data-stu-id="ad140-161">to</span></span>](#to-arrayemailaddressdetailsrecipients) | <span data-ttu-id="ad140-162">Элемент</span><span class="sxs-lookup"><span data-stu-id="ad140-162">Member</span></span> |
| [<span data-ttu-id="ad140-163">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="ad140-163">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="ad140-164">Метод</span><span class="sxs-lookup"><span data-stu-id="ad140-164">Method</span></span> |
| [<span data-ttu-id="ad140-165">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="ad140-165">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="ad140-166">Метод</span><span class="sxs-lookup"><span data-stu-id="ad140-166">Method</span></span> |
| [<span data-ttu-id="ad140-167">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="ad140-167">displayReplyAllForm</span></span>](#displayreplyallformformdata-callback) | <span data-ttu-id="ad140-168">Метод</span><span class="sxs-lookup"><span data-stu-id="ad140-168">Method</span></span> |
| [<span data-ttu-id="ad140-169">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="ad140-169">displayReplyForm</span></span>](#displayreplyformformdata-callback) | <span data-ttu-id="ad140-170">Метод</span><span class="sxs-lookup"><span data-stu-id="ad140-170">Method</span></span> |
| [<span data-ttu-id="ad140-171">getEntities</span><span class="sxs-lookup"><span data-stu-id="ad140-171">getEntities</span></span>](#getentities--entities) | <span data-ttu-id="ad140-172">Метод</span><span class="sxs-lookup"><span data-stu-id="ad140-172">Method</span></span> |
| [<span data-ttu-id="ad140-173">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="ad140-173">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="ad140-174">Метод</span><span class="sxs-lookup"><span data-stu-id="ad140-174">Method</span></span> |
| [<span data-ttu-id="ad140-175">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="ad140-175">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="ad140-176">Метод</span><span class="sxs-lookup"><span data-stu-id="ad140-176">Method</span></span> |
| [<span data-ttu-id="ad140-177">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="ad140-177">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="ad140-178">Метод</span><span class="sxs-lookup"><span data-stu-id="ad140-178">Method</span></span> |
| [<span data-ttu-id="ad140-179">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="ad140-179">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="ad140-180">Метод</span><span class="sxs-lookup"><span data-stu-id="ad140-180">Method</span></span> |
| [<span data-ttu-id="ad140-181">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="ad140-181">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="ad140-182">Метод</span><span class="sxs-lookup"><span data-stu-id="ad140-182">Method</span></span> |
| [<span data-ttu-id="ad140-183">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="ad140-183">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="ad140-184">Метод</span><span class="sxs-lookup"><span data-stu-id="ad140-184">Method</span></span> |
| [<span data-ttu-id="ad140-185">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="ad140-185">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="ad140-186">Метод</span><span class="sxs-lookup"><span data-stu-id="ad140-186">Method</span></span> |
| [<span data-ttu-id="ad140-187">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="ad140-187">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="ad140-188">Метод</span><span class="sxs-lookup"><span data-stu-id="ad140-188">Method</span></span> |

### <a name="example"></a><span data-ttu-id="ad140-189">Пример</span><span class="sxs-lookup"><span data-stu-id="ad140-189">Example</span></span>

<span data-ttu-id="ad140-190">В примере кода JavaScript, приведенном ниже, показано, как получить доступ к свойству `subject` текущего элемента в Outlook.</span><span class="sxs-lookup"><span data-stu-id="ad140-190">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="ad140-191">Members</span><span class="sxs-lookup"><span data-stu-id="ad140-191">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetailsviewoutlook-js-12"></a><span data-ttu-id="ad140-192">attachments: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.2)></span><span class="sxs-lookup"><span data-stu-id="ad140-192">attachments: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.2)></span></span>

<span data-ttu-id="ad140-p103">Получает массив вложений для элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="ad140-p103">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="ad140-195">Outlook блокирует определенные типы файлов из-за потенциальных проблем с безопасностью, поэтому они не возвращаются.</span><span class="sxs-lookup"><span data-stu-id="ad140-195">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="ad140-196">Дополнительные сведения см. в статье [Блокировка вложений в Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="ad140-196">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="ad140-197">Тип</span><span class="sxs-lookup"><span data-stu-id="ad140-197">Type</span></span>

*   <span data-ttu-id="ad140-198">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.2)></span><span class="sxs-lookup"><span data-stu-id="ad140-198">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.2)></span></span>

##### <a name="requirements"></a><span data-ttu-id="ad140-199">Требования</span><span class="sxs-lookup"><span data-stu-id="ad140-199">Requirements</span></span>

|<span data-ttu-id="ad140-200">Требование</span><span class="sxs-lookup"><span data-stu-id="ad140-200">Requirement</span></span>| <span data-ttu-id="ad140-201">Значение</span><span class="sxs-lookup"><span data-stu-id="ad140-201">Value</span></span>|
|---|---|
|[<span data-ttu-id="ad140-202">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="ad140-202">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ad140-203">1.0</span><span class="sxs-lookup"><span data-stu-id="ad140-203">1.0</span></span>|
|[<span data-ttu-id="ad140-204">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ad140-204">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ad140-205">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ad140-205">ReadItem</span></span>|
|[<span data-ttu-id="ad140-206">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ad140-206">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ad140-207">Чтение</span><span class="sxs-lookup"><span data-stu-id="ad140-207">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ad140-208">Пример</span><span class="sxs-lookup"><span data-stu-id="ad140-208">Example</span></span>

<span data-ttu-id="ad140-209">С помощью приведенного ниже кода можно создать HTML-строку с подробными сведениями обо всех вложениях для текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="ad140-209">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

#### <a name="bcc-recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-12"></a><span data-ttu-id="ad140-210">bcc: [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="ad140-210">bcc: [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

<span data-ttu-id="ad140-211">Получает объект, который предоставляет методы для получения или обновления получателей скрытой копии сообщения.</span><span class="sxs-lookup"><span data-stu-id="ad140-211">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="ad140-212">Только в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="ad140-212">Compose mode only.</span></span>

<span data-ttu-id="ad140-213">Коллекция может включать не более 100 элементов по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="ad140-213">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="ad140-214">Однако для компьютеров с Windows и Mac OS действуют указанные ниже ограничения.</span><span class="sxs-lookup"><span data-stu-id="ad140-214">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="ad140-215">Максимальное количество элементов — 500.</span><span class="sxs-lookup"><span data-stu-id="ad140-215">Get 500 members maximum.</span></span>
- <span data-ttu-id="ad140-216">Установить ограничение количества элементов на вызов — не более 100, общего количества — не более 500.</span><span class="sxs-lookup"><span data-stu-id="ad140-216">Set a maximum of 100 members per call, up to 500 members total.</span></span>

##### <a name="type"></a><span data-ttu-id="ad140-217">Тип</span><span class="sxs-lookup"><span data-stu-id="ad140-217">Type</span></span>

*   [<span data-ttu-id="ad140-218">Получатели</span><span class="sxs-lookup"><span data-stu-id="ad140-218">Recipients</span></span>](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)

##### <a name="requirements"></a><span data-ttu-id="ad140-219">Requirements</span><span class="sxs-lookup"><span data-stu-id="ad140-219">Requirements</span></span>

|<span data-ttu-id="ad140-220">Требование</span><span class="sxs-lookup"><span data-stu-id="ad140-220">Requirement</span></span>| <span data-ttu-id="ad140-221">Значение</span><span class="sxs-lookup"><span data-stu-id="ad140-221">Value</span></span>|
|---|---|
|[<span data-ttu-id="ad140-222">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="ad140-222">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ad140-223">1.1</span><span class="sxs-lookup"><span data-stu-id="ad140-223">1.1</span></span>|
|[<span data-ttu-id="ad140-224">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ad140-224">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ad140-225">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ad140-225">ReadItem</span></span>|
|[<span data-ttu-id="ad140-226">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ad140-226">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ad140-227">Создание</span><span class="sxs-lookup"><span data-stu-id="ad140-227">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="ad140-228">Пример</span><span class="sxs-lookup"><span data-stu-id="ad140-228">Example</span></span>

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

#### <a name="body-bodyjavascriptapioutlookofficebodyviewoutlook-js-12"></a><span data-ttu-id="ad140-229">body: [Body](/javascript/api/outlook/office.body?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="ad140-229">body: [Body](/javascript/api/outlook/office.body?view=outlook-js-1.2)</span></span>

<span data-ttu-id="ad140-230">Получает объект, предоставляющий методы для работы с основным текстом элемента.</span><span class="sxs-lookup"><span data-stu-id="ad140-230">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="ad140-231">Тип</span><span class="sxs-lookup"><span data-stu-id="ad140-231">Type</span></span>

*   [<span data-ttu-id="ad140-232">Body</span><span class="sxs-lookup"><span data-stu-id="ad140-232">Body</span></span>](/javascript/api/outlook/office.body?view=outlook-js-1.2)

##### <a name="requirements"></a><span data-ttu-id="ad140-233">Requirements</span><span class="sxs-lookup"><span data-stu-id="ad140-233">Requirements</span></span>

|<span data-ttu-id="ad140-234">Требование</span><span class="sxs-lookup"><span data-stu-id="ad140-234">Requirement</span></span>| <span data-ttu-id="ad140-235">Значение</span><span class="sxs-lookup"><span data-stu-id="ad140-235">Value</span></span>|
|---|---|
|[<span data-ttu-id="ad140-236">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="ad140-236">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ad140-237">1.1</span><span class="sxs-lookup"><span data-stu-id="ad140-237">1.1</span></span>|
|[<span data-ttu-id="ad140-238">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ad140-238">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ad140-239">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ad140-239">ReadItem</span></span>|
|[<span data-ttu-id="ad140-240">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ad140-240">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ad140-241">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="ad140-241">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ad140-242">Пример</span><span class="sxs-lookup"><span data-stu-id="ad140-242">Example</span></span>

<span data-ttu-id="ad140-243">В этом примере возвращается текст сообщения в формате обычного текста.</span><span class="sxs-lookup"><span data-stu-id="ad140-243">This example gets the body of the message in plain text.</span></span>

```js
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="ad140-244">Ниже приведен пример итогового параметра, переданного функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="ad140-244">The following is an example of the result parameter passed to the callback function.</span></span>

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

#### <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-12recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-12"></a><span data-ttu-id="ad140-245">cc: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="ad140-245">cc: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

<span data-ttu-id="ad140-246">Предоставляет доступ к получателям копии сообщения.</span><span class="sxs-lookup"><span data-stu-id="ad140-246">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="ad140-247">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="ad140-247">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="ad140-248">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="ad140-248">Read mode</span></span>

<span data-ttu-id="ad140-249">Свойство `cc` возвращает массив, который содержит объект `EmailAddressDetails` для каждого получателя, указанного в строке **Копия** сообщения.</span><span class="sxs-lookup"><span data-stu-id="ad140-249">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message.</span></span> <span data-ttu-id="ad140-250">Коллекция может включать не более 100 элементов по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="ad140-250">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="ad140-251">Однако на компьютерах с Windows и Mac OS может быть до 500 элементов.</span><span class="sxs-lookup"><span data-stu-id="ad140-251">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="ad140-252">Режим создания</span><span class="sxs-lookup"><span data-stu-id="ad140-252">Compose mode</span></span>

<span data-ttu-id="ad140-253">Свойство `cc` возвращает объект `Recipients`, предоставляющий методы для получения или обновления получателей, которые указаны в строке **Копия** сообщения.</span><span class="sxs-lookup"><span data-stu-id="ad140-253">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span> <span data-ttu-id="ad140-254">Коллекция может включать не более 100 элементов по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="ad140-254">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="ad140-255">Однако для компьютеров с Windows и Mac OS действуют указанные ниже ограничения.</span><span class="sxs-lookup"><span data-stu-id="ad140-255">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="ad140-256">Максимальное количество элементов — 500.</span><span class="sxs-lookup"><span data-stu-id="ad140-256">Get 500 members maximum.</span></span>
- <span data-ttu-id="ad140-257">Установить ограничение количества элементов на вызов — не более 100, общего количества — не более 500.</span><span class="sxs-lookup"><span data-stu-id="ad140-257">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="ad140-258">Тип</span><span class="sxs-lookup"><span data-stu-id="ad140-258">Type</span></span>

*   <span data-ttu-id="ad140-259">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="ad140-259">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

##### <a name="requirements"></a><span data-ttu-id="ad140-260">Requirements</span><span class="sxs-lookup"><span data-stu-id="ad140-260">Requirements</span></span>

|<span data-ttu-id="ad140-261">Требование</span><span class="sxs-lookup"><span data-stu-id="ad140-261">Requirement</span></span>| <span data-ttu-id="ad140-262">Значение</span><span class="sxs-lookup"><span data-stu-id="ad140-262">Value</span></span>|
|---|---|
|[<span data-ttu-id="ad140-263">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="ad140-263">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ad140-264">1.0</span><span class="sxs-lookup"><span data-stu-id="ad140-264">1.0</span></span>|
|[<span data-ttu-id="ad140-265">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ad140-265">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ad140-266">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ad140-266">ReadItem</span></span>|
|[<span data-ttu-id="ad140-267">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ad140-267">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ad140-268">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="ad140-268">Compose or Read</span></span>|

<br>

---
---

#### <a name="nullable-conversationid-string"></a><span data-ttu-id="ad140-269">(nullable) conversationId: String</span><span class="sxs-lookup"><span data-stu-id="ad140-269">(nullable) conversationId: String</span></span>

<span data-ttu-id="ad140-270">Получает идентификатор разговора по электронной почте, содержащего конкретное сообщение.</span><span class="sxs-lookup"><span data-stu-id="ad140-270">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="ad140-p110">Вы можете получить целочисленное значение этого свойства, если ваше почтовое приложение активируется в формах просмотра или формах создания ответов. Если пользователь изменит тему ответа, после его отправки идентификатор беседы будет изменен, и полученное ранее значение будет недействительным.</span><span class="sxs-lookup"><span data-stu-id="ad140-p110">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="ad140-p111">Это свойство имеет значение NULL для нового элемента в форме создания. Свойство `conversationId` вернет значение, если пользователь задаст тему и сохранит элемент.</span><span class="sxs-lookup"><span data-stu-id="ad140-p111">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="ad140-275">Тип</span><span class="sxs-lookup"><span data-stu-id="ad140-275">Type</span></span>

*   <span data-ttu-id="ad140-276">String</span><span class="sxs-lookup"><span data-stu-id="ad140-276">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="ad140-277">Requirements</span><span class="sxs-lookup"><span data-stu-id="ad140-277">Requirements</span></span>

|<span data-ttu-id="ad140-278">Требование</span><span class="sxs-lookup"><span data-stu-id="ad140-278">Requirement</span></span>| <span data-ttu-id="ad140-279">Значение</span><span class="sxs-lookup"><span data-stu-id="ad140-279">Value</span></span>|
|---|---|
|[<span data-ttu-id="ad140-280">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="ad140-280">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ad140-281">1.0</span><span class="sxs-lookup"><span data-stu-id="ad140-281">1.0</span></span>|
|[<span data-ttu-id="ad140-282">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ad140-282">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ad140-283">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ad140-283">ReadItem</span></span>|
|[<span data-ttu-id="ad140-284">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ad140-284">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ad140-285">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="ad140-285">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ad140-286">Пример</span><span class="sxs-lookup"><span data-stu-id="ad140-286">Example</span></span>

```js
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

<br>

---
---

#### <a name="datetimecreated-date"></a><span data-ttu-id="ad140-287">dateTimeCreated: Date</span><span class="sxs-lookup"><span data-stu-id="ad140-287">dateTimeCreated: Date</span></span>

<span data-ttu-id="ad140-p112">Получает дату и время создания элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="ad140-p112">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="ad140-290">Тип</span><span class="sxs-lookup"><span data-stu-id="ad140-290">Type</span></span>

*   <span data-ttu-id="ad140-291">Дата</span><span class="sxs-lookup"><span data-stu-id="ad140-291">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="ad140-292">Требования</span><span class="sxs-lookup"><span data-stu-id="ad140-292">Requirements</span></span>

|<span data-ttu-id="ad140-293">Требование</span><span class="sxs-lookup"><span data-stu-id="ad140-293">Requirement</span></span>| <span data-ttu-id="ad140-294">Значение</span><span class="sxs-lookup"><span data-stu-id="ad140-294">Value</span></span>|
|---|---|
|[<span data-ttu-id="ad140-295">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="ad140-295">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ad140-296">1.0</span><span class="sxs-lookup"><span data-stu-id="ad140-296">1.0</span></span>|
|[<span data-ttu-id="ad140-297">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ad140-297">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ad140-298">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ad140-298">ReadItem</span></span>|
|[<span data-ttu-id="ad140-299">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ad140-299">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ad140-300">Чтение</span><span class="sxs-lookup"><span data-stu-id="ad140-300">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ad140-301">Пример</span><span class="sxs-lookup"><span data-stu-id="ad140-301">Example</span></span>

```js
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

<br>

---
---

#### <a name="datetimemodified-date"></a><span data-ttu-id="ad140-302">dateTimeModified: Date</span><span class="sxs-lookup"><span data-stu-id="ad140-302">dateTimeModified: Date</span></span>

<span data-ttu-id="ad140-p113">Получает дату и время последнего изменения элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="ad140-p113">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="ad140-305">Этот элемент не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="ad140-305">This member is not supported in Outlook on iOS or Android.</span></span>

##### <a name="type"></a><span data-ttu-id="ad140-306">Тип</span><span class="sxs-lookup"><span data-stu-id="ad140-306">Type</span></span>

*   <span data-ttu-id="ad140-307">Дата</span><span class="sxs-lookup"><span data-stu-id="ad140-307">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="ad140-308">Требования</span><span class="sxs-lookup"><span data-stu-id="ad140-308">Requirements</span></span>

|<span data-ttu-id="ad140-309">Требование</span><span class="sxs-lookup"><span data-stu-id="ad140-309">Requirement</span></span>| <span data-ttu-id="ad140-310">Значение</span><span class="sxs-lookup"><span data-stu-id="ad140-310">Value</span></span>|
|---|---|
|[<span data-ttu-id="ad140-311">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="ad140-311">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ad140-312">1.0</span><span class="sxs-lookup"><span data-stu-id="ad140-312">1.0</span></span>|
|[<span data-ttu-id="ad140-313">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ad140-313">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ad140-314">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ad140-314">ReadItem</span></span>|
|[<span data-ttu-id="ad140-315">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ad140-315">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ad140-316">Чтение</span><span class="sxs-lookup"><span data-stu-id="ad140-316">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ad140-317">Пример</span><span class="sxs-lookup"><span data-stu-id="ad140-317">Example</span></span>

```js
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

<br>

---
---

#### <a name="end-datetimejavascriptapioutlookofficetimeviewoutlook-js-12"></a><span data-ttu-id="ad140-318">end: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="ad140-318">end: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.2)</span></span>

<span data-ttu-id="ad140-319">Получает или задает дату и время окончания встречи.</span><span class="sxs-lookup"><span data-stu-id="ad140-319">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="ad140-p114">Свойство `end` представлено в виде значения даты и времени в формате UTC. Преобразовать значение свойства end в местные значения даты и времени клиента можно с помощью метода [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime).</span><span class="sxs-lookup"><span data-stu-id="ad140-p114">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="ad140-322">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="ad140-322">Read mode</span></span>

<span data-ttu-id="ad140-323">Свойство `end` возвращает объект `Date`.</span><span class="sxs-lookup"><span data-stu-id="ad140-323">The `end` property returns a `Date` object.</span></span>

```js
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="ad140-324">Режим создания</span><span class="sxs-lookup"><span data-stu-id="ad140-324">Compose mode</span></span>

<span data-ttu-id="ad140-325">Свойство `end` возвращает объект `Time`.</span><span class="sxs-lookup"><span data-stu-id="ad140-325">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="ad140-326">Если вы задаете время окончания с помощью метода [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.2#setasync-datetime--options--callback-), необходимо использовать метод [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) для преобразования местного времени на клиенте в формат UTC для сервера.</span><span class="sxs-lookup"><span data-stu-id="ad140-326">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.2#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="ad140-327">В примере ниже показано, как с помощью метода [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.2#setasync-datetime--options--callback-) объекта `Time` задать время окончания встречи.</span><span class="sxs-lookup"><span data-stu-id="ad140-327">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.2#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="ad140-328">Тип</span><span class="sxs-lookup"><span data-stu-id="ad140-328">Type</span></span>

*   <span data-ttu-id="ad140-329">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="ad140-329">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.2)</span></span>

##### <a name="requirements"></a><span data-ttu-id="ad140-330">Requirements</span><span class="sxs-lookup"><span data-stu-id="ad140-330">Requirements</span></span>

|<span data-ttu-id="ad140-331">Требование</span><span class="sxs-lookup"><span data-stu-id="ad140-331">Requirement</span></span>| <span data-ttu-id="ad140-332">Значение</span><span class="sxs-lookup"><span data-stu-id="ad140-332">Value</span></span>|
|---|---|
|[<span data-ttu-id="ad140-333">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="ad140-333">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ad140-334">1.0</span><span class="sxs-lookup"><span data-stu-id="ad140-334">1.0</span></span>|
|[<span data-ttu-id="ad140-335">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ad140-335">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ad140-336">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ad140-336">ReadItem</span></span>|
|[<span data-ttu-id="ad140-337">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ad140-337">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ad140-338">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="ad140-338">Compose or Read</span></span>|

<br>

---
---

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-12"></a><span data-ttu-id="ad140-339">from: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="ad140-339">from: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)</span></span>

<span data-ttu-id="ad140-p115">Получает электронный адрес отправителя сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="ad140-p115">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="ad140-p116">Свойства `from` и [`sender`](#sender-emailaddressdetails) представляют одно лицо, если сообщение не отправлено представителем. В противном случае свойство `from` представляет лицо, делегировавшее полномочия, а свойство sender — представителя.</span><span class="sxs-lookup"><span data-stu-id="ad140-p116">The `from` and [`sender`](#sender-emailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="ad140-344">Свойству `recipientType`, принадлежащему объекту `EmailAddressDetails` в свойстве `from`, задано значение `undefined`.</span><span class="sxs-lookup"><span data-stu-id="ad140-344">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="ad140-345">Тип</span><span class="sxs-lookup"><span data-stu-id="ad140-345">Type</span></span>

*   [<span data-ttu-id="ad140-346">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="ad140-346">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)

##### <a name="requirements"></a><span data-ttu-id="ad140-347">Requirements</span><span class="sxs-lookup"><span data-stu-id="ad140-347">Requirements</span></span>

|<span data-ttu-id="ad140-348">Требование</span><span class="sxs-lookup"><span data-stu-id="ad140-348">Requirement</span></span>| <span data-ttu-id="ad140-349">Значение</span><span class="sxs-lookup"><span data-stu-id="ad140-349">Value</span></span>|
|---|---|
|[<span data-ttu-id="ad140-350">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="ad140-350">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ad140-351">1.0</span><span class="sxs-lookup"><span data-stu-id="ad140-351">1.0</span></span>|
|[<span data-ttu-id="ad140-352">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ad140-352">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ad140-353">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ad140-353">ReadItem</span></span>|
|[<span data-ttu-id="ad140-354">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ad140-354">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ad140-355">Чтение</span><span class="sxs-lookup"><span data-stu-id="ad140-355">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ad140-356">Пример</span><span class="sxs-lookup"><span data-stu-id="ad140-356">Example</span></span>

```js
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

<br>

---
---

#### <a name="internetmessageid-string"></a><span data-ttu-id="ad140-357">internetMessageId: String</span><span class="sxs-lookup"><span data-stu-id="ad140-357">internetMessageId: String</span></span>

<span data-ttu-id="ad140-p117">Получает идентификатор интернет-сообщения для электронного сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="ad140-p117">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="ad140-360">Тип</span><span class="sxs-lookup"><span data-stu-id="ad140-360">Type</span></span>

*   <span data-ttu-id="ad140-361">String</span><span class="sxs-lookup"><span data-stu-id="ad140-361">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="ad140-362">Requirements</span><span class="sxs-lookup"><span data-stu-id="ad140-362">Requirements</span></span>

|<span data-ttu-id="ad140-363">Требование</span><span class="sxs-lookup"><span data-stu-id="ad140-363">Requirement</span></span>| <span data-ttu-id="ad140-364">Значение</span><span class="sxs-lookup"><span data-stu-id="ad140-364">Value</span></span>|
|---|---|
|[<span data-ttu-id="ad140-365">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="ad140-365">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ad140-366">1.0</span><span class="sxs-lookup"><span data-stu-id="ad140-366">1.0</span></span>|
|[<span data-ttu-id="ad140-367">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ad140-367">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ad140-368">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ad140-368">ReadItem</span></span>|
|[<span data-ttu-id="ad140-369">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ad140-369">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ad140-370">Чтение</span><span class="sxs-lookup"><span data-stu-id="ad140-370">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ad140-371">Пример</span><span class="sxs-lookup"><span data-stu-id="ad140-371">Example</span></span>

```js
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

<br>

---
---

#### <a name="itemclass-string"></a><span data-ttu-id="ad140-372">itemClass: String</span><span class="sxs-lookup"><span data-stu-id="ad140-372">itemClass: String</span></span>

<span data-ttu-id="ad140-p118">Получает класс элемента веб-служб Exchange для выбранного элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="ad140-p118">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="ad140-p119">Свойство `itemClass` указывает класс сообщения выбранного элемента. Ниже приводятся классы сообщения по умолчанию для элемента сообщения или встречи.</span><span class="sxs-lookup"><span data-stu-id="ad140-p119">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="ad140-377">Тип</span><span class="sxs-lookup"><span data-stu-id="ad140-377">Type</span></span> | <span data-ttu-id="ad140-378">Описание</span><span class="sxs-lookup"><span data-stu-id="ad140-378">Description</span></span> | <span data-ttu-id="ad140-379">Класс элемента</span><span class="sxs-lookup"><span data-stu-id="ad140-379">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="ad140-380">Элементы встречи</span><span class="sxs-lookup"><span data-stu-id="ad140-380">Appointment items</span></span> | <span data-ttu-id="ad140-381">Это элементы календаря для класса элемента `IPM.Appointment` или `IPM.Appointment.Occurrence`.</span><span class="sxs-lookup"><span data-stu-id="ad140-381">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurrence` |
| <span data-ttu-id="ad140-382">Элементы сообщения</span><span class="sxs-lookup"><span data-stu-id="ad140-382">Message items</span></span> | <span data-ttu-id="ad140-383">Сюда входят электронные сообщения, для которых по умолчанию задан класс сообщения `IPM.Note`, а также приглашения на собрания, ответы на них и уведомления об их отмене, использующие `IPM.Schedule.Meeting` в качестве базового класса сообщения.</span><span class="sxs-lookup"><span data-stu-id="ad140-383">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="ad140-384">Можно создавать настраиваемые классы сообщения, расширяющие классы сообщения по умолчанию, например настраиваемый класс сообщения о встрече `IPM.Appointment.Contoso`.</span><span class="sxs-lookup"><span data-stu-id="ad140-384">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="ad140-385">Тип</span><span class="sxs-lookup"><span data-stu-id="ad140-385">Type</span></span>

*   <span data-ttu-id="ad140-386">String</span><span class="sxs-lookup"><span data-stu-id="ad140-386">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="ad140-387">Requirements</span><span class="sxs-lookup"><span data-stu-id="ad140-387">Requirements</span></span>

|<span data-ttu-id="ad140-388">Требование</span><span class="sxs-lookup"><span data-stu-id="ad140-388">Requirement</span></span>| <span data-ttu-id="ad140-389">Значение</span><span class="sxs-lookup"><span data-stu-id="ad140-389">Value</span></span>|
|---|---|
|[<span data-ttu-id="ad140-390">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="ad140-390">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ad140-391">1.0</span><span class="sxs-lookup"><span data-stu-id="ad140-391">1.0</span></span>|
|[<span data-ttu-id="ad140-392">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ad140-392">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ad140-393">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ad140-393">ReadItem</span></span>|
|[<span data-ttu-id="ad140-394">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ad140-394">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ad140-395">Чтение</span><span class="sxs-lookup"><span data-stu-id="ad140-395">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ad140-396">Пример</span><span class="sxs-lookup"><span data-stu-id="ad140-396">Example</span></span>

```js
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

<br>

---
---

#### <a name="nullable-itemid-string"></a><span data-ttu-id="ad140-397">(nullable) itemId: String</span><span class="sxs-lookup"><span data-stu-id="ad140-397">(nullable) itemId: String</span></span>

<span data-ttu-id="ad140-p120">Получает [идентификатор элемента веб-служб Exchange](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange) для текущего элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="ad140-p120">Gets the [Exchange Web Services item identifier](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange) for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="ad140-400">Идентификатор, возвращаемый свойством `itemId`, совпадает с [идентификатором элемента веб-служб Exchange](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange).</span><span class="sxs-lookup"><span data-stu-id="ad140-400">The identifier returned by the `itemId` property is the same as the [Exchange Web Services item identifier](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange).</span></span> <span data-ttu-id="ad140-401">Свойство `itemId` не совпадает с идентификатором записи Outlook, а также идентификатором, который используется REST API Outlook.</span><span class="sxs-lookup"><span data-stu-id="ad140-401">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="ad140-402">Перед выполнением вызовов API REST, использующих это значение, его `Office.context.mailbox.convertToRestId`необходимо преобразовать с помощью, которое доступно в наборе требований 1,3.</span><span class="sxs-lookup"><span data-stu-id="ad140-402">Before making REST API calls using this value, it should be converted using `Office.context.mailbox.convertToRestId`, which is available starting in requirement set 1.3.</span></span> <span data-ttu-id="ad140-403">Дополнительные сведения см. в статье [Использование REST API Outlook из надстройки Outlook](/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="ad140-403">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

##### <a name="type"></a><span data-ttu-id="ad140-404">Тип</span><span class="sxs-lookup"><span data-stu-id="ad140-404">Type</span></span>

*   <span data-ttu-id="ad140-405">String</span><span class="sxs-lookup"><span data-stu-id="ad140-405">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="ad140-406">Requirements</span><span class="sxs-lookup"><span data-stu-id="ad140-406">Requirements</span></span>

|<span data-ttu-id="ad140-407">Требование</span><span class="sxs-lookup"><span data-stu-id="ad140-407">Requirement</span></span>| <span data-ttu-id="ad140-408">Значение</span><span class="sxs-lookup"><span data-stu-id="ad140-408">Value</span></span>|
|---|---|
|[<span data-ttu-id="ad140-409">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="ad140-409">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ad140-410">1.0</span><span class="sxs-lookup"><span data-stu-id="ad140-410">1.0</span></span>|
|[<span data-ttu-id="ad140-411">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ad140-411">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ad140-412">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ad140-412">ReadItem</span></span>|
|[<span data-ttu-id="ad140-413">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ad140-413">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ad140-414">Чтение</span><span class="sxs-lookup"><span data-stu-id="ad140-414">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ad140-415">Пример</span><span class="sxs-lookup"><span data-stu-id="ad140-415">Example</span></span>

<span data-ttu-id="ad140-p122">Указанный ниже код проверяет наличие идентификатора элемента. Если свойство `itemId` возвращает значение `null` или `undefined`, элемент будет сохранен в хранилище, а из асинхронного результата будет получен идентификатор элемента.</span><span class="sxs-lookup"><span data-stu-id="ad140-p122">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

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

#### <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtypeviewoutlook-js-12"></a><span data-ttu-id="ad140-418">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="ad140-418">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.2)</span></span>

<span data-ttu-id="ad140-419">Получает тип элемента, который представляет экземпляр.</span><span class="sxs-lookup"><span data-stu-id="ad140-419">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="ad140-420">Свойство `itemType` возвращает одно из значений перечисления `ItemType`, которое указывает, является ли экземпляр объекта `item` сообщением или встречей.</span><span class="sxs-lookup"><span data-stu-id="ad140-420">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="ad140-421">Тип</span><span class="sxs-lookup"><span data-stu-id="ad140-421">Type</span></span>

*   [<span data-ttu-id="ad140-422">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="ad140-422">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.2)

##### <a name="requirements"></a><span data-ttu-id="ad140-423">Требования</span><span class="sxs-lookup"><span data-stu-id="ad140-423">Requirements</span></span>

|<span data-ttu-id="ad140-424">Требование</span><span class="sxs-lookup"><span data-stu-id="ad140-424">Requirement</span></span>| <span data-ttu-id="ad140-425">Значение</span><span class="sxs-lookup"><span data-stu-id="ad140-425">Value</span></span>|
|---|---|
|[<span data-ttu-id="ad140-426">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="ad140-426">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ad140-427">1.0</span><span class="sxs-lookup"><span data-stu-id="ad140-427">1.0</span></span>|
|[<span data-ttu-id="ad140-428">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ad140-428">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ad140-429">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ad140-429">ReadItem</span></span>|
|[<span data-ttu-id="ad140-430">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ad140-430">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ad140-431">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="ad140-431">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ad140-432">Пример</span><span class="sxs-lookup"><span data-stu-id="ad140-432">Example</span></span>

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

#### <a name="location-stringlocationjavascriptapioutlookofficelocationviewoutlook-js-12"></a><span data-ttu-id="ad140-433">location: String|[Location](/javascript/api/outlook/office.location?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="ad140-433">location: String|[Location](/javascript/api/outlook/office.location?view=outlook-js-1.2)</span></span>

<span data-ttu-id="ad140-434">Получает или задает место встречи.</span><span class="sxs-lookup"><span data-stu-id="ad140-434">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="ad140-435">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="ad140-435">Read mode</span></span>

<span data-ttu-id="ad140-436">Свойство `location` возвращает строку, содержащую сведения о месте встречи.</span><span class="sxs-lookup"><span data-stu-id="ad140-436">The `location` property returns a string that contains the location of the appointment.</span></span>

```js
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="ad140-437">Режим создания</span><span class="sxs-lookup"><span data-stu-id="ad140-437">Compose mode</span></span>

<span data-ttu-id="ad140-438">Свойство `location` возвращает объект `Location`, предоставляющий методы, которые используются для получения и задания места встречи.</span><span class="sxs-lookup"><span data-stu-id="ad140-438">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```js
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="ad140-439">Тип</span><span class="sxs-lookup"><span data-stu-id="ad140-439">Type</span></span>

*   <span data-ttu-id="ad140-440">String | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="ad140-440">String | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.2)</span></span>

##### <a name="requirements"></a><span data-ttu-id="ad140-441">Требования</span><span class="sxs-lookup"><span data-stu-id="ad140-441">Requirements</span></span>

|<span data-ttu-id="ad140-442">Требование</span><span class="sxs-lookup"><span data-stu-id="ad140-442">Requirement</span></span>| <span data-ttu-id="ad140-443">Значение</span><span class="sxs-lookup"><span data-stu-id="ad140-443">Value</span></span>|
|---|---|
|[<span data-ttu-id="ad140-444">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="ad140-444">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ad140-445">1.0</span><span class="sxs-lookup"><span data-stu-id="ad140-445">1.0</span></span>|
|[<span data-ttu-id="ad140-446">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ad140-446">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ad140-447">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ad140-447">ReadItem</span></span>|
|[<span data-ttu-id="ad140-448">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ad140-448">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ad140-449">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="ad140-449">Compose or Read</span></span>|

<br>

---
---

#### <a name="normalizedsubject-string"></a><span data-ttu-id="ad140-450">normalizedSubject: String</span><span class="sxs-lookup"><span data-stu-id="ad140-450">normalizedSubject: String</span></span>

<span data-ttu-id="ad140-p123">Получает тему элемента со всеми удаленными префиксами (включая `RE:` и `FWD:`). Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="ad140-p123">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="ad140-p124">Свойство normalizedSubject получает тему элемента со стандартными префиксами (такими как `RE:` и `FW:`), добавляемыми почтовыми программами. Для получения темы элемента с неизмененными префиксами используйте свойство [`subject`](#subject-stringsubject).</span><span class="sxs-lookup"><span data-stu-id="ad140-p124">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="ad140-455">Тип</span><span class="sxs-lookup"><span data-stu-id="ad140-455">Type</span></span>

*   <span data-ttu-id="ad140-456">String</span><span class="sxs-lookup"><span data-stu-id="ad140-456">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="ad140-457">Requirements</span><span class="sxs-lookup"><span data-stu-id="ad140-457">Requirements</span></span>

|<span data-ttu-id="ad140-458">Требование</span><span class="sxs-lookup"><span data-stu-id="ad140-458">Requirement</span></span>| <span data-ttu-id="ad140-459">Значение</span><span class="sxs-lookup"><span data-stu-id="ad140-459">Value</span></span>|
|---|---|
|[<span data-ttu-id="ad140-460">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="ad140-460">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ad140-461">1.0</span><span class="sxs-lookup"><span data-stu-id="ad140-461">1.0</span></span>|
|[<span data-ttu-id="ad140-462">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ad140-462">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ad140-463">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ad140-463">ReadItem</span></span>|
|[<span data-ttu-id="ad140-464">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ad140-464">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ad140-465">Чтение</span><span class="sxs-lookup"><span data-stu-id="ad140-465">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ad140-466">Пример</span><span class="sxs-lookup"><span data-stu-id="ad140-466">Example</span></span>

```js
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

<br>

---
---

#### <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-12recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-12"></a><span data-ttu-id="ad140-467">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="ad140-467">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

<span data-ttu-id="ad140-468">Предоставляет доступ к необязательным участникам события.</span><span class="sxs-lookup"><span data-stu-id="ad140-468">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="ad140-469">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="ad140-469">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="ad140-470">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="ad140-470">Read mode</span></span>

<span data-ttu-id="ad140-471">Свойство `optionalAttendees` возвращает массив, содержащий объект `EmailAddressDetails` для каждого необязательного участника собрания.</span><span class="sxs-lookup"><span data-stu-id="ad140-471">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span> <span data-ttu-id="ad140-472">Коллекция может включать не более 100 элементов по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="ad140-472">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="ad140-473">Однако на компьютерах с Windows и Mac OS может быть до 500 элементов.</span><span class="sxs-lookup"><span data-stu-id="ad140-473">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="ad140-474">Режим создания</span><span class="sxs-lookup"><span data-stu-id="ad140-474">Compose mode</span></span>

<span data-ttu-id="ad140-475">Свойство `optionalAttendees` возвращает объект `Recipients`, который предоставляет методы для получения или обновления необязательных участников собрания.</span><span class="sxs-lookup"><span data-stu-id="ad140-475">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span> <span data-ttu-id="ad140-476">Коллекция может включать не более 100 элементов по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="ad140-476">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="ad140-477">Однако для компьютеров с Windows и Mac OS действуют указанные ниже ограничения.</span><span class="sxs-lookup"><span data-stu-id="ad140-477">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="ad140-478">Максимальное количество элементов — 500.</span><span class="sxs-lookup"><span data-stu-id="ad140-478">Get 500 members maximum.</span></span>
- <span data-ttu-id="ad140-479">Установить ограничение количества элементов на вызов — не более 100, общего количества — не более 500.</span><span class="sxs-lookup"><span data-stu-id="ad140-479">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="ad140-480">Тип</span><span class="sxs-lookup"><span data-stu-id="ad140-480">Type</span></span>

*   <span data-ttu-id="ad140-481">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="ad140-481">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

##### <a name="requirements"></a><span data-ttu-id="ad140-482">Requirements</span><span class="sxs-lookup"><span data-stu-id="ad140-482">Requirements</span></span>

|<span data-ttu-id="ad140-483">Требование</span><span class="sxs-lookup"><span data-stu-id="ad140-483">Requirement</span></span>| <span data-ttu-id="ad140-484">Значение</span><span class="sxs-lookup"><span data-stu-id="ad140-484">Value</span></span>|
|---|---|
|[<span data-ttu-id="ad140-485">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="ad140-485">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ad140-486">1.0</span><span class="sxs-lookup"><span data-stu-id="ad140-486">1.0</span></span>|
|[<span data-ttu-id="ad140-487">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ad140-487">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ad140-488">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ad140-488">ReadItem</span></span>|
|[<span data-ttu-id="ad140-489">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ad140-489">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ad140-490">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="ad140-490">Compose or Read</span></span>|

<br>

---
---

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-12"></a><span data-ttu-id="ad140-491">organizer: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="ad140-491">organizer: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)</span></span>

<span data-ttu-id="ad140-p128">Получает электронный адрес организатора указанного собрания. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="ad140-p128">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="ad140-494">Тип</span><span class="sxs-lookup"><span data-stu-id="ad140-494">Type</span></span>

*   [<span data-ttu-id="ad140-495">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="ad140-495">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)

##### <a name="requirements"></a><span data-ttu-id="ad140-496">Requirements</span><span class="sxs-lookup"><span data-stu-id="ad140-496">Requirements</span></span>

|<span data-ttu-id="ad140-497">Требование</span><span class="sxs-lookup"><span data-stu-id="ad140-497">Requirement</span></span>| <span data-ttu-id="ad140-498">Значение</span><span class="sxs-lookup"><span data-stu-id="ad140-498">Value</span></span>|
|---|---|
|[<span data-ttu-id="ad140-499">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="ad140-499">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ad140-500">1.0</span><span class="sxs-lookup"><span data-stu-id="ad140-500">1.0</span></span>|
|[<span data-ttu-id="ad140-501">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ad140-501">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ad140-502">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ad140-502">ReadItem</span></span>|
|[<span data-ttu-id="ad140-503">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ad140-503">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ad140-504">Чтение</span><span class="sxs-lookup"><span data-stu-id="ad140-504">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ad140-505">Пример</span><span class="sxs-lookup"><span data-stu-id="ad140-505">Example</span></span>

```js
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

<br>

---
---

#### <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-12recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-12"></a><span data-ttu-id="ad140-506">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="ad140-506">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

<span data-ttu-id="ad140-507">Предоставляет доступ к обязательным участникам события.</span><span class="sxs-lookup"><span data-stu-id="ad140-507">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="ad140-508">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="ad140-508">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="ad140-509">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="ad140-509">Read mode</span></span>

<span data-ttu-id="ad140-510">Свойство `requiredAttendees` возвращает массив, содержащий объект `EmailAddressDetails` для каждого обязательного участника собрания.</span><span class="sxs-lookup"><span data-stu-id="ad140-510">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span> <span data-ttu-id="ad140-511">Коллекция может включать не более 100 элементов по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="ad140-511">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="ad140-512">Однако на компьютерах с Windows и Mac OS может быть до 500 элементов.</span><span class="sxs-lookup"><span data-stu-id="ad140-512">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="ad140-513">Режим создания</span><span class="sxs-lookup"><span data-stu-id="ad140-513">Compose mode</span></span>

<span data-ttu-id="ad140-514">Свойство `requiredAttendees` возвращает объект `Recipients`, предоставляющий методы, с помощью которых можно получить или обновить сведения об обязательных участниках собрания.</span><span class="sxs-lookup"><span data-stu-id="ad140-514">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span> <span data-ttu-id="ad140-515">Коллекция может включать не более 100 элементов по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="ad140-515">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="ad140-516">Однако для компьютеров с Windows и Mac OS действуют указанные ниже ограничения.</span><span class="sxs-lookup"><span data-stu-id="ad140-516">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="ad140-517">Максимальное количество элементов — 500.</span><span class="sxs-lookup"><span data-stu-id="ad140-517">Get 500 members maximum.</span></span>
- <span data-ttu-id="ad140-518">Установить ограничение количества элементов на вызов — не более 100, общего количества — не более 500.</span><span class="sxs-lookup"><span data-stu-id="ad140-518">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="ad140-519">Тип</span><span class="sxs-lookup"><span data-stu-id="ad140-519">Type</span></span>

*   <span data-ttu-id="ad140-520">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="ad140-520">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

##### <a name="requirements"></a><span data-ttu-id="ad140-521">Requirements</span><span class="sxs-lookup"><span data-stu-id="ad140-521">Requirements</span></span>

|<span data-ttu-id="ad140-522">Требование</span><span class="sxs-lookup"><span data-stu-id="ad140-522">Requirement</span></span>| <span data-ttu-id="ad140-523">Значение</span><span class="sxs-lookup"><span data-stu-id="ad140-523">Value</span></span>|
|---|---|
|[<span data-ttu-id="ad140-524">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="ad140-524">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ad140-525">1.0</span><span class="sxs-lookup"><span data-stu-id="ad140-525">1.0</span></span>|
|[<span data-ttu-id="ad140-526">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ad140-526">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ad140-527">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ad140-527">ReadItem</span></span>|
|[<span data-ttu-id="ad140-528">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ad140-528">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ad140-529">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="ad140-529">Compose or Read</span></span>|

<br>

---
---

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-12"></a><span data-ttu-id="ad140-530">sender: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="ad140-530">sender: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)</span></span>

<span data-ttu-id="ad140-p132">Получает электронный адрес отправителя электронного сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="ad140-p132">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="ad140-p133">Свойства [`from`](#from-emailaddressdetails) и `sender` представляют одно лицо, если сообщение не отправлено представителем. В противном случае свойство `from` представляет лицо, делегировавшее полномочия, а свойство sender — представителя.</span><span class="sxs-lookup"><span data-stu-id="ad140-p133">The [`from`](#from-emailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="ad140-535">Свойству `recipientType`, принадлежащему объекту `EmailAddressDetails` в свойстве `sender`, задано значение `undefined`.</span><span class="sxs-lookup"><span data-stu-id="ad140-535">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="ad140-536">Тип</span><span class="sxs-lookup"><span data-stu-id="ad140-536">Type</span></span>

*   [<span data-ttu-id="ad140-537">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="ad140-537">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)

##### <a name="requirements"></a><span data-ttu-id="ad140-538">Requirements</span><span class="sxs-lookup"><span data-stu-id="ad140-538">Requirements</span></span>

|<span data-ttu-id="ad140-539">Требование</span><span class="sxs-lookup"><span data-stu-id="ad140-539">Requirement</span></span>| <span data-ttu-id="ad140-540">Значение</span><span class="sxs-lookup"><span data-stu-id="ad140-540">Value</span></span>|
|---|---|
|[<span data-ttu-id="ad140-541">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="ad140-541">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ad140-542">1.0</span><span class="sxs-lookup"><span data-stu-id="ad140-542">1.0</span></span>|
|[<span data-ttu-id="ad140-543">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ad140-543">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ad140-544">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ad140-544">ReadItem</span></span>|
|[<span data-ttu-id="ad140-545">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ad140-545">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ad140-546">Чтение</span><span class="sxs-lookup"><span data-stu-id="ad140-546">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ad140-547">Пример</span><span class="sxs-lookup"><span data-stu-id="ad140-547">Example</span></span>

```js
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

<br>

---
---

#### <a name="start-datetimejavascriptapioutlookofficetimeviewoutlook-js-12"></a><span data-ttu-id="ad140-548">start: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="ad140-548">start: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.2)</span></span>

<span data-ttu-id="ad140-549">Получает или задает дату и время начала встречи.</span><span class="sxs-lookup"><span data-stu-id="ad140-549">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="ad140-p134">Свойство `start` представлено в виде значения даты и времени в формате UTC. Это значение можно преобразовать в местные значения даты и времени клиента с помощью метода [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime).</span><span class="sxs-lookup"><span data-stu-id="ad140-p134">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="ad140-552">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="ad140-552">Read mode</span></span>

<span data-ttu-id="ad140-553">Свойство `start` возвращает объект `Date`.</span><span class="sxs-lookup"><span data-stu-id="ad140-553">The `start` property returns a `Date` object.</span></span>

```js
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="ad140-554">Режим создания</span><span class="sxs-lookup"><span data-stu-id="ad140-554">Compose mode</span></span>

<span data-ttu-id="ad140-555">Свойство `start` возвращает объект `Time`.</span><span class="sxs-lookup"><span data-stu-id="ad140-555">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="ad140-556">Если вы задаете время начала с помощью метода [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.2#setasync-datetime--options--callback-), необходимо использовать метод [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) для преобразования местного времени на клиенте в формат UTC для сервера.</span><span class="sxs-lookup"><span data-stu-id="ad140-556">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.2#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>
<span data-ttu-id="ad140-557">В примере ниже с помощью метода [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.2#setasync-datetime--options--callback-) объекта `Time` задается время начала встречи в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="ad140-557">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.2#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="ad140-558">Тип</span><span class="sxs-lookup"><span data-stu-id="ad140-558">Type</span></span>

*   <span data-ttu-id="ad140-559">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="ad140-559">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.2)</span></span>

##### <a name="requirements"></a><span data-ttu-id="ad140-560">Requirements</span><span class="sxs-lookup"><span data-stu-id="ad140-560">Requirements</span></span>

|<span data-ttu-id="ad140-561">Требование</span><span class="sxs-lookup"><span data-stu-id="ad140-561">Requirement</span></span>| <span data-ttu-id="ad140-562">Значение</span><span class="sxs-lookup"><span data-stu-id="ad140-562">Value</span></span>|
|---|---|
|[<span data-ttu-id="ad140-563">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="ad140-563">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ad140-564">1.0</span><span class="sxs-lookup"><span data-stu-id="ad140-564">1.0</span></span>|
|[<span data-ttu-id="ad140-565">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ad140-565">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ad140-566">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ad140-566">ReadItem</span></span>|
|[<span data-ttu-id="ad140-567">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ad140-567">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ad140-568">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="ad140-568">Compose or Read</span></span>|

<br>

---
---

#### <a name="subject-stringsubjectjavascriptapioutlookofficesubjectviewoutlook-js-12"></a><span data-ttu-id="ad140-569">subject: String|[Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="ad140-569">subject: String|[Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.2)</span></span>

<span data-ttu-id="ad140-570">Получает или задает описание, которое отображается в поле темы элемента.</span><span class="sxs-lookup"><span data-stu-id="ad140-570">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="ad140-571">Свойство `subject` получает или задает всю тему элемента для отправки с почтового сервера.</span><span class="sxs-lookup"><span data-stu-id="ad140-571">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="ad140-572">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="ad140-572">Read mode</span></span>

<span data-ttu-id="ad140-p136">Свойство `subject` возвращает строку. С помощью свойства [`normalizedSubject`](#normalizedsubject-string) можно получить тему без начальных префиксов, таких как `RE:` и `FW:`.</span><span class="sxs-lookup"><span data-stu-id="ad140-p136">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```js
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="ad140-575">Режим создания</span><span class="sxs-lookup"><span data-stu-id="ad140-575">Compose mode</span></span>

<span data-ttu-id="ad140-576">Свойство `subject` возвращает объект `Subject`, который предоставляет методы для получения и задания темы.</span><span class="sxs-lookup"><span data-stu-id="ad140-576">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```js
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="ad140-577">Тип</span><span class="sxs-lookup"><span data-stu-id="ad140-577">Type</span></span>

*   <span data-ttu-id="ad140-578">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="ad140-578">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.2)</span></span>

##### <a name="requirements"></a><span data-ttu-id="ad140-579">Requirements</span><span class="sxs-lookup"><span data-stu-id="ad140-579">Requirements</span></span>

|<span data-ttu-id="ad140-580">Требование</span><span class="sxs-lookup"><span data-stu-id="ad140-580">Requirement</span></span>| <span data-ttu-id="ad140-581">Значение</span><span class="sxs-lookup"><span data-stu-id="ad140-581">Value</span></span>|
|---|---|
|[<span data-ttu-id="ad140-582">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="ad140-582">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ad140-583">1.0</span><span class="sxs-lookup"><span data-stu-id="ad140-583">1.0</span></span>|
|[<span data-ttu-id="ad140-584">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ad140-584">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ad140-585">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ad140-585">ReadItem</span></span>|
|[<span data-ttu-id="ad140-586">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ad140-586">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ad140-587">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="ad140-587">Compose or Read</span></span>|

<br>

---
---

#### <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-12recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-12"></a><span data-ttu-id="ad140-588">to: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="ad140-588">to: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

<span data-ttu-id="ad140-589">Предоставляет доступ к получателям, указанным в строке **Кому** сообщения.</span><span class="sxs-lookup"><span data-stu-id="ad140-589">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="ad140-590">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="ad140-590">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="ad140-591">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="ad140-591">Read mode</span></span>

<span data-ttu-id="ad140-592">Свойство `to` возвращает массив, содержащий объект `EmailAddressDetails` для каждого получателя в строке **Кому** сообщения.</span><span class="sxs-lookup"><span data-stu-id="ad140-592">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message.</span></span> <span data-ttu-id="ad140-593">Коллекция может включать не более 100 элементов по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="ad140-593">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="ad140-594">Однако на компьютерах с Windows и Mac OS может быть до 500 элементов.</span><span class="sxs-lookup"><span data-stu-id="ad140-594">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="ad140-595">Режим создания</span><span class="sxs-lookup"><span data-stu-id="ad140-595">Compose mode</span></span>

<span data-ttu-id="ad140-596">Свойство `to` возвращает объект `Recipients`, предоставляющий методы для получения или обновления получателей, которые указаны в строке **Кому** сообщения.</span><span class="sxs-lookup"><span data-stu-id="ad140-596">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span> <span data-ttu-id="ad140-597">Коллекция может включать не более 100 элементов по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="ad140-597">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="ad140-598">Однако для компьютеров с Windows и Mac OS действуют указанные ниже ограничения.</span><span class="sxs-lookup"><span data-stu-id="ad140-598">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="ad140-599">Максимальное количество элементов — 500.</span><span class="sxs-lookup"><span data-stu-id="ad140-599">Get 500 members maximum.</span></span>
- <span data-ttu-id="ad140-600">Установить ограничение количества элементов на вызов — не более 100, общего количества — не более 500.</span><span class="sxs-lookup"><span data-stu-id="ad140-600">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="ad140-601">Тип</span><span class="sxs-lookup"><span data-stu-id="ad140-601">Type</span></span>

*   <span data-ttu-id="ad140-602">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="ad140-602">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

##### <a name="requirements"></a><span data-ttu-id="ad140-603">Requirements</span><span class="sxs-lookup"><span data-stu-id="ad140-603">Requirements</span></span>

|<span data-ttu-id="ad140-604">Требование</span><span class="sxs-lookup"><span data-stu-id="ad140-604">Requirement</span></span>| <span data-ttu-id="ad140-605">Значение</span><span class="sxs-lookup"><span data-stu-id="ad140-605">Value</span></span>|
|---|---|
|[<span data-ttu-id="ad140-606">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="ad140-606">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ad140-607">1.0</span><span class="sxs-lookup"><span data-stu-id="ad140-607">1.0</span></span>|
|[<span data-ttu-id="ad140-608">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ad140-608">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ad140-609">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ad140-609">ReadItem</span></span>|
|[<span data-ttu-id="ad140-610">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ad140-610">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ad140-611">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="ad140-611">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="ad140-612">Методы</span><span class="sxs-lookup"><span data-stu-id="ad140-612">Methods</span></span>

#### <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="ad140-613">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="ad140-613">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="ad140-614">Добавляет файл в сообщение или встречу в качестве вложения.</span><span class="sxs-lookup"><span data-stu-id="ad140-614">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="ad140-615">Метод `addFileAttachmentAsync` передает файл по указанному универсальному коду ресурса (URI) и вкладывает его в элемент в форме создания.</span><span class="sxs-lookup"><span data-stu-id="ad140-615">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="ad140-616">Идентификатор можно использовать с методом [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback), чтобы удалить вложение, добавленное во время текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="ad140-616">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ad140-617">Параметры</span><span class="sxs-lookup"><span data-stu-id="ad140-617">Parameters</span></span>

|<span data-ttu-id="ad140-618">Имя</span><span class="sxs-lookup"><span data-stu-id="ad140-618">Name</span></span>| <span data-ttu-id="ad140-619">Тип</span><span class="sxs-lookup"><span data-stu-id="ad140-619">Type</span></span>| <span data-ttu-id="ad140-620">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="ad140-620">Attributes</span></span>| <span data-ttu-id="ad140-621">Описание</span><span class="sxs-lookup"><span data-stu-id="ad140-621">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="ad140-622">String</span><span class="sxs-lookup"><span data-stu-id="ad140-622">String</span></span>||<span data-ttu-id="ad140-p140">Универсальный код ресурса (URI), представляющий расположение файла, который нужно вложить в сообщение или встречу. Максимальная длина — 2048 символов.</span><span class="sxs-lookup"><span data-stu-id="ad140-p140">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="ad140-625">String</span><span class="sxs-lookup"><span data-stu-id="ad140-625">String</span></span>||<span data-ttu-id="ad140-p141">Имя вложения, которое отображается при передаче вложения. Максимальная длина — 255 символов.</span><span class="sxs-lookup"><span data-stu-id="ad140-p141">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="ad140-628">Объект</span><span class="sxs-lookup"><span data-stu-id="ad140-628">Object</span></span>| <span data-ttu-id="ad140-629">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="ad140-629">&lt;optional&gt;</span></span>|<span data-ttu-id="ad140-630">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="ad140-630">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="ad140-631">Object</span><span class="sxs-lookup"><span data-stu-id="ad140-631">Object</span></span>| <span data-ttu-id="ad140-632">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="ad140-632">&lt;optional&gt;</span></span>|<span data-ttu-id="ad140-633">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="ad140-633">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="ad140-634">функция</span><span class="sxs-lookup"><span data-stu-id="ad140-634">function</span></span>| <span data-ttu-id="ad140-635">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="ad140-635">&lt;optional&gt;</span></span>|<span data-ttu-id="ad140-636">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="ad140-636">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="ad140-637">После успешного выполнения идентификатор вложения будет представлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="ad140-637">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="ad140-638">Если добавить вложение не удастся, объект `asyncResult` будет содержать объект `Error` с описанием ошибки.</span><span class="sxs-lookup"><span data-stu-id="ad140-638">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="ad140-639">Ошибки</span><span class="sxs-lookup"><span data-stu-id="ad140-639">Errors</span></span>

| <span data-ttu-id="ad140-640">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="ad140-640">Error code</span></span> | <span data-ttu-id="ad140-641">Описание</span><span class="sxs-lookup"><span data-stu-id="ad140-641">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="ad140-642">Вложение превышает максимальный размер.</span><span class="sxs-lookup"><span data-stu-id="ad140-642">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="ad140-643">Расширение вложения не поддерживается.</span><span class="sxs-lookup"><span data-stu-id="ad140-643">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="ad140-644">Сообщение или встреча содержат слишком много вложений.</span><span class="sxs-lookup"><span data-stu-id="ad140-644">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="ad140-645">Требования</span><span class="sxs-lookup"><span data-stu-id="ad140-645">Requirements</span></span>

|<span data-ttu-id="ad140-646">Требование</span><span class="sxs-lookup"><span data-stu-id="ad140-646">Requirement</span></span>| <span data-ttu-id="ad140-647">Значение</span><span class="sxs-lookup"><span data-stu-id="ad140-647">Value</span></span>|
|---|---|
|[<span data-ttu-id="ad140-648">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="ad140-648">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ad140-649">1.1</span><span class="sxs-lookup"><span data-stu-id="ad140-649">1.1</span></span>|
|[<span data-ttu-id="ad140-650">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ad140-650">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ad140-651">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="ad140-651">ReadWriteItem</span></span>|
|[<span data-ttu-id="ad140-652">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ad140-652">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ad140-653">Создание</span><span class="sxs-lookup"><span data-stu-id="ad140-653">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="ad140-654">Пример</span><span class="sxs-lookup"><span data-stu-id="ad140-654">Example</span></span>

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

<br>

---
---

#### <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="ad140-655">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="ad140-655">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="ad140-656">Добавляет к сообщению элемент Exchange, например сообщение, в виде вложения.</span><span class="sxs-lookup"><span data-stu-id="ad140-656">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="ad140-p142">С помощью метода `addItemAttachmentAsync` можно в элемент формы создания вложить элемент с указанным идентификатором Exchange. Если указать метод обратного вызова, то этот метод вызывается с помощью параметра `asyncResult`, который содержит идентификатор вложения или код, указывающий на ошибки, которые произошли при вложении элемента. При необходимости можно использовать параметр `options` для передачи сведений о состоянии методу обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="ad140-p142">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="ad140-660">Идентификатор можно использовать с методом [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback), чтобы удалить вложение, добавленное во время текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="ad140-660">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="ad140-661">Если ваша надстройка Office выполняется в Outlook в Интернете, метод `addItemAttachmentAsync` обеспечивает вложение элементов в элементы, отличные от редактируемого. Однако это действие не рекомендуется выполнять, так как оно не поддерживается.</span><span class="sxs-lookup"><span data-stu-id="ad140-661">If your Office Add-in is running in Outlook on the web, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ad140-662">Параметры</span><span class="sxs-lookup"><span data-stu-id="ad140-662">Parameters</span></span>

|<span data-ttu-id="ad140-663">Имя</span><span class="sxs-lookup"><span data-stu-id="ad140-663">Name</span></span>| <span data-ttu-id="ad140-664">Тип</span><span class="sxs-lookup"><span data-stu-id="ad140-664">Type</span></span>| <span data-ttu-id="ad140-665">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="ad140-665">Attributes</span></span>| <span data-ttu-id="ad140-666">Описание</span><span class="sxs-lookup"><span data-stu-id="ad140-666">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="ad140-667">String</span><span class="sxs-lookup"><span data-stu-id="ad140-667">String</span></span>||<span data-ttu-id="ad140-p143">Идентификатор Exchange для вкладываемого элемента. Максимальная длина — 100 символов.</span><span class="sxs-lookup"><span data-stu-id="ad140-p143">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="ad140-670">String</span><span class="sxs-lookup"><span data-stu-id="ad140-670">String</span></span>||<span data-ttu-id="ad140-671">Тема вкладываемого элемента.</span><span class="sxs-lookup"><span data-stu-id="ad140-671">The subject of the item to be attached.</span></span> <span data-ttu-id="ad140-672">Максимальная длина: 255 символов.</span><span class="sxs-lookup"><span data-stu-id="ad140-672">The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="ad140-673">Object</span><span class="sxs-lookup"><span data-stu-id="ad140-673">Object</span></span>| <span data-ttu-id="ad140-674">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="ad140-674">&lt;optional&gt;</span></span>|<span data-ttu-id="ad140-675">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="ad140-675">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="ad140-676">Объект</span><span class="sxs-lookup"><span data-stu-id="ad140-676">Object</span></span>| <span data-ttu-id="ad140-677">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="ad140-677">&lt;optional&gt;</span></span>|<span data-ttu-id="ad140-678">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="ad140-678">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="ad140-679">функция</span><span class="sxs-lookup"><span data-stu-id="ad140-679">function</span></span>| <span data-ttu-id="ad140-680">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="ad140-680">&lt;optional&gt;</span></span>|<span data-ttu-id="ad140-681">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="ad140-681">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="ad140-682">После успешного выполнения идентификатор вложения будет представлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="ad140-682">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="ad140-683">Если добавить вложение не удастся, объект `asyncResult` будет содержать объект `Error` с описанием ошибки.</span><span class="sxs-lookup"><span data-stu-id="ad140-683">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="ad140-684">Ошибки</span><span class="sxs-lookup"><span data-stu-id="ad140-684">Errors</span></span>

| <span data-ttu-id="ad140-685">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="ad140-685">Error code</span></span> | <span data-ttu-id="ad140-686">Описание</span><span class="sxs-lookup"><span data-stu-id="ad140-686">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="ad140-687">Сообщение или встреча содержат слишком много вложений.</span><span class="sxs-lookup"><span data-stu-id="ad140-687">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="ad140-688">Requirements</span><span class="sxs-lookup"><span data-stu-id="ad140-688">Requirements</span></span>

|<span data-ttu-id="ad140-689">Требование</span><span class="sxs-lookup"><span data-stu-id="ad140-689">Requirement</span></span>| <span data-ttu-id="ad140-690">Значение</span><span class="sxs-lookup"><span data-stu-id="ad140-690">Value</span></span>|
|---|---|
|[<span data-ttu-id="ad140-691">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="ad140-691">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ad140-692">1.1</span><span class="sxs-lookup"><span data-stu-id="ad140-692">1.1</span></span>|
|[<span data-ttu-id="ad140-693">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ad140-693">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ad140-694">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="ad140-694">ReadWriteItem</span></span>|
|[<span data-ttu-id="ad140-695">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ad140-695">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ad140-696">Создание</span><span class="sxs-lookup"><span data-stu-id="ad140-696">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="ad140-697">Пример</span><span class="sxs-lookup"><span data-stu-id="ad140-697">Example</span></span>

<span data-ttu-id="ad140-698">В следующем примере существующий элемент Outlook добавляется в виде вложения с именем `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="ad140-698">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="ad140-699">displayReplyAllForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="ad140-699">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="ad140-700">Отображает форму ответа, включающую отправителя и всех получателей выбранного сообщения или организатора и всех участников выбранной встречи.</span><span class="sxs-lookup"><span data-stu-id="ad140-700">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="ad140-701">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="ad140-701">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="ad140-702">В Outlook в Интернете форма ответа отображается в виде всплывающей формы в представлении с 3 либо 1 или 2 столбцами.</span><span class="sxs-lookup"><span data-stu-id="ad140-702">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="ad140-703">Если любой строковый параметр превышает указанные для него ограничения, `displayReplyAllForm` возвращает исключение.</span><span class="sxs-lookup"><span data-stu-id="ad140-703">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="ad140-p145">Если в параметре `formData.attachments` указаны вложения, Outlook в Интернете и классические клиенты пытаются скачать их и вложить в форму ответа. Если какие-либо вложения добавить не удается, в форме отображается сообщение об ошибке. Если сообщения об ошибках не предусмотрены, то они не отображаются.</span><span class="sxs-lookup"><span data-stu-id="ad140-p145">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ad140-707">Параметры</span><span class="sxs-lookup"><span data-stu-id="ad140-707">Parameters</span></span>

|<span data-ttu-id="ad140-708">Имя</span><span class="sxs-lookup"><span data-stu-id="ad140-708">Name</span></span>| <span data-ttu-id="ad140-709">Тип</span><span class="sxs-lookup"><span data-stu-id="ad140-709">Type</span></span>| <span data-ttu-id="ad140-710">Описание</span><span class="sxs-lookup"><span data-stu-id="ad140-710">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="ad140-711">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="ad140-711">String &#124; Object</span></span>| |<span data-ttu-id="ad140-p146">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="ad140-p146">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="ad140-714">**ИЛИ**</span><span class="sxs-lookup"><span data-stu-id="ad140-714">**OR**</span></span><br/><span data-ttu-id="ad140-p147">Объект, который содержит текст или данные вложения и функцию обратного вызова. Ниже представлено определение этого объекта.</span><span class="sxs-lookup"><span data-stu-id="ad140-p147">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="ad140-717">String</span><span class="sxs-lookup"><span data-stu-id="ad140-717">String</span></span> | <span data-ttu-id="ad140-718">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="ad140-718">&lt;optional&gt;</span></span> | <span data-ttu-id="ad140-p148">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="ad140-p148">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="ad140-721">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="ad140-721">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="ad140-722">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="ad140-722">&lt;optional&gt;</span></span> | <span data-ttu-id="ad140-723">Массив объектов JSON, представляющих собой вложенные файлы или элементы.</span><span class="sxs-lookup"><span data-stu-id="ad140-723">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="ad140-724">String</span><span class="sxs-lookup"><span data-stu-id="ad140-724">String</span></span> | | <span data-ttu-id="ad140-p149">Указывает тип вложения. Допустимые значения: `file` для вложенного файла и `item` для вложенного элемента.</span><span class="sxs-lookup"><span data-stu-id="ad140-p149">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="ad140-727">Строка</span><span class="sxs-lookup"><span data-stu-id="ad140-727">String</span></span> | | <span data-ttu-id="ad140-728">Строка, содержащая имя вложения, длиной до 255 символов.</span><span class="sxs-lookup"><span data-stu-id="ad140-728">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="ad140-729">String</span><span class="sxs-lookup"><span data-stu-id="ad140-729">String</span></span> | | <span data-ttu-id="ad140-p150">Используется, только если свойству `type` задано значение `file`. URI расположения файла.</span><span class="sxs-lookup"><span data-stu-id="ad140-p150">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="ad140-732">String</span><span class="sxs-lookup"><span data-stu-id="ad140-732">String</span></span> | | <span data-ttu-id="ad140-p151">Используется, только если свойству `type` задано значение `item`. Идентификатор вложения EWS. Это строка длиной до 100 символов.</span><span class="sxs-lookup"><span data-stu-id="ad140-p151">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="ad140-736">function</span><span class="sxs-lookup"><span data-stu-id="ad140-736">function</span></span> | <span data-ttu-id="ad140-737">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="ad140-737">&lt;optional&gt;</span></span> | <span data-ttu-id="ad140-738">По завершении работы метода функция, переданная параметру `callback`, вызывается с помощью одного параметра `asyncResult`, представляющего собой объект [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="ad140-738">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="ad140-739">Требования</span><span class="sxs-lookup"><span data-stu-id="ad140-739">Requirements</span></span>

|<span data-ttu-id="ad140-740">Требование</span><span class="sxs-lookup"><span data-stu-id="ad140-740">Requirement</span></span>| <span data-ttu-id="ad140-741">Значение</span><span class="sxs-lookup"><span data-stu-id="ad140-741">Value</span></span>|
|---|---|
|[<span data-ttu-id="ad140-742">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="ad140-742">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ad140-743">1.0</span><span class="sxs-lookup"><span data-stu-id="ad140-743">1.0</span></span>|
|[<span data-ttu-id="ad140-744">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ad140-744">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ad140-745">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ad140-745">ReadItem</span></span>|
|[<span data-ttu-id="ad140-746">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ad140-746">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ad140-747">Чтение</span><span class="sxs-lookup"><span data-stu-id="ad140-747">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="ad140-748">Примеры</span><span class="sxs-lookup"><span data-stu-id="ad140-748">Examples</span></span>

<span data-ttu-id="ad140-749">Приведенный ниже код передает строку в функцию `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="ad140-749">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="ad140-750">Ответ с пустым текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="ad140-750">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="ad140-751">Ответ только с текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="ad140-751">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="ad140-752">Ответ с текстом сообщения и вложенным файлом.</span><span class="sxs-lookup"><span data-stu-id="ad140-752">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="ad140-753">Ответ с текстом сообщения и вложенным элементом.</span><span class="sxs-lookup"><span data-stu-id="ad140-753">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="ad140-754">Ответ с текстом сообщения, вложенным файлом, вложенным элементом и обратным вызовом.</span><span class="sxs-lookup"><span data-stu-id="ad140-754">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="ad140-755">displayReplyForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="ad140-755">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="ad140-756">Отображает форму ответа, включающую только отправителя выбранного сообщения или организатора выбранной встречи.</span><span class="sxs-lookup"><span data-stu-id="ad140-756">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="ad140-757">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="ad140-757">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="ad140-758">В Outlook в Интернете форма ответа отображается в виде всплывающей формы в представлении с 3 либо 1 или 2 столбцами.</span><span class="sxs-lookup"><span data-stu-id="ad140-758">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="ad140-759">Если любой строковый параметр превышает указанные для него ограничения, `displayReplyForm` возвращает исключение.</span><span class="sxs-lookup"><span data-stu-id="ad140-759">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="ad140-p152">Если в параметре `formData.attachments` указаны вложения, Outlook в Интернете и классические клиенты пытаются скачать их и вложить в форму ответа. Если какие-либо вложения добавить не удается, в форме отображается сообщение об ошибке. Если сообщения об ошибках не предусмотрены, то они не отображаются.</span><span class="sxs-lookup"><span data-stu-id="ad140-p152">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ad140-763">Параметры</span><span class="sxs-lookup"><span data-stu-id="ad140-763">Parameters</span></span>

|<span data-ttu-id="ad140-764">Имя</span><span class="sxs-lookup"><span data-stu-id="ad140-764">Name</span></span>| <span data-ttu-id="ad140-765">Тип</span><span class="sxs-lookup"><span data-stu-id="ad140-765">Type</span></span>| <span data-ttu-id="ad140-766">Описание</span><span class="sxs-lookup"><span data-stu-id="ad140-766">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="ad140-767">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="ad140-767">String &#124; Object</span></span>| | <span data-ttu-id="ad140-p153">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="ad140-p153">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="ad140-770">**ИЛИ**</span><span class="sxs-lookup"><span data-stu-id="ad140-770">**OR**</span></span><br/><span data-ttu-id="ad140-p154">Объект, который содержит текст или данные вложения и функцию обратного вызова. Ниже представлено определение этого объекта.</span><span class="sxs-lookup"><span data-stu-id="ad140-p154">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="ad140-773">String</span><span class="sxs-lookup"><span data-stu-id="ad140-773">String</span></span> | <span data-ttu-id="ad140-774">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="ad140-774">&lt;optional&gt;</span></span> | <span data-ttu-id="ad140-p155">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="ad140-p155">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="ad140-777">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="ad140-777">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="ad140-778">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="ad140-778">&lt;optional&gt;</span></span> | <span data-ttu-id="ad140-779">Массив объектов JSON, представляющих собой вложенные файлы или элементы.</span><span class="sxs-lookup"><span data-stu-id="ad140-779">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="ad140-780">String</span><span class="sxs-lookup"><span data-stu-id="ad140-780">String</span></span> | | <span data-ttu-id="ad140-p156">Указывает тип вложения. Допустимые значения: `file` для вложенного файла и `item` для вложенного элемента.</span><span class="sxs-lookup"><span data-stu-id="ad140-p156">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="ad140-783">Строка</span><span class="sxs-lookup"><span data-stu-id="ad140-783">String</span></span> | | <span data-ttu-id="ad140-784">Строка, содержащая имя вложения, длиной до 255 символов.</span><span class="sxs-lookup"><span data-stu-id="ad140-784">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="ad140-785">Строка</span><span class="sxs-lookup"><span data-stu-id="ad140-785">String</span></span> | | <span data-ttu-id="ad140-p157">Используется, только если свойству `type` задано значение `file`. URI расположения файла.</span><span class="sxs-lookup"><span data-stu-id="ad140-p157">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="ad140-788">String</span><span class="sxs-lookup"><span data-stu-id="ad140-788">String</span></span> | | <span data-ttu-id="ad140-p158">Используется, только если свойству `type` задано значение `item`. Идентификатор вложения EWS. Это строка длиной до 100 символов.</span><span class="sxs-lookup"><span data-stu-id="ad140-p158">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="ad140-792">function</span><span class="sxs-lookup"><span data-stu-id="ad140-792">function</span></span> | <span data-ttu-id="ad140-793">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="ad140-793">&lt;optional&gt;</span></span> | <span data-ttu-id="ad140-794">По завершении работы метода функция, переданная параметру `callback`, вызывается с помощью одного параметра `asyncResult`, представляющего собой объект [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="ad140-794">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="ad140-795">Requirements</span><span class="sxs-lookup"><span data-stu-id="ad140-795">Requirements</span></span>

|<span data-ttu-id="ad140-796">Требование</span><span class="sxs-lookup"><span data-stu-id="ad140-796">Requirement</span></span>| <span data-ttu-id="ad140-797">Значение</span><span class="sxs-lookup"><span data-stu-id="ad140-797">Value</span></span>|
|---|---|
|[<span data-ttu-id="ad140-798">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="ad140-798">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ad140-799">1.0</span><span class="sxs-lookup"><span data-stu-id="ad140-799">1.0</span></span>|
|[<span data-ttu-id="ad140-800">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ad140-800">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ad140-801">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ad140-801">ReadItem</span></span>|
|[<span data-ttu-id="ad140-802">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ad140-802">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ad140-803">Чтение</span><span class="sxs-lookup"><span data-stu-id="ad140-803">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="ad140-804">Примеры</span><span class="sxs-lookup"><span data-stu-id="ad140-804">Examples</span></span>

<span data-ttu-id="ad140-805">Приведенный ниже код передает строку в функцию `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="ad140-805">The following code passes a string to the `displayReplyForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="ad140-806">Ответ с пустым текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="ad140-806">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="ad140-807">Ответ только с текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="ad140-807">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="ad140-808">Ответ с текстом сообщения и вложенным файлом.</span><span class="sxs-lookup"><span data-stu-id="ad140-808">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="ad140-809">Ответ с текстом сообщения и вложенным элементом.</span><span class="sxs-lookup"><span data-stu-id="ad140-809">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="ad140-810">Ответ с текстом сообщения, вложенным файлом, вложенным элементом и обратным вызовом.</span><span class="sxs-lookup"><span data-stu-id="ad140-810">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlookofficeentitiesviewoutlook-js-12"></a><span data-ttu-id="ad140-811">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.2)}</span><span class="sxs-lookup"><span data-stu-id="ad140-811">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.2)}</span></span>

<span data-ttu-id="ad140-812">Получает сущности, обнаруженные в теле выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="ad140-812">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="ad140-813">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="ad140-813">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="ad140-814">Требования</span><span class="sxs-lookup"><span data-stu-id="ad140-814">Requirements</span></span>

|<span data-ttu-id="ad140-815">Требование</span><span class="sxs-lookup"><span data-stu-id="ad140-815">Requirement</span></span>| <span data-ttu-id="ad140-816">Значение</span><span class="sxs-lookup"><span data-stu-id="ad140-816">Value</span></span>|
|---|---|
|[<span data-ttu-id="ad140-817">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="ad140-817">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ad140-818">1.0</span><span class="sxs-lookup"><span data-stu-id="ad140-818">1.0</span></span>|
|[<span data-ttu-id="ad140-819">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ad140-819">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ad140-820">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ad140-820">ReadItem</span></span>|
|[<span data-ttu-id="ad140-821">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ad140-821">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ad140-822">Чтение</span><span class="sxs-lookup"><span data-stu-id="ad140-822">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="ad140-823">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="ad140-823">Returns:</span></span>

<span data-ttu-id="ad140-824">Тип: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="ad140-824">Type: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.2)</span></span>

##### <a name="example"></a><span data-ttu-id="ad140-825">Пример</span><span class="sxs-lookup"><span data-stu-id="ad140-825">Example</span></span>

<span data-ttu-id="ad140-826">Ниже приведен пример получения доступа к сущностям контактов в теле текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="ad140-826">The following example accesses the contacts entities in the current item's body.</span></span>

```js
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

<br>

---
---

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-12meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-12phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-12tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-12"></a><span data-ttu-id="ad140-827">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.2)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.2)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.2)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.2))>}</span><span class="sxs-lookup"><span data-stu-id="ad140-827">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.2)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.2)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.2)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.2))>}</span></span>

<span data-ttu-id="ad140-828">Получает массив всех сущностей указанного типа, обнаруженных в теле выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="ad140-828">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="ad140-829">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="ad140-829">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ad140-830">Параметры</span><span class="sxs-lookup"><span data-stu-id="ad140-830">Parameters</span></span>

|<span data-ttu-id="ad140-831">Имя</span><span class="sxs-lookup"><span data-stu-id="ad140-831">Name</span></span>| <span data-ttu-id="ad140-832">Тип</span><span class="sxs-lookup"><span data-stu-id="ad140-832">Type</span></span>| <span data-ttu-id="ad140-833">Описание</span><span class="sxs-lookup"><span data-stu-id="ad140-833">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="ad140-834">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="ad140-834">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook/office.mailboxenums.entitytype?view=outlook-js-1.2)|<span data-ttu-id="ad140-835">Одно из значений перечисления EntityType.</span><span class="sxs-lookup"><span data-stu-id="ad140-835">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ad140-836">Requirements</span><span class="sxs-lookup"><span data-stu-id="ad140-836">Requirements</span></span>

|<span data-ttu-id="ad140-837">Требование</span><span class="sxs-lookup"><span data-stu-id="ad140-837">Requirement</span></span>| <span data-ttu-id="ad140-838">Значение</span><span class="sxs-lookup"><span data-stu-id="ad140-838">Value</span></span>|
|---|---|
|[<span data-ttu-id="ad140-839">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="ad140-839">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ad140-840">1.0</span><span class="sxs-lookup"><span data-stu-id="ad140-840">1.0</span></span>|
|[<span data-ttu-id="ad140-841">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ad140-841">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ad140-842">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="ad140-842">Restricted</span></span>|
|[<span data-ttu-id="ad140-843">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ad140-843">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ad140-844">Чтение</span><span class="sxs-lookup"><span data-stu-id="ad140-844">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="ad140-845">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="ad140-845">Returns:</span></span>

<span data-ttu-id="ad140-846">Если значение, переданное в `entityType`, не является допустимым членом перечисления `EntityType`, метод возвращает значение NULL.</span><span class="sxs-lookup"><span data-stu-id="ad140-846">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="ad140-847">Если в теле элемента отсутствуют сущности указанного типа, метод возвращает пустой массив.</span><span class="sxs-lookup"><span data-stu-id="ad140-847">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="ad140-848">В противном случае тип объектов в возвращаемом массиве зависит от типа сущности, запрошенной в параметре `entityType`.</span><span class="sxs-lookup"><span data-stu-id="ad140-848">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="ad140-849">Хотя минимальный уровень разрешений для использования этого метода — **Restricted**, для некоторых типов сущностей требуется доступ на уровне **ReadItem**, как указано в приведенной ниже таблице.</span><span class="sxs-lookup"><span data-stu-id="ad140-849">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="ad140-850">Значение параметра `entityType`</span><span class="sxs-lookup"><span data-stu-id="ad140-850">Value of `entityType`</span></span> | <span data-ttu-id="ad140-851">Тип объектов в возвращаемом массиве</span><span class="sxs-lookup"><span data-stu-id="ad140-851">Type of objects in returned array</span></span> | <span data-ttu-id="ad140-852">Необходимый уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ad140-852">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="ad140-853">String</span><span class="sxs-lookup"><span data-stu-id="ad140-853">String</span></span> | <span data-ttu-id="ad140-854">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="ad140-854">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="ad140-855">Contact</span><span class="sxs-lookup"><span data-stu-id="ad140-855">Contact</span></span> | <span data-ttu-id="ad140-856">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="ad140-856">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="ad140-857">String</span><span class="sxs-lookup"><span data-stu-id="ad140-857">String</span></span> | <span data-ttu-id="ad140-858">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="ad140-858">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="ad140-859">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="ad140-859">MeetingSuggestion</span></span> | <span data-ttu-id="ad140-860">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="ad140-860">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="ad140-861">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="ad140-861">PhoneNumber</span></span> | <span data-ttu-id="ad140-862">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="ad140-862">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="ad140-863">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="ad140-863">TaskSuggestion</span></span> | <span data-ttu-id="ad140-864">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="ad140-864">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="ad140-865">String</span><span class="sxs-lookup"><span data-stu-id="ad140-865">String</span></span> | <span data-ttu-id="ad140-866">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="ad140-866">**Restricted**</span></span> |

<span data-ttu-id="ad140-867">Тип: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.2)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.2)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.2)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.2))></span><span class="sxs-lookup"><span data-stu-id="ad140-867">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.2)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.2)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.2)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.2))></span></span>

##### <a name="example"></a><span data-ttu-id="ad140-868">Пример</span><span class="sxs-lookup"><span data-stu-id="ad140-868">Example</span></span>

<span data-ttu-id="ad140-869">В примере ниже показано, как получить доступ к массиву строк, которые представляют собой почтовые адреса в теле текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="ad140-869">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-12meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-12phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-12tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-12"></a><span data-ttu-id="ad140-870">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.2)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.2)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.2)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.2))>}</span><span class="sxs-lookup"><span data-stu-id="ad140-870">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.2)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.2)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.2)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.2))>}</span></span>

<span data-ttu-id="ad140-871">Возвращает известные сущности в выбранном элементе, которые проходят через именованный фильтр, определяемый в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="ad140-871">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="ad140-872">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="ad140-872">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="ad140-873">Метод `getFilteredEntitiesByName` возвращает сущности, соответствующие регулярному выражению, которое определяется в элементе правила [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) в XML-файле манифеста, с использованием указанного значения элемента `FilterName`.</span><span class="sxs-lookup"><span data-stu-id="ad140-873">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ad140-874">Параметры</span><span class="sxs-lookup"><span data-stu-id="ad140-874">Parameters</span></span>

|<span data-ttu-id="ad140-875">Имя</span><span class="sxs-lookup"><span data-stu-id="ad140-875">Name</span></span>| <span data-ttu-id="ad140-876">Тип</span><span class="sxs-lookup"><span data-stu-id="ad140-876">Type</span></span>| <span data-ttu-id="ad140-877">Описание</span><span class="sxs-lookup"><span data-stu-id="ad140-877">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="ad140-878">String</span><span class="sxs-lookup"><span data-stu-id="ad140-878">String</span></span>|<span data-ttu-id="ad140-879">Имя элемента правила `ItemHasKnownEntity`, определяющее соответствующий фильтр.</span><span class="sxs-lookup"><span data-stu-id="ad140-879">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ad140-880">Requirements</span><span class="sxs-lookup"><span data-stu-id="ad140-880">Requirements</span></span>

|<span data-ttu-id="ad140-881">Требование</span><span class="sxs-lookup"><span data-stu-id="ad140-881">Requirement</span></span>| <span data-ttu-id="ad140-882">Значение</span><span class="sxs-lookup"><span data-stu-id="ad140-882">Value</span></span>|
|---|---|
|[<span data-ttu-id="ad140-883">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="ad140-883">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ad140-884">1.0</span><span class="sxs-lookup"><span data-stu-id="ad140-884">1.0</span></span>|
|[<span data-ttu-id="ad140-885">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ad140-885">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ad140-886">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ad140-886">ReadItem</span></span>|
|[<span data-ttu-id="ad140-887">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ad140-887">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ad140-888">Чтение</span><span class="sxs-lookup"><span data-stu-id="ad140-888">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="ad140-889">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="ad140-889">Returns:</span></span>

<span data-ttu-id="ad140-p160">Если в манифесте нет элемента `ItemHasKnownEntity` со значением `FilterName`, соответствующим параметру `name`, метод возвращает `null`. Если параметр `name` соответствует элементу `ItemHasKnownEntity` в манифесте, но при этом в текущем элементе нет соответствующих сущностей, метод возвращает пустой массив.</span><span class="sxs-lookup"><span data-stu-id="ad140-p160">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="ad140-892">Тип: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.2)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.2)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.2)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.2))></span><span class="sxs-lookup"><span data-stu-id="ad140-892">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.2)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.2)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.2)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.2))></span></span>

<br>

---
---

#### <a name="getregexmatches--object"></a><span data-ttu-id="ad140-893">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="ad140-893">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="ad140-894">Возвращает строковые значения в выбранном элементе, которые соответствуют регулярным выражениям, определенным в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="ad140-894">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="ad140-895">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="ad140-895">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="ad140-p161">Метод `getRegExMatches` возвращает строки, соответствующие регулярному выражению, которое определяется в каждом элементе правила `ItemHasRegularExpressionMatch` или `ItemHasKnownEntity` в XML-файле манифеста. Для правила `ItemHasRegularExpressionMatch` соответствующую строку должно содержать свойство элемента, указанного этим правилом. Простой тип `PropertyName` определяет поддерживаемые свойства.</span><span class="sxs-lookup"><span data-stu-id="ad140-p161">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="ad140-899">Например, рассмотрим манифест надстройки, который содержит указанный ниже элемент `Rule`.</span><span class="sxs-lookup"><span data-stu-id="ad140-899">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="ad140-900">Объект, возвращаемый методом `getRegExMatches`, будет содержать два свойства: `fruits` и `veggies`.</span><span class="sxs-lookup"><span data-stu-id="ad140-900">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

> [!NOTE]
> <span data-ttu-id="ad140-p162">Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты.</span><span class="sxs-lookup"><span data-stu-id="ad140-p162">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="requirements"></a><span data-ttu-id="ad140-903">Requirements</span><span class="sxs-lookup"><span data-stu-id="ad140-903">Requirements</span></span>

|<span data-ttu-id="ad140-904">Требование</span><span class="sxs-lookup"><span data-stu-id="ad140-904">Requirement</span></span>| <span data-ttu-id="ad140-905">Значение</span><span class="sxs-lookup"><span data-stu-id="ad140-905">Value</span></span>|
|---|---|
|[<span data-ttu-id="ad140-906">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="ad140-906">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ad140-907">1.0</span><span class="sxs-lookup"><span data-stu-id="ad140-907">1.0</span></span>|
|[<span data-ttu-id="ad140-908">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ad140-908">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ad140-909">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ad140-909">ReadItem</span></span>|
|[<span data-ttu-id="ad140-910">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ad140-910">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ad140-911">Чтение</span><span class="sxs-lookup"><span data-stu-id="ad140-911">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="ad140-912">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="ad140-912">Returns:</span></span>

<span data-ttu-id="ad140-p163">Объект, содержащий массив строк, которые соответствуют регулярным выражениям, определяемым в XML-файле манифеста. Имя каждого массива равно соответствующему значению атрибута `RegExName` подходящего правила `ItemHasRegularExpressionMatch` или атрибута `FilterName` соответствующего правила `ItemHasKnownEntity`.</span><span class="sxs-lookup"><span data-stu-id="ad140-p163">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<span data-ttu-id="ad140-915">Тип: Object</span><span class="sxs-lookup"><span data-stu-id="ad140-915">Type: Object</span></span>

##### <a name="example"></a><span data-ttu-id="ad140-916">Пример</span><span class="sxs-lookup"><span data-stu-id="ad140-916">Example</span></span>

<span data-ttu-id="ad140-917">В примере ниже показано, как получить доступ к массиву совпадений для <rule>элементов регулярного выражения `fruits` и `veggies`, которые указаны в манифесте</rule>.</span><span class="sxs-lookup"><span data-stu-id="ad140-917">The following example shows how to access the array of matches for the regular expression <rule>elements `fruits` and `veggies`, which are specified in the manifest.</rule></span></span>

```js
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

<br>

---
---

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="ad140-918">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="ad140-918">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="ad140-919">Возвращает строковые значения в выбранном элементе, которые соответствуют именованному регулярному выражению, определенному в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="ad140-919">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="ad140-920">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="ad140-920">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="ad140-921">Метод `getRegExMatchesByName` возвращает строки, соответствующие регулярному выражению, которое определяется в элементе правила `ItemHasRegularExpressionMatch` в XML-файле манифеста, с использованием указанного значения элемента `RegExName`.</span><span class="sxs-lookup"><span data-stu-id="ad140-921">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="ad140-p164">Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты.</span><span class="sxs-lookup"><span data-stu-id="ad140-p164">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ad140-924">Параметры</span><span class="sxs-lookup"><span data-stu-id="ad140-924">Parameters</span></span>

|<span data-ttu-id="ad140-925">Имя</span><span class="sxs-lookup"><span data-stu-id="ad140-925">Name</span></span>| <span data-ttu-id="ad140-926">Тип</span><span class="sxs-lookup"><span data-stu-id="ad140-926">Type</span></span>| <span data-ttu-id="ad140-927">Описание</span><span class="sxs-lookup"><span data-stu-id="ad140-927">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="ad140-928">String</span><span class="sxs-lookup"><span data-stu-id="ad140-928">String</span></span>|<span data-ttu-id="ad140-929">Имя элемента правила `ItemHasRegularExpressionMatch`, определяющее соответствующий фильтр.</span><span class="sxs-lookup"><span data-stu-id="ad140-929">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ad140-930">Requirements</span><span class="sxs-lookup"><span data-stu-id="ad140-930">Requirements</span></span>

|<span data-ttu-id="ad140-931">Требование</span><span class="sxs-lookup"><span data-stu-id="ad140-931">Requirement</span></span>| <span data-ttu-id="ad140-932">Значение</span><span class="sxs-lookup"><span data-stu-id="ad140-932">Value</span></span>|
|---|---|
|[<span data-ttu-id="ad140-933">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="ad140-933">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ad140-934">1.0</span><span class="sxs-lookup"><span data-stu-id="ad140-934">1.0</span></span>|
|[<span data-ttu-id="ad140-935">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ad140-935">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ad140-936">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ad140-936">ReadItem</span></span>|
|[<span data-ttu-id="ad140-937">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ad140-937">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ad140-938">Чтение</span><span class="sxs-lookup"><span data-stu-id="ad140-938">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="ad140-939">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="ad140-939">Returns:</span></span>

<span data-ttu-id="ad140-940">Массив строк, соответствующих регулярному выражению, определяемому в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="ad140-940">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<span data-ttu-id="ad140-941">Тип: Array.< String ></span><span class="sxs-lookup"><span data-stu-id="ad140-941">Type: Array.< String ></span></span>

##### <a name="example"></a><span data-ttu-id="ad140-942">Пример</span><span class="sxs-lookup"><span data-stu-id="ad140-942">Example</span></span>

```js
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

<br>

---
---

#### <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="ad140-943">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="ad140-943">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="ad140-944">Асинхронно возвращает данные, выбранные в теме или тексте сообщения.</span><span class="sxs-lookup"><span data-stu-id="ad140-944">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="ad140-p165">Если выделенный фрагмент отсутствует, но курсор находится в тексте или теме, метод возвращает пустую строку для выбранных данных. Если выбраны не текст и не тема, метод возвращает ошибку `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="ad140-p165">If there is no selection but the cursor is in the body or subject, the method returns an empty string for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ad140-947">Параметры</span><span class="sxs-lookup"><span data-stu-id="ad140-947">Parameters</span></span>

|<span data-ttu-id="ad140-948">Имя</span><span class="sxs-lookup"><span data-stu-id="ad140-948">Name</span></span>| <span data-ttu-id="ad140-949">Тип</span><span class="sxs-lookup"><span data-stu-id="ad140-949">Type</span></span>| <span data-ttu-id="ad140-950">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="ad140-950">Attributes</span></span>| <span data-ttu-id="ad140-951">Описание</span><span class="sxs-lookup"><span data-stu-id="ad140-951">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="ad140-952">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="ad140-952">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="ad140-p166">Запрашивает формат данных. Если задано значение Text, метод возвращает обычный текст как строку, удаляя все имеющиеся HTML-теги. Если задано значение HTML, метод возвращает выделенный текст (обычный текст или HTML).</span><span class="sxs-lookup"><span data-stu-id="ad140-p166">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="ad140-956">Object</span><span class="sxs-lookup"><span data-stu-id="ad140-956">Object</span></span>| <span data-ttu-id="ad140-957">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="ad140-957">&lt;optional&gt;</span></span>|<span data-ttu-id="ad140-958">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="ad140-958">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="ad140-959">Объект</span><span class="sxs-lookup"><span data-stu-id="ad140-959">Object</span></span>| <span data-ttu-id="ad140-960">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="ad140-960">&lt;optional&gt;</span></span>|<span data-ttu-id="ad140-961">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="ad140-961">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="ad140-962">функция</span><span class="sxs-lookup"><span data-stu-id="ad140-962">function</span></span>||<span data-ttu-id="ad140-963">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="ad140-963">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="ad140-964">Чтобы получить доступ к выбранным данным из метода обратного вызова, вызовите `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="ad140-964">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="ad140-965">Для доступа к исходному свойству, представляющему собой источник выбранных данных, вызовите параметр `asyncResult.value.sourceProperty`, который может иметь значение `body` или `subject`.</span><span class="sxs-lookup"><span data-stu-id="ad140-965">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ad140-966">Requirements</span><span class="sxs-lookup"><span data-stu-id="ad140-966">Requirements</span></span>

|<span data-ttu-id="ad140-967">Требование</span><span class="sxs-lookup"><span data-stu-id="ad140-967">Requirement</span></span>| <span data-ttu-id="ad140-968">Значение</span><span class="sxs-lookup"><span data-stu-id="ad140-968">Value</span></span>|
|---|---|
|[<span data-ttu-id="ad140-969">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="ad140-969">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ad140-970">1.2</span><span class="sxs-lookup"><span data-stu-id="ad140-970">1.2</span></span>|
|[<span data-ttu-id="ad140-971">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ad140-971">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ad140-972">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ad140-972">ReadItem</span></span>|
|[<span data-ttu-id="ad140-973">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ad140-973">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ad140-974">Создание</span><span class="sxs-lookup"><span data-stu-id="ad140-974">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="ad140-975">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="ad140-975">Returns:</span></span>

<span data-ttu-id="ad140-976">Выбранные данные в виде строки с форматом, определенным в параметре `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="ad140-976">The selected data as a string with format determined by `coercionType`.</span></span>

<span data-ttu-id="ad140-977">Тип: строка</span><span class="sxs-lookup"><span data-stu-id="ad140-977">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="ad140-978">Пример</span><span class="sxs-lookup"><span data-stu-id="ad140-978">Example</span></span>

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

#### <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="ad140-979">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="ad140-979">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="ad140-980">Асинхронно загружает настраиваемые свойства для надстройки для выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="ad140-980">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="ad140-p168">Настраиваемые свойства сохраняются в виде пар "ключ-значение" для каждого приложения и каждого элемента. Этот метод возвращает объект `CustomProperties` при обратном вызове, который предоставляет методы для доступа к настраиваемым свойствам, характерным для текущего элемента и текущей надстройки. Настраиваемые свойства не шифруются для элемента, поэтому этот способ хранения не является безопасным.</span><span class="sxs-lookup"><span data-stu-id="ad140-p168">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ad140-984">Параметры</span><span class="sxs-lookup"><span data-stu-id="ad140-984">Parameters</span></span>

|<span data-ttu-id="ad140-985">Имя</span><span class="sxs-lookup"><span data-stu-id="ad140-985">Name</span></span>| <span data-ttu-id="ad140-986">Тип</span><span class="sxs-lookup"><span data-stu-id="ad140-986">Type</span></span>| <span data-ttu-id="ad140-987">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="ad140-987">Attributes</span></span>| <span data-ttu-id="ad140-988">Описание</span><span class="sxs-lookup"><span data-stu-id="ad140-988">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="ad140-989">function</span><span class="sxs-lookup"><span data-stu-id="ad140-989">function</span></span>||<span data-ttu-id="ad140-990">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="ad140-990">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="ad140-991">Настраиваемые свойства предоставляются в виде объекта [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.2) в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="ad140-991">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.2) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="ad140-992">Этот объект позволяет получить, задать и удалить настраиваемые свойства для элемента, а также сохранить изменения, внесенные в набор настраиваемых свойств, на сервере.</span><span class="sxs-lookup"><span data-stu-id="ad140-992">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="ad140-993">Объект</span><span class="sxs-lookup"><span data-stu-id="ad140-993">Object</span></span>| <span data-ttu-id="ad140-994">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="ad140-994">&lt;optional&gt;</span></span>|<span data-ttu-id="ad140-995">Разработчики могут указать любой объект, к которому необходимо получить доступ, в функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="ad140-995">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="ad140-996">Доступ к этому объекту можно получить с помощью свойства `asyncResult.asyncContext` в функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="ad140-996">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ad140-997">Требования</span><span class="sxs-lookup"><span data-stu-id="ad140-997">Requirements</span></span>

|<span data-ttu-id="ad140-998">Требование</span><span class="sxs-lookup"><span data-stu-id="ad140-998">Requirement</span></span>| <span data-ttu-id="ad140-999">Значение</span><span class="sxs-lookup"><span data-stu-id="ad140-999">Value</span></span>|
|---|---|
|[<span data-ttu-id="ad140-1000">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="ad140-1000">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ad140-1001">1.0</span><span class="sxs-lookup"><span data-stu-id="ad140-1001">1.0</span></span>|
|[<span data-ttu-id="ad140-1002">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ad140-1002">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ad140-1003">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ad140-1003">ReadItem</span></span>|
|[<span data-ttu-id="ad140-1004">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ad140-1004">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ad140-1005">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="ad140-1005">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ad140-1006">Пример</span><span class="sxs-lookup"><span data-stu-id="ad140-1006">Example</span></span>

<span data-ttu-id="ad140-p171">Приведенный ниже пример кода показывает, как асинхронно загружать настраиваемые свойства, характерные для текущего элемента, с помощью метода `loadCustomPropertiesAsync`. Этот пример также показывает, как сохранять эти свойства на сервере с помощью метода `CustomProperties.saveAsync`. После загрузки настраиваемых свойств в этом примере кода метод `CustomProperties.get` используется для считывания настраиваемого свойства `myProp`, метод `CustomProperties.set` — для записи настраиваемого свойства `otherProp`, а метод `saveAsync` — для сохранения настраиваемых свойств.</span><span class="sxs-lookup"><span data-stu-id="ad140-p171">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

<br>

---
---

#### <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="ad140-1010">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="ad140-1010">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="ad140-1011">Удаляет вложение из сообщения или встречи.</span><span class="sxs-lookup"><span data-stu-id="ad140-1011">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="ad140-1012">Метод `removeAttachmentAsync` удаляет из элемента вложение с указанным идентификатором.</span><span class="sxs-lookup"><span data-stu-id="ad140-1012">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="ad140-1013">Идентификатор вложения рекомендуется использовать для удаления вложения, только если оно добавлено тем же почтовым приложением в ходе текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="ad140-1013">As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session.</span></span> <span data-ttu-id="ad140-1014">В Outlook в Интернете и на мобильных устройствах идентификатор вложения действителен только в течение одного сеанса.</span><span class="sxs-lookup"><span data-stu-id="ad140-1014">In Outlook on the web and mobile devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="ad140-1015">Сеанс завершается, когда пользователь закрывает приложение или начинает создавать элемент во встроенной форме, а затем переходит из формы в отдельное окно.</span><span class="sxs-lookup"><span data-stu-id="ad140-1015">A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ad140-1016">Параметры</span><span class="sxs-lookup"><span data-stu-id="ad140-1016">Parameters</span></span>

|<span data-ttu-id="ad140-1017">Имя</span><span class="sxs-lookup"><span data-stu-id="ad140-1017">Name</span></span>| <span data-ttu-id="ad140-1018">Тип</span><span class="sxs-lookup"><span data-stu-id="ad140-1018">Type</span></span>| <span data-ttu-id="ad140-1019">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="ad140-1019">Attributes</span></span>| <span data-ttu-id="ad140-1020">Описание</span><span class="sxs-lookup"><span data-stu-id="ad140-1020">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="ad140-1021">String</span><span class="sxs-lookup"><span data-stu-id="ad140-1021">String</span></span>||<span data-ttu-id="ad140-1022">Идентификатор удаляемого вложения.</span><span class="sxs-lookup"><span data-stu-id="ad140-1022">The identifier of the attachment to remove.</span></span>|
|`options`| <span data-ttu-id="ad140-1023">Объект</span><span class="sxs-lookup"><span data-stu-id="ad140-1023">Object</span></span>| <span data-ttu-id="ad140-1024">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="ad140-1024">&lt;optional&gt;</span></span>|<span data-ttu-id="ad140-1025">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="ad140-1025">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="ad140-1026">Объект</span><span class="sxs-lookup"><span data-stu-id="ad140-1026">Object</span></span>| <span data-ttu-id="ad140-1027">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="ad140-1027">&lt;optional&gt;</span></span>|<span data-ttu-id="ad140-1028">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="ad140-1028">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="ad140-1029">функция</span><span class="sxs-lookup"><span data-stu-id="ad140-1029">function</span></span>| <span data-ttu-id="ad140-1030">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="ad140-1030">&lt;optional&gt;</span></span>|<span data-ttu-id="ad140-1031">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="ad140-1031">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="ad140-1032">Если удалить вложение не удается, свойство `asyncResult.error` содержит код ошибки с указанием ее причины.</span><span class="sxs-lookup"><span data-stu-id="ad140-1032">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="ad140-1033">Ошибки</span><span class="sxs-lookup"><span data-stu-id="ad140-1033">Errors</span></span>

| <span data-ttu-id="ad140-1034">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="ad140-1034">Error code</span></span> | <span data-ttu-id="ad140-1035">Описание</span><span class="sxs-lookup"><span data-stu-id="ad140-1035">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="ad140-1036">Идентификатор вложения не существует.</span><span class="sxs-lookup"><span data-stu-id="ad140-1036">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="ad140-1037">Requirements</span><span class="sxs-lookup"><span data-stu-id="ad140-1037">Requirements</span></span>

|<span data-ttu-id="ad140-1038">Требование</span><span class="sxs-lookup"><span data-stu-id="ad140-1038">Requirement</span></span>| <span data-ttu-id="ad140-1039">Значение</span><span class="sxs-lookup"><span data-stu-id="ad140-1039">Value</span></span>|
|---|---|
|[<span data-ttu-id="ad140-1040">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="ad140-1040">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ad140-1041">1.1</span><span class="sxs-lookup"><span data-stu-id="ad140-1041">1.1</span></span>|
|[<span data-ttu-id="ad140-1042">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ad140-1042">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ad140-1043">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="ad140-1043">ReadWriteItem</span></span>|
|[<span data-ttu-id="ad140-1044">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ad140-1044">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ad140-1045">Создание</span><span class="sxs-lookup"><span data-stu-id="ad140-1045">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="ad140-1046">Пример</span><span class="sxs-lookup"><span data-stu-id="ad140-1046">Example</span></span>

<span data-ttu-id="ad140-1047">Указанный ниже код удаляет вложение с идентификатором "0".</span><span class="sxs-lookup"><span data-stu-id="ad140-1047">The following code removes an attachment with an identifier of '0'.</span></span>

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

#### <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="ad140-1048">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="ad140-1048">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="ad140-1049">Асинхронно вставляет данные в текст или тему сообщения.</span><span class="sxs-lookup"><span data-stu-id="ad140-1049">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="ad140-p173">Метод `setSelectedDataAsync` вставляет указанную строку в местоположение курсора в теме или тексте элемента либо, если текст выделен в редакторе, он заменяет выделенный текст. Если курсор находится вне текста или темы элемента, возвращается ошибка. После вставки курсор помещается в конец вставленного содержимого.</span><span class="sxs-lookup"><span data-stu-id="ad140-p173">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ad140-1053">Параметры</span><span class="sxs-lookup"><span data-stu-id="ad140-1053">Parameters</span></span>

|<span data-ttu-id="ad140-1054">Имя</span><span class="sxs-lookup"><span data-stu-id="ad140-1054">Name</span></span>| <span data-ttu-id="ad140-1055">Тип</span><span class="sxs-lookup"><span data-stu-id="ad140-1055">Type</span></span>| <span data-ttu-id="ad140-1056">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="ad140-1056">Attributes</span></span>| <span data-ttu-id="ad140-1057">Описание</span><span class="sxs-lookup"><span data-stu-id="ad140-1057">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="ad140-1058">String</span><span class="sxs-lookup"><span data-stu-id="ad140-1058">String</span></span>||<span data-ttu-id="ad140-p174">Вставляемые данные. Объем данных не должен превышать 1 000 000 символов. Если передано больше 1 000 000 символов, возвращается исключение `ArgumentOutOfRange`.</span><span class="sxs-lookup"><span data-stu-id="ad140-p174">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="ad140-1062">Object</span><span class="sxs-lookup"><span data-stu-id="ad140-1062">Object</span></span>| <span data-ttu-id="ad140-1063">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="ad140-1063">&lt;optional&gt;</span></span>|<span data-ttu-id="ad140-1064">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="ad140-1064">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="ad140-1065">Объект</span><span class="sxs-lookup"><span data-stu-id="ad140-1065">Object</span></span>| <span data-ttu-id="ad140-1066">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="ad140-1066">&lt;optional&gt;</span></span>|<span data-ttu-id="ad140-1067">В методе обратного вызова разработчики могут указать любой объект, к которому необходимо получить доступ.</span><span class="sxs-lookup"><span data-stu-id="ad140-1067">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="ad140-1068">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="ad140-1068">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="ad140-1069">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="ad140-1069">&lt;optional&gt;</span></span>|<span data-ttu-id="ad140-1070">Если задано значение `text`, текущий стиль применяется в Outlook в Интернете и классических клиентах.</span><span class="sxs-lookup"><span data-stu-id="ad140-1070">If `text`, the current style is applied in Outlook on the web and desktop clients.</span></span> <span data-ttu-id="ad140-1071">Если поле представляет собой редактор HTML, вставляются только текстовые данные, даже если они имеют формат HTML.</span><span class="sxs-lookup"><span data-stu-id="ad140-1071">If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="ad140-1072">Если задано значение `html` и поле (не тема) поддерживает HTML, в Outlook в Интернете применяется текущий стиль, а в классических клиентах Outlook — стиль по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="ad140-1072">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook on the web and the default style is applied in Outlook desktop clients.</span></span> <span data-ttu-id="ad140-1073">Если поле является текстовым, возвращается ошибка `InvalidDataFormat`.</span><span class="sxs-lookup"><span data-stu-id="ad140-1073">If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="ad140-1074">Если свойство `coercionType` не задано, результат зависит от поля: если поле имеет формат HTML, используется текст в формате HTML, а если поле текстовое, применяется обычный текст.</span><span class="sxs-lookup"><span data-stu-id="ad140-1074">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="ad140-1075">функция</span><span class="sxs-lookup"><span data-stu-id="ad140-1075">function</span></span>||<span data-ttu-id="ad140-1076">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="ad140-1076">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="ad140-1077">Требования</span><span class="sxs-lookup"><span data-stu-id="ad140-1077">Requirements</span></span>

|<span data-ttu-id="ad140-1078">Требование</span><span class="sxs-lookup"><span data-stu-id="ad140-1078">Requirement</span></span>| <span data-ttu-id="ad140-1079">Значение</span><span class="sxs-lookup"><span data-stu-id="ad140-1079">Value</span></span>|
|---|---|
|[<span data-ttu-id="ad140-1080">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="ad140-1080">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ad140-1081">1.2</span><span class="sxs-lookup"><span data-stu-id="ad140-1081">1.2</span></span>|
|[<span data-ttu-id="ad140-1082">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ad140-1082">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ad140-1083">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="ad140-1083">ReadWriteItem</span></span>|
|[<span data-ttu-id="ad140-1084">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ad140-1084">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ad140-1085">Создание</span><span class="sxs-lookup"><span data-stu-id="ad140-1085">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="ad140-1086">Пример</span><span class="sxs-lookup"><span data-stu-id="ad140-1086">Example</span></span>

```js
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
