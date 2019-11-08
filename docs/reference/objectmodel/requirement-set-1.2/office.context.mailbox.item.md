---
title: Office. Context. Mailbox. Item — набор требований 1,2
description: ''
ms.date: 11/06/2019
localization_priority: Normal
ms.openlocfilehash: 50cc2bcf338d2fb2fee5e32e0cd408c72c138214
ms.sourcegitcommit: 08c0b9ff319c391922fa43d3c2e9783cf6b53b1b
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 11/08/2019
ms.locfileid: "38066272"
---
# <a name="item"></a><span data-ttu-id="5e27f-102">item</span><span class="sxs-lookup"><span data-stu-id="5e27f-102">item</span></span>

### <span data-ttu-id="5e27f-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md). item</span><span class="sxs-lookup"><span data-stu-id="5e27f-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md). item</span></span>

<span data-ttu-id="5e27f-p102">Пространство имен `item` используется для доступа к выбранному в данный момент сообщению, приглашению на собрание или описанию встречи. Вы можете определить тип пространства имен `item` с помощью свойства [itemType](#itemtype-officemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="5e27f-p102">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="5e27f-107">Requirements</span><span class="sxs-lookup"><span data-stu-id="5e27f-107">Requirements</span></span>

|<span data-ttu-id="5e27f-108">Требование</span><span class="sxs-lookup"><span data-stu-id="5e27f-108">Requirement</span></span>| <span data-ttu-id="5e27f-109">Значение</span><span class="sxs-lookup"><span data-stu-id="5e27f-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="5e27f-110">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5e27f-110">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5e27f-111">1.0</span><span class="sxs-lookup"><span data-stu-id="5e27f-111">1.0</span></span>|
|[<span data-ttu-id="5e27f-112">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5e27f-112">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5e27f-113">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="5e27f-113">Restricted</span></span>|
|[<span data-ttu-id="5e27f-114">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5e27f-114">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="5e27f-115">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="5e27f-115">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="5e27f-116">Элементы и методы</span><span class="sxs-lookup"><span data-stu-id="5e27f-116">Members and methods</span></span>

| <span data-ttu-id="5e27f-117">Элемент	</span><span class="sxs-lookup"><span data-stu-id="5e27f-117">Member</span></span> | <span data-ttu-id="5e27f-118">Тип</span><span class="sxs-lookup"><span data-stu-id="5e27f-118">Type</span></span> |
|--------|------|
| [<span data-ttu-id="5e27f-119">attachments</span><span class="sxs-lookup"><span data-stu-id="5e27f-119">attachments</span></span>](#attachments-arrayattachmentdetails) | <span data-ttu-id="5e27f-120">Элемент</span><span class="sxs-lookup"><span data-stu-id="5e27f-120">Member</span></span> |
| [<span data-ttu-id="5e27f-121">bcc</span><span class="sxs-lookup"><span data-stu-id="5e27f-121">bcc</span></span>](#bcc-recipients) | <span data-ttu-id="5e27f-122">Элемент</span><span class="sxs-lookup"><span data-stu-id="5e27f-122">Member</span></span> |
| [<span data-ttu-id="5e27f-123">body</span><span class="sxs-lookup"><span data-stu-id="5e27f-123">body</span></span>](#body-body) | <span data-ttu-id="5e27f-124">Элемент</span><span class="sxs-lookup"><span data-stu-id="5e27f-124">Member</span></span> |
| [<span data-ttu-id="5e27f-125">cc</span><span class="sxs-lookup"><span data-stu-id="5e27f-125">cc</span></span>](#cc-arrayemailaddressdetailsrecipients) | <span data-ttu-id="5e27f-126">Элемент</span><span class="sxs-lookup"><span data-stu-id="5e27f-126">Member</span></span> |
| [<span data-ttu-id="5e27f-127">conversationId</span><span class="sxs-lookup"><span data-stu-id="5e27f-127">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="5e27f-128">Элемент</span><span class="sxs-lookup"><span data-stu-id="5e27f-128">Member</span></span> |
| [<span data-ttu-id="5e27f-129">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="5e27f-129">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="5e27f-130">Элемент</span><span class="sxs-lookup"><span data-stu-id="5e27f-130">Member</span></span> |
| [<span data-ttu-id="5e27f-131">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="5e27f-131">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="5e27f-132">Элемент</span><span class="sxs-lookup"><span data-stu-id="5e27f-132">Member</span></span> |
| [<span data-ttu-id="5e27f-133">end</span><span class="sxs-lookup"><span data-stu-id="5e27f-133">end</span></span>](#end-datetime) | <span data-ttu-id="5e27f-134">Элемент</span><span class="sxs-lookup"><span data-stu-id="5e27f-134">Member</span></span> |
| [<span data-ttu-id="5e27f-135">from</span><span class="sxs-lookup"><span data-stu-id="5e27f-135">from</span></span>](#from-emailaddressdetails) | <span data-ttu-id="5e27f-136">Элемент</span><span class="sxs-lookup"><span data-stu-id="5e27f-136">Member</span></span> |
| [<span data-ttu-id="5e27f-137">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="5e27f-137">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="5e27f-138">Элемент</span><span class="sxs-lookup"><span data-stu-id="5e27f-138">Member</span></span> |
| [<span data-ttu-id="5e27f-139">itemClass</span><span class="sxs-lookup"><span data-stu-id="5e27f-139">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="5e27f-140">Элемент</span><span class="sxs-lookup"><span data-stu-id="5e27f-140">Member</span></span> |
| [<span data-ttu-id="5e27f-141">itemId</span><span class="sxs-lookup"><span data-stu-id="5e27f-141">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="5e27f-142">Элемент</span><span class="sxs-lookup"><span data-stu-id="5e27f-142">Member</span></span> |
| [<span data-ttu-id="5e27f-143">itemType</span><span class="sxs-lookup"><span data-stu-id="5e27f-143">itemType</span></span>](#itemtype-officemailboxenumsitemtype) | <span data-ttu-id="5e27f-144">Элемент</span><span class="sxs-lookup"><span data-stu-id="5e27f-144">Member</span></span> |
| [<span data-ttu-id="5e27f-145">location</span><span class="sxs-lookup"><span data-stu-id="5e27f-145">location</span></span>](#location-stringlocation) | <span data-ttu-id="5e27f-146">Элемент</span><span class="sxs-lookup"><span data-stu-id="5e27f-146">Member</span></span> |
| [<span data-ttu-id="5e27f-147">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="5e27f-147">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="5e27f-148">Элемент</span><span class="sxs-lookup"><span data-stu-id="5e27f-148">Member</span></span> |
| [<span data-ttu-id="5e27f-149">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="5e27f-149">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="5e27f-150">Элемент</span><span class="sxs-lookup"><span data-stu-id="5e27f-150">Member</span></span> |
| [<span data-ttu-id="5e27f-151">organizer</span><span class="sxs-lookup"><span data-stu-id="5e27f-151">organizer</span></span>](#organizer-emailaddressdetails) | <span data-ttu-id="5e27f-152">Элемент</span><span class="sxs-lookup"><span data-stu-id="5e27f-152">Member</span></span> |
| [<span data-ttu-id="5e27f-153">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="5e27f-153">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="5e27f-154">Member</span><span class="sxs-lookup"><span data-stu-id="5e27f-154">Member</span></span> |
| [<span data-ttu-id="5e27f-155">sender</span><span class="sxs-lookup"><span data-stu-id="5e27f-155">sender</span></span>](#sender-emailaddressdetails) | <span data-ttu-id="5e27f-156">Элемент</span><span class="sxs-lookup"><span data-stu-id="5e27f-156">Member</span></span> |
| [<span data-ttu-id="5e27f-157">start</span><span class="sxs-lookup"><span data-stu-id="5e27f-157">start</span></span>](#start-datetime) | <span data-ttu-id="5e27f-158">Элемент</span><span class="sxs-lookup"><span data-stu-id="5e27f-158">Member</span></span> |
| [<span data-ttu-id="5e27f-159">subject</span><span class="sxs-lookup"><span data-stu-id="5e27f-159">subject</span></span>](#subject-stringsubject) | <span data-ttu-id="5e27f-160">Элемент</span><span class="sxs-lookup"><span data-stu-id="5e27f-160">Member</span></span> |
| [<span data-ttu-id="5e27f-161">to</span><span class="sxs-lookup"><span data-stu-id="5e27f-161">to</span></span>](#to-arrayemailaddressdetailsrecipients) | <span data-ttu-id="5e27f-162">Элемент</span><span class="sxs-lookup"><span data-stu-id="5e27f-162">Member</span></span> |
| [<span data-ttu-id="5e27f-163">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="5e27f-163">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="5e27f-164">Метод</span><span class="sxs-lookup"><span data-stu-id="5e27f-164">Method</span></span> |
| [<span data-ttu-id="5e27f-165">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="5e27f-165">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="5e27f-166">Метод</span><span class="sxs-lookup"><span data-stu-id="5e27f-166">Method</span></span> |
| [<span data-ttu-id="5e27f-167">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="5e27f-167">displayReplyAllForm</span></span>](#displayreplyallformformdata-callback) | <span data-ttu-id="5e27f-168">Метод</span><span class="sxs-lookup"><span data-stu-id="5e27f-168">Method</span></span> |
| [<span data-ttu-id="5e27f-169">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="5e27f-169">displayReplyForm</span></span>](#displayreplyformformdata-callback) | <span data-ttu-id="5e27f-170">Метод</span><span class="sxs-lookup"><span data-stu-id="5e27f-170">Method</span></span> |
| [<span data-ttu-id="5e27f-171">getEntities</span><span class="sxs-lookup"><span data-stu-id="5e27f-171">getEntities</span></span>](#getentities--entities) | <span data-ttu-id="5e27f-172">Метод</span><span class="sxs-lookup"><span data-stu-id="5e27f-172">Method</span></span> |
| [<span data-ttu-id="5e27f-173">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="5e27f-173">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="5e27f-174">Метод</span><span class="sxs-lookup"><span data-stu-id="5e27f-174">Method</span></span> |
| [<span data-ttu-id="5e27f-175">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="5e27f-175">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="5e27f-176">Метод</span><span class="sxs-lookup"><span data-stu-id="5e27f-176">Method</span></span> |
| [<span data-ttu-id="5e27f-177">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="5e27f-177">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="5e27f-178">Метод</span><span class="sxs-lookup"><span data-stu-id="5e27f-178">Method</span></span> |
| [<span data-ttu-id="5e27f-179">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="5e27f-179">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="5e27f-180">Метод</span><span class="sxs-lookup"><span data-stu-id="5e27f-180">Method</span></span> |
| [<span data-ttu-id="5e27f-181">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="5e27f-181">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="5e27f-182">Метод</span><span class="sxs-lookup"><span data-stu-id="5e27f-182">Method</span></span> |
| [<span data-ttu-id="5e27f-183">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="5e27f-183">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="5e27f-184">Метод</span><span class="sxs-lookup"><span data-stu-id="5e27f-184">Method</span></span> |
| [<span data-ttu-id="5e27f-185">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="5e27f-185">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="5e27f-186">Метод</span><span class="sxs-lookup"><span data-stu-id="5e27f-186">Method</span></span> |
| [<span data-ttu-id="5e27f-187">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="5e27f-187">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="5e27f-188">Метод</span><span class="sxs-lookup"><span data-stu-id="5e27f-188">Method</span></span> |

### <a name="example"></a><span data-ttu-id="5e27f-189">Пример</span><span class="sxs-lookup"><span data-stu-id="5e27f-189">Example</span></span>

<span data-ttu-id="5e27f-190">В примере кода JavaScript, приведенном ниже, показано, как получить доступ к свойству `subject` текущего элемента в Outlook.</span><span class="sxs-lookup"><span data-stu-id="5e27f-190">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="5e27f-191">Members</span><span class="sxs-lookup"><span data-stu-id="5e27f-191">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetailsviewoutlook-js-12"></a><span data-ttu-id="5e27f-192">attachments: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.2)></span><span class="sxs-lookup"><span data-stu-id="5e27f-192">attachments: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.2)></span></span>

<span data-ttu-id="5e27f-p103">Получает массив вложений для элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="5e27f-p103">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="5e27f-195">Outlook блокирует определенные типы файлов из-за потенциальных проблем с безопасностью, поэтому они не возвращаются.</span><span class="sxs-lookup"><span data-stu-id="5e27f-195">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="5e27f-196">Дополнительные сведения см. в статье [Блокировка вложений в Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="5e27f-196">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="5e27f-197">Тип</span><span class="sxs-lookup"><span data-stu-id="5e27f-197">Type</span></span>

*   <span data-ttu-id="5e27f-198">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.2)></span><span class="sxs-lookup"><span data-stu-id="5e27f-198">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.2)></span></span>

##### <a name="requirements"></a><span data-ttu-id="5e27f-199">Требования</span><span class="sxs-lookup"><span data-stu-id="5e27f-199">Requirements</span></span>

|<span data-ttu-id="5e27f-200">Требование</span><span class="sxs-lookup"><span data-stu-id="5e27f-200">Requirement</span></span>| <span data-ttu-id="5e27f-201">Значение</span><span class="sxs-lookup"><span data-stu-id="5e27f-201">Value</span></span>|
|---|---|
|[<span data-ttu-id="5e27f-202">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5e27f-202">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5e27f-203">1.0</span><span class="sxs-lookup"><span data-stu-id="5e27f-203">1.0</span></span>|
|[<span data-ttu-id="5e27f-204">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5e27f-204">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5e27f-205">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5e27f-205">ReadItem</span></span>|
|[<span data-ttu-id="5e27f-206">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5e27f-206">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="5e27f-207">Чтение</span><span class="sxs-lookup"><span data-stu-id="5e27f-207">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="5e27f-208">Пример</span><span class="sxs-lookup"><span data-stu-id="5e27f-208">Example</span></span>

<span data-ttu-id="5e27f-209">С помощью приведенного ниже кода можно создать HTML-строку с подробными сведениями обо всех вложениях для текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="5e27f-209">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

#### <a name="bcc-recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-12"></a><span data-ttu-id="5e27f-210">bcc: [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="5e27f-210">bcc: [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

<span data-ttu-id="5e27f-211">Получает объект, который предоставляет методы для получения или обновления получателей скрытой копии сообщения.</span><span class="sxs-lookup"><span data-stu-id="5e27f-211">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="5e27f-212">Только в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="5e27f-212">Compose mode only.</span></span>

<span data-ttu-id="5e27f-213">Коллекция может включать не более 100 элементов по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="5e27f-213">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="5e27f-214">Однако для компьютеров с Windows и Mac OS действуют указанные ниже ограничения.</span><span class="sxs-lookup"><span data-stu-id="5e27f-214">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="5e27f-215">Максимальное количество элементов — 500.</span><span class="sxs-lookup"><span data-stu-id="5e27f-215">Get 500 members maximum.</span></span>
- <span data-ttu-id="5e27f-216">Установить ограничение количества элементов на вызов — не более 100, общего количества — не более 500.</span><span class="sxs-lookup"><span data-stu-id="5e27f-216">Set a maximum of 100 members per call, up to 500 members total.</span></span>

##### <a name="type"></a><span data-ttu-id="5e27f-217">Тип</span><span class="sxs-lookup"><span data-stu-id="5e27f-217">Type</span></span>

*   [<span data-ttu-id="5e27f-218">Получатели</span><span class="sxs-lookup"><span data-stu-id="5e27f-218">Recipients</span></span>](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)

##### <a name="requirements"></a><span data-ttu-id="5e27f-219">Requirements</span><span class="sxs-lookup"><span data-stu-id="5e27f-219">Requirements</span></span>

|<span data-ttu-id="5e27f-220">Требование</span><span class="sxs-lookup"><span data-stu-id="5e27f-220">Requirement</span></span>| <span data-ttu-id="5e27f-221">Значение</span><span class="sxs-lookup"><span data-stu-id="5e27f-221">Value</span></span>|
|---|---|
|[<span data-ttu-id="5e27f-222">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5e27f-222">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5e27f-223">1.1</span><span class="sxs-lookup"><span data-stu-id="5e27f-223">1.1</span></span>|
|[<span data-ttu-id="5e27f-224">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5e27f-224">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5e27f-225">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5e27f-225">ReadItem</span></span>|
|[<span data-ttu-id="5e27f-226">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5e27f-226">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="5e27f-227">Создание</span><span class="sxs-lookup"><span data-stu-id="5e27f-227">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="5e27f-228">Пример</span><span class="sxs-lookup"><span data-stu-id="5e27f-228">Example</span></span>

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

#### <a name="body-bodyjavascriptapioutlookofficebodyviewoutlook-js-12"></a><span data-ttu-id="5e27f-229">body: [Body](/javascript/api/outlook/office.body?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="5e27f-229">body: [Body](/javascript/api/outlook/office.body?view=outlook-js-1.2)</span></span>

<span data-ttu-id="5e27f-230">Получает объект, предоставляющий методы для работы с основным текстом элемента.</span><span class="sxs-lookup"><span data-stu-id="5e27f-230">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="5e27f-231">Тип</span><span class="sxs-lookup"><span data-stu-id="5e27f-231">Type</span></span>

*   [<span data-ttu-id="5e27f-232">Body</span><span class="sxs-lookup"><span data-stu-id="5e27f-232">Body</span></span>](/javascript/api/outlook/office.body?view=outlook-js-1.2)

##### <a name="requirements"></a><span data-ttu-id="5e27f-233">Requirements</span><span class="sxs-lookup"><span data-stu-id="5e27f-233">Requirements</span></span>

|<span data-ttu-id="5e27f-234">Требование</span><span class="sxs-lookup"><span data-stu-id="5e27f-234">Requirement</span></span>| <span data-ttu-id="5e27f-235">Значение</span><span class="sxs-lookup"><span data-stu-id="5e27f-235">Value</span></span>|
|---|---|
|[<span data-ttu-id="5e27f-236">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5e27f-236">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5e27f-237">1.1</span><span class="sxs-lookup"><span data-stu-id="5e27f-237">1.1</span></span>|
|[<span data-ttu-id="5e27f-238">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5e27f-238">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5e27f-239">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5e27f-239">ReadItem</span></span>|
|[<span data-ttu-id="5e27f-240">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5e27f-240">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="5e27f-241">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="5e27f-241">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="5e27f-242">Пример</span><span class="sxs-lookup"><span data-stu-id="5e27f-242">Example</span></span>

<span data-ttu-id="5e27f-243">В этом примере возвращается текст сообщения в формате обычного текста.</span><span class="sxs-lookup"><span data-stu-id="5e27f-243">This example gets the body of the message in plain text.</span></span>

```js
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="5e27f-244">Ниже приведен пример итогового параметра, переданного функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="5e27f-244">The following is an example of the result parameter passed to the callback function.</span></span>

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

#### <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-12recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-12"></a><span data-ttu-id="5e27f-245">cc: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="5e27f-245">cc: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

<span data-ttu-id="5e27f-246">Предоставляет доступ к получателям копии сообщения.</span><span class="sxs-lookup"><span data-stu-id="5e27f-246">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="5e27f-247">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="5e27f-247">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="5e27f-248">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="5e27f-248">Read mode</span></span>

<span data-ttu-id="5e27f-249">Свойство `cc` возвращает массив, который содержит объект `EmailAddressDetails` для каждого получателя, указанного в строке **Копия** сообщения.</span><span class="sxs-lookup"><span data-stu-id="5e27f-249">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message.</span></span> <span data-ttu-id="5e27f-250">Коллекция может включать не более 100 элементов по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="5e27f-250">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="5e27f-251">Однако на компьютерах с Windows и Mac OS может быть до 500 элементов.</span><span class="sxs-lookup"><span data-stu-id="5e27f-251">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="5e27f-252">Режим создания</span><span class="sxs-lookup"><span data-stu-id="5e27f-252">Compose mode</span></span>

<span data-ttu-id="5e27f-253">Свойство `cc` возвращает объект `Recipients`, предоставляющий методы для получения или обновления получателей, которые указаны в строке **Копия** сообщения.</span><span class="sxs-lookup"><span data-stu-id="5e27f-253">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span> <span data-ttu-id="5e27f-254">Коллекция может включать не более 100 элементов по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="5e27f-254">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="5e27f-255">Однако для компьютеров с Windows и Mac OS действуют указанные ниже ограничения.</span><span class="sxs-lookup"><span data-stu-id="5e27f-255">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="5e27f-256">Максимальное количество элементов — 500.</span><span class="sxs-lookup"><span data-stu-id="5e27f-256">Get 500 members maximum.</span></span>
- <span data-ttu-id="5e27f-257">Установить ограничение количества элементов на вызов — не более 100, общего количества — не более 500.</span><span class="sxs-lookup"><span data-stu-id="5e27f-257">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="5e27f-258">Тип</span><span class="sxs-lookup"><span data-stu-id="5e27f-258">Type</span></span>

*   <span data-ttu-id="5e27f-259">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="5e27f-259">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

##### <a name="requirements"></a><span data-ttu-id="5e27f-260">Requirements</span><span class="sxs-lookup"><span data-stu-id="5e27f-260">Requirements</span></span>

|<span data-ttu-id="5e27f-261">Требование</span><span class="sxs-lookup"><span data-stu-id="5e27f-261">Requirement</span></span>| <span data-ttu-id="5e27f-262">Значение</span><span class="sxs-lookup"><span data-stu-id="5e27f-262">Value</span></span>|
|---|---|
|[<span data-ttu-id="5e27f-263">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="5e27f-263">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5e27f-264">1.0</span><span class="sxs-lookup"><span data-stu-id="5e27f-264">1.0</span></span>|
|[<span data-ttu-id="5e27f-265">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5e27f-265">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5e27f-266">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5e27f-266">ReadItem</span></span>|
|[<span data-ttu-id="5e27f-267">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5e27f-267">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="5e27f-268">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="5e27f-268">Compose or Read</span></span>|

<br>

---
---

#### <a name="nullable-conversationid-string"></a><span data-ttu-id="5e27f-269">(nullable) conversationId: String</span><span class="sxs-lookup"><span data-stu-id="5e27f-269">(nullable) conversationId: String</span></span>

<span data-ttu-id="5e27f-270">Получает идентификатор разговора по электронной почте, содержащего конкретное сообщение.</span><span class="sxs-lookup"><span data-stu-id="5e27f-270">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="5e27f-p110">Вы можете получить целочисленное значение этого свойства, если ваше почтовое приложение активируется в формах просмотра или формах создания ответов. Если пользователь изменит тему ответа, после его отправки идентификатор беседы будет изменен, и полученное ранее значение будет недействительным.</span><span class="sxs-lookup"><span data-stu-id="5e27f-p110">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="5e27f-p111">Это свойство имеет значение NULL для нового элемента в форме создания. Свойство `conversationId` вернет значение, если пользователь задаст тему и сохранит элемент.</span><span class="sxs-lookup"><span data-stu-id="5e27f-p111">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="5e27f-275">Тип</span><span class="sxs-lookup"><span data-stu-id="5e27f-275">Type</span></span>

*   <span data-ttu-id="5e27f-276">String</span><span class="sxs-lookup"><span data-stu-id="5e27f-276">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="5e27f-277">Requirements</span><span class="sxs-lookup"><span data-stu-id="5e27f-277">Requirements</span></span>

|<span data-ttu-id="5e27f-278">Требование</span><span class="sxs-lookup"><span data-stu-id="5e27f-278">Requirement</span></span>| <span data-ttu-id="5e27f-279">Значение</span><span class="sxs-lookup"><span data-stu-id="5e27f-279">Value</span></span>|
|---|---|
|[<span data-ttu-id="5e27f-280">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="5e27f-280">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5e27f-281">1.0</span><span class="sxs-lookup"><span data-stu-id="5e27f-281">1.0</span></span>|
|[<span data-ttu-id="5e27f-282">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5e27f-282">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5e27f-283">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5e27f-283">ReadItem</span></span>|
|[<span data-ttu-id="5e27f-284">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5e27f-284">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="5e27f-285">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="5e27f-285">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="5e27f-286">Пример</span><span class="sxs-lookup"><span data-stu-id="5e27f-286">Example</span></span>

```js
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

<br>

---
---

#### <a name="datetimecreated-date"></a><span data-ttu-id="5e27f-287">dateTimeCreated: Date</span><span class="sxs-lookup"><span data-stu-id="5e27f-287">dateTimeCreated: Date</span></span>

<span data-ttu-id="5e27f-p112">Получает дату и время создания элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="5e27f-p112">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="5e27f-290">Тип</span><span class="sxs-lookup"><span data-stu-id="5e27f-290">Type</span></span>

*   <span data-ttu-id="5e27f-291">Дата</span><span class="sxs-lookup"><span data-stu-id="5e27f-291">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="5e27f-292">Требования</span><span class="sxs-lookup"><span data-stu-id="5e27f-292">Requirements</span></span>

|<span data-ttu-id="5e27f-293">Требование</span><span class="sxs-lookup"><span data-stu-id="5e27f-293">Requirement</span></span>| <span data-ttu-id="5e27f-294">Значение</span><span class="sxs-lookup"><span data-stu-id="5e27f-294">Value</span></span>|
|---|---|
|[<span data-ttu-id="5e27f-295">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="5e27f-295">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5e27f-296">1.0</span><span class="sxs-lookup"><span data-stu-id="5e27f-296">1.0</span></span>|
|[<span data-ttu-id="5e27f-297">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5e27f-297">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5e27f-298">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5e27f-298">ReadItem</span></span>|
|[<span data-ttu-id="5e27f-299">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5e27f-299">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="5e27f-300">Чтение</span><span class="sxs-lookup"><span data-stu-id="5e27f-300">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="5e27f-301">Пример</span><span class="sxs-lookup"><span data-stu-id="5e27f-301">Example</span></span>

```js
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

<br>

---
---

#### <a name="datetimemodified-date"></a><span data-ttu-id="5e27f-302">dateTimeModified: Date</span><span class="sxs-lookup"><span data-stu-id="5e27f-302">dateTimeModified: Date</span></span>

<span data-ttu-id="5e27f-p113">Получает дату и время последнего изменения элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="5e27f-p113">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="5e27f-305">Этот элемент не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="5e27f-305">This member is not supported in Outlook on iOS or Android.</span></span>

##### <a name="type"></a><span data-ttu-id="5e27f-306">Тип</span><span class="sxs-lookup"><span data-stu-id="5e27f-306">Type</span></span>

*   <span data-ttu-id="5e27f-307">Дата</span><span class="sxs-lookup"><span data-stu-id="5e27f-307">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="5e27f-308">Требования</span><span class="sxs-lookup"><span data-stu-id="5e27f-308">Requirements</span></span>

|<span data-ttu-id="5e27f-309">Требование</span><span class="sxs-lookup"><span data-stu-id="5e27f-309">Requirement</span></span>| <span data-ttu-id="5e27f-310">Значение</span><span class="sxs-lookup"><span data-stu-id="5e27f-310">Value</span></span>|
|---|---|
|[<span data-ttu-id="5e27f-311">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="5e27f-311">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5e27f-312">1.0</span><span class="sxs-lookup"><span data-stu-id="5e27f-312">1.0</span></span>|
|[<span data-ttu-id="5e27f-313">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5e27f-313">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5e27f-314">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5e27f-314">ReadItem</span></span>|
|[<span data-ttu-id="5e27f-315">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5e27f-315">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="5e27f-316">Чтение</span><span class="sxs-lookup"><span data-stu-id="5e27f-316">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="5e27f-317">Пример</span><span class="sxs-lookup"><span data-stu-id="5e27f-317">Example</span></span>

```js
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

<br>

---
---

#### <a name="end-datetimejavascriptapioutlookofficetimeviewoutlook-js-12"></a><span data-ttu-id="5e27f-318">end: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="5e27f-318">end: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.2)</span></span>

<span data-ttu-id="5e27f-319">Получает или задает дату и время окончания встречи.</span><span class="sxs-lookup"><span data-stu-id="5e27f-319">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="5e27f-p114">Свойство `end` представлено в виде значения даты и времени в формате UTC. Преобразовать значение свойства end в местные значения даты и времени клиента можно с помощью метода [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime).</span><span class="sxs-lookup"><span data-stu-id="5e27f-p114">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="5e27f-322">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="5e27f-322">Read mode</span></span>

<span data-ttu-id="5e27f-323">Свойство `end` возвращает объект `Date`.</span><span class="sxs-lookup"><span data-stu-id="5e27f-323">The `end` property returns a `Date` object.</span></span>

```js
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="5e27f-324">Режим создания</span><span class="sxs-lookup"><span data-stu-id="5e27f-324">Compose mode</span></span>

<span data-ttu-id="5e27f-325">Свойство `end` возвращает объект `Time`.</span><span class="sxs-lookup"><span data-stu-id="5e27f-325">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="5e27f-326">Если вы задаете время окончания с помощью метода [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.2#setasync-datetime--options--callback-), необходимо использовать метод [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) для преобразования местного времени на клиенте в формат UTC для сервера.</span><span class="sxs-lookup"><span data-stu-id="5e27f-326">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.2#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="5e27f-327">В примере ниже показано, как с помощью метода [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.2#setasync-datetime--options--callback-) объекта `Time` задать время окончания встречи.</span><span class="sxs-lookup"><span data-stu-id="5e27f-327">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.2#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="5e27f-328">Тип</span><span class="sxs-lookup"><span data-stu-id="5e27f-328">Type</span></span>

*   <span data-ttu-id="5e27f-329">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="5e27f-329">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.2)</span></span>

##### <a name="requirements"></a><span data-ttu-id="5e27f-330">Requirements</span><span class="sxs-lookup"><span data-stu-id="5e27f-330">Requirements</span></span>

|<span data-ttu-id="5e27f-331">Требование</span><span class="sxs-lookup"><span data-stu-id="5e27f-331">Requirement</span></span>| <span data-ttu-id="5e27f-332">Значение</span><span class="sxs-lookup"><span data-stu-id="5e27f-332">Value</span></span>|
|---|---|
|[<span data-ttu-id="5e27f-333">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5e27f-333">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5e27f-334">1.0</span><span class="sxs-lookup"><span data-stu-id="5e27f-334">1.0</span></span>|
|[<span data-ttu-id="5e27f-335">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5e27f-335">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5e27f-336">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5e27f-336">ReadItem</span></span>|
|[<span data-ttu-id="5e27f-337">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5e27f-337">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="5e27f-338">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="5e27f-338">Compose or Read</span></span>|

<br>

---
---

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-12"></a><span data-ttu-id="5e27f-339">from: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="5e27f-339">from: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)</span></span>

<span data-ttu-id="5e27f-p115">Получает электронный адрес отправителя сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="5e27f-p115">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="5e27f-p116">Свойства `from` и [`sender`](#sender-emailaddressdetails) представляют одно лицо, если сообщение не отправлено представителем. В противном случае свойство `from` представляет лицо, делегировавшее полномочия, а свойство sender — представителя.</span><span class="sxs-lookup"><span data-stu-id="5e27f-p116">The `from` and [`sender`](#sender-emailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="5e27f-344">Свойству `recipientType`, принадлежащему объекту `EmailAddressDetails` в свойстве `from`, задано значение `undefined`.</span><span class="sxs-lookup"><span data-stu-id="5e27f-344">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="5e27f-345">Тип</span><span class="sxs-lookup"><span data-stu-id="5e27f-345">Type</span></span>

*   [<span data-ttu-id="5e27f-346">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="5e27f-346">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)

##### <a name="requirements"></a><span data-ttu-id="5e27f-347">Requirements</span><span class="sxs-lookup"><span data-stu-id="5e27f-347">Requirements</span></span>

|<span data-ttu-id="5e27f-348">Требование</span><span class="sxs-lookup"><span data-stu-id="5e27f-348">Requirement</span></span>| <span data-ttu-id="5e27f-349">Значение</span><span class="sxs-lookup"><span data-stu-id="5e27f-349">Value</span></span>|
|---|---|
|[<span data-ttu-id="5e27f-350">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5e27f-350">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5e27f-351">1.0</span><span class="sxs-lookup"><span data-stu-id="5e27f-351">1.0</span></span>|
|[<span data-ttu-id="5e27f-352">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5e27f-352">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5e27f-353">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5e27f-353">ReadItem</span></span>|
|[<span data-ttu-id="5e27f-354">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5e27f-354">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="5e27f-355">Чтение</span><span class="sxs-lookup"><span data-stu-id="5e27f-355">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="5e27f-356">Пример</span><span class="sxs-lookup"><span data-stu-id="5e27f-356">Example</span></span>

```js
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

<br>

---
---

#### <a name="internetmessageid-string"></a><span data-ttu-id="5e27f-357">internetMessageId: String</span><span class="sxs-lookup"><span data-stu-id="5e27f-357">internetMessageId: String</span></span>

<span data-ttu-id="5e27f-p117">Получает идентификатор интернет-сообщения для электронного сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="5e27f-p117">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="5e27f-360">Тип</span><span class="sxs-lookup"><span data-stu-id="5e27f-360">Type</span></span>

*   <span data-ttu-id="5e27f-361">String</span><span class="sxs-lookup"><span data-stu-id="5e27f-361">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="5e27f-362">Requirements</span><span class="sxs-lookup"><span data-stu-id="5e27f-362">Requirements</span></span>

|<span data-ttu-id="5e27f-363">Требование</span><span class="sxs-lookup"><span data-stu-id="5e27f-363">Requirement</span></span>| <span data-ttu-id="5e27f-364">Значение</span><span class="sxs-lookup"><span data-stu-id="5e27f-364">Value</span></span>|
|---|---|
|[<span data-ttu-id="5e27f-365">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5e27f-365">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5e27f-366">1.0</span><span class="sxs-lookup"><span data-stu-id="5e27f-366">1.0</span></span>|
|[<span data-ttu-id="5e27f-367">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5e27f-367">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5e27f-368">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5e27f-368">ReadItem</span></span>|
|[<span data-ttu-id="5e27f-369">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5e27f-369">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="5e27f-370">Чтение</span><span class="sxs-lookup"><span data-stu-id="5e27f-370">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="5e27f-371">Пример</span><span class="sxs-lookup"><span data-stu-id="5e27f-371">Example</span></span>

```js
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

<br>

---
---

#### <a name="itemclass-string"></a><span data-ttu-id="5e27f-372">itemClass: String</span><span class="sxs-lookup"><span data-stu-id="5e27f-372">itemClass: String</span></span>

<span data-ttu-id="5e27f-p118">Получает класс элемента веб-служб Exchange для выбранного элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="5e27f-p118">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="5e27f-p119">Свойство `itemClass` указывает класс сообщения выбранного элемента. Ниже приводятся классы сообщения по умолчанию для элемента сообщения или встречи.</span><span class="sxs-lookup"><span data-stu-id="5e27f-p119">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="5e27f-377">Тип</span><span class="sxs-lookup"><span data-stu-id="5e27f-377">Type</span></span> | <span data-ttu-id="5e27f-378">Описание</span><span class="sxs-lookup"><span data-stu-id="5e27f-378">Description</span></span> | <span data-ttu-id="5e27f-379">Класс элемента</span><span class="sxs-lookup"><span data-stu-id="5e27f-379">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="5e27f-380">Элементы встречи</span><span class="sxs-lookup"><span data-stu-id="5e27f-380">Appointment items</span></span> | <span data-ttu-id="5e27f-381">Это элементы календаря для класса элемента `IPM.Appointment` или `IPM.Appointment.Occurrence`.</span><span class="sxs-lookup"><span data-stu-id="5e27f-381">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurrence` |
| <span data-ttu-id="5e27f-382">Элементы сообщения</span><span class="sxs-lookup"><span data-stu-id="5e27f-382">Message items</span></span> | <span data-ttu-id="5e27f-383">Сюда входят электронные сообщения, для которых по умолчанию задан класс сообщения `IPM.Note`, а также приглашения на собрания, ответы на них и уведомления об их отмене, использующие `IPM.Schedule.Meeting` в качестве базового класса сообщения.</span><span class="sxs-lookup"><span data-stu-id="5e27f-383">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="5e27f-384">Можно создавать настраиваемые классы сообщения, расширяющие классы сообщения по умолчанию, например настраиваемый класс сообщения о встрече `IPM.Appointment.Contoso`.</span><span class="sxs-lookup"><span data-stu-id="5e27f-384">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="5e27f-385">Тип</span><span class="sxs-lookup"><span data-stu-id="5e27f-385">Type</span></span>

*   <span data-ttu-id="5e27f-386">String</span><span class="sxs-lookup"><span data-stu-id="5e27f-386">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="5e27f-387">Requirements</span><span class="sxs-lookup"><span data-stu-id="5e27f-387">Requirements</span></span>

|<span data-ttu-id="5e27f-388">Требование</span><span class="sxs-lookup"><span data-stu-id="5e27f-388">Requirement</span></span>| <span data-ttu-id="5e27f-389">Значение</span><span class="sxs-lookup"><span data-stu-id="5e27f-389">Value</span></span>|
|---|---|
|[<span data-ttu-id="5e27f-390">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5e27f-390">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5e27f-391">1.0</span><span class="sxs-lookup"><span data-stu-id="5e27f-391">1.0</span></span>|
|[<span data-ttu-id="5e27f-392">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5e27f-392">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5e27f-393">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5e27f-393">ReadItem</span></span>|
|[<span data-ttu-id="5e27f-394">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5e27f-394">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="5e27f-395">Чтение</span><span class="sxs-lookup"><span data-stu-id="5e27f-395">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="5e27f-396">Пример</span><span class="sxs-lookup"><span data-stu-id="5e27f-396">Example</span></span>

```js
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

<br>

---
---

#### <a name="nullable-itemid-string"></a><span data-ttu-id="5e27f-397">(nullable) itemId: String</span><span class="sxs-lookup"><span data-stu-id="5e27f-397">(nullable) itemId: String</span></span>

<span data-ttu-id="5e27f-p120">Получает [идентификатор элемента веб-служб Exchange](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange) для текущего элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="5e27f-p120">Gets the [Exchange Web Services item identifier](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange) for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="5e27f-400">Идентификатор, возвращаемый свойством `itemId`, совпадает с [идентификатором элемента веб-служб Exchange](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange).</span><span class="sxs-lookup"><span data-stu-id="5e27f-400">The identifier returned by the `itemId` property is the same as the [Exchange Web Services item identifier](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange).</span></span> <span data-ttu-id="5e27f-401">Свойство `itemId` не совпадает с идентификатором записи Outlook, а также идентификатором, который используется REST API Outlook.</span><span class="sxs-lookup"><span data-stu-id="5e27f-401">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="5e27f-402">Перед выполнением вызовов API REST, использующих это значение, его `Office.context.mailbox.convertToRestId`необходимо преобразовать с помощью, которое доступно в наборе требований 1,3.</span><span class="sxs-lookup"><span data-stu-id="5e27f-402">Before making REST API calls using this value, it should be converted using `Office.context.mailbox.convertToRestId`, which is available starting in requirement set 1.3.</span></span> <span data-ttu-id="5e27f-403">Дополнительные сведения см. в статье [Использование REST API Outlook из надстройки Outlook](/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="5e27f-403">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

##### <a name="type"></a><span data-ttu-id="5e27f-404">Тип</span><span class="sxs-lookup"><span data-stu-id="5e27f-404">Type</span></span>

*   <span data-ttu-id="5e27f-405">String</span><span class="sxs-lookup"><span data-stu-id="5e27f-405">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="5e27f-406">Requirements</span><span class="sxs-lookup"><span data-stu-id="5e27f-406">Requirements</span></span>

|<span data-ttu-id="5e27f-407">Требование</span><span class="sxs-lookup"><span data-stu-id="5e27f-407">Requirement</span></span>| <span data-ttu-id="5e27f-408">Значение</span><span class="sxs-lookup"><span data-stu-id="5e27f-408">Value</span></span>|
|---|---|
|[<span data-ttu-id="5e27f-409">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5e27f-409">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5e27f-410">1.0</span><span class="sxs-lookup"><span data-stu-id="5e27f-410">1.0</span></span>|
|[<span data-ttu-id="5e27f-411">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5e27f-411">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5e27f-412">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5e27f-412">ReadItem</span></span>|
|[<span data-ttu-id="5e27f-413">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5e27f-413">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="5e27f-414">Чтение</span><span class="sxs-lookup"><span data-stu-id="5e27f-414">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="5e27f-415">Пример</span><span class="sxs-lookup"><span data-stu-id="5e27f-415">Example</span></span>

<span data-ttu-id="5e27f-p122">Указанный ниже код проверяет наличие идентификатора элемента. Если свойство `itemId` возвращает значение `null` или `undefined`, элемент будет сохранен в хранилище, а из асинхронного результата будет получен идентификатор элемента.</span><span class="sxs-lookup"><span data-stu-id="5e27f-p122">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

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

#### <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtypeviewoutlook-js-12"></a><span data-ttu-id="5e27f-418">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="5e27f-418">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.2)</span></span>

<span data-ttu-id="5e27f-419">Получает тип элемента, который представляет экземпляр.</span><span class="sxs-lookup"><span data-stu-id="5e27f-419">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="5e27f-420">Свойство `itemType` возвращает одно из значений перечисления `ItemType`, которое указывает, является ли экземпляр объекта `item` сообщением или встречей.</span><span class="sxs-lookup"><span data-stu-id="5e27f-420">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="5e27f-421">Тип</span><span class="sxs-lookup"><span data-stu-id="5e27f-421">Type</span></span>

*   [<span data-ttu-id="5e27f-422">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="5e27f-422">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.2)

##### <a name="requirements"></a><span data-ttu-id="5e27f-423">Требования</span><span class="sxs-lookup"><span data-stu-id="5e27f-423">Requirements</span></span>

|<span data-ttu-id="5e27f-424">Требование</span><span class="sxs-lookup"><span data-stu-id="5e27f-424">Requirement</span></span>| <span data-ttu-id="5e27f-425">Значение</span><span class="sxs-lookup"><span data-stu-id="5e27f-425">Value</span></span>|
|---|---|
|[<span data-ttu-id="5e27f-426">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5e27f-426">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5e27f-427">1.0</span><span class="sxs-lookup"><span data-stu-id="5e27f-427">1.0</span></span>|
|[<span data-ttu-id="5e27f-428">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5e27f-428">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5e27f-429">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5e27f-429">ReadItem</span></span>|
|[<span data-ttu-id="5e27f-430">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5e27f-430">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="5e27f-431">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="5e27f-431">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="5e27f-432">Пример</span><span class="sxs-lookup"><span data-stu-id="5e27f-432">Example</span></span>

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

#### <a name="location-stringlocationjavascriptapioutlookofficelocationviewoutlook-js-12"></a><span data-ttu-id="5e27f-433">location: String|[Location](/javascript/api/outlook/office.location?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="5e27f-433">location: String|[Location](/javascript/api/outlook/office.location?view=outlook-js-1.2)</span></span>

<span data-ttu-id="5e27f-434">Получает или задает место встречи.</span><span class="sxs-lookup"><span data-stu-id="5e27f-434">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="5e27f-435">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="5e27f-435">Read mode</span></span>

<span data-ttu-id="5e27f-436">Свойство `location` возвращает строку, содержащую сведения о месте встречи.</span><span class="sxs-lookup"><span data-stu-id="5e27f-436">The `location` property returns a string that contains the location of the appointment.</span></span>

```js
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="5e27f-437">Режим создания</span><span class="sxs-lookup"><span data-stu-id="5e27f-437">Compose mode</span></span>

<span data-ttu-id="5e27f-438">Свойство `location` возвращает объект `Location`, предоставляющий методы, которые используются для получения и задания места встречи.</span><span class="sxs-lookup"><span data-stu-id="5e27f-438">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```js
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="5e27f-439">Тип</span><span class="sxs-lookup"><span data-stu-id="5e27f-439">Type</span></span>

*   <span data-ttu-id="5e27f-440">String | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="5e27f-440">String | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.2)</span></span>

##### <a name="requirements"></a><span data-ttu-id="5e27f-441">Требования</span><span class="sxs-lookup"><span data-stu-id="5e27f-441">Requirements</span></span>

|<span data-ttu-id="5e27f-442">Требование</span><span class="sxs-lookup"><span data-stu-id="5e27f-442">Requirement</span></span>| <span data-ttu-id="5e27f-443">Значение</span><span class="sxs-lookup"><span data-stu-id="5e27f-443">Value</span></span>|
|---|---|
|[<span data-ttu-id="5e27f-444">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5e27f-444">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5e27f-445">1.0</span><span class="sxs-lookup"><span data-stu-id="5e27f-445">1.0</span></span>|
|[<span data-ttu-id="5e27f-446">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5e27f-446">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5e27f-447">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5e27f-447">ReadItem</span></span>|
|[<span data-ttu-id="5e27f-448">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5e27f-448">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="5e27f-449">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="5e27f-449">Compose or Read</span></span>|

<br>

---
---

#### <a name="normalizedsubject-string"></a><span data-ttu-id="5e27f-450">normalizedSubject: String</span><span class="sxs-lookup"><span data-stu-id="5e27f-450">normalizedSubject: String</span></span>

<span data-ttu-id="5e27f-p123">Получает тему элемента со всеми удаленными префиксами (включая `RE:` и `FWD:`). Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="5e27f-p123">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="5e27f-p124">Свойство normalizedSubject получает тему элемента со стандартными префиксами (такими как `RE:` и `FW:`), добавляемыми почтовыми программами. Для получения темы элемента с неизмененными префиксами используйте свойство [`subject`](#subject-stringsubject).</span><span class="sxs-lookup"><span data-stu-id="5e27f-p124">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="5e27f-455">Тип</span><span class="sxs-lookup"><span data-stu-id="5e27f-455">Type</span></span>

*   <span data-ttu-id="5e27f-456">String</span><span class="sxs-lookup"><span data-stu-id="5e27f-456">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="5e27f-457">Requirements</span><span class="sxs-lookup"><span data-stu-id="5e27f-457">Requirements</span></span>

|<span data-ttu-id="5e27f-458">Требование</span><span class="sxs-lookup"><span data-stu-id="5e27f-458">Requirement</span></span>| <span data-ttu-id="5e27f-459">Значение</span><span class="sxs-lookup"><span data-stu-id="5e27f-459">Value</span></span>|
|---|---|
|[<span data-ttu-id="5e27f-460">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5e27f-460">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5e27f-461">1.0</span><span class="sxs-lookup"><span data-stu-id="5e27f-461">1.0</span></span>|
|[<span data-ttu-id="5e27f-462">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5e27f-462">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5e27f-463">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5e27f-463">ReadItem</span></span>|
|[<span data-ttu-id="5e27f-464">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5e27f-464">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="5e27f-465">Чтение</span><span class="sxs-lookup"><span data-stu-id="5e27f-465">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="5e27f-466">Пример</span><span class="sxs-lookup"><span data-stu-id="5e27f-466">Example</span></span>

```js
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

<br>

---
---

#### <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-12recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-12"></a><span data-ttu-id="5e27f-467">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="5e27f-467">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

<span data-ttu-id="5e27f-468">Предоставляет доступ к необязательным участникам события.</span><span class="sxs-lookup"><span data-stu-id="5e27f-468">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="5e27f-469">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="5e27f-469">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="5e27f-470">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="5e27f-470">Read mode</span></span>

<span data-ttu-id="5e27f-471">Свойство `optionalAttendees` возвращает массив, содержащий объект `EmailAddressDetails` для каждого необязательного участника собрания.</span><span class="sxs-lookup"><span data-stu-id="5e27f-471">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span> <span data-ttu-id="5e27f-472">Коллекция может включать не более 100 элементов по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="5e27f-472">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="5e27f-473">Однако на компьютерах с Windows и Mac OS может быть до 500 элементов.</span><span class="sxs-lookup"><span data-stu-id="5e27f-473">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="5e27f-474">Режим создания</span><span class="sxs-lookup"><span data-stu-id="5e27f-474">Compose mode</span></span>

<span data-ttu-id="5e27f-475">Свойство `optionalAttendees` возвращает объект `Recipients`, который предоставляет методы для получения или обновления необязательных участников собрания.</span><span class="sxs-lookup"><span data-stu-id="5e27f-475">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span> <span data-ttu-id="5e27f-476">Коллекция может включать не более 100 элементов по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="5e27f-476">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="5e27f-477">Однако для компьютеров с Windows и Mac OS действуют указанные ниже ограничения.</span><span class="sxs-lookup"><span data-stu-id="5e27f-477">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="5e27f-478">Максимальное количество элементов — 500.</span><span class="sxs-lookup"><span data-stu-id="5e27f-478">Get 500 members maximum.</span></span>
- <span data-ttu-id="5e27f-479">Установить ограничение количества элементов на вызов — не более 100, общего количества — не более 500.</span><span class="sxs-lookup"><span data-stu-id="5e27f-479">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="5e27f-480">Тип</span><span class="sxs-lookup"><span data-stu-id="5e27f-480">Type</span></span>

*   <span data-ttu-id="5e27f-481">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="5e27f-481">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

##### <a name="requirements"></a><span data-ttu-id="5e27f-482">Requirements</span><span class="sxs-lookup"><span data-stu-id="5e27f-482">Requirements</span></span>

|<span data-ttu-id="5e27f-483">Требование</span><span class="sxs-lookup"><span data-stu-id="5e27f-483">Requirement</span></span>| <span data-ttu-id="5e27f-484">Значение</span><span class="sxs-lookup"><span data-stu-id="5e27f-484">Value</span></span>|
|---|---|
|[<span data-ttu-id="5e27f-485">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5e27f-485">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5e27f-486">1.0</span><span class="sxs-lookup"><span data-stu-id="5e27f-486">1.0</span></span>|
|[<span data-ttu-id="5e27f-487">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5e27f-487">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5e27f-488">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5e27f-488">ReadItem</span></span>|
|[<span data-ttu-id="5e27f-489">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5e27f-489">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="5e27f-490">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="5e27f-490">Compose or Read</span></span>|

<br>

---
---

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-12"></a><span data-ttu-id="5e27f-491">organizer: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="5e27f-491">organizer: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)</span></span>

<span data-ttu-id="5e27f-p128">Получает электронный адрес организатора указанного собрания. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="5e27f-p128">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="5e27f-494">Тип</span><span class="sxs-lookup"><span data-stu-id="5e27f-494">Type</span></span>

*   [<span data-ttu-id="5e27f-495">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="5e27f-495">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)

##### <a name="requirements"></a><span data-ttu-id="5e27f-496">Requirements</span><span class="sxs-lookup"><span data-stu-id="5e27f-496">Requirements</span></span>

|<span data-ttu-id="5e27f-497">Требование</span><span class="sxs-lookup"><span data-stu-id="5e27f-497">Requirement</span></span>| <span data-ttu-id="5e27f-498">Значение</span><span class="sxs-lookup"><span data-stu-id="5e27f-498">Value</span></span>|
|---|---|
|[<span data-ttu-id="5e27f-499">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5e27f-499">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5e27f-500">1.0</span><span class="sxs-lookup"><span data-stu-id="5e27f-500">1.0</span></span>|
|[<span data-ttu-id="5e27f-501">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5e27f-501">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5e27f-502">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5e27f-502">ReadItem</span></span>|
|[<span data-ttu-id="5e27f-503">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5e27f-503">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="5e27f-504">Чтение</span><span class="sxs-lookup"><span data-stu-id="5e27f-504">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="5e27f-505">Пример</span><span class="sxs-lookup"><span data-stu-id="5e27f-505">Example</span></span>

```js
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

<br>

---
---

#### <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-12recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-12"></a><span data-ttu-id="5e27f-506">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="5e27f-506">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

<span data-ttu-id="5e27f-507">Предоставляет доступ к обязательным участникам события.</span><span class="sxs-lookup"><span data-stu-id="5e27f-507">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="5e27f-508">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="5e27f-508">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="5e27f-509">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="5e27f-509">Read mode</span></span>

<span data-ttu-id="5e27f-510">Свойство `requiredAttendees` возвращает массив, содержащий объект `EmailAddressDetails` для каждого обязательного участника собрания.</span><span class="sxs-lookup"><span data-stu-id="5e27f-510">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span> <span data-ttu-id="5e27f-511">Коллекция может включать не более 100 элементов по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="5e27f-511">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="5e27f-512">Однако на компьютерах с Windows и Mac OS может быть до 500 элементов.</span><span class="sxs-lookup"><span data-stu-id="5e27f-512">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="5e27f-513">Режим создания</span><span class="sxs-lookup"><span data-stu-id="5e27f-513">Compose mode</span></span>

<span data-ttu-id="5e27f-514">Свойство `requiredAttendees` возвращает объект `Recipients`, предоставляющий методы, с помощью которых можно получить или обновить сведения об обязательных участниках собрания.</span><span class="sxs-lookup"><span data-stu-id="5e27f-514">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span> <span data-ttu-id="5e27f-515">Коллекция может включать не более 100 элементов по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="5e27f-515">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="5e27f-516">Однако для компьютеров с Windows и Mac OS действуют указанные ниже ограничения.</span><span class="sxs-lookup"><span data-stu-id="5e27f-516">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="5e27f-517">Максимальное количество элементов — 500.</span><span class="sxs-lookup"><span data-stu-id="5e27f-517">Get 500 members maximum.</span></span>
- <span data-ttu-id="5e27f-518">Установить ограничение количества элементов на вызов — не более 100, общего количества — не более 500.</span><span class="sxs-lookup"><span data-stu-id="5e27f-518">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="5e27f-519">Тип</span><span class="sxs-lookup"><span data-stu-id="5e27f-519">Type</span></span>

*   <span data-ttu-id="5e27f-520">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="5e27f-520">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

##### <a name="requirements"></a><span data-ttu-id="5e27f-521">Requirements</span><span class="sxs-lookup"><span data-stu-id="5e27f-521">Requirements</span></span>

|<span data-ttu-id="5e27f-522">Требование</span><span class="sxs-lookup"><span data-stu-id="5e27f-522">Requirement</span></span>| <span data-ttu-id="5e27f-523">Значение</span><span class="sxs-lookup"><span data-stu-id="5e27f-523">Value</span></span>|
|---|---|
|[<span data-ttu-id="5e27f-524">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5e27f-524">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5e27f-525">1.0</span><span class="sxs-lookup"><span data-stu-id="5e27f-525">1.0</span></span>|
|[<span data-ttu-id="5e27f-526">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5e27f-526">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5e27f-527">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5e27f-527">ReadItem</span></span>|
|[<span data-ttu-id="5e27f-528">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5e27f-528">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="5e27f-529">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="5e27f-529">Compose or Read</span></span>|

<br>

---
---

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-12"></a><span data-ttu-id="5e27f-530">sender: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="5e27f-530">sender: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)</span></span>

<span data-ttu-id="5e27f-p132">Получает электронный адрес отправителя электронного сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="5e27f-p132">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="5e27f-p133">Свойства [`from`](#from-emailaddressdetails) и `sender` представляют одно лицо, если сообщение не отправлено представителем. В противном случае свойство `from` представляет лицо, делегировавшее полномочия, а свойство sender — представителя.</span><span class="sxs-lookup"><span data-stu-id="5e27f-p133">The [`from`](#from-emailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="5e27f-535">Свойству `recipientType`, принадлежащему объекту `EmailAddressDetails` в свойстве `sender`, задано значение `undefined`.</span><span class="sxs-lookup"><span data-stu-id="5e27f-535">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="5e27f-536">Тип</span><span class="sxs-lookup"><span data-stu-id="5e27f-536">Type</span></span>

*   [<span data-ttu-id="5e27f-537">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="5e27f-537">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)

##### <a name="requirements"></a><span data-ttu-id="5e27f-538">Requirements</span><span class="sxs-lookup"><span data-stu-id="5e27f-538">Requirements</span></span>

|<span data-ttu-id="5e27f-539">Требование</span><span class="sxs-lookup"><span data-stu-id="5e27f-539">Requirement</span></span>| <span data-ttu-id="5e27f-540">Значение</span><span class="sxs-lookup"><span data-stu-id="5e27f-540">Value</span></span>|
|---|---|
|[<span data-ttu-id="5e27f-541">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5e27f-541">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5e27f-542">1.0</span><span class="sxs-lookup"><span data-stu-id="5e27f-542">1.0</span></span>|
|[<span data-ttu-id="5e27f-543">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5e27f-543">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5e27f-544">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5e27f-544">ReadItem</span></span>|
|[<span data-ttu-id="5e27f-545">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5e27f-545">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="5e27f-546">Чтение</span><span class="sxs-lookup"><span data-stu-id="5e27f-546">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="5e27f-547">Пример</span><span class="sxs-lookup"><span data-stu-id="5e27f-547">Example</span></span>

```js
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

<br>

---
---

#### <a name="start-datetimejavascriptapioutlookofficetimeviewoutlook-js-12"></a><span data-ttu-id="5e27f-548">start: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="5e27f-548">start: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.2)</span></span>

<span data-ttu-id="5e27f-549">Получает или задает дату и время начала встречи.</span><span class="sxs-lookup"><span data-stu-id="5e27f-549">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="5e27f-p134">Свойство `start` представлено в виде значения даты и времени в формате UTC. Это значение можно преобразовать в местные значения даты и времени клиента с помощью метода [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime).</span><span class="sxs-lookup"><span data-stu-id="5e27f-p134">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="5e27f-552">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="5e27f-552">Read mode</span></span>

<span data-ttu-id="5e27f-553">Свойство `start` возвращает объект `Date`.</span><span class="sxs-lookup"><span data-stu-id="5e27f-553">The `start` property returns a `Date` object.</span></span>

```js
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="5e27f-554">Режим создания</span><span class="sxs-lookup"><span data-stu-id="5e27f-554">Compose mode</span></span>

<span data-ttu-id="5e27f-555">Свойство `start` возвращает объект `Time`.</span><span class="sxs-lookup"><span data-stu-id="5e27f-555">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="5e27f-556">Если вы задаете время начала с помощью метода [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.2#setasync-datetime--options--callback-), необходимо использовать метод [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) для преобразования местного времени на клиенте в формат UTC для сервера.</span><span class="sxs-lookup"><span data-stu-id="5e27f-556">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.2#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>
<span data-ttu-id="5e27f-557">В примере ниже с помощью метода [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.2#setasync-datetime--options--callback-) объекта `Time` задается время начала встречи в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="5e27f-557">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.2#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="5e27f-558">Тип</span><span class="sxs-lookup"><span data-stu-id="5e27f-558">Type</span></span>

*   <span data-ttu-id="5e27f-559">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="5e27f-559">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.2)</span></span>

##### <a name="requirements"></a><span data-ttu-id="5e27f-560">Requirements</span><span class="sxs-lookup"><span data-stu-id="5e27f-560">Requirements</span></span>

|<span data-ttu-id="5e27f-561">Требование</span><span class="sxs-lookup"><span data-stu-id="5e27f-561">Requirement</span></span>| <span data-ttu-id="5e27f-562">Значение</span><span class="sxs-lookup"><span data-stu-id="5e27f-562">Value</span></span>|
|---|---|
|[<span data-ttu-id="5e27f-563">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="5e27f-563">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5e27f-564">1.0</span><span class="sxs-lookup"><span data-stu-id="5e27f-564">1.0</span></span>|
|[<span data-ttu-id="5e27f-565">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5e27f-565">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5e27f-566">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5e27f-566">ReadItem</span></span>|
|[<span data-ttu-id="5e27f-567">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5e27f-567">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="5e27f-568">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="5e27f-568">Compose or Read</span></span>|

<br>

---
---

#### <a name="subject-stringsubjectjavascriptapioutlookofficesubjectviewoutlook-js-12"></a><span data-ttu-id="5e27f-569">subject: String|[Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="5e27f-569">subject: String|[Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.2)</span></span>

<span data-ttu-id="5e27f-570">Получает или задает описание, которое отображается в поле темы элемента.</span><span class="sxs-lookup"><span data-stu-id="5e27f-570">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="5e27f-571">Свойство `subject` получает или задает всю тему элемента для отправки с почтового сервера.</span><span class="sxs-lookup"><span data-stu-id="5e27f-571">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="5e27f-572">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="5e27f-572">Read mode</span></span>

<span data-ttu-id="5e27f-p136">Свойство `subject` возвращает строку. С помощью свойства [`normalizedSubject`](#normalizedsubject-string) можно получить тему без начальных префиксов, таких как `RE:` и `FW:`.</span><span class="sxs-lookup"><span data-stu-id="5e27f-p136">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```js
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="5e27f-575">Режим создания</span><span class="sxs-lookup"><span data-stu-id="5e27f-575">Compose mode</span></span>

<span data-ttu-id="5e27f-576">Свойство `subject` возвращает объект `Subject`, который предоставляет методы для получения и задания темы.</span><span class="sxs-lookup"><span data-stu-id="5e27f-576">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```js
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="5e27f-577">Тип</span><span class="sxs-lookup"><span data-stu-id="5e27f-577">Type</span></span>

*   <span data-ttu-id="5e27f-578">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="5e27f-578">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.2)</span></span>

##### <a name="requirements"></a><span data-ttu-id="5e27f-579">Requirements</span><span class="sxs-lookup"><span data-stu-id="5e27f-579">Requirements</span></span>

|<span data-ttu-id="5e27f-580">Требование</span><span class="sxs-lookup"><span data-stu-id="5e27f-580">Requirement</span></span>| <span data-ttu-id="5e27f-581">Значение</span><span class="sxs-lookup"><span data-stu-id="5e27f-581">Value</span></span>|
|---|---|
|[<span data-ttu-id="5e27f-582">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="5e27f-582">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5e27f-583">1.0</span><span class="sxs-lookup"><span data-stu-id="5e27f-583">1.0</span></span>|
|[<span data-ttu-id="5e27f-584">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5e27f-584">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5e27f-585">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5e27f-585">ReadItem</span></span>|
|[<span data-ttu-id="5e27f-586">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5e27f-586">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="5e27f-587">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="5e27f-587">Compose or Read</span></span>|

<br>

---
---

#### <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-12recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-12"></a><span data-ttu-id="5e27f-588">to: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="5e27f-588">to: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

<span data-ttu-id="5e27f-589">Предоставляет доступ к получателям, указанным в строке **Кому** сообщения.</span><span class="sxs-lookup"><span data-stu-id="5e27f-589">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="5e27f-590">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="5e27f-590">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="5e27f-591">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="5e27f-591">Read mode</span></span>

<span data-ttu-id="5e27f-592">Свойство `to` возвращает массив, содержащий объект `EmailAddressDetails` для каждого получателя в строке **Кому** сообщения.</span><span class="sxs-lookup"><span data-stu-id="5e27f-592">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message.</span></span> <span data-ttu-id="5e27f-593">Коллекция может включать не более 100 элементов по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="5e27f-593">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="5e27f-594">Однако на компьютерах с Windows и Mac OS может быть до 500 элементов.</span><span class="sxs-lookup"><span data-stu-id="5e27f-594">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="5e27f-595">Режим создания</span><span class="sxs-lookup"><span data-stu-id="5e27f-595">Compose mode</span></span>

<span data-ttu-id="5e27f-596">Свойство `to` возвращает объект `Recipients`, предоставляющий методы для получения или обновления получателей, которые указаны в строке **Кому** сообщения.</span><span class="sxs-lookup"><span data-stu-id="5e27f-596">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span> <span data-ttu-id="5e27f-597">Коллекция может включать не более 100 элементов по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="5e27f-597">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="5e27f-598">Однако для компьютеров с Windows и Mac OS действуют указанные ниже ограничения.</span><span class="sxs-lookup"><span data-stu-id="5e27f-598">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="5e27f-599">Максимальное количество элементов — 500.</span><span class="sxs-lookup"><span data-stu-id="5e27f-599">Get 500 members maximum.</span></span>
- <span data-ttu-id="5e27f-600">Установить ограничение количества элементов на вызов — не более 100, общего количества — не более 500.</span><span class="sxs-lookup"><span data-stu-id="5e27f-600">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="5e27f-601">Тип</span><span class="sxs-lookup"><span data-stu-id="5e27f-601">Type</span></span>

*   <span data-ttu-id="5e27f-602">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="5e27f-602">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

##### <a name="requirements"></a><span data-ttu-id="5e27f-603">Requirements</span><span class="sxs-lookup"><span data-stu-id="5e27f-603">Requirements</span></span>

|<span data-ttu-id="5e27f-604">Требование</span><span class="sxs-lookup"><span data-stu-id="5e27f-604">Requirement</span></span>| <span data-ttu-id="5e27f-605">Значение</span><span class="sxs-lookup"><span data-stu-id="5e27f-605">Value</span></span>|
|---|---|
|[<span data-ttu-id="5e27f-606">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="5e27f-606">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5e27f-607">1.0</span><span class="sxs-lookup"><span data-stu-id="5e27f-607">1.0</span></span>|
|[<span data-ttu-id="5e27f-608">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5e27f-608">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5e27f-609">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5e27f-609">ReadItem</span></span>|
|[<span data-ttu-id="5e27f-610">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5e27f-610">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="5e27f-611">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="5e27f-611">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="5e27f-612">Методы</span><span class="sxs-lookup"><span data-stu-id="5e27f-612">Methods</span></span>

#### <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="5e27f-613">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="5e27f-613">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="5e27f-614">Добавляет файл в сообщение или встречу в качестве вложения.</span><span class="sxs-lookup"><span data-stu-id="5e27f-614">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="5e27f-615">Метод `addFileAttachmentAsync` передает файл по указанному универсальному коду ресурса (URI) и вкладывает его в элемент в форме создания.</span><span class="sxs-lookup"><span data-stu-id="5e27f-615">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="5e27f-616">Идентификатор можно использовать с методом [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback), чтобы удалить вложение, добавленное во время текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="5e27f-616">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="5e27f-617">Параметры</span><span class="sxs-lookup"><span data-stu-id="5e27f-617">Parameters</span></span>

|<span data-ttu-id="5e27f-618">Имя</span><span class="sxs-lookup"><span data-stu-id="5e27f-618">Name</span></span>| <span data-ttu-id="5e27f-619">Тип</span><span class="sxs-lookup"><span data-stu-id="5e27f-619">Type</span></span>| <span data-ttu-id="5e27f-620">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="5e27f-620">Attributes</span></span>| <span data-ttu-id="5e27f-621">Описание</span><span class="sxs-lookup"><span data-stu-id="5e27f-621">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="5e27f-622">String</span><span class="sxs-lookup"><span data-stu-id="5e27f-622">String</span></span>||<span data-ttu-id="5e27f-p140">Универсальный код ресурса (URI), представляющий расположение файла, который нужно вложить в сообщение или встречу. Максимальная длина — 2048 символов.</span><span class="sxs-lookup"><span data-stu-id="5e27f-p140">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="5e27f-625">String</span><span class="sxs-lookup"><span data-stu-id="5e27f-625">String</span></span>||<span data-ttu-id="5e27f-p141">Имя вложения, которое отображается при передаче вложения. Максимальная длина — 255 символов.</span><span class="sxs-lookup"><span data-stu-id="5e27f-p141">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="5e27f-628">Object</span><span class="sxs-lookup"><span data-stu-id="5e27f-628">Object</span></span>| <span data-ttu-id="5e27f-629">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="5e27f-629">&lt;optional&gt;</span></span>|<span data-ttu-id="5e27f-630">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="5e27f-630">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="5e27f-631">Object</span><span class="sxs-lookup"><span data-stu-id="5e27f-631">Object</span></span>| <span data-ttu-id="5e27f-632">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="5e27f-632">&lt;optional&gt;</span></span>|<span data-ttu-id="5e27f-633">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="5e27f-633">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="5e27f-634">функция</span><span class="sxs-lookup"><span data-stu-id="5e27f-634">function</span></span>| <span data-ttu-id="5e27f-635">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="5e27f-635">&lt;optional&gt;</span></span>|<span data-ttu-id="5e27f-636">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="5e27f-636">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="5e27f-637">После успешного выполнения идентификатор вложения будет представлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="5e27f-637">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="5e27f-638">Если добавить вложение не удастся, объект `asyncResult` будет содержать объект `Error` с описанием ошибки.</span><span class="sxs-lookup"><span data-stu-id="5e27f-638">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="5e27f-639">Ошибки</span><span class="sxs-lookup"><span data-stu-id="5e27f-639">Errors</span></span>

| <span data-ttu-id="5e27f-640">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="5e27f-640">Error code</span></span> | <span data-ttu-id="5e27f-641">Описание</span><span class="sxs-lookup"><span data-stu-id="5e27f-641">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="5e27f-642">Вложение превышает максимальный размер.</span><span class="sxs-lookup"><span data-stu-id="5e27f-642">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="5e27f-643">Расширение вложения не поддерживается.</span><span class="sxs-lookup"><span data-stu-id="5e27f-643">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="5e27f-644">Сообщение или встреча содержат слишком много вложений.</span><span class="sxs-lookup"><span data-stu-id="5e27f-644">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="5e27f-645">Требования</span><span class="sxs-lookup"><span data-stu-id="5e27f-645">Requirements</span></span>

|<span data-ttu-id="5e27f-646">Требование</span><span class="sxs-lookup"><span data-stu-id="5e27f-646">Requirement</span></span>| <span data-ttu-id="5e27f-647">Значение</span><span class="sxs-lookup"><span data-stu-id="5e27f-647">Value</span></span>|
|---|---|
|[<span data-ttu-id="5e27f-648">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5e27f-648">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5e27f-649">1.1</span><span class="sxs-lookup"><span data-stu-id="5e27f-649">1.1</span></span>|
|[<span data-ttu-id="5e27f-650">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5e27f-650">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5e27f-651">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="5e27f-651">ReadWriteItem</span></span>|
|[<span data-ttu-id="5e27f-652">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5e27f-652">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="5e27f-653">Создание</span><span class="sxs-lookup"><span data-stu-id="5e27f-653">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="5e27f-654">Пример</span><span class="sxs-lookup"><span data-stu-id="5e27f-654">Example</span></span>

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

#### <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="5e27f-655">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="5e27f-655">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="5e27f-656">Добавляет к сообщению элемент Exchange, например сообщение, в виде вложения.</span><span class="sxs-lookup"><span data-stu-id="5e27f-656">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="5e27f-p142">С помощью метода `addItemAttachmentAsync` можно в элемент формы создания вложить элемент с указанным идентификатором Exchange. Если указать метод обратного вызова, то этот метод вызывается с помощью параметра `asyncResult`, который содержит идентификатор вложения или код, указывающий на ошибки, которые произошли при вложении элемента. При необходимости можно использовать параметр `options` для передачи сведений о состоянии методу обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="5e27f-p142">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="5e27f-660">Идентификатор можно использовать с методом [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback), чтобы удалить вложение, добавленное во время текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="5e27f-660">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="5e27f-661">Если ваша надстройка Office выполняется в Outlook в Интернете, метод `addItemAttachmentAsync` обеспечивает вложение элементов в элементы, отличные от редактируемого. Однако это действие не рекомендуется выполнять, так как оно не поддерживается.</span><span class="sxs-lookup"><span data-stu-id="5e27f-661">If your Office Add-in is running in Outlook on the web, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="5e27f-662">Параметры</span><span class="sxs-lookup"><span data-stu-id="5e27f-662">Parameters</span></span>

|<span data-ttu-id="5e27f-663">Имя</span><span class="sxs-lookup"><span data-stu-id="5e27f-663">Name</span></span>| <span data-ttu-id="5e27f-664">Тип</span><span class="sxs-lookup"><span data-stu-id="5e27f-664">Type</span></span>| <span data-ttu-id="5e27f-665">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="5e27f-665">Attributes</span></span>| <span data-ttu-id="5e27f-666">Описание</span><span class="sxs-lookup"><span data-stu-id="5e27f-666">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="5e27f-667">String</span><span class="sxs-lookup"><span data-stu-id="5e27f-667">String</span></span>||<span data-ttu-id="5e27f-p143">Идентификатор Exchange для вкладываемого элемента. Максимальная длина — 100 символов.</span><span class="sxs-lookup"><span data-stu-id="5e27f-p143">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="5e27f-670">String</span><span class="sxs-lookup"><span data-stu-id="5e27f-670">String</span></span>||<span data-ttu-id="5e27f-671">Тема вкладываемого элемента.</span><span class="sxs-lookup"><span data-stu-id="5e27f-671">The subject of the item to be attached.</span></span> <span data-ttu-id="5e27f-672">Максимальная длина: 255 символов.</span><span class="sxs-lookup"><span data-stu-id="5e27f-672">The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="5e27f-673">Object</span><span class="sxs-lookup"><span data-stu-id="5e27f-673">Object</span></span>| <span data-ttu-id="5e27f-674">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="5e27f-674">&lt;optional&gt;</span></span>|<span data-ttu-id="5e27f-675">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="5e27f-675">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="5e27f-676">Object</span><span class="sxs-lookup"><span data-stu-id="5e27f-676">Object</span></span>| <span data-ttu-id="5e27f-677">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="5e27f-677">&lt;optional&gt;</span></span>|<span data-ttu-id="5e27f-678">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="5e27f-678">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="5e27f-679">функция</span><span class="sxs-lookup"><span data-stu-id="5e27f-679">function</span></span>| <span data-ttu-id="5e27f-680">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="5e27f-680">&lt;optional&gt;</span></span>|<span data-ttu-id="5e27f-681">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="5e27f-681">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="5e27f-682">После успешного выполнения идентификатор вложения будет представлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="5e27f-682">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="5e27f-683">Если добавить вложение не удастся, объект `asyncResult` будет содержать объект `Error` с описанием ошибки.</span><span class="sxs-lookup"><span data-stu-id="5e27f-683">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="5e27f-684">Ошибки</span><span class="sxs-lookup"><span data-stu-id="5e27f-684">Errors</span></span>

| <span data-ttu-id="5e27f-685">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="5e27f-685">Error code</span></span> | <span data-ttu-id="5e27f-686">Описание</span><span class="sxs-lookup"><span data-stu-id="5e27f-686">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="5e27f-687">Сообщение или встреча содержат слишком много вложений.</span><span class="sxs-lookup"><span data-stu-id="5e27f-687">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="5e27f-688">Requirements</span><span class="sxs-lookup"><span data-stu-id="5e27f-688">Requirements</span></span>

|<span data-ttu-id="5e27f-689">Требование</span><span class="sxs-lookup"><span data-stu-id="5e27f-689">Requirement</span></span>| <span data-ttu-id="5e27f-690">Значение</span><span class="sxs-lookup"><span data-stu-id="5e27f-690">Value</span></span>|
|---|---|
|[<span data-ttu-id="5e27f-691">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5e27f-691">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5e27f-692">1.1</span><span class="sxs-lookup"><span data-stu-id="5e27f-692">1.1</span></span>|
|[<span data-ttu-id="5e27f-693">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5e27f-693">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5e27f-694">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="5e27f-694">ReadWriteItem</span></span>|
|[<span data-ttu-id="5e27f-695">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5e27f-695">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="5e27f-696">Создание</span><span class="sxs-lookup"><span data-stu-id="5e27f-696">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="5e27f-697">Пример</span><span class="sxs-lookup"><span data-stu-id="5e27f-697">Example</span></span>

<span data-ttu-id="5e27f-698">В следующем примере существующий элемент Outlook добавляется в виде вложения с именем `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="5e27f-698">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="5e27f-699">displayReplyAllForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="5e27f-699">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="5e27f-700">Отображает форму ответа, включающую отправителя и всех получателей выбранного сообщения или организатора и всех участников выбранной встречи.</span><span class="sxs-lookup"><span data-stu-id="5e27f-700">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="5e27f-701">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="5e27f-701">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="5e27f-702">В Outlook в Интернете форма ответа отображается в виде всплывающей формы в представлении с 3 либо 1 или 2 столбцами.</span><span class="sxs-lookup"><span data-stu-id="5e27f-702">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="5e27f-703">Если любой строковый параметр превышает указанные для него ограничения, `displayReplyAllForm` возвращает исключение.</span><span class="sxs-lookup"><span data-stu-id="5e27f-703">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="5e27f-p145">Если в параметре `formData.attachments` указаны вложения, Outlook в Интернете и классические клиенты пытаются скачать их и вложить в форму ответа. Если какие-либо вложения добавить не удается, в форме отображается сообщение об ошибке. Если сообщения об ошибках не предусмотрены, то они не отображаются.</span><span class="sxs-lookup"><span data-stu-id="5e27f-p145">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="5e27f-707">Параметры</span><span class="sxs-lookup"><span data-stu-id="5e27f-707">Parameters</span></span>

|<span data-ttu-id="5e27f-708">Имя</span><span class="sxs-lookup"><span data-stu-id="5e27f-708">Name</span></span>| <span data-ttu-id="5e27f-709">Тип</span><span class="sxs-lookup"><span data-stu-id="5e27f-709">Type</span></span>| <span data-ttu-id="5e27f-710">Описание</span><span class="sxs-lookup"><span data-stu-id="5e27f-710">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="5e27f-711">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="5e27f-711">String &#124; Object</span></span>| |<span data-ttu-id="5e27f-p146">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="5e27f-p146">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="5e27f-714">**ИЛИ**</span><span class="sxs-lookup"><span data-stu-id="5e27f-714">**OR**</span></span><br/><span data-ttu-id="5e27f-p147">Объект, который содержит текст или данные вложения и функцию обратного вызова. Ниже представлено определение этого объекта.</span><span class="sxs-lookup"><span data-stu-id="5e27f-p147">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="5e27f-717">String</span><span class="sxs-lookup"><span data-stu-id="5e27f-717">String</span></span> | <span data-ttu-id="5e27f-718">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="5e27f-718">&lt;optional&gt;</span></span> | <span data-ttu-id="5e27f-p148">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="5e27f-p148">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="5e27f-721">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="5e27f-721">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="5e27f-722">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="5e27f-722">&lt;optional&gt;</span></span> | <span data-ttu-id="5e27f-723">Массив объектов JSON, представляющих собой вложенные файлы или элементы.</span><span class="sxs-lookup"><span data-stu-id="5e27f-723">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="5e27f-724">String</span><span class="sxs-lookup"><span data-stu-id="5e27f-724">String</span></span> | | <span data-ttu-id="5e27f-p149">Указывает тип вложения. Допустимые значения: `file` для вложенного файла и `item` для вложенного элемента.</span><span class="sxs-lookup"><span data-stu-id="5e27f-p149">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="5e27f-727">Строка</span><span class="sxs-lookup"><span data-stu-id="5e27f-727">String</span></span> | | <span data-ttu-id="5e27f-728">Строка, содержащая имя вложения, длиной до 255 символов.</span><span class="sxs-lookup"><span data-stu-id="5e27f-728">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="5e27f-729">String</span><span class="sxs-lookup"><span data-stu-id="5e27f-729">String</span></span> | | <span data-ttu-id="5e27f-p150">Используется, только если свойству `type` задано значение `file`. URI расположения файла.</span><span class="sxs-lookup"><span data-stu-id="5e27f-p150">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="5e27f-732">String</span><span class="sxs-lookup"><span data-stu-id="5e27f-732">String</span></span> | | <span data-ttu-id="5e27f-p151">Используется, только если свойству `type` задано значение `item`. Идентификатор вложения EWS. Это строка длиной до 100 символов.</span><span class="sxs-lookup"><span data-stu-id="5e27f-p151">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="5e27f-736">function</span><span class="sxs-lookup"><span data-stu-id="5e27f-736">function</span></span> | <span data-ttu-id="5e27f-737">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="5e27f-737">&lt;optional&gt;</span></span> | <span data-ttu-id="5e27f-738">По завершении работы метода функция, переданная параметру `callback`, вызывается с помощью одного параметра `asyncResult`, представляющего собой объект [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="5e27f-738">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="5e27f-739">Требования</span><span class="sxs-lookup"><span data-stu-id="5e27f-739">Requirements</span></span>

|<span data-ttu-id="5e27f-740">Требование</span><span class="sxs-lookup"><span data-stu-id="5e27f-740">Requirement</span></span>| <span data-ttu-id="5e27f-741">Значение</span><span class="sxs-lookup"><span data-stu-id="5e27f-741">Value</span></span>|
|---|---|
|[<span data-ttu-id="5e27f-742">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="5e27f-742">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5e27f-743">1.0</span><span class="sxs-lookup"><span data-stu-id="5e27f-743">1.0</span></span>|
|[<span data-ttu-id="5e27f-744">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5e27f-744">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5e27f-745">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5e27f-745">ReadItem</span></span>|
|[<span data-ttu-id="5e27f-746">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5e27f-746">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="5e27f-747">Чтение</span><span class="sxs-lookup"><span data-stu-id="5e27f-747">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="5e27f-748">Примеры</span><span class="sxs-lookup"><span data-stu-id="5e27f-748">Examples</span></span>

<span data-ttu-id="5e27f-749">Приведенный ниже код передает строку в функцию `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="5e27f-749">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="5e27f-750">Ответ с пустым текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="5e27f-750">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="5e27f-751">Ответ только с текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="5e27f-751">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="5e27f-752">Ответ с текстом сообщения и вложенным файлом.</span><span class="sxs-lookup"><span data-stu-id="5e27f-752">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="5e27f-753">Ответ с текстом сообщения и вложенным элементом.</span><span class="sxs-lookup"><span data-stu-id="5e27f-753">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="5e27f-754">Ответ с текстом сообщения, вложенным файлом, вложенным элементом и обратным вызовом.</span><span class="sxs-lookup"><span data-stu-id="5e27f-754">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="5e27f-755">displayReplyForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="5e27f-755">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="5e27f-756">Отображает форму ответа, включающую только отправителя выбранного сообщения или организатора выбранной встречи.</span><span class="sxs-lookup"><span data-stu-id="5e27f-756">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="5e27f-757">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="5e27f-757">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="5e27f-758">В Outlook в Интернете форма ответа отображается в виде всплывающей формы в представлении с 3 либо 1 или 2 столбцами.</span><span class="sxs-lookup"><span data-stu-id="5e27f-758">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="5e27f-759">Если любой строковый параметр превышает указанные для него ограничения, `displayReplyForm` возвращает исключение.</span><span class="sxs-lookup"><span data-stu-id="5e27f-759">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="5e27f-p152">Если в параметре `formData.attachments` указаны вложения, Outlook в Интернете и классические клиенты пытаются скачать их и вложить в форму ответа. Если какие-либо вложения добавить не удается, в форме отображается сообщение об ошибке. Если сообщения об ошибках не предусмотрены, то они не отображаются.</span><span class="sxs-lookup"><span data-stu-id="5e27f-p152">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="5e27f-763">Параметры</span><span class="sxs-lookup"><span data-stu-id="5e27f-763">Parameters</span></span>

|<span data-ttu-id="5e27f-764">Имя</span><span class="sxs-lookup"><span data-stu-id="5e27f-764">Name</span></span>| <span data-ttu-id="5e27f-765">Тип</span><span class="sxs-lookup"><span data-stu-id="5e27f-765">Type</span></span>| <span data-ttu-id="5e27f-766">Описание</span><span class="sxs-lookup"><span data-stu-id="5e27f-766">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="5e27f-767">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="5e27f-767">String &#124; Object</span></span>| | <span data-ttu-id="5e27f-p153">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="5e27f-p153">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="5e27f-770">**ИЛИ**</span><span class="sxs-lookup"><span data-stu-id="5e27f-770">**OR**</span></span><br/><span data-ttu-id="5e27f-p154">Объект, который содержит текст или данные вложения и функцию обратного вызова. Ниже представлено определение этого объекта.</span><span class="sxs-lookup"><span data-stu-id="5e27f-p154">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="5e27f-773">String</span><span class="sxs-lookup"><span data-stu-id="5e27f-773">String</span></span> | <span data-ttu-id="5e27f-774">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="5e27f-774">&lt;optional&gt;</span></span> | <span data-ttu-id="5e27f-p155">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="5e27f-p155">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="5e27f-777">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="5e27f-777">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="5e27f-778">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="5e27f-778">&lt;optional&gt;</span></span> | <span data-ttu-id="5e27f-779">Массив объектов JSON, представляющих собой вложенные файлы или элементы.</span><span class="sxs-lookup"><span data-stu-id="5e27f-779">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="5e27f-780">String</span><span class="sxs-lookup"><span data-stu-id="5e27f-780">String</span></span> | | <span data-ttu-id="5e27f-p156">Указывает тип вложения. Допустимые значения: `file` для вложенного файла и `item` для вложенного элемента.</span><span class="sxs-lookup"><span data-stu-id="5e27f-p156">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="5e27f-783">Строка</span><span class="sxs-lookup"><span data-stu-id="5e27f-783">String</span></span> | | <span data-ttu-id="5e27f-784">Строка, содержащая имя вложения, длиной до 255 символов.</span><span class="sxs-lookup"><span data-stu-id="5e27f-784">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="5e27f-785">Строка</span><span class="sxs-lookup"><span data-stu-id="5e27f-785">String</span></span> | | <span data-ttu-id="5e27f-p157">Используется, только если свойству `type` задано значение `file`. URI расположения файла.</span><span class="sxs-lookup"><span data-stu-id="5e27f-p157">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="5e27f-788">String</span><span class="sxs-lookup"><span data-stu-id="5e27f-788">String</span></span> | | <span data-ttu-id="5e27f-p158">Используется, только если свойству `type` задано значение `item`. Идентификатор вложения EWS. Это строка длиной до 100 символов.</span><span class="sxs-lookup"><span data-stu-id="5e27f-p158">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="5e27f-792">function</span><span class="sxs-lookup"><span data-stu-id="5e27f-792">function</span></span> | <span data-ttu-id="5e27f-793">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="5e27f-793">&lt;optional&gt;</span></span> | <span data-ttu-id="5e27f-794">По завершении работы метода функция, переданная параметру `callback`, вызывается с помощью одного параметра `asyncResult`, представляющего собой объект [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="5e27f-794">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="5e27f-795">Requirements</span><span class="sxs-lookup"><span data-stu-id="5e27f-795">Requirements</span></span>

|<span data-ttu-id="5e27f-796">Требование</span><span class="sxs-lookup"><span data-stu-id="5e27f-796">Requirement</span></span>| <span data-ttu-id="5e27f-797">Значение</span><span class="sxs-lookup"><span data-stu-id="5e27f-797">Value</span></span>|
|---|---|
|[<span data-ttu-id="5e27f-798">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5e27f-798">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5e27f-799">1.0</span><span class="sxs-lookup"><span data-stu-id="5e27f-799">1.0</span></span>|
|[<span data-ttu-id="5e27f-800">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5e27f-800">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5e27f-801">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5e27f-801">ReadItem</span></span>|
|[<span data-ttu-id="5e27f-802">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5e27f-802">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="5e27f-803">Чтение</span><span class="sxs-lookup"><span data-stu-id="5e27f-803">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="5e27f-804">Примеры</span><span class="sxs-lookup"><span data-stu-id="5e27f-804">Examples</span></span>

<span data-ttu-id="5e27f-805">Приведенный ниже код передает строку в функцию `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="5e27f-805">The following code passes a string to the `displayReplyForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="5e27f-806">Ответ с пустым текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="5e27f-806">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="5e27f-807">Ответ только с текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="5e27f-807">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="5e27f-808">Ответ с текстом сообщения и вложенным файлом.</span><span class="sxs-lookup"><span data-stu-id="5e27f-808">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="5e27f-809">Ответ с текстом сообщения и вложенным элементом.</span><span class="sxs-lookup"><span data-stu-id="5e27f-809">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="5e27f-810">Ответ с текстом сообщения, вложенным файлом, вложенным элементом и обратным вызовом.</span><span class="sxs-lookup"><span data-stu-id="5e27f-810">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlookofficeentitiesviewoutlook-js-12"></a><span data-ttu-id="5e27f-811">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.2)}</span><span class="sxs-lookup"><span data-stu-id="5e27f-811">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.2)}</span></span>

<span data-ttu-id="5e27f-812">Получает сущности, обнаруженные в теле выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="5e27f-812">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="5e27f-813">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="5e27f-813">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="5e27f-814">Требования</span><span class="sxs-lookup"><span data-stu-id="5e27f-814">Requirements</span></span>

|<span data-ttu-id="5e27f-815">Требование</span><span class="sxs-lookup"><span data-stu-id="5e27f-815">Requirement</span></span>| <span data-ttu-id="5e27f-816">Значение</span><span class="sxs-lookup"><span data-stu-id="5e27f-816">Value</span></span>|
|---|---|
|[<span data-ttu-id="5e27f-817">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5e27f-817">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5e27f-818">1.0</span><span class="sxs-lookup"><span data-stu-id="5e27f-818">1.0</span></span>|
|[<span data-ttu-id="5e27f-819">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5e27f-819">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5e27f-820">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5e27f-820">ReadItem</span></span>|
|[<span data-ttu-id="5e27f-821">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5e27f-821">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="5e27f-822">Чтение</span><span class="sxs-lookup"><span data-stu-id="5e27f-822">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="5e27f-823">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="5e27f-823">Returns:</span></span>

<span data-ttu-id="5e27f-824">Тип: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="5e27f-824">Type: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.2)</span></span>

##### <a name="example"></a><span data-ttu-id="5e27f-825">Пример</span><span class="sxs-lookup"><span data-stu-id="5e27f-825">Example</span></span>

<span data-ttu-id="5e27f-826">Ниже приведен пример получения доступа к сущностям контактов в теле текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="5e27f-826">The following example accesses the contacts entities in the current item's body.</span></span>

```js
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

<br>

---
---

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-12meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-12phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-12tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-12"></a><span data-ttu-id="5e27f-827">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.2)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.2)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.2)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.2))>}</span><span class="sxs-lookup"><span data-stu-id="5e27f-827">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.2)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.2)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.2)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.2))>}</span></span>

<span data-ttu-id="5e27f-828">Получает массив всех сущностей указанного типа, обнаруженных в теле выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="5e27f-828">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="5e27f-829">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="5e27f-829">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="5e27f-830">Параметры</span><span class="sxs-lookup"><span data-stu-id="5e27f-830">Parameters</span></span>

|<span data-ttu-id="5e27f-831">Имя</span><span class="sxs-lookup"><span data-stu-id="5e27f-831">Name</span></span>| <span data-ttu-id="5e27f-832">Тип</span><span class="sxs-lookup"><span data-stu-id="5e27f-832">Type</span></span>| <span data-ttu-id="5e27f-833">Описание</span><span class="sxs-lookup"><span data-stu-id="5e27f-833">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="5e27f-834">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="5e27f-834">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook/office.mailboxenums.entitytype?view=outlook-js-1.2)|<span data-ttu-id="5e27f-835">Одно из значений перечисления EntityType.</span><span class="sxs-lookup"><span data-stu-id="5e27f-835">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="5e27f-836">Requirements</span><span class="sxs-lookup"><span data-stu-id="5e27f-836">Requirements</span></span>

|<span data-ttu-id="5e27f-837">Требование</span><span class="sxs-lookup"><span data-stu-id="5e27f-837">Requirement</span></span>| <span data-ttu-id="5e27f-838">Значение</span><span class="sxs-lookup"><span data-stu-id="5e27f-838">Value</span></span>|
|---|---|
|[<span data-ttu-id="5e27f-839">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5e27f-839">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5e27f-840">1.0</span><span class="sxs-lookup"><span data-stu-id="5e27f-840">1.0</span></span>|
|[<span data-ttu-id="5e27f-841">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5e27f-841">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5e27f-842">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="5e27f-842">Restricted</span></span>|
|[<span data-ttu-id="5e27f-843">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5e27f-843">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="5e27f-844">Чтение</span><span class="sxs-lookup"><span data-stu-id="5e27f-844">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="5e27f-845">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="5e27f-845">Returns:</span></span>

<span data-ttu-id="5e27f-846">Если значение, переданное в `entityType`, не является допустимым членом перечисления `EntityType`, метод возвращает значение NULL.</span><span class="sxs-lookup"><span data-stu-id="5e27f-846">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="5e27f-847">Если в теле элемента отсутствуют сущности указанного типа, метод возвращает пустой массив.</span><span class="sxs-lookup"><span data-stu-id="5e27f-847">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="5e27f-848">В противном случае тип объектов в возвращаемом массиве зависит от типа сущности, запрошенной в параметре `entityType`.</span><span class="sxs-lookup"><span data-stu-id="5e27f-848">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="5e27f-849">Хотя минимальный уровень разрешений для использования этого метода — **Restricted**, для некоторых типов сущностей требуется доступ на уровне **ReadItem**, как указано в приведенной ниже таблице.</span><span class="sxs-lookup"><span data-stu-id="5e27f-849">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="5e27f-850">Значение параметра `entityType`</span><span class="sxs-lookup"><span data-stu-id="5e27f-850">Value of `entityType`</span></span> | <span data-ttu-id="5e27f-851">Тип объектов в возвращаемом массиве</span><span class="sxs-lookup"><span data-stu-id="5e27f-851">Type of objects in returned array</span></span> | <span data-ttu-id="5e27f-852">Необходимый уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5e27f-852">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="5e27f-853">String</span><span class="sxs-lookup"><span data-stu-id="5e27f-853">String</span></span> | <span data-ttu-id="5e27f-854">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="5e27f-854">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="5e27f-855">Contact</span><span class="sxs-lookup"><span data-stu-id="5e27f-855">Contact</span></span> | <span data-ttu-id="5e27f-856">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="5e27f-856">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="5e27f-857">String</span><span class="sxs-lookup"><span data-stu-id="5e27f-857">String</span></span> | <span data-ttu-id="5e27f-858">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="5e27f-858">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="5e27f-859">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="5e27f-859">MeetingSuggestion</span></span> | <span data-ttu-id="5e27f-860">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="5e27f-860">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="5e27f-861">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="5e27f-861">PhoneNumber</span></span> | <span data-ttu-id="5e27f-862">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="5e27f-862">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="5e27f-863">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="5e27f-863">TaskSuggestion</span></span> | <span data-ttu-id="5e27f-864">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="5e27f-864">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="5e27f-865">String</span><span class="sxs-lookup"><span data-stu-id="5e27f-865">String</span></span> | <span data-ttu-id="5e27f-866">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="5e27f-866">**Restricted**</span></span> |

<span data-ttu-id="5e27f-867">Тип: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.2)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.2)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.2)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.2))></span><span class="sxs-lookup"><span data-stu-id="5e27f-867">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.2)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.2)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.2)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.2))></span></span>

##### <a name="example"></a><span data-ttu-id="5e27f-868">Пример</span><span class="sxs-lookup"><span data-stu-id="5e27f-868">Example</span></span>

<span data-ttu-id="5e27f-869">В примере ниже показано, как получить доступ к массиву строк, которые представляют собой почтовые адреса в теле текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="5e27f-869">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-12meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-12phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-12tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-12"></a><span data-ttu-id="5e27f-870">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.2)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.2)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.2)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.2))>}</span><span class="sxs-lookup"><span data-stu-id="5e27f-870">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.2)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.2)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.2)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.2))>}</span></span>

<span data-ttu-id="5e27f-871">Возвращает известные сущности в выбранном элементе, которые проходят через именованный фильтр, определяемый в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="5e27f-871">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="5e27f-872">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="5e27f-872">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="5e27f-873">Метод `getFilteredEntitiesByName` возвращает сущности, соответствующие регулярному выражению, которое определяется в элементе правила [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) в XML-файле манифеста, с использованием указанного значения элемента `FilterName`.</span><span class="sxs-lookup"><span data-stu-id="5e27f-873">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="5e27f-874">Параметры</span><span class="sxs-lookup"><span data-stu-id="5e27f-874">Parameters</span></span>

|<span data-ttu-id="5e27f-875">Имя</span><span class="sxs-lookup"><span data-stu-id="5e27f-875">Name</span></span>| <span data-ttu-id="5e27f-876">Тип</span><span class="sxs-lookup"><span data-stu-id="5e27f-876">Type</span></span>| <span data-ttu-id="5e27f-877">Описание</span><span class="sxs-lookup"><span data-stu-id="5e27f-877">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="5e27f-878">String</span><span class="sxs-lookup"><span data-stu-id="5e27f-878">String</span></span>|<span data-ttu-id="5e27f-879">Имя элемента правила `ItemHasKnownEntity`, определяющее соответствующий фильтр.</span><span class="sxs-lookup"><span data-stu-id="5e27f-879">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="5e27f-880">Requirements</span><span class="sxs-lookup"><span data-stu-id="5e27f-880">Requirements</span></span>

|<span data-ttu-id="5e27f-881">Требование</span><span class="sxs-lookup"><span data-stu-id="5e27f-881">Requirement</span></span>| <span data-ttu-id="5e27f-882">Значение</span><span class="sxs-lookup"><span data-stu-id="5e27f-882">Value</span></span>|
|---|---|
|[<span data-ttu-id="5e27f-883">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5e27f-883">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5e27f-884">1.0</span><span class="sxs-lookup"><span data-stu-id="5e27f-884">1.0</span></span>|
|[<span data-ttu-id="5e27f-885">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5e27f-885">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5e27f-886">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5e27f-886">ReadItem</span></span>|
|[<span data-ttu-id="5e27f-887">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5e27f-887">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="5e27f-888">Чтение</span><span class="sxs-lookup"><span data-stu-id="5e27f-888">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="5e27f-889">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="5e27f-889">Returns:</span></span>

<span data-ttu-id="5e27f-p160">Если в манифесте нет элемента `ItemHasKnownEntity` со значением `FilterName`, соответствующим параметру `name`, метод возвращает `null`. Если параметр `name` соответствует элементу `ItemHasKnownEntity` в манифесте, но при этом в текущем элементе нет соответствующих сущностей, метод возвращает пустой массив.</span><span class="sxs-lookup"><span data-stu-id="5e27f-p160">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="5e27f-892">Тип: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.2)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.2)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.2)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.2))></span><span class="sxs-lookup"><span data-stu-id="5e27f-892">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.2)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.2)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.2)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.2))></span></span>

<br>

---
---

#### <a name="getregexmatches--object"></a><span data-ttu-id="5e27f-893">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="5e27f-893">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="5e27f-894">Возвращает строковые значения в выбранном элементе, которые соответствуют регулярным выражениям, определенным в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="5e27f-894">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="5e27f-895">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="5e27f-895">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="5e27f-p161">Метод `getRegExMatches` возвращает строки, соответствующие регулярному выражению, которое определяется в каждом элементе правила `ItemHasRegularExpressionMatch` или `ItemHasKnownEntity` в XML-файле манифеста. Для правила `ItemHasRegularExpressionMatch` соответствующую строку должно содержать свойство элемента, указанного этим правилом. Простой тип `PropertyName` определяет поддерживаемые свойства.</span><span class="sxs-lookup"><span data-stu-id="5e27f-p161">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="5e27f-899">Например, рассмотрим манифест надстройки, который содержит указанный ниже элемент `Rule`.</span><span class="sxs-lookup"><span data-stu-id="5e27f-899">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="5e27f-900">Объект, возвращаемый методом `getRegExMatches`, будет содержать два свойства: `fruits` и `veggies`.</span><span class="sxs-lookup"><span data-stu-id="5e27f-900">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

> [!NOTE]
> <span data-ttu-id="5e27f-p162">Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты.</span><span class="sxs-lookup"><span data-stu-id="5e27f-p162">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="requirements"></a><span data-ttu-id="5e27f-903">Requirements</span><span class="sxs-lookup"><span data-stu-id="5e27f-903">Requirements</span></span>

|<span data-ttu-id="5e27f-904">Требование</span><span class="sxs-lookup"><span data-stu-id="5e27f-904">Requirement</span></span>| <span data-ttu-id="5e27f-905">Значение</span><span class="sxs-lookup"><span data-stu-id="5e27f-905">Value</span></span>|
|---|---|
|[<span data-ttu-id="5e27f-906">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5e27f-906">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5e27f-907">1.0</span><span class="sxs-lookup"><span data-stu-id="5e27f-907">1.0</span></span>|
|[<span data-ttu-id="5e27f-908">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5e27f-908">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5e27f-909">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5e27f-909">ReadItem</span></span>|
|[<span data-ttu-id="5e27f-910">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5e27f-910">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="5e27f-911">Чтение</span><span class="sxs-lookup"><span data-stu-id="5e27f-911">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="5e27f-912">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="5e27f-912">Returns:</span></span>

<span data-ttu-id="5e27f-p163">Объект, содержащий массив строк, которые соответствуют регулярным выражениям, определяемым в XML-файле манифеста. Имя каждого массива равно соответствующему значению атрибута `RegExName` подходящего правила `ItemHasRegularExpressionMatch` или атрибута `FilterName` соответствующего правила `ItemHasKnownEntity`.</span><span class="sxs-lookup"><span data-stu-id="5e27f-p163">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<span data-ttu-id="5e27f-915">Тип: Object</span><span class="sxs-lookup"><span data-stu-id="5e27f-915">Type: Object</span></span>

##### <a name="example"></a><span data-ttu-id="5e27f-916">Пример</span><span class="sxs-lookup"><span data-stu-id="5e27f-916">Example</span></span>

<span data-ttu-id="5e27f-917">В примере ниже показано, как получить доступ к массиву совпадений для <rule>элементов регулярного выражения `fruits` и `veggies`, которые указаны в манифесте</rule>.</span><span class="sxs-lookup"><span data-stu-id="5e27f-917">The following example shows how to access the array of matches for the regular expression <rule>elements `fruits` and `veggies`, which are specified in the manifest.</rule></span></span>

```js
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

<br>

---
---

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="5e27f-918">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="5e27f-918">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="5e27f-919">Возвращает строковые значения в выбранном элементе, которые соответствуют именованному регулярному выражению, определенному в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="5e27f-919">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="5e27f-920">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="5e27f-920">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="5e27f-921">Метод `getRegExMatchesByName` возвращает строки, соответствующие регулярному выражению, которое определяется в элементе правила `ItemHasRegularExpressionMatch` в XML-файле манифеста, с использованием указанного значения элемента `RegExName`.</span><span class="sxs-lookup"><span data-stu-id="5e27f-921">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="5e27f-p164">Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты.</span><span class="sxs-lookup"><span data-stu-id="5e27f-p164">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="5e27f-924">Параметры</span><span class="sxs-lookup"><span data-stu-id="5e27f-924">Parameters</span></span>

|<span data-ttu-id="5e27f-925">Имя</span><span class="sxs-lookup"><span data-stu-id="5e27f-925">Name</span></span>| <span data-ttu-id="5e27f-926">Тип</span><span class="sxs-lookup"><span data-stu-id="5e27f-926">Type</span></span>| <span data-ttu-id="5e27f-927">Описание</span><span class="sxs-lookup"><span data-stu-id="5e27f-927">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="5e27f-928">String</span><span class="sxs-lookup"><span data-stu-id="5e27f-928">String</span></span>|<span data-ttu-id="5e27f-929">Имя элемента правила `ItemHasRegularExpressionMatch`, определяющее соответствующий фильтр.</span><span class="sxs-lookup"><span data-stu-id="5e27f-929">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="5e27f-930">Requirements</span><span class="sxs-lookup"><span data-stu-id="5e27f-930">Requirements</span></span>

|<span data-ttu-id="5e27f-931">Требование</span><span class="sxs-lookup"><span data-stu-id="5e27f-931">Requirement</span></span>| <span data-ttu-id="5e27f-932">Значение</span><span class="sxs-lookup"><span data-stu-id="5e27f-932">Value</span></span>|
|---|---|
|[<span data-ttu-id="5e27f-933">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5e27f-933">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5e27f-934">1.0</span><span class="sxs-lookup"><span data-stu-id="5e27f-934">1.0</span></span>|
|[<span data-ttu-id="5e27f-935">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5e27f-935">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5e27f-936">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5e27f-936">ReadItem</span></span>|
|[<span data-ttu-id="5e27f-937">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5e27f-937">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="5e27f-938">Чтение</span><span class="sxs-lookup"><span data-stu-id="5e27f-938">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="5e27f-939">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="5e27f-939">Returns:</span></span>

<span data-ttu-id="5e27f-940">Массив строк, соответствующих регулярному выражению, определяемому в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="5e27f-940">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<span data-ttu-id="5e27f-941">Тип: Array.< String ></span><span class="sxs-lookup"><span data-stu-id="5e27f-941">Type: Array.< String ></span></span>

##### <a name="example"></a><span data-ttu-id="5e27f-942">Пример</span><span class="sxs-lookup"><span data-stu-id="5e27f-942">Example</span></span>

```js
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

<br>

---
---

#### <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="5e27f-943">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="5e27f-943">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="5e27f-944">Асинхронно возвращает данные, выбранные в теме или тексте сообщения.</span><span class="sxs-lookup"><span data-stu-id="5e27f-944">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="5e27f-945">Если выделенный фрагмент отсутствует, но курсор находится в основном тексте или теме, метод возвращает пустую строку для выбранных данных.</span><span class="sxs-lookup"><span data-stu-id="5e27f-945">If there is no selection but the cursor is in the body or subject, the method returns an empty string for the selected data.</span></span> <span data-ttu-id="5e27f-946">Если выбраны не текст и не тема, метод возвращает ошибку `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="5e27f-946">If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

> [!NOTE]
> <span data-ttu-id="5e27f-947">В Outlook в Интернете метод возвращает строку null, если текст не выделен, но курсор находится в тексте.</span><span class="sxs-lookup"><span data-stu-id="5e27f-947">In Outlook on the web, the method returns the string "null" if no text is selected but the cursor is in the body.</span></span> <span data-ttu-id="5e27f-948">Чтобы проверить эту ситуацию, ознакомьтесь с приведенным далее в этом разделе.</span><span class="sxs-lookup"><span data-stu-id="5e27f-948">To check for this situation, see the example later in this section.</span></span>

##### <a name="parameters"></a><span data-ttu-id="5e27f-949">Параметры</span><span class="sxs-lookup"><span data-stu-id="5e27f-949">Parameters</span></span>

|<span data-ttu-id="5e27f-950">Имя</span><span class="sxs-lookup"><span data-stu-id="5e27f-950">Name</span></span>| <span data-ttu-id="5e27f-951">Тип</span><span class="sxs-lookup"><span data-stu-id="5e27f-951">Type</span></span>| <span data-ttu-id="5e27f-952">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="5e27f-952">Attributes</span></span>| <span data-ttu-id="5e27f-953">Описание</span><span class="sxs-lookup"><span data-stu-id="5e27f-953">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="5e27f-954">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="5e27f-954">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="5e27f-p167">Запрашивает формат данных. Если задано значение Text, метод возвращает обычный текст как строку, удаляя все имеющиеся HTML-теги. Если задано значение HTML, метод возвращает выделенный текст (обычный текст или HTML).</span><span class="sxs-lookup"><span data-stu-id="5e27f-p167">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="5e27f-958">Object</span><span class="sxs-lookup"><span data-stu-id="5e27f-958">Object</span></span>| <span data-ttu-id="5e27f-959">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="5e27f-959">&lt;optional&gt;</span></span>|<span data-ttu-id="5e27f-960">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="5e27f-960">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="5e27f-961">Object</span><span class="sxs-lookup"><span data-stu-id="5e27f-961">Object</span></span>| <span data-ttu-id="5e27f-962">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="5e27f-962">&lt;optional&gt;</span></span>|<span data-ttu-id="5e27f-963">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="5e27f-963">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="5e27f-964">функция</span><span class="sxs-lookup"><span data-stu-id="5e27f-964">function</span></span>||<span data-ttu-id="5e27f-965">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="5e27f-965">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="5e27f-966">Чтобы получить доступ к выбранным данным из метода обратного вызова, вызовите `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="5e27f-966">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="5e27f-967">Для доступа к исходному свойству, представляющему собой источник выбранных данных, вызовите параметр `asyncResult.value.sourceProperty`, который может иметь значение `body` или `subject`.</span><span class="sxs-lookup"><span data-stu-id="5e27f-967">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="5e27f-968">Requirements</span><span class="sxs-lookup"><span data-stu-id="5e27f-968">Requirements</span></span>

|<span data-ttu-id="5e27f-969">Требование</span><span class="sxs-lookup"><span data-stu-id="5e27f-969">Requirement</span></span>| <span data-ttu-id="5e27f-970">Значение</span><span class="sxs-lookup"><span data-stu-id="5e27f-970">Value</span></span>|
|---|---|
|[<span data-ttu-id="5e27f-971">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="5e27f-971">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5e27f-972">1.2</span><span class="sxs-lookup"><span data-stu-id="5e27f-972">1.2</span></span>|
|[<span data-ttu-id="5e27f-973">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5e27f-973">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5e27f-974">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5e27f-974">ReadItem</span></span>|
|[<span data-ttu-id="5e27f-975">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5e27f-975">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="5e27f-976">Создание</span><span class="sxs-lookup"><span data-stu-id="5e27f-976">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="5e27f-977">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="5e27f-977">Returns:</span></span>

<span data-ttu-id="5e27f-978">Выбранные данные в виде строки с форматом, определенным в параметре `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="5e27f-978">The selected data as a string with format determined by `coercionType`.</span></span>

<span data-ttu-id="5e27f-979">Тип: строка</span><span class="sxs-lookup"><span data-stu-id="5e27f-979">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="5e27f-980">Пример</span><span class="sxs-lookup"><span data-stu-id="5e27f-980">Example</span></span>

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

#### <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="5e27f-981">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="5e27f-981">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="5e27f-982">Асинхронно загружает настраиваемые свойства для надстройки для выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="5e27f-982">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="5e27f-p169">Настраиваемые свойства сохраняются в виде пар "ключ-значение" для каждого приложения и каждого элемента. Этот метод возвращает объект `CustomProperties` при обратном вызове, который предоставляет методы для доступа к настраиваемым свойствам, характерным для текущего элемента и текущей надстройки. Настраиваемые свойства не шифруются для элемента, поэтому этот способ хранения не является безопасным.</span><span class="sxs-lookup"><span data-stu-id="5e27f-p169">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="5e27f-986">Параметры</span><span class="sxs-lookup"><span data-stu-id="5e27f-986">Parameters</span></span>

|<span data-ttu-id="5e27f-987">Имя</span><span class="sxs-lookup"><span data-stu-id="5e27f-987">Name</span></span>| <span data-ttu-id="5e27f-988">Тип</span><span class="sxs-lookup"><span data-stu-id="5e27f-988">Type</span></span>| <span data-ttu-id="5e27f-989">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="5e27f-989">Attributes</span></span>| <span data-ttu-id="5e27f-990">Описание</span><span class="sxs-lookup"><span data-stu-id="5e27f-990">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="5e27f-991">function</span><span class="sxs-lookup"><span data-stu-id="5e27f-991">function</span></span>||<span data-ttu-id="5e27f-992">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="5e27f-992">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="5e27f-993">Настраиваемые свойства предоставляются в виде объекта [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.2) в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="5e27f-993">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.2) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="5e27f-994">Этот объект позволяет получить, задать и удалить настраиваемые свойства для элемента, а также сохранить изменения, внесенные в набор настраиваемых свойств, на сервере.</span><span class="sxs-lookup"><span data-stu-id="5e27f-994">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="5e27f-995">Объект</span><span class="sxs-lookup"><span data-stu-id="5e27f-995">Object</span></span>| <span data-ttu-id="5e27f-996">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="5e27f-996">&lt;optional&gt;</span></span>|<span data-ttu-id="5e27f-997">Разработчики могут указать любой объект, к которому необходимо получить доступ, в функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="5e27f-997">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="5e27f-998">Доступ к этому объекту можно получить с помощью свойства `asyncResult.asyncContext` в функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="5e27f-998">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="5e27f-999">Требования</span><span class="sxs-lookup"><span data-stu-id="5e27f-999">Requirements</span></span>

|<span data-ttu-id="5e27f-1000">Требование</span><span class="sxs-lookup"><span data-stu-id="5e27f-1000">Requirement</span></span>| <span data-ttu-id="5e27f-1001">Значение</span><span class="sxs-lookup"><span data-stu-id="5e27f-1001">Value</span></span>|
|---|---|
|[<span data-ttu-id="5e27f-1002">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5e27f-1002">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5e27f-1003">1.0</span><span class="sxs-lookup"><span data-stu-id="5e27f-1003">1.0</span></span>|
|[<span data-ttu-id="5e27f-1004">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5e27f-1004">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5e27f-1005">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5e27f-1005">ReadItem</span></span>|
|[<span data-ttu-id="5e27f-1006">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5e27f-1006">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="5e27f-1007">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="5e27f-1007">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="5e27f-1008">Пример</span><span class="sxs-lookup"><span data-stu-id="5e27f-1008">Example</span></span>

<span data-ttu-id="5e27f-p172">Приведенный ниже пример кода показывает, как асинхронно загружать настраиваемые свойства, характерные для текущего элемента, с помощью метода `loadCustomPropertiesAsync`. Этот пример также показывает, как сохранять эти свойства на сервере с помощью метода `CustomProperties.saveAsync`. После загрузки настраиваемых свойств в этом примере кода метод `CustomProperties.get` используется для считывания настраиваемого свойства `myProp`, метод `CustomProperties.set` — для записи настраиваемого свойства `otherProp`, а метод `saveAsync` — для сохранения настраиваемых свойств.</span><span class="sxs-lookup"><span data-stu-id="5e27f-p172">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

#### <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="5e27f-1012">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="5e27f-1012">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="5e27f-1013">Удаляет вложение из сообщения или встречи.</span><span class="sxs-lookup"><span data-stu-id="5e27f-1013">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="5e27f-1014">Метод `removeAttachmentAsync` удаляет из элемента вложение с указанным идентификатором.</span><span class="sxs-lookup"><span data-stu-id="5e27f-1014">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="5e27f-1015">Идентификатор вложения рекомендуется использовать для удаления вложения, только если оно добавлено тем же почтовым приложением в ходе текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="5e27f-1015">As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session.</span></span> <span data-ttu-id="5e27f-1016">В Outlook в Интернете и на мобильных устройствах идентификатор вложения действителен только в течение одного сеанса.</span><span class="sxs-lookup"><span data-stu-id="5e27f-1016">In Outlook on the web and mobile devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="5e27f-1017">Сеанс завершается, когда пользователь закрывает приложение или начинает создавать элемент во встроенной форме, а затем переходит из формы в отдельное окно.</span><span class="sxs-lookup"><span data-stu-id="5e27f-1017">A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="5e27f-1018">Параметры</span><span class="sxs-lookup"><span data-stu-id="5e27f-1018">Parameters</span></span>

|<span data-ttu-id="5e27f-1019">Имя</span><span class="sxs-lookup"><span data-stu-id="5e27f-1019">Name</span></span>| <span data-ttu-id="5e27f-1020">Тип</span><span class="sxs-lookup"><span data-stu-id="5e27f-1020">Type</span></span>| <span data-ttu-id="5e27f-1021">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="5e27f-1021">Attributes</span></span>| <span data-ttu-id="5e27f-1022">Описание</span><span class="sxs-lookup"><span data-stu-id="5e27f-1022">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="5e27f-1023">String</span><span class="sxs-lookup"><span data-stu-id="5e27f-1023">String</span></span>||<span data-ttu-id="5e27f-1024">Идентификатор удаляемого вложения.</span><span class="sxs-lookup"><span data-stu-id="5e27f-1024">The identifier of the attachment to remove.</span></span>|
|`options`| <span data-ttu-id="5e27f-1025">Object</span><span class="sxs-lookup"><span data-stu-id="5e27f-1025">Object</span></span>| <span data-ttu-id="5e27f-1026">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="5e27f-1026">&lt;optional&gt;</span></span>|<span data-ttu-id="5e27f-1027">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="5e27f-1027">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="5e27f-1028">Object</span><span class="sxs-lookup"><span data-stu-id="5e27f-1028">Object</span></span>| <span data-ttu-id="5e27f-1029">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="5e27f-1029">&lt;optional&gt;</span></span>|<span data-ttu-id="5e27f-1030">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="5e27f-1030">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="5e27f-1031">функция</span><span class="sxs-lookup"><span data-stu-id="5e27f-1031">function</span></span>| <span data-ttu-id="5e27f-1032">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="5e27f-1032">&lt;optional&gt;</span></span>|<span data-ttu-id="5e27f-1033">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="5e27f-1033">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="5e27f-1034">Если удалить вложение не удается, свойство `asyncResult.error` содержит код ошибки с указанием ее причины.</span><span class="sxs-lookup"><span data-stu-id="5e27f-1034">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="5e27f-1035">Ошибки</span><span class="sxs-lookup"><span data-stu-id="5e27f-1035">Errors</span></span>

| <span data-ttu-id="5e27f-1036">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="5e27f-1036">Error code</span></span> | <span data-ttu-id="5e27f-1037">Описание</span><span class="sxs-lookup"><span data-stu-id="5e27f-1037">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="5e27f-1038">Идентификатор вложения не существует.</span><span class="sxs-lookup"><span data-stu-id="5e27f-1038">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="5e27f-1039">Requirements</span><span class="sxs-lookup"><span data-stu-id="5e27f-1039">Requirements</span></span>

|<span data-ttu-id="5e27f-1040">Требование</span><span class="sxs-lookup"><span data-stu-id="5e27f-1040">Requirement</span></span>| <span data-ttu-id="5e27f-1041">Значение</span><span class="sxs-lookup"><span data-stu-id="5e27f-1041">Value</span></span>|
|---|---|
|[<span data-ttu-id="5e27f-1042">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5e27f-1042">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5e27f-1043">1.1</span><span class="sxs-lookup"><span data-stu-id="5e27f-1043">1.1</span></span>|
|[<span data-ttu-id="5e27f-1044">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5e27f-1044">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5e27f-1045">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="5e27f-1045">ReadWriteItem</span></span>|
|[<span data-ttu-id="5e27f-1046">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5e27f-1046">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="5e27f-1047">Создание</span><span class="sxs-lookup"><span data-stu-id="5e27f-1047">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="5e27f-1048">Пример</span><span class="sxs-lookup"><span data-stu-id="5e27f-1048">Example</span></span>

<span data-ttu-id="5e27f-1049">Указанный ниже код удаляет вложение с идентификатором "0".</span><span class="sxs-lookup"><span data-stu-id="5e27f-1049">The following code removes an attachment with an identifier of '0'.</span></span>

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

#### <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="5e27f-1050">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="5e27f-1050">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="5e27f-1051">Асинхронно вставляет данные в текст или тему сообщения.</span><span class="sxs-lookup"><span data-stu-id="5e27f-1051">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="5e27f-p174">Метод `setSelectedDataAsync` вставляет указанную строку в местоположение курсора в теме или тексте элемента либо, если текст выделен в редакторе, он заменяет выделенный текст. Если курсор находится вне текста или темы элемента, возвращается ошибка. После вставки курсор помещается в конец вставленного содержимого.</span><span class="sxs-lookup"><span data-stu-id="5e27f-p174">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="5e27f-1055">Параметры</span><span class="sxs-lookup"><span data-stu-id="5e27f-1055">Parameters</span></span>

|<span data-ttu-id="5e27f-1056">Имя</span><span class="sxs-lookup"><span data-stu-id="5e27f-1056">Name</span></span>| <span data-ttu-id="5e27f-1057">Тип</span><span class="sxs-lookup"><span data-stu-id="5e27f-1057">Type</span></span>| <span data-ttu-id="5e27f-1058">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="5e27f-1058">Attributes</span></span>| <span data-ttu-id="5e27f-1059">Описание</span><span class="sxs-lookup"><span data-stu-id="5e27f-1059">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="5e27f-1060">String</span><span class="sxs-lookup"><span data-stu-id="5e27f-1060">String</span></span>||<span data-ttu-id="5e27f-p175">Вставляемые данные. Объем данных не должен превышать 1 000 000 символов. Если передано больше 1 000 000 символов, возвращается исключение `ArgumentOutOfRange`.</span><span class="sxs-lookup"><span data-stu-id="5e27f-p175">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="5e27f-1064">Object</span><span class="sxs-lookup"><span data-stu-id="5e27f-1064">Object</span></span>| <span data-ttu-id="5e27f-1065">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="5e27f-1065">&lt;optional&gt;</span></span>|<span data-ttu-id="5e27f-1066">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="5e27f-1066">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="5e27f-1067">Object</span><span class="sxs-lookup"><span data-stu-id="5e27f-1067">Object</span></span>| <span data-ttu-id="5e27f-1068">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="5e27f-1068">&lt;optional&gt;</span></span>|<span data-ttu-id="5e27f-1069">В методе обратного вызова разработчики могут указать любой объект, к которому необходимо получить доступ.</span><span class="sxs-lookup"><span data-stu-id="5e27f-1069">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="5e27f-1070">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="5e27f-1070">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="5e27f-1071">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="5e27f-1071">&lt;optional&gt;</span></span>|<span data-ttu-id="5e27f-1072">Если задано значение `text`, текущий стиль применяется в Outlook в Интернете и классических клиентах.</span><span class="sxs-lookup"><span data-stu-id="5e27f-1072">If `text`, the current style is applied in Outlook on the web and desktop clients.</span></span> <span data-ttu-id="5e27f-1073">Если поле представляет собой редактор HTML, вставляются только текстовые данные, даже если они имеют формат HTML.</span><span class="sxs-lookup"><span data-stu-id="5e27f-1073">If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="5e27f-1074">Если задано значение `html` и поле (не тема) поддерживает HTML, в Outlook в Интернете применяется текущий стиль, а в классических клиентах Outlook — стиль по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="5e27f-1074">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook on the web and the default style is applied in Outlook desktop clients.</span></span> <span data-ttu-id="5e27f-1075">Если поле является текстовым, возвращается ошибка `InvalidDataFormat`.</span><span class="sxs-lookup"><span data-stu-id="5e27f-1075">If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="5e27f-1076">Если свойство `coercionType` не задано, результат зависит от поля: если поле имеет формат HTML, используется текст в формате HTML, а если поле текстовое, применяется обычный текст.</span><span class="sxs-lookup"><span data-stu-id="5e27f-1076">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="5e27f-1077">функция</span><span class="sxs-lookup"><span data-stu-id="5e27f-1077">function</span></span>||<span data-ttu-id="5e27f-1078">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="5e27f-1078">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="5e27f-1079">Требования</span><span class="sxs-lookup"><span data-stu-id="5e27f-1079">Requirements</span></span>

|<span data-ttu-id="5e27f-1080">Требование</span><span class="sxs-lookup"><span data-stu-id="5e27f-1080">Requirement</span></span>| <span data-ttu-id="5e27f-1081">Значение</span><span class="sxs-lookup"><span data-stu-id="5e27f-1081">Value</span></span>|
|---|---|
|[<span data-ttu-id="5e27f-1082">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="5e27f-1082">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5e27f-1083">1.2</span><span class="sxs-lookup"><span data-stu-id="5e27f-1083">1.2</span></span>|
|[<span data-ttu-id="5e27f-1084">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5e27f-1084">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5e27f-1085">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="5e27f-1085">ReadWriteItem</span></span>|
|[<span data-ttu-id="5e27f-1086">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5e27f-1086">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="5e27f-1087">Создание</span><span class="sxs-lookup"><span data-stu-id="5e27f-1087">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="5e27f-1088">Пример</span><span class="sxs-lookup"><span data-stu-id="5e27f-1088">Example</span></span>

```js
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
