---
title: Office. Context. Mailbox. Item — набор требований 1,2
description: ''
ms.date: 08/08/2019
localization_priority: Normal
ms.openlocfilehash: 7475e62c26d24ed9d191ca89934dd5d183b477fa
ms.sourcegitcommit: 1fb99b1b4e63868a0e81a928c69a34c42bf7e209
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/30/2019
ms.locfileid: "36696416"
---
# <a name="item"></a><span data-ttu-id="33c8e-102">item</span><span class="sxs-lookup"><span data-stu-id="33c8e-102">item</span></span>

### <span data-ttu-id="33c8e-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md). item</span><span class="sxs-lookup"><span data-stu-id="33c8e-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md). item</span></span>

<span data-ttu-id="33c8e-p102">Пространство имен `item` используется для доступа к выбранному в данный момент сообщению, приглашению на собрание или описанию встречи. Вы можете определить тип пространства имен `item` с помощью свойства [itemType](#itemtype-officemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="33c8e-p102">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="33c8e-107">Requirements</span><span class="sxs-lookup"><span data-stu-id="33c8e-107">Requirements</span></span>

|<span data-ttu-id="33c8e-108">Требование</span><span class="sxs-lookup"><span data-stu-id="33c8e-108">Requirement</span></span>| <span data-ttu-id="33c8e-109">Значение</span><span class="sxs-lookup"><span data-stu-id="33c8e-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="33c8e-110">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="33c8e-110">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="33c8e-111">1.0</span><span class="sxs-lookup"><span data-stu-id="33c8e-111">1.0</span></span>|
|[<span data-ttu-id="33c8e-112">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="33c8e-112">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="33c8e-113">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="33c8e-113">Restricted</span></span>|
|[<span data-ttu-id="33c8e-114">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="33c8e-114">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="33c8e-115">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="33c8e-115">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="33c8e-116">Элементы и методы</span><span class="sxs-lookup"><span data-stu-id="33c8e-116">Members and methods</span></span>

| <span data-ttu-id="33c8e-117">Элемент	</span><span class="sxs-lookup"><span data-stu-id="33c8e-117">Member</span></span> | <span data-ttu-id="33c8e-118">Тип</span><span class="sxs-lookup"><span data-stu-id="33c8e-118">Type</span></span> |
|--------|------|
| [<span data-ttu-id="33c8e-119">attachments</span><span class="sxs-lookup"><span data-stu-id="33c8e-119">attachments</span></span>](#attachments-arrayattachmentdetails) | <span data-ttu-id="33c8e-120">Member</span><span class="sxs-lookup"><span data-stu-id="33c8e-120">Member</span></span> |
| [<span data-ttu-id="33c8e-121">bcc</span><span class="sxs-lookup"><span data-stu-id="33c8e-121">bcc</span></span>](#bcc-recipients) | <span data-ttu-id="33c8e-122">Member</span><span class="sxs-lookup"><span data-stu-id="33c8e-122">Member</span></span> |
| [<span data-ttu-id="33c8e-123">body</span><span class="sxs-lookup"><span data-stu-id="33c8e-123">body</span></span>](#body-body) | <span data-ttu-id="33c8e-124">Member</span><span class="sxs-lookup"><span data-stu-id="33c8e-124">Member</span></span> |
| [<span data-ttu-id="33c8e-125">cc</span><span class="sxs-lookup"><span data-stu-id="33c8e-125">cc</span></span>](#cc-arrayemailaddressdetailsrecipients) | <span data-ttu-id="33c8e-126">Member</span><span class="sxs-lookup"><span data-stu-id="33c8e-126">Member</span></span> |
| [<span data-ttu-id="33c8e-127">conversationId</span><span class="sxs-lookup"><span data-stu-id="33c8e-127">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="33c8e-128">Member</span><span class="sxs-lookup"><span data-stu-id="33c8e-128">Member</span></span> |
| [<span data-ttu-id="33c8e-129">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="33c8e-129">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="33c8e-130">Member</span><span class="sxs-lookup"><span data-stu-id="33c8e-130">Member</span></span> |
| [<span data-ttu-id="33c8e-131">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="33c8e-131">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="33c8e-132">Member</span><span class="sxs-lookup"><span data-stu-id="33c8e-132">Member</span></span> |
| [<span data-ttu-id="33c8e-133">end</span><span class="sxs-lookup"><span data-stu-id="33c8e-133">end</span></span>](#end-datetime) | <span data-ttu-id="33c8e-134">Member</span><span class="sxs-lookup"><span data-stu-id="33c8e-134">Member</span></span> |
| [<span data-ttu-id="33c8e-135">from</span><span class="sxs-lookup"><span data-stu-id="33c8e-135">from</span></span>](#from-emailaddressdetails) | <span data-ttu-id="33c8e-136">Member</span><span class="sxs-lookup"><span data-stu-id="33c8e-136">Member</span></span> |
| [<span data-ttu-id="33c8e-137">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="33c8e-137">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="33c8e-138">Member</span><span class="sxs-lookup"><span data-stu-id="33c8e-138">Member</span></span> |
| [<span data-ttu-id="33c8e-139">itemClass</span><span class="sxs-lookup"><span data-stu-id="33c8e-139">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="33c8e-140">Member</span><span class="sxs-lookup"><span data-stu-id="33c8e-140">Member</span></span> |
| [<span data-ttu-id="33c8e-141">itemId</span><span class="sxs-lookup"><span data-stu-id="33c8e-141">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="33c8e-142">Member</span><span class="sxs-lookup"><span data-stu-id="33c8e-142">Member</span></span> |
| [<span data-ttu-id="33c8e-143">itemType</span><span class="sxs-lookup"><span data-stu-id="33c8e-143">itemType</span></span>](#itemtype-officemailboxenumsitemtype) | <span data-ttu-id="33c8e-144">Member</span><span class="sxs-lookup"><span data-stu-id="33c8e-144">Member</span></span> |
| [<span data-ttu-id="33c8e-145">location</span><span class="sxs-lookup"><span data-stu-id="33c8e-145">location</span></span>](#location-stringlocation) | <span data-ttu-id="33c8e-146">Member</span><span class="sxs-lookup"><span data-stu-id="33c8e-146">Member</span></span> |
| [<span data-ttu-id="33c8e-147">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="33c8e-147">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="33c8e-148">Member</span><span class="sxs-lookup"><span data-stu-id="33c8e-148">Member</span></span> |
| [<span data-ttu-id="33c8e-149">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="33c8e-149">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="33c8e-150">Member</span><span class="sxs-lookup"><span data-stu-id="33c8e-150">Member</span></span> |
| [<span data-ttu-id="33c8e-151">organizer</span><span class="sxs-lookup"><span data-stu-id="33c8e-151">organizer</span></span>](#organizer-emailaddressdetails) | <span data-ttu-id="33c8e-152">Member</span><span class="sxs-lookup"><span data-stu-id="33c8e-152">Member</span></span> |
| [<span data-ttu-id="33c8e-153">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="33c8e-153">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="33c8e-154">Member</span><span class="sxs-lookup"><span data-stu-id="33c8e-154">Member</span></span> |
| [<span data-ttu-id="33c8e-155">sender</span><span class="sxs-lookup"><span data-stu-id="33c8e-155">sender</span></span>](#sender-emailaddressdetails) | <span data-ttu-id="33c8e-156">Member</span><span class="sxs-lookup"><span data-stu-id="33c8e-156">Member</span></span> |
| [<span data-ttu-id="33c8e-157">start</span><span class="sxs-lookup"><span data-stu-id="33c8e-157">start</span></span>](#start-datetime) | <span data-ttu-id="33c8e-158">Member</span><span class="sxs-lookup"><span data-stu-id="33c8e-158">Member</span></span> |
| [<span data-ttu-id="33c8e-159">subject</span><span class="sxs-lookup"><span data-stu-id="33c8e-159">subject</span></span>](#subject-stringsubject) | <span data-ttu-id="33c8e-160">Элемент</span><span class="sxs-lookup"><span data-stu-id="33c8e-160">Member</span></span> |
| [<span data-ttu-id="33c8e-161">to</span><span class="sxs-lookup"><span data-stu-id="33c8e-161">to</span></span>](#to-arrayemailaddressdetailsrecipients) | <span data-ttu-id="33c8e-162">Элемент</span><span class="sxs-lookup"><span data-stu-id="33c8e-162">Member</span></span> |
| [<span data-ttu-id="33c8e-163">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="33c8e-163">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="33c8e-164">Метод</span><span class="sxs-lookup"><span data-stu-id="33c8e-164">Method</span></span> |
| [<span data-ttu-id="33c8e-165">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="33c8e-165">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="33c8e-166">Метод</span><span class="sxs-lookup"><span data-stu-id="33c8e-166">Method</span></span> |
| [<span data-ttu-id="33c8e-167">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="33c8e-167">displayReplyAllForm</span></span>](#displayreplyallformformdata-callback) | <span data-ttu-id="33c8e-168">Метод</span><span class="sxs-lookup"><span data-stu-id="33c8e-168">Method</span></span> |
| [<span data-ttu-id="33c8e-169">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="33c8e-169">displayReplyForm</span></span>](#displayreplyformformdata-callback) | <span data-ttu-id="33c8e-170">Метод</span><span class="sxs-lookup"><span data-stu-id="33c8e-170">Method</span></span> |
| [<span data-ttu-id="33c8e-171">getEntities</span><span class="sxs-lookup"><span data-stu-id="33c8e-171">getEntities</span></span>](#getentities--entities) | <span data-ttu-id="33c8e-172">Метод</span><span class="sxs-lookup"><span data-stu-id="33c8e-172">Method</span></span> |
| [<span data-ttu-id="33c8e-173">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="33c8e-173">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="33c8e-174">Метод</span><span class="sxs-lookup"><span data-stu-id="33c8e-174">Method</span></span> |
| [<span data-ttu-id="33c8e-175">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="33c8e-175">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="33c8e-176">Метод</span><span class="sxs-lookup"><span data-stu-id="33c8e-176">Method</span></span> |
| [<span data-ttu-id="33c8e-177">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="33c8e-177">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="33c8e-178">Метод</span><span class="sxs-lookup"><span data-stu-id="33c8e-178">Method</span></span> |
| [<span data-ttu-id="33c8e-179">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="33c8e-179">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="33c8e-180">Метод</span><span class="sxs-lookup"><span data-stu-id="33c8e-180">Method</span></span> |
| [<span data-ttu-id="33c8e-181">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="33c8e-181">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="33c8e-182">Метод</span><span class="sxs-lookup"><span data-stu-id="33c8e-182">Method</span></span> |
| [<span data-ttu-id="33c8e-183">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="33c8e-183">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="33c8e-184">Метод</span><span class="sxs-lookup"><span data-stu-id="33c8e-184">Method</span></span> |
| [<span data-ttu-id="33c8e-185">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="33c8e-185">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="33c8e-186">Метод</span><span class="sxs-lookup"><span data-stu-id="33c8e-186">Method</span></span> |
| [<span data-ttu-id="33c8e-187">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="33c8e-187">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="33c8e-188">Метод</span><span class="sxs-lookup"><span data-stu-id="33c8e-188">Method</span></span> |

### <a name="example"></a><span data-ttu-id="33c8e-189">Пример</span><span class="sxs-lookup"><span data-stu-id="33c8e-189">Example</span></span>

<span data-ttu-id="33c8e-190">В примере кода JavaScript, приведенном ниже, показано, как получить доступ к свойству `subject` текущего элемента в Outlook.</span><span class="sxs-lookup"><span data-stu-id="33c8e-190">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="33c8e-191">Элементы</span><span class="sxs-lookup"><span data-stu-id="33c8e-191">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetailsviewoutlook-js-12"></a><span data-ttu-id="33c8e-192">вложения: Array. <[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.2)></span><span class="sxs-lookup"><span data-stu-id="33c8e-192">attachments: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.2)></span></span>

<span data-ttu-id="33c8e-p103">Получает массив вложений для элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="33c8e-p103">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="33c8e-195">Outlook блокирует определенные типы файлов из-за потенциальных проблем с безопасностью, поэтому они не возвращаются.</span><span class="sxs-lookup"><span data-stu-id="33c8e-195">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="33c8e-196">Дополнительные сведения см. в статье [Блокировка вложений в Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="33c8e-196">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="33c8e-197">Тип</span><span class="sxs-lookup"><span data-stu-id="33c8e-197">Type</span></span>

*   <span data-ttu-id="33c8e-198">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.2)></span><span class="sxs-lookup"><span data-stu-id="33c8e-198">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.2)></span></span>

##### <a name="requirements"></a><span data-ttu-id="33c8e-199">Требования</span><span class="sxs-lookup"><span data-stu-id="33c8e-199">Requirements</span></span>

|<span data-ttu-id="33c8e-200">Требование</span><span class="sxs-lookup"><span data-stu-id="33c8e-200">Requirement</span></span>| <span data-ttu-id="33c8e-201">Значение</span><span class="sxs-lookup"><span data-stu-id="33c8e-201">Value</span></span>|
|---|---|
|[<span data-ttu-id="33c8e-202">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="33c8e-202">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="33c8e-203">1.0</span><span class="sxs-lookup"><span data-stu-id="33c8e-203">1.0</span></span>|
|[<span data-ttu-id="33c8e-204">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="33c8e-204">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="33c8e-205">ReadItem</span><span class="sxs-lookup"><span data-stu-id="33c8e-205">ReadItem</span></span>|
|[<span data-ttu-id="33c8e-206">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="33c8e-206">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="33c8e-207">Чтение</span><span class="sxs-lookup"><span data-stu-id="33c8e-207">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="33c8e-208">Пример</span><span class="sxs-lookup"><span data-stu-id="33c8e-208">Example</span></span>

<span data-ttu-id="33c8e-209">С помощью приведенного ниже кода можно создать HTML-строку с подробными сведениями обо всех вложениях для текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="33c8e-209">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

#### <a name="bcc-recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-12"></a><span data-ttu-id="33c8e-210">СК: [получатели](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="33c8e-210">bcc: [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

<span data-ttu-id="33c8e-211">Получает объект, который предоставляет методы для получения или обновления получателей скрытой копии сообщения.</span><span class="sxs-lookup"><span data-stu-id="33c8e-211">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="33c8e-212">Только в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="33c8e-212">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="33c8e-213">Тип</span><span class="sxs-lookup"><span data-stu-id="33c8e-213">Type</span></span>

*   [<span data-ttu-id="33c8e-214">Получатели</span><span class="sxs-lookup"><span data-stu-id="33c8e-214">Recipients</span></span>](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)

##### <a name="requirements"></a><span data-ttu-id="33c8e-215">Требования</span><span class="sxs-lookup"><span data-stu-id="33c8e-215">Requirements</span></span>

|<span data-ttu-id="33c8e-216">Требование</span><span class="sxs-lookup"><span data-stu-id="33c8e-216">Requirement</span></span>| <span data-ttu-id="33c8e-217">Значение</span><span class="sxs-lookup"><span data-stu-id="33c8e-217">Value</span></span>|
|---|---|
|[<span data-ttu-id="33c8e-218">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="33c8e-218">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="33c8e-219">1.1</span><span class="sxs-lookup"><span data-stu-id="33c8e-219">1.1</span></span>|
|[<span data-ttu-id="33c8e-220">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="33c8e-220">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="33c8e-221">ReadItem</span><span class="sxs-lookup"><span data-stu-id="33c8e-221">ReadItem</span></span>|
|[<span data-ttu-id="33c8e-222">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="33c8e-222">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="33c8e-223">Создание</span><span class="sxs-lookup"><span data-stu-id="33c8e-223">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="33c8e-224">Пример</span><span class="sxs-lookup"><span data-stu-id="33c8e-224">Example</span></span>

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

#### <a name="body-bodyjavascriptapioutlookofficebodyviewoutlook-js-12"></a><span data-ttu-id="33c8e-225">основной текст: [Body](/javascript/api/outlook/office.body?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="33c8e-225">body: [Body](/javascript/api/outlook/office.body?view=outlook-js-1.2)</span></span>

<span data-ttu-id="33c8e-226">Получает объект, предоставляющий методы для работы с основным текстом элемента.</span><span class="sxs-lookup"><span data-stu-id="33c8e-226">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="33c8e-227">Тип</span><span class="sxs-lookup"><span data-stu-id="33c8e-227">Type</span></span>

*   [<span data-ttu-id="33c8e-228">Body</span><span class="sxs-lookup"><span data-stu-id="33c8e-228">Body</span></span>](/javascript/api/outlook/office.body?view=outlook-js-1.2)

##### <a name="requirements"></a><span data-ttu-id="33c8e-229">Требования</span><span class="sxs-lookup"><span data-stu-id="33c8e-229">Requirements</span></span>

|<span data-ttu-id="33c8e-230">Требование</span><span class="sxs-lookup"><span data-stu-id="33c8e-230">Requirement</span></span>| <span data-ttu-id="33c8e-231">Значение</span><span class="sxs-lookup"><span data-stu-id="33c8e-231">Value</span></span>|
|---|---|
|[<span data-ttu-id="33c8e-232">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="33c8e-232">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="33c8e-233">1.1</span><span class="sxs-lookup"><span data-stu-id="33c8e-233">1.1</span></span>|
|[<span data-ttu-id="33c8e-234">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="33c8e-234">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="33c8e-235">ReadItem</span><span class="sxs-lookup"><span data-stu-id="33c8e-235">ReadItem</span></span>|
|[<span data-ttu-id="33c8e-236">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="33c8e-236">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="33c8e-237">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="33c8e-237">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="33c8e-238">Пример</span><span class="sxs-lookup"><span data-stu-id="33c8e-238">Example</span></span>

<span data-ttu-id="33c8e-239">В этом примере возвращается текст сообщения в формате обычного текста.</span><span class="sxs-lookup"><span data-stu-id="33c8e-239">This example gets the body of the message in plain text.</span></span>

```js
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="33c8e-240">Ниже приведен пример итогового параметра, переданного функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="33c8e-240">The following is an example of the result parameter passed to the callback function.</span></span>

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

#### <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-12recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-12"></a><span data-ttu-id="33c8e-241">CC: Array. <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)>|[получатели](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="33c8e-241">cc: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

<span data-ttu-id="33c8e-242">Предоставляет доступ к получателям копии сообщения.</span><span class="sxs-lookup"><span data-stu-id="33c8e-242">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="33c8e-243">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="33c8e-243">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="33c8e-244">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="33c8e-244">Read mode</span></span>

<span data-ttu-id="33c8e-p107">Свойство `cc` возвращает массив, который содержит объект `EmailAddressDetails` для каждого получателя, указанного в строке **Копия** сообщения. Коллекция может включать не более 100 элементов.</span><span class="sxs-lookup"><span data-stu-id="33c8e-p107">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="33c8e-247">Режим создания</span><span class="sxs-lookup"><span data-stu-id="33c8e-247">Compose mode</span></span>

<span data-ttu-id="33c8e-248">Свойство `cc` возвращает объект `Recipients`, предоставляющий методы для получения или обновления получателей, которые указаны в строке **Копия** сообщения.</span><span class="sxs-lookup"><span data-stu-id="33c8e-248">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

```js
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="33c8e-249">Тип</span><span class="sxs-lookup"><span data-stu-id="33c8e-249">Type</span></span>

*   <span data-ttu-id="33c8e-250">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="33c8e-250">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

##### <a name="requirements"></a><span data-ttu-id="33c8e-251">Требования</span><span class="sxs-lookup"><span data-stu-id="33c8e-251">Requirements</span></span>

|<span data-ttu-id="33c8e-252">Требование</span><span class="sxs-lookup"><span data-stu-id="33c8e-252">Requirement</span></span>| <span data-ttu-id="33c8e-253">Значение</span><span class="sxs-lookup"><span data-stu-id="33c8e-253">Value</span></span>|
|---|---|
|[<span data-ttu-id="33c8e-254">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="33c8e-254">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="33c8e-255">1.0</span><span class="sxs-lookup"><span data-stu-id="33c8e-255">1.0</span></span>|
|[<span data-ttu-id="33c8e-256">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="33c8e-256">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="33c8e-257">ReadItem</span><span class="sxs-lookup"><span data-stu-id="33c8e-257">ReadItem</span></span>|
|[<span data-ttu-id="33c8e-258">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="33c8e-258">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="33c8e-259">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="33c8e-259">Compose or Read</span></span>|

<br>

---
---

#### <a name="nullable-conversationid-string"></a><span data-ttu-id="33c8e-260">(Nullable) conversationId: строка</span><span class="sxs-lookup"><span data-stu-id="33c8e-260">(nullable) conversationId: String</span></span>

<span data-ttu-id="33c8e-261">Получает идентификатор разговора по электронной почте, содержащего конкретное сообщение.</span><span class="sxs-lookup"><span data-stu-id="33c8e-261">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="33c8e-p108">Вы можете получить целочисленное значение этого свойства, если ваше почтовое приложение активируется в формах просмотра или формах создания ответов. Если пользователь изменит тему ответа, после его отправки идентификатор беседы будет изменен, и полученное ранее значение будет недействительным.</span><span class="sxs-lookup"><span data-stu-id="33c8e-p108">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="33c8e-p109">Это свойство имеет значение NULL для нового элемента в форме создания. Свойство `conversationId` вернет значение, если пользователь задаст тему и сохранит элемент.</span><span class="sxs-lookup"><span data-stu-id="33c8e-p109">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="33c8e-266">Тип</span><span class="sxs-lookup"><span data-stu-id="33c8e-266">Type</span></span>

*   <span data-ttu-id="33c8e-267">String</span><span class="sxs-lookup"><span data-stu-id="33c8e-267">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="33c8e-268">Требования</span><span class="sxs-lookup"><span data-stu-id="33c8e-268">Requirements</span></span>

|<span data-ttu-id="33c8e-269">Требование</span><span class="sxs-lookup"><span data-stu-id="33c8e-269">Requirement</span></span>| <span data-ttu-id="33c8e-270">Значение</span><span class="sxs-lookup"><span data-stu-id="33c8e-270">Value</span></span>|
|---|---|
|[<span data-ttu-id="33c8e-271">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="33c8e-271">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="33c8e-272">1.0</span><span class="sxs-lookup"><span data-stu-id="33c8e-272">1.0</span></span>|
|[<span data-ttu-id="33c8e-273">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="33c8e-273">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="33c8e-274">ReadItem</span><span class="sxs-lookup"><span data-stu-id="33c8e-274">ReadItem</span></span>|
|[<span data-ttu-id="33c8e-275">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="33c8e-275">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="33c8e-276">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="33c8e-276">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="33c8e-277">Пример</span><span class="sxs-lookup"><span data-stu-id="33c8e-277">Example</span></span>

```js
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

<br>

---
---

#### <a name="datetimecreated-date"></a><span data-ttu-id="33c8e-278">dateTimeCreated: Дата</span><span class="sxs-lookup"><span data-stu-id="33c8e-278">dateTimeCreated: Date</span></span>

<span data-ttu-id="33c8e-p110">Получает дату и время создания элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="33c8e-p110">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="33c8e-281">Тип</span><span class="sxs-lookup"><span data-stu-id="33c8e-281">Type</span></span>

*   <span data-ttu-id="33c8e-282">Дата</span><span class="sxs-lookup"><span data-stu-id="33c8e-282">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="33c8e-283">Требования</span><span class="sxs-lookup"><span data-stu-id="33c8e-283">Requirements</span></span>

|<span data-ttu-id="33c8e-284">Требование</span><span class="sxs-lookup"><span data-stu-id="33c8e-284">Requirement</span></span>| <span data-ttu-id="33c8e-285">Значение</span><span class="sxs-lookup"><span data-stu-id="33c8e-285">Value</span></span>|
|---|---|
|[<span data-ttu-id="33c8e-286">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="33c8e-286">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="33c8e-287">1.0</span><span class="sxs-lookup"><span data-stu-id="33c8e-287">1.0</span></span>|
|[<span data-ttu-id="33c8e-288">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="33c8e-288">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="33c8e-289">ReadItem</span><span class="sxs-lookup"><span data-stu-id="33c8e-289">ReadItem</span></span>|
|[<span data-ttu-id="33c8e-290">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="33c8e-290">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="33c8e-291">Чтение</span><span class="sxs-lookup"><span data-stu-id="33c8e-291">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="33c8e-292">Пример</span><span class="sxs-lookup"><span data-stu-id="33c8e-292">Example</span></span>

```js
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

<br>

---
---

#### <a name="datetimemodified-date"></a><span data-ttu-id="33c8e-293">dateTimeModified: Дата</span><span class="sxs-lookup"><span data-stu-id="33c8e-293">dateTimeModified: Date</span></span>

<span data-ttu-id="33c8e-294">Получает дату и время последнего изменения элемента.</span><span class="sxs-lookup"><span data-stu-id="33c8e-294">Gets the date and time that an item was last modified.</span></span> <span data-ttu-id="33c8e-295">Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="33c8e-295">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="33c8e-296">Этот элемент не поддерживается в Outlook на iOS или Android.</span><span class="sxs-lookup"><span data-stu-id="33c8e-296">This member is not supported in Outlook on iOS or Android.</span></span>

##### <a name="type"></a><span data-ttu-id="33c8e-297">Тип</span><span class="sxs-lookup"><span data-stu-id="33c8e-297">Type</span></span>

*   <span data-ttu-id="33c8e-298">Дата</span><span class="sxs-lookup"><span data-stu-id="33c8e-298">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="33c8e-299">Требования</span><span class="sxs-lookup"><span data-stu-id="33c8e-299">Requirements</span></span>

|<span data-ttu-id="33c8e-300">Требование</span><span class="sxs-lookup"><span data-stu-id="33c8e-300">Requirement</span></span>| <span data-ttu-id="33c8e-301">Значение</span><span class="sxs-lookup"><span data-stu-id="33c8e-301">Value</span></span>|
|---|---|
|[<span data-ttu-id="33c8e-302">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="33c8e-302">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="33c8e-303">1.0</span><span class="sxs-lookup"><span data-stu-id="33c8e-303">1.0</span></span>|
|[<span data-ttu-id="33c8e-304">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="33c8e-304">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="33c8e-305">ReadItem</span><span class="sxs-lookup"><span data-stu-id="33c8e-305">ReadItem</span></span>|
|[<span data-ttu-id="33c8e-306">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="33c8e-306">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="33c8e-307">Чтение</span><span class="sxs-lookup"><span data-stu-id="33c8e-307">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="33c8e-308">Пример</span><span class="sxs-lookup"><span data-stu-id="33c8e-308">Example</span></span>

```js
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

<br>

---
---

#### <a name="end-datetimejavascriptapioutlookofficetimeviewoutlook-js-12"></a><span data-ttu-id="33c8e-309">конец: Дата | [Time (время](/javascript/api/outlook/office.time?view=outlook-js-1.2) )</span><span class="sxs-lookup"><span data-stu-id="33c8e-309">end: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.2)</span></span>

<span data-ttu-id="33c8e-310">Получает или задает дату и время окончания встречи.</span><span class="sxs-lookup"><span data-stu-id="33c8e-310">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="33c8e-p112">Свойство `end` представлено в виде значения даты и времени в формате UTC. Преобразовать значение свойства end в местные значения даты и времени клиента можно с помощью метода [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime).</span><span class="sxs-lookup"><span data-stu-id="33c8e-p112">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="33c8e-313">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="33c8e-313">Read mode</span></span>

<span data-ttu-id="33c8e-314">Свойство `end` возвращает объект `Date`.</span><span class="sxs-lookup"><span data-stu-id="33c8e-314">The `end` property returns a `Date` object.</span></span>

```js
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="33c8e-315">Режим создания</span><span class="sxs-lookup"><span data-stu-id="33c8e-315">Compose mode</span></span>

<span data-ttu-id="33c8e-316">Свойство `end` возвращает объект `Time`.</span><span class="sxs-lookup"><span data-stu-id="33c8e-316">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="33c8e-317">Если вы задаете время окончания с помощью метода [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.2#setasync-datetime--options--callback-), необходимо использовать метод [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) для преобразования местного времени на клиенте в формат UTC для сервера.</span><span class="sxs-lookup"><span data-stu-id="33c8e-317">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.2#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="33c8e-318">В примере ниже показано, как с помощью метода [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.2#setasync-datetime--options--callback-) объекта `Time` задать время окончания встречи.</span><span class="sxs-lookup"><span data-stu-id="33c8e-318">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.2#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="33c8e-319">Тип</span><span class="sxs-lookup"><span data-stu-id="33c8e-319">Type</span></span>

*   <span data-ttu-id="33c8e-320">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="33c8e-320">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.2)</span></span>

##### <a name="requirements"></a><span data-ttu-id="33c8e-321">Требования</span><span class="sxs-lookup"><span data-stu-id="33c8e-321">Requirements</span></span>

|<span data-ttu-id="33c8e-322">Требование</span><span class="sxs-lookup"><span data-stu-id="33c8e-322">Requirement</span></span>| <span data-ttu-id="33c8e-323">Значение</span><span class="sxs-lookup"><span data-stu-id="33c8e-323">Value</span></span>|
|---|---|
|[<span data-ttu-id="33c8e-324">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="33c8e-324">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="33c8e-325">1.0</span><span class="sxs-lookup"><span data-stu-id="33c8e-325">1.0</span></span>|
|[<span data-ttu-id="33c8e-326">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="33c8e-326">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="33c8e-327">ReadItem</span><span class="sxs-lookup"><span data-stu-id="33c8e-327">ReadItem</span></span>|
|[<span data-ttu-id="33c8e-328">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="33c8e-328">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="33c8e-329">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="33c8e-329">Compose or Read</span></span>|

<br>

---
---

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-12"></a><span data-ttu-id="33c8e-330">от: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="33c8e-330">from: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)</span></span>

<span data-ttu-id="33c8e-p113">Получает электронный адрес отправителя сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="33c8e-p113">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="33c8e-p114">Свойства `from` и [`sender`](#sender-emailaddressdetails) представляют одно лицо, если сообщение не отправлено представителем. В противном случае свойство `from` представляет лицо, делегировавшее полномочия, а свойство sender — представителя.</span><span class="sxs-lookup"><span data-stu-id="33c8e-p114">The `from` and [`sender`](#sender-emailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="33c8e-335">Свойству `recipientType`, принадлежащему объекту `EmailAddressDetails` в свойстве `from`, задано значение `undefined`.</span><span class="sxs-lookup"><span data-stu-id="33c8e-335">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="33c8e-336">Тип</span><span class="sxs-lookup"><span data-stu-id="33c8e-336">Type</span></span>

*   [<span data-ttu-id="33c8e-337">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="33c8e-337">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)

##### <a name="requirements"></a><span data-ttu-id="33c8e-338">Требования</span><span class="sxs-lookup"><span data-stu-id="33c8e-338">Requirements</span></span>

|<span data-ttu-id="33c8e-339">Требование</span><span class="sxs-lookup"><span data-stu-id="33c8e-339">Requirement</span></span>| <span data-ttu-id="33c8e-340">Значение</span><span class="sxs-lookup"><span data-stu-id="33c8e-340">Value</span></span>|
|---|---|
|[<span data-ttu-id="33c8e-341">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="33c8e-341">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="33c8e-342">1.0</span><span class="sxs-lookup"><span data-stu-id="33c8e-342">1.0</span></span>|
|[<span data-ttu-id="33c8e-343">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="33c8e-343">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="33c8e-344">ReadItem</span><span class="sxs-lookup"><span data-stu-id="33c8e-344">ReadItem</span></span>|
|[<span data-ttu-id="33c8e-345">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="33c8e-345">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="33c8e-346">Чтение</span><span class="sxs-lookup"><span data-stu-id="33c8e-346">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="33c8e-347">Пример</span><span class="sxs-lookup"><span data-stu-id="33c8e-347">Example</span></span>

```js
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

<br>

---
---

#### <a name="internetmessageid-string"></a><span data-ttu-id="33c8e-348">internetMessageId: строка</span><span class="sxs-lookup"><span data-stu-id="33c8e-348">internetMessageId: String</span></span>

<span data-ttu-id="33c8e-p115">Получает идентификатор интернет-сообщения для электронного сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="33c8e-p115">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="33c8e-351">Тип</span><span class="sxs-lookup"><span data-stu-id="33c8e-351">Type</span></span>

*   <span data-ttu-id="33c8e-352">String</span><span class="sxs-lookup"><span data-stu-id="33c8e-352">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="33c8e-353">Требования</span><span class="sxs-lookup"><span data-stu-id="33c8e-353">Requirements</span></span>

|<span data-ttu-id="33c8e-354">Требование</span><span class="sxs-lookup"><span data-stu-id="33c8e-354">Requirement</span></span>| <span data-ttu-id="33c8e-355">Значение</span><span class="sxs-lookup"><span data-stu-id="33c8e-355">Value</span></span>|
|---|---|
|[<span data-ttu-id="33c8e-356">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="33c8e-356">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="33c8e-357">1.0</span><span class="sxs-lookup"><span data-stu-id="33c8e-357">1.0</span></span>|
|[<span data-ttu-id="33c8e-358">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="33c8e-358">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="33c8e-359">ReadItem</span><span class="sxs-lookup"><span data-stu-id="33c8e-359">ReadItem</span></span>|
|[<span data-ttu-id="33c8e-360">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="33c8e-360">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="33c8e-361">Чтение</span><span class="sxs-lookup"><span data-stu-id="33c8e-361">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="33c8e-362">Пример</span><span class="sxs-lookup"><span data-stu-id="33c8e-362">Example</span></span>

```js
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

<br>

---
---

#### <a name="itemclass-string"></a><span data-ttu-id="33c8e-363">itemClass: строка</span><span class="sxs-lookup"><span data-stu-id="33c8e-363">itemClass: String</span></span>

<span data-ttu-id="33c8e-p116">Получает класс элемента веб-служб Exchange для выбранного элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="33c8e-p116">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="33c8e-p117">Свойство `itemClass` указывает класс сообщения выбранного элемента. Ниже приводятся классы сообщения по умолчанию для элемента сообщения или встречи.</span><span class="sxs-lookup"><span data-stu-id="33c8e-p117">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="33c8e-368">Тип</span><span class="sxs-lookup"><span data-stu-id="33c8e-368">Type</span></span> | <span data-ttu-id="33c8e-369">Описание</span><span class="sxs-lookup"><span data-stu-id="33c8e-369">Description</span></span> | <span data-ttu-id="33c8e-370">Класс элемента</span><span class="sxs-lookup"><span data-stu-id="33c8e-370">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="33c8e-371">Элементы встречи</span><span class="sxs-lookup"><span data-stu-id="33c8e-371">Appointment items</span></span> | <span data-ttu-id="33c8e-372">Это элементы календаря для класса элемента `IPM.Appointment` или `IPM.Appointment.Occurrence`.</span><span class="sxs-lookup"><span data-stu-id="33c8e-372">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurrence` |
| <span data-ttu-id="33c8e-373">Элементы сообщения</span><span class="sxs-lookup"><span data-stu-id="33c8e-373">Message items</span></span> | <span data-ttu-id="33c8e-374">Сюда входят электронные сообщения, для которых по умолчанию задан класс сообщения `IPM.Note`, а также приглашения на собрания, ответы на них и уведомления об их отмене, использующие `IPM.Schedule.Meeting` в качестве базового класса сообщения.</span><span class="sxs-lookup"><span data-stu-id="33c8e-374">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="33c8e-375">Можно создавать настраиваемые классы сообщения, расширяющие классы сообщения по умолчанию, например настраиваемый класс сообщения о встрече `IPM.Appointment.Contoso`.</span><span class="sxs-lookup"><span data-stu-id="33c8e-375">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="33c8e-376">Тип</span><span class="sxs-lookup"><span data-stu-id="33c8e-376">Type</span></span>

*   <span data-ttu-id="33c8e-377">String</span><span class="sxs-lookup"><span data-stu-id="33c8e-377">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="33c8e-378">Требования</span><span class="sxs-lookup"><span data-stu-id="33c8e-378">Requirements</span></span>

|<span data-ttu-id="33c8e-379">Требование</span><span class="sxs-lookup"><span data-stu-id="33c8e-379">Requirement</span></span>| <span data-ttu-id="33c8e-380">Значение</span><span class="sxs-lookup"><span data-stu-id="33c8e-380">Value</span></span>|
|---|---|
|[<span data-ttu-id="33c8e-381">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="33c8e-381">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="33c8e-382">1.0</span><span class="sxs-lookup"><span data-stu-id="33c8e-382">1.0</span></span>|
|[<span data-ttu-id="33c8e-383">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="33c8e-383">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="33c8e-384">ReadItem</span><span class="sxs-lookup"><span data-stu-id="33c8e-384">ReadItem</span></span>|
|[<span data-ttu-id="33c8e-385">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="33c8e-385">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="33c8e-386">Чтение</span><span class="sxs-lookup"><span data-stu-id="33c8e-386">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="33c8e-387">Пример</span><span class="sxs-lookup"><span data-stu-id="33c8e-387">Example</span></span>

```js
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

<br>

---
---

#### <a name="nullable-itemid-string"></a><span data-ttu-id="33c8e-388">(Nullable) itemId: строка</span><span class="sxs-lookup"><span data-stu-id="33c8e-388">(nullable) itemId: String</span></span>

<span data-ttu-id="33c8e-389">Получает идентификатор элемента веб-служб Exchange для текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="33c8e-389">Gets the Exchange Web Services item identifier for the current item.</span></span> <span data-ttu-id="33c8e-390">Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="33c8e-390">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="33c8e-391">Идентификатор, возвращаемый свойством `itemId`, совпадает с идентификатором элемента веб-служб Exchange.</span><span class="sxs-lookup"><span data-stu-id="33c8e-391">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="33c8e-392">Свойство `itemId` не совпадает с идентификатором записи Outlook, а также идентификатором, который используется REST API Outlook.</span><span class="sxs-lookup"><span data-stu-id="33c8e-392">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="33c8e-393">Перед выполнением вызовов API REST, использующих это значение, его `Office.context.mailbox.convertToRestId`необходимо преобразовать с помощью, которое доступно в наборе требований 1,3.</span><span class="sxs-lookup"><span data-stu-id="33c8e-393">Before making REST API calls using this value, it should be converted using `Office.context.mailbox.convertToRestId`, which is available starting in requirement set 1.3.</span></span> <span data-ttu-id="33c8e-394">Дополнительные сведения см. в статье [Использование REST API Outlook из надстройки Outlook](/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="33c8e-394">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

##### <a name="type"></a><span data-ttu-id="33c8e-395">Тип</span><span class="sxs-lookup"><span data-stu-id="33c8e-395">Type</span></span>

*   <span data-ttu-id="33c8e-396">String</span><span class="sxs-lookup"><span data-stu-id="33c8e-396">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="33c8e-397">Требования</span><span class="sxs-lookup"><span data-stu-id="33c8e-397">Requirements</span></span>

|<span data-ttu-id="33c8e-398">Требование</span><span class="sxs-lookup"><span data-stu-id="33c8e-398">Requirement</span></span>| <span data-ttu-id="33c8e-399">Значение</span><span class="sxs-lookup"><span data-stu-id="33c8e-399">Value</span></span>|
|---|---|
|[<span data-ttu-id="33c8e-400">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="33c8e-400">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="33c8e-401">1.0</span><span class="sxs-lookup"><span data-stu-id="33c8e-401">1.0</span></span>|
|[<span data-ttu-id="33c8e-402">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="33c8e-402">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="33c8e-403">ReadItem</span><span class="sxs-lookup"><span data-stu-id="33c8e-403">ReadItem</span></span>|
|[<span data-ttu-id="33c8e-404">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="33c8e-404">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="33c8e-405">Чтение</span><span class="sxs-lookup"><span data-stu-id="33c8e-405">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="33c8e-406">Пример</span><span class="sxs-lookup"><span data-stu-id="33c8e-406">Example</span></span>

<span data-ttu-id="33c8e-p120">Указанный ниже код проверяет наличие идентификатора элемента. Если свойство `itemId` возвращает значение `null` или `undefined`, элемент будет сохранен в хранилище, а из асинхронного результата будет получен идентификатор элемента.</span><span class="sxs-lookup"><span data-stu-id="33c8e-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

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

#### <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtypeviewoutlook-js-12"></a><span data-ttu-id="33c8e-409">itemType: [Office. MailboxEnums. ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="33c8e-409">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.2)</span></span>

<span data-ttu-id="33c8e-410">Получает тип элемента, который представляет экземпляр.</span><span class="sxs-lookup"><span data-stu-id="33c8e-410">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="33c8e-411">Свойство `itemType` возвращает одно из значений перечисления `ItemType`, которое указывает, является ли экземпляр объекта `item` сообщением или встречей.</span><span class="sxs-lookup"><span data-stu-id="33c8e-411">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="33c8e-412">Тип</span><span class="sxs-lookup"><span data-stu-id="33c8e-412">Type</span></span>

*   [<span data-ttu-id="33c8e-413">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="33c8e-413">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.2)

##### <a name="requirements"></a><span data-ttu-id="33c8e-414">Требования</span><span class="sxs-lookup"><span data-stu-id="33c8e-414">Requirements</span></span>

|<span data-ttu-id="33c8e-415">Требование</span><span class="sxs-lookup"><span data-stu-id="33c8e-415">Requirement</span></span>| <span data-ttu-id="33c8e-416">Значение</span><span class="sxs-lookup"><span data-stu-id="33c8e-416">Value</span></span>|
|---|---|
|[<span data-ttu-id="33c8e-417">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="33c8e-417">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="33c8e-418">1.0</span><span class="sxs-lookup"><span data-stu-id="33c8e-418">1.0</span></span>|
|[<span data-ttu-id="33c8e-419">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="33c8e-419">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="33c8e-420">ReadItem</span><span class="sxs-lookup"><span data-stu-id="33c8e-420">ReadItem</span></span>|
|[<span data-ttu-id="33c8e-421">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="33c8e-421">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="33c8e-422">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="33c8e-422">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="33c8e-423">Пример</span><span class="sxs-lookup"><span data-stu-id="33c8e-423">Example</span></span>

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

#### <a name="location-stringlocationjavascriptapioutlookofficelocationviewoutlook-js-12"></a><span data-ttu-id="33c8e-424">Местоположение: строка | [Location (расположение](/javascript/api/outlook/office.location?view=outlook-js-1.2) )</span><span class="sxs-lookup"><span data-stu-id="33c8e-424">location: String|[Location](/javascript/api/outlook/office.location?view=outlook-js-1.2)</span></span>

<span data-ttu-id="33c8e-425">Получает или задает место встречи.</span><span class="sxs-lookup"><span data-stu-id="33c8e-425">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="33c8e-426">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="33c8e-426">Read mode</span></span>

<span data-ttu-id="33c8e-427">Свойство `location` возвращает строку, содержащую сведения о месте встречи.</span><span class="sxs-lookup"><span data-stu-id="33c8e-427">The `location` property returns a string that contains the location of the appointment.</span></span>

```js
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="33c8e-428">Режим создания</span><span class="sxs-lookup"><span data-stu-id="33c8e-428">Compose mode</span></span>

<span data-ttu-id="33c8e-429">Свойство `location` возвращает объект `Location`, предоставляющий методы, которые используются для получения и задания места встречи.</span><span class="sxs-lookup"><span data-stu-id="33c8e-429">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```js
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="33c8e-430">Тип</span><span class="sxs-lookup"><span data-stu-id="33c8e-430">Type</span></span>

*   <span data-ttu-id="33c8e-431">String | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="33c8e-431">String | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.2)</span></span>

##### <a name="requirements"></a><span data-ttu-id="33c8e-432">Требования</span><span class="sxs-lookup"><span data-stu-id="33c8e-432">Requirements</span></span>

|<span data-ttu-id="33c8e-433">Требование</span><span class="sxs-lookup"><span data-stu-id="33c8e-433">Requirement</span></span>| <span data-ttu-id="33c8e-434">Значение</span><span class="sxs-lookup"><span data-stu-id="33c8e-434">Value</span></span>|
|---|---|
|[<span data-ttu-id="33c8e-435">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="33c8e-435">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="33c8e-436">1.0</span><span class="sxs-lookup"><span data-stu-id="33c8e-436">1.0</span></span>|
|[<span data-ttu-id="33c8e-437">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="33c8e-437">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="33c8e-438">ReadItem</span><span class="sxs-lookup"><span data-stu-id="33c8e-438">ReadItem</span></span>|
|[<span data-ttu-id="33c8e-439">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="33c8e-439">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="33c8e-440">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="33c8e-440">Compose or Read</span></span>|

<br>

---
---

#### <a name="normalizedsubject-string"></a><span data-ttu-id="33c8e-441">normalizedSubject: строка</span><span class="sxs-lookup"><span data-stu-id="33c8e-441">normalizedSubject: String</span></span>

<span data-ttu-id="33c8e-p121">Получает тему элемента со всеми удаленными префиксами (включая `RE:` и `FWD:`). Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="33c8e-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="33c8e-p122">Свойство normalizedSubject получает тему элемента со стандартными префиксами (такими как `RE:` и `FW:`), добавляемыми почтовыми программами. Для получения темы элемента с неизмененными префиксами используйте свойство [`subject`](#subject-stringsubject).</span><span class="sxs-lookup"><span data-stu-id="33c8e-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="33c8e-446">Тип</span><span class="sxs-lookup"><span data-stu-id="33c8e-446">Type</span></span>

*   <span data-ttu-id="33c8e-447">String</span><span class="sxs-lookup"><span data-stu-id="33c8e-447">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="33c8e-448">Требования</span><span class="sxs-lookup"><span data-stu-id="33c8e-448">Requirements</span></span>

|<span data-ttu-id="33c8e-449">Требование</span><span class="sxs-lookup"><span data-stu-id="33c8e-449">Requirement</span></span>| <span data-ttu-id="33c8e-450">Значение</span><span class="sxs-lookup"><span data-stu-id="33c8e-450">Value</span></span>|
|---|---|
|[<span data-ttu-id="33c8e-451">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="33c8e-451">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="33c8e-452">1.0</span><span class="sxs-lookup"><span data-stu-id="33c8e-452">1.0</span></span>|
|[<span data-ttu-id="33c8e-453">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="33c8e-453">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="33c8e-454">ReadItem</span><span class="sxs-lookup"><span data-stu-id="33c8e-454">ReadItem</span></span>|
|[<span data-ttu-id="33c8e-455">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="33c8e-455">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="33c8e-456">Чтение</span><span class="sxs-lookup"><span data-stu-id="33c8e-456">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="33c8e-457">Пример</span><span class="sxs-lookup"><span data-stu-id="33c8e-457">Example</span></span>

```js
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

<br>

---
---

#### <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-12recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-12"></a><span data-ttu-id="33c8e-458">optionalAttendees: Array. <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)>|[получатели](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="33c8e-458">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

<span data-ttu-id="33c8e-459">Предоставляет доступ к необязательным участникам события.</span><span class="sxs-lookup"><span data-stu-id="33c8e-459">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="33c8e-460">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="33c8e-460">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="33c8e-461">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="33c8e-461">Read mode</span></span>

<span data-ttu-id="33c8e-462">Свойство `optionalAttendees` возвращает массив, содержащий объект `EmailAddressDetails` для каждого необязательного участника собрания.</span><span class="sxs-lookup"><span data-stu-id="33c8e-462">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

```js
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="33c8e-463">Режим создания</span><span class="sxs-lookup"><span data-stu-id="33c8e-463">Compose mode</span></span>

<span data-ttu-id="33c8e-464">Свойство `optionalAttendees` возвращает объект `Recipients`, который предоставляет методы для получения или обновления необязательных участников собрания.</span><span class="sxs-lookup"><span data-stu-id="33c8e-464">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

```js
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="33c8e-465">Тип</span><span class="sxs-lookup"><span data-stu-id="33c8e-465">Type</span></span>

*   <span data-ttu-id="33c8e-466">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="33c8e-466">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

##### <a name="requirements"></a><span data-ttu-id="33c8e-467">Требования</span><span class="sxs-lookup"><span data-stu-id="33c8e-467">Requirements</span></span>

|<span data-ttu-id="33c8e-468">Требование</span><span class="sxs-lookup"><span data-stu-id="33c8e-468">Requirement</span></span>| <span data-ttu-id="33c8e-469">Значение</span><span class="sxs-lookup"><span data-stu-id="33c8e-469">Value</span></span>|
|---|---|
|[<span data-ttu-id="33c8e-470">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="33c8e-470">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="33c8e-471">1.0</span><span class="sxs-lookup"><span data-stu-id="33c8e-471">1.0</span></span>|
|[<span data-ttu-id="33c8e-472">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="33c8e-472">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="33c8e-473">ReadItem</span><span class="sxs-lookup"><span data-stu-id="33c8e-473">ReadItem</span></span>|
|[<span data-ttu-id="33c8e-474">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="33c8e-474">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="33c8e-475">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="33c8e-475">Compose or Read</span></span>|

<br>

---
---

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-12"></a><span data-ttu-id="33c8e-476">Организатор: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="33c8e-476">organizer: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)</span></span>

<span data-ttu-id="33c8e-p124">Получает электронный адрес организатора указанного собрания. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="33c8e-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="33c8e-479">Тип</span><span class="sxs-lookup"><span data-stu-id="33c8e-479">Type</span></span>

*   [<span data-ttu-id="33c8e-480">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="33c8e-480">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)

##### <a name="requirements"></a><span data-ttu-id="33c8e-481">Требования</span><span class="sxs-lookup"><span data-stu-id="33c8e-481">Requirements</span></span>

|<span data-ttu-id="33c8e-482">Требование</span><span class="sxs-lookup"><span data-stu-id="33c8e-482">Requirement</span></span>| <span data-ttu-id="33c8e-483">Значение</span><span class="sxs-lookup"><span data-stu-id="33c8e-483">Value</span></span>|
|---|---|
|[<span data-ttu-id="33c8e-484">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="33c8e-484">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="33c8e-485">1.0</span><span class="sxs-lookup"><span data-stu-id="33c8e-485">1.0</span></span>|
|[<span data-ttu-id="33c8e-486">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="33c8e-486">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="33c8e-487">ReadItem</span><span class="sxs-lookup"><span data-stu-id="33c8e-487">ReadItem</span></span>|
|[<span data-ttu-id="33c8e-488">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="33c8e-488">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="33c8e-489">Чтение</span><span class="sxs-lookup"><span data-stu-id="33c8e-489">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="33c8e-490">Пример</span><span class="sxs-lookup"><span data-stu-id="33c8e-490">Example</span></span>

```js
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

<br>

---
---

#### <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-12recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-12"></a><span data-ttu-id="33c8e-491">requiredAttendees: Array. <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)>|[получатели](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="33c8e-491">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

<span data-ttu-id="33c8e-492">Предоставляет доступ к обязательным участникам события.</span><span class="sxs-lookup"><span data-stu-id="33c8e-492">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="33c8e-493">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="33c8e-493">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="33c8e-494">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="33c8e-494">Read mode</span></span>

<span data-ttu-id="33c8e-495">Свойство `requiredAttendees` возвращает массив, содержащий объект `EmailAddressDetails` для каждого обязательного участника собрания.</span><span class="sxs-lookup"><span data-stu-id="33c8e-495">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

```js
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="33c8e-496">Режим создания</span><span class="sxs-lookup"><span data-stu-id="33c8e-496">Compose mode</span></span>

<span data-ttu-id="33c8e-497">Свойство `requiredAttendees` возвращает объект `Recipients`, предоставляющий методы, с помощью которых можно получить или обновить сведения об обязательных участниках собрания.</span><span class="sxs-lookup"><span data-stu-id="33c8e-497">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

```js
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="33c8e-498">Тип</span><span class="sxs-lookup"><span data-stu-id="33c8e-498">Type</span></span>

*   <span data-ttu-id="33c8e-499">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="33c8e-499">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

##### <a name="requirements"></a><span data-ttu-id="33c8e-500">Требования</span><span class="sxs-lookup"><span data-stu-id="33c8e-500">Requirements</span></span>

|<span data-ttu-id="33c8e-501">Требование</span><span class="sxs-lookup"><span data-stu-id="33c8e-501">Requirement</span></span>| <span data-ttu-id="33c8e-502">Значение</span><span class="sxs-lookup"><span data-stu-id="33c8e-502">Value</span></span>|
|---|---|
|[<span data-ttu-id="33c8e-503">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="33c8e-503">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="33c8e-504">1.0</span><span class="sxs-lookup"><span data-stu-id="33c8e-504">1.0</span></span>|
|[<span data-ttu-id="33c8e-505">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="33c8e-505">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="33c8e-506">ReadItem</span><span class="sxs-lookup"><span data-stu-id="33c8e-506">ReadItem</span></span>|
|[<span data-ttu-id="33c8e-507">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="33c8e-507">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="33c8e-508">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="33c8e-508">Compose or Read</span></span>|

<br>

---
---

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-12"></a><span data-ttu-id="33c8e-509">Отправитель: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="33c8e-509">sender: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)</span></span>

<span data-ttu-id="33c8e-p126">Получает электронный адрес отправителя электронного сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="33c8e-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="33c8e-p127">Свойства [`from`](#from-emailaddressdetails) и `sender` представляют одно лицо, если сообщение не отправлено представителем. В противном случае свойство `from` представляет лицо, делегировавшее полномочия, а свойство sender — представителя.</span><span class="sxs-lookup"><span data-stu-id="33c8e-p127">The [`from`](#from-emailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="33c8e-514">Свойству `recipientType`, принадлежащему объекту `EmailAddressDetails` в свойстве `sender`, задано значение `undefined`.</span><span class="sxs-lookup"><span data-stu-id="33c8e-514">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="33c8e-515">Тип</span><span class="sxs-lookup"><span data-stu-id="33c8e-515">Type</span></span>

*   [<span data-ttu-id="33c8e-516">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="33c8e-516">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)

##### <a name="requirements"></a><span data-ttu-id="33c8e-517">Требования</span><span class="sxs-lookup"><span data-stu-id="33c8e-517">Requirements</span></span>

|<span data-ttu-id="33c8e-518">Требование</span><span class="sxs-lookup"><span data-stu-id="33c8e-518">Requirement</span></span>| <span data-ttu-id="33c8e-519">Значение</span><span class="sxs-lookup"><span data-stu-id="33c8e-519">Value</span></span>|
|---|---|
|[<span data-ttu-id="33c8e-520">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="33c8e-520">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="33c8e-521">1.0</span><span class="sxs-lookup"><span data-stu-id="33c8e-521">1.0</span></span>|
|[<span data-ttu-id="33c8e-522">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="33c8e-522">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="33c8e-523">ReadItem</span><span class="sxs-lookup"><span data-stu-id="33c8e-523">ReadItem</span></span>|
|[<span data-ttu-id="33c8e-524">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="33c8e-524">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="33c8e-525">Чтение</span><span class="sxs-lookup"><span data-stu-id="33c8e-525">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="33c8e-526">Пример</span><span class="sxs-lookup"><span data-stu-id="33c8e-526">Example</span></span>

```js
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

<br>

---
---

#### <a name="start-datetimejavascriptapioutlookofficetimeviewoutlook-js-12"></a><span data-ttu-id="33c8e-527">Начало: Дата | [Time (время](/javascript/api/outlook/office.time?view=outlook-js-1.2) )</span><span class="sxs-lookup"><span data-stu-id="33c8e-527">start: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.2)</span></span>

<span data-ttu-id="33c8e-528">Получает или задает дату и время начала встречи.</span><span class="sxs-lookup"><span data-stu-id="33c8e-528">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="33c8e-p128">Свойство `start` представлено в виде значения даты и времени в формате UTC. Это значение можно преобразовать в местные значения даты и времени клиента с помощью метода [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime).</span><span class="sxs-lookup"><span data-stu-id="33c8e-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="33c8e-531">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="33c8e-531">Read mode</span></span>

<span data-ttu-id="33c8e-532">Свойство `start` возвращает объект `Date`.</span><span class="sxs-lookup"><span data-stu-id="33c8e-532">The `start` property returns a `Date` object.</span></span>

```js
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="33c8e-533">Режим создания</span><span class="sxs-lookup"><span data-stu-id="33c8e-533">Compose mode</span></span>

<span data-ttu-id="33c8e-534">Свойство `start` возвращает объект `Time`.</span><span class="sxs-lookup"><span data-stu-id="33c8e-534">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="33c8e-535">Если вы задаете время начала с помощью метода [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.2#setasync-datetime--options--callback-), необходимо использовать метод [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) для преобразования местного времени на клиенте в формат UTC для сервера.</span><span class="sxs-lookup"><span data-stu-id="33c8e-535">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.2#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>
<span data-ttu-id="33c8e-536">В примере ниже с помощью метода [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.2#setasync-datetime--options--callback-) объекта `Time` задается время начала встречи в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="33c8e-536">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.2#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="33c8e-537">Тип</span><span class="sxs-lookup"><span data-stu-id="33c8e-537">Type</span></span>

*   <span data-ttu-id="33c8e-538">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="33c8e-538">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.2)</span></span>

##### <a name="requirements"></a><span data-ttu-id="33c8e-539">Требования</span><span class="sxs-lookup"><span data-stu-id="33c8e-539">Requirements</span></span>

|<span data-ttu-id="33c8e-540">Требование</span><span class="sxs-lookup"><span data-stu-id="33c8e-540">Requirement</span></span>| <span data-ttu-id="33c8e-541">Значение</span><span class="sxs-lookup"><span data-stu-id="33c8e-541">Value</span></span>|
|---|---|
|[<span data-ttu-id="33c8e-542">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="33c8e-542">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="33c8e-543">1.0</span><span class="sxs-lookup"><span data-stu-id="33c8e-543">1.0</span></span>|
|[<span data-ttu-id="33c8e-544">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="33c8e-544">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="33c8e-545">ReadItem</span><span class="sxs-lookup"><span data-stu-id="33c8e-545">ReadItem</span></span>|
|[<span data-ttu-id="33c8e-546">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="33c8e-546">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="33c8e-547">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="33c8e-547">Compose or Read</span></span>|

<br>

---
---

#### <a name="subject-stringsubjectjavascriptapioutlookofficesubjectviewoutlook-js-12"></a><span data-ttu-id="33c8e-548">Тема: строка | [Subject (тема](/javascript/api/outlook/office.subject?view=outlook-js-1.2) )</span><span class="sxs-lookup"><span data-stu-id="33c8e-548">subject: String|[Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.2)</span></span>

<span data-ttu-id="33c8e-549">Получает или задает описание, которое отображается в поле темы элемента.</span><span class="sxs-lookup"><span data-stu-id="33c8e-549">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="33c8e-550">Свойство `subject` получает или задает всю тему элемента для отправки с почтового сервера.</span><span class="sxs-lookup"><span data-stu-id="33c8e-550">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="33c8e-551">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="33c8e-551">Read mode</span></span>

<span data-ttu-id="33c8e-p130">Свойство `subject` возвращает строку. С помощью свойства [`normalizedSubject`](#normalizedsubject-string) можно получить тему без начальных префиксов, таких как `RE:` и `FW:`.</span><span class="sxs-lookup"><span data-stu-id="33c8e-p130">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```js
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="33c8e-554">Режим создания</span><span class="sxs-lookup"><span data-stu-id="33c8e-554">Compose mode</span></span>

<span data-ttu-id="33c8e-555">Свойство `subject` возвращает объект `Subject`, который предоставляет методы для получения и задания темы.</span><span class="sxs-lookup"><span data-stu-id="33c8e-555">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```js
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="33c8e-556">Тип</span><span class="sxs-lookup"><span data-stu-id="33c8e-556">Type</span></span>

*   <span data-ttu-id="33c8e-557">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="33c8e-557">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.2)</span></span>

##### <a name="requirements"></a><span data-ttu-id="33c8e-558">Требования</span><span class="sxs-lookup"><span data-stu-id="33c8e-558">Requirements</span></span>

|<span data-ttu-id="33c8e-559">Требование</span><span class="sxs-lookup"><span data-stu-id="33c8e-559">Requirement</span></span>| <span data-ttu-id="33c8e-560">Значение</span><span class="sxs-lookup"><span data-stu-id="33c8e-560">Value</span></span>|
|---|---|
|[<span data-ttu-id="33c8e-561">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="33c8e-561">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="33c8e-562">1.0</span><span class="sxs-lookup"><span data-stu-id="33c8e-562">1.0</span></span>|
|[<span data-ttu-id="33c8e-563">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="33c8e-563">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="33c8e-564">ReadItem</span><span class="sxs-lookup"><span data-stu-id="33c8e-564">ReadItem</span></span>|
|[<span data-ttu-id="33c8e-565">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="33c8e-565">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="33c8e-566">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="33c8e-566">Compose or Read</span></span>|

<br>

---
---

#### <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-12recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-12"></a><span data-ttu-id="33c8e-567">Кому: Array. <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)>|[получатели](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="33c8e-567">to: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

<span data-ttu-id="33c8e-568">Предоставляет доступ к получателям, указанным в строке **Кому** сообщения.</span><span class="sxs-lookup"><span data-stu-id="33c8e-568">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="33c8e-569">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="33c8e-569">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="33c8e-570">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="33c8e-570">Read mode</span></span>

<span data-ttu-id="33c8e-p132">Свойство `to` возвращает массив, содержащий объект `EmailAddressDetails` для каждого получателя в строке **Кому** сообщения. Коллекция может включать не более 100 элементов.</span><span class="sxs-lookup"><span data-stu-id="33c8e-p132">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="33c8e-573">Режим создания</span><span class="sxs-lookup"><span data-stu-id="33c8e-573">Compose mode</span></span>

<span data-ttu-id="33c8e-574">Свойство `to` возвращает объект `Recipients`, предоставляющий методы для получения или обновления получателей, которые указаны в строке **Кому** сообщения.</span><span class="sxs-lookup"><span data-stu-id="33c8e-574">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

```js
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="33c8e-575">Тип</span><span class="sxs-lookup"><span data-stu-id="33c8e-575">Type</span></span>

*   <span data-ttu-id="33c8e-576">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="33c8e-576">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

##### <a name="requirements"></a><span data-ttu-id="33c8e-577">Требования</span><span class="sxs-lookup"><span data-stu-id="33c8e-577">Requirements</span></span>

|<span data-ttu-id="33c8e-578">Требование</span><span class="sxs-lookup"><span data-stu-id="33c8e-578">Requirement</span></span>| <span data-ttu-id="33c8e-579">Значение</span><span class="sxs-lookup"><span data-stu-id="33c8e-579">Value</span></span>|
|---|---|
|[<span data-ttu-id="33c8e-580">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="33c8e-580">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="33c8e-581">1.0</span><span class="sxs-lookup"><span data-stu-id="33c8e-581">1.0</span></span>|
|[<span data-ttu-id="33c8e-582">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="33c8e-582">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="33c8e-583">ReadItem</span><span class="sxs-lookup"><span data-stu-id="33c8e-583">ReadItem</span></span>|
|[<span data-ttu-id="33c8e-584">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="33c8e-584">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="33c8e-585">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="33c8e-585">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="33c8e-586">Методы</span><span class="sxs-lookup"><span data-stu-id="33c8e-586">Methods</span></span>

#### <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="33c8e-587">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="33c8e-587">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="33c8e-588">Добавляет файл в сообщение или встречу в качестве вложения.</span><span class="sxs-lookup"><span data-stu-id="33c8e-588">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="33c8e-589">Метод `addFileAttachmentAsync` передает файл по указанному универсальному коду ресурса (URI) и вкладывает его в элемент в форме создания.</span><span class="sxs-lookup"><span data-stu-id="33c8e-589">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="33c8e-590">Идентификатор можно использовать с методом [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback), чтобы удалить вложение, добавленное во время текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="33c8e-590">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="33c8e-591">Параметры</span><span class="sxs-lookup"><span data-stu-id="33c8e-591">Parameters</span></span>

|<span data-ttu-id="33c8e-592">Имя</span><span class="sxs-lookup"><span data-stu-id="33c8e-592">Name</span></span>| <span data-ttu-id="33c8e-593">Тип</span><span class="sxs-lookup"><span data-stu-id="33c8e-593">Type</span></span>| <span data-ttu-id="33c8e-594">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="33c8e-594">Attributes</span></span>| <span data-ttu-id="33c8e-595">Описание</span><span class="sxs-lookup"><span data-stu-id="33c8e-595">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="33c8e-596">String</span><span class="sxs-lookup"><span data-stu-id="33c8e-596">String</span></span>||<span data-ttu-id="33c8e-p133">Универсальный код ресурса (URI), представляющий расположение файла, который нужно вложить в сообщение или встречу. Максимальная длина — 2048 символов.</span><span class="sxs-lookup"><span data-stu-id="33c8e-p133">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="33c8e-599">String.</span><span class="sxs-lookup"><span data-stu-id="33c8e-599">String</span></span>||<span data-ttu-id="33c8e-p134">Имя вложения, которое отображается при передаче вложения. Максимальная длина — 255 символов.</span><span class="sxs-lookup"><span data-stu-id="33c8e-p134">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="33c8e-602">Объект</span><span class="sxs-lookup"><span data-stu-id="33c8e-602">Object</span></span>| <span data-ttu-id="33c8e-603">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="33c8e-603">&lt;optional&gt;</span></span>|<span data-ttu-id="33c8e-604">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="33c8e-604">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="33c8e-605">Объект</span><span class="sxs-lookup"><span data-stu-id="33c8e-605">Object</span></span>| <span data-ttu-id="33c8e-606">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="33c8e-606">&lt;optional&gt;</span></span>|<span data-ttu-id="33c8e-607">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="33c8e-607">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="33c8e-608">функция</span><span class="sxs-lookup"><span data-stu-id="33c8e-608">function</span></span>| <span data-ttu-id="33c8e-609">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="33c8e-609">&lt;optional&gt;</span></span>|<span data-ttu-id="33c8e-610">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="33c8e-610">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="33c8e-611">После успешного выполнения идентификатор вложения будет представлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="33c8e-611">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="33c8e-612">Если добавить вложение не удастся, объект `asyncResult` будет содержать объект `Error` с описанием ошибки.</span><span class="sxs-lookup"><span data-stu-id="33c8e-612">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="33c8e-613">Ошибки</span><span class="sxs-lookup"><span data-stu-id="33c8e-613">Errors</span></span>

| <span data-ttu-id="33c8e-614">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="33c8e-614">Error code</span></span> | <span data-ttu-id="33c8e-615">Описание</span><span class="sxs-lookup"><span data-stu-id="33c8e-615">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="33c8e-616">Вложение превышает максимальный размер.</span><span class="sxs-lookup"><span data-stu-id="33c8e-616">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="33c8e-617">Расширение вложения не поддерживается.</span><span class="sxs-lookup"><span data-stu-id="33c8e-617">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="33c8e-618">Сообщение или встреча содержат слишком много вложений.</span><span class="sxs-lookup"><span data-stu-id="33c8e-618">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="33c8e-619">Требования</span><span class="sxs-lookup"><span data-stu-id="33c8e-619">Requirements</span></span>

|<span data-ttu-id="33c8e-620">Требование</span><span class="sxs-lookup"><span data-stu-id="33c8e-620">Requirement</span></span>| <span data-ttu-id="33c8e-621">Значение</span><span class="sxs-lookup"><span data-stu-id="33c8e-621">Value</span></span>|
|---|---|
|[<span data-ttu-id="33c8e-622">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="33c8e-622">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="33c8e-623">1.1</span><span class="sxs-lookup"><span data-stu-id="33c8e-623">1.1</span></span>|
|[<span data-ttu-id="33c8e-624">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="33c8e-624">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="33c8e-625">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="33c8e-625">ReadWriteItem</span></span>|
|[<span data-ttu-id="33c8e-626">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="33c8e-626">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="33c8e-627">Создание</span><span class="sxs-lookup"><span data-stu-id="33c8e-627">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="33c8e-628">Пример</span><span class="sxs-lookup"><span data-stu-id="33c8e-628">Example</span></span>

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

#### <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="33c8e-629">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="33c8e-629">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="33c8e-630">Добавляет к сообщению элемент Exchange, например сообщение, в виде вложения.</span><span class="sxs-lookup"><span data-stu-id="33c8e-630">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="33c8e-p135">С помощью метода `addItemAttachmentAsync` можно в элемент формы создания вложить элемент с указанным идентификатором Exchange. Если указать метод обратного вызова, то этот метод вызывается с помощью параметра `asyncResult`, который содержит идентификатор вложения или код, указывающий на ошибки, которые произошли при вложении элемента. При необходимости можно использовать параметр `options` для передачи сведений о состоянии методу обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="33c8e-p135">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="33c8e-634">Идентификатор можно использовать с методом [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback), чтобы удалить вложение, добавленное во время текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="33c8e-634">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="33c8e-635">Если ваша надстройка Office работает в Outlook в Интернете, `addItemAttachmentAsync` метод может присоединять элементы к элементам, отличным от редактируемого элемента; Однако это не поддерживается и не рекомендуется.</span><span class="sxs-lookup"><span data-stu-id="33c8e-635">If your Office Add-in is running in Outlook on the web, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="33c8e-636">Параметры</span><span class="sxs-lookup"><span data-stu-id="33c8e-636">Parameters</span></span>

|<span data-ttu-id="33c8e-637">Имя</span><span class="sxs-lookup"><span data-stu-id="33c8e-637">Name</span></span>| <span data-ttu-id="33c8e-638">Тип</span><span class="sxs-lookup"><span data-stu-id="33c8e-638">Type</span></span>| <span data-ttu-id="33c8e-639">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="33c8e-639">Attributes</span></span>| <span data-ttu-id="33c8e-640">Описание</span><span class="sxs-lookup"><span data-stu-id="33c8e-640">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="33c8e-641">String</span><span class="sxs-lookup"><span data-stu-id="33c8e-641">String</span></span>||<span data-ttu-id="33c8e-p136">Идентификатор Exchange для вкладываемого элемента. Максимальная длина — 100 символов.</span><span class="sxs-lookup"><span data-stu-id="33c8e-p136">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="33c8e-644">String</span><span class="sxs-lookup"><span data-stu-id="33c8e-644">String</span></span>||<span data-ttu-id="33c8e-645">Тема вкладываемого элемента.</span><span class="sxs-lookup"><span data-stu-id="33c8e-645">The subject of the item to be attached.</span></span> <span data-ttu-id="33c8e-646">Максимальная длина: 255 символов.</span><span class="sxs-lookup"><span data-stu-id="33c8e-646">The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="33c8e-647">Object</span><span class="sxs-lookup"><span data-stu-id="33c8e-647">Object</span></span>| <span data-ttu-id="33c8e-648">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="33c8e-648">&lt;optional&gt;</span></span>|<span data-ttu-id="33c8e-649">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="33c8e-649">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="33c8e-650">Объект</span><span class="sxs-lookup"><span data-stu-id="33c8e-650">Object</span></span>| <span data-ttu-id="33c8e-651">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="33c8e-651">&lt;optional&gt;</span></span>|<span data-ttu-id="33c8e-652">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="33c8e-652">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="33c8e-653">функция</span><span class="sxs-lookup"><span data-stu-id="33c8e-653">function</span></span>| <span data-ttu-id="33c8e-654">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="33c8e-654">&lt;optional&gt;</span></span>|<span data-ttu-id="33c8e-655">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="33c8e-655">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="33c8e-656">После успешного выполнения идентификатор вложения будет представлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="33c8e-656">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="33c8e-657">Если добавить вложение не удастся, объект `asyncResult` будет содержать объект `Error` с описанием ошибки.</span><span class="sxs-lookup"><span data-stu-id="33c8e-657">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="33c8e-658">Ошибки</span><span class="sxs-lookup"><span data-stu-id="33c8e-658">Errors</span></span>

| <span data-ttu-id="33c8e-659">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="33c8e-659">Error code</span></span> | <span data-ttu-id="33c8e-660">Описание</span><span class="sxs-lookup"><span data-stu-id="33c8e-660">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="33c8e-661">Сообщение или встреча содержат слишком много вложений.</span><span class="sxs-lookup"><span data-stu-id="33c8e-661">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="33c8e-662">Требования</span><span class="sxs-lookup"><span data-stu-id="33c8e-662">Requirements</span></span>

|<span data-ttu-id="33c8e-663">Требование</span><span class="sxs-lookup"><span data-stu-id="33c8e-663">Requirement</span></span>| <span data-ttu-id="33c8e-664">Значение</span><span class="sxs-lookup"><span data-stu-id="33c8e-664">Value</span></span>|
|---|---|
|[<span data-ttu-id="33c8e-665">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="33c8e-665">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="33c8e-666">1.1</span><span class="sxs-lookup"><span data-stu-id="33c8e-666">1.1</span></span>|
|[<span data-ttu-id="33c8e-667">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="33c8e-667">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="33c8e-668">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="33c8e-668">ReadWriteItem</span></span>|
|[<span data-ttu-id="33c8e-669">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="33c8e-669">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="33c8e-670">Создание</span><span class="sxs-lookup"><span data-stu-id="33c8e-670">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="33c8e-671">Пример</span><span class="sxs-lookup"><span data-stu-id="33c8e-671">Example</span></span>

<span data-ttu-id="33c8e-672">В следующем примере существующий элемент Outlook добавляется в виде вложения с именем `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="33c8e-672">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="33c8e-673">displayReplyAllForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="33c8e-673">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="33c8e-674">Отображает форму ответа, включающую отправителя и всех получателей выбранного сообщения или организатора и всех участников выбранной встречи.</span><span class="sxs-lookup"><span data-stu-id="33c8e-674">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="33c8e-675">Этот метод не поддерживается в Outlook на iOS или Android.</span><span class="sxs-lookup"><span data-stu-id="33c8e-675">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="33c8e-676">В Outlook в Интернете форма ответа отображается в виде всплывающей формы в представлении из трех столбцов и всплывающей формы в представлении с 2 или 1 столбца.</span><span class="sxs-lookup"><span data-stu-id="33c8e-676">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="33c8e-677">Если любой строковый параметр превышает указанные для него ограничения, `displayReplyAllForm` возвращает исключение.</span><span class="sxs-lookup"><span data-stu-id="33c8e-677">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="33c8e-678">Если в `formData.attachments` параметре указаны вложения, Outlook в Интернете и клиенте для настольных компьютеров пытаются скачать все вложения и присоединить их к форме ответа.</span><span class="sxs-lookup"><span data-stu-id="33c8e-678">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form.</span></span> <span data-ttu-id="33c8e-679">Если какие-либо вложения добавить не удается, в форме отображается сообщение об ошибке.</span><span class="sxs-lookup"><span data-stu-id="33c8e-679">If any attachments fail to be added, an error is shown in the form UI.</span></span> <span data-ttu-id="33c8e-680">Если сообщения об ошибках не предусмотрены, то они не отображаются.</span><span class="sxs-lookup"><span data-stu-id="33c8e-680">If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="33c8e-681">Параметры</span><span class="sxs-lookup"><span data-stu-id="33c8e-681">Parameters</span></span>

|<span data-ttu-id="33c8e-682">Имя</span><span class="sxs-lookup"><span data-stu-id="33c8e-682">Name</span></span>| <span data-ttu-id="33c8e-683">Тип</span><span class="sxs-lookup"><span data-stu-id="33c8e-683">Type</span></span>| <span data-ttu-id="33c8e-684">Описание</span><span class="sxs-lookup"><span data-stu-id="33c8e-684">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="33c8e-685">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="33c8e-685">String &#124; Object</span></span>| |<span data-ttu-id="33c8e-p139">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="33c8e-p139">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="33c8e-688">**ИЛИ**</span><span class="sxs-lookup"><span data-stu-id="33c8e-688">**OR**</span></span><br/><span data-ttu-id="33c8e-p140">Объект, который содержит текст или данные вложения и функцию обратного вызова. Ниже представлено определение этого объекта.</span><span class="sxs-lookup"><span data-stu-id="33c8e-p140">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="33c8e-691">String.</span><span class="sxs-lookup"><span data-stu-id="33c8e-691">String</span></span> | <span data-ttu-id="33c8e-692">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="33c8e-692">&lt;optional&gt;</span></span> | <span data-ttu-id="33c8e-p141">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="33c8e-p141">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="33c8e-695">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="33c8e-695">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="33c8e-696">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="33c8e-696">&lt;optional&gt;</span></span> | <span data-ttu-id="33c8e-697">Массив объектов JSON, представляющих собой вложенные файлы или элементы.</span><span class="sxs-lookup"><span data-stu-id="33c8e-697">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="33c8e-698">String.</span><span class="sxs-lookup"><span data-stu-id="33c8e-698">String</span></span> | | <span data-ttu-id="33c8e-p142">Указывает тип вложения. Допустимые значения: `file` для вложенного файла и `item` для вложенного элемента.</span><span class="sxs-lookup"><span data-stu-id="33c8e-p142">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="33c8e-701">Строка</span><span class="sxs-lookup"><span data-stu-id="33c8e-701">String</span></span> | | <span data-ttu-id="33c8e-702">Строка, содержащая имя вложения, длиной до 255 символов.</span><span class="sxs-lookup"><span data-stu-id="33c8e-702">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="33c8e-703">String</span><span class="sxs-lookup"><span data-stu-id="33c8e-703">String</span></span> | | <span data-ttu-id="33c8e-p143">Используется, только если свойству `type` задано значение `file`. URI расположения файла.</span><span class="sxs-lookup"><span data-stu-id="33c8e-p143">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="33c8e-706">String</span><span class="sxs-lookup"><span data-stu-id="33c8e-706">String</span></span> | | <span data-ttu-id="33c8e-p144">Используется, только если свойству `type` задано значение `item`. Идентификатор вложения EWS. Это строка длиной до 100 символов.</span><span class="sxs-lookup"><span data-stu-id="33c8e-p144">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="33c8e-710">function</span><span class="sxs-lookup"><span data-stu-id="33c8e-710">function</span></span> | <span data-ttu-id="33c8e-711">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="33c8e-711">&lt;optional&gt;</span></span> | <span data-ttu-id="33c8e-712">По завершении работы метода функция, переданная параметру `callback`, вызывается с помощью одного параметра `asyncResult`, представляющего собой объект [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="33c8e-712">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="33c8e-713">Требования</span><span class="sxs-lookup"><span data-stu-id="33c8e-713">Requirements</span></span>

|<span data-ttu-id="33c8e-714">Требование</span><span class="sxs-lookup"><span data-stu-id="33c8e-714">Requirement</span></span>| <span data-ttu-id="33c8e-715">Значение</span><span class="sxs-lookup"><span data-stu-id="33c8e-715">Value</span></span>|
|---|---|
|[<span data-ttu-id="33c8e-716">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="33c8e-716">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="33c8e-717">1.0</span><span class="sxs-lookup"><span data-stu-id="33c8e-717">1.0</span></span>|
|[<span data-ttu-id="33c8e-718">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="33c8e-718">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="33c8e-719">ReadItem</span><span class="sxs-lookup"><span data-stu-id="33c8e-719">ReadItem</span></span>|
|[<span data-ttu-id="33c8e-720">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="33c8e-720">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="33c8e-721">Чтение</span><span class="sxs-lookup"><span data-stu-id="33c8e-721">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="33c8e-722">Примеры</span><span class="sxs-lookup"><span data-stu-id="33c8e-722">Examples</span></span>

<span data-ttu-id="33c8e-723">Приведенный ниже код передает строку в функцию `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="33c8e-723">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="33c8e-724">Ответ с пустым текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="33c8e-724">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="33c8e-725">Ответ только с текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="33c8e-725">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="33c8e-726">Ответ с текстом сообщения и вложенным файлом.</span><span class="sxs-lookup"><span data-stu-id="33c8e-726">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="33c8e-727">Ответ с текстом сообщения и вложенным элементом.</span><span class="sxs-lookup"><span data-stu-id="33c8e-727">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="33c8e-728">Ответ с текстом сообщения, вложенным файлом, вложенным элементом и обратным вызовом.</span><span class="sxs-lookup"><span data-stu-id="33c8e-728">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="33c8e-729">displayReplyForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="33c8e-729">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="33c8e-730">Отображает форму ответа, включающую только отправителя выбранного сообщения или организатора выбранной встречи.</span><span class="sxs-lookup"><span data-stu-id="33c8e-730">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="33c8e-731">Этот метод не поддерживается в Outlook на iOS или Android.</span><span class="sxs-lookup"><span data-stu-id="33c8e-731">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="33c8e-732">В Outlook в Интернете форма ответа отображается в виде всплывающей формы в представлении из трех столбцов и всплывающей формы в представлении с 2 или 1 столбца.</span><span class="sxs-lookup"><span data-stu-id="33c8e-732">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="33c8e-733">Если любой строковый параметр превышает указанные для него ограничения, `displayReplyForm` возвращает исключение.</span><span class="sxs-lookup"><span data-stu-id="33c8e-733">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="33c8e-734">Если в `formData.attachments` параметре указаны вложения, Outlook в Интернете и клиенте для настольных компьютеров пытаются скачать все вложения и присоединить их к форме ответа.</span><span class="sxs-lookup"><span data-stu-id="33c8e-734">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form.</span></span> <span data-ttu-id="33c8e-735">Если какие-либо вложения добавить не удается, в форме отображается сообщение об ошибке.</span><span class="sxs-lookup"><span data-stu-id="33c8e-735">If any attachments fail to be added, an error is shown in the form UI.</span></span> <span data-ttu-id="33c8e-736">Если сообщения об ошибках не предусмотрены, то они не отображаются.</span><span class="sxs-lookup"><span data-stu-id="33c8e-736">If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="33c8e-737">Параметры</span><span class="sxs-lookup"><span data-stu-id="33c8e-737">Parameters</span></span>

|<span data-ttu-id="33c8e-738">Имя</span><span class="sxs-lookup"><span data-stu-id="33c8e-738">Name</span></span>| <span data-ttu-id="33c8e-739">Тип</span><span class="sxs-lookup"><span data-stu-id="33c8e-739">Type</span></span>| <span data-ttu-id="33c8e-740">Описание</span><span class="sxs-lookup"><span data-stu-id="33c8e-740">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="33c8e-741">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="33c8e-741">String &#124; Object</span></span>| | <span data-ttu-id="33c8e-p146">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="33c8e-p146">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="33c8e-744">**ИЛИ**</span><span class="sxs-lookup"><span data-stu-id="33c8e-744">**OR**</span></span><br/><span data-ttu-id="33c8e-p147">Объект, который содержит текст или данные вложения и функцию обратного вызова. Ниже представлено определение этого объекта.</span><span class="sxs-lookup"><span data-stu-id="33c8e-p147">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="33c8e-747">String.</span><span class="sxs-lookup"><span data-stu-id="33c8e-747">String</span></span> | <span data-ttu-id="33c8e-748">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="33c8e-748">&lt;optional&gt;</span></span> | <span data-ttu-id="33c8e-p148">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="33c8e-p148">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="33c8e-751">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="33c8e-751">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="33c8e-752">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="33c8e-752">&lt;optional&gt;</span></span> | <span data-ttu-id="33c8e-753">Массив объектов JSON, представляющих собой вложенные файлы или элементы.</span><span class="sxs-lookup"><span data-stu-id="33c8e-753">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="33c8e-754">String.</span><span class="sxs-lookup"><span data-stu-id="33c8e-754">String</span></span> | | <span data-ttu-id="33c8e-p149">Указывает тип вложения. Допустимые значения: `file` для вложенного файла и `item` для вложенного элемента.</span><span class="sxs-lookup"><span data-stu-id="33c8e-p149">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="33c8e-757">Строка</span><span class="sxs-lookup"><span data-stu-id="33c8e-757">String</span></span> | | <span data-ttu-id="33c8e-758">Строка, содержащая имя вложения, длиной до 255 символов.</span><span class="sxs-lookup"><span data-stu-id="33c8e-758">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="33c8e-759">String</span><span class="sxs-lookup"><span data-stu-id="33c8e-759">String</span></span> | | <span data-ttu-id="33c8e-p150">Используется, только если свойству `type` задано значение `file`. URI расположения файла.</span><span class="sxs-lookup"><span data-stu-id="33c8e-p150">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="33c8e-762">String</span><span class="sxs-lookup"><span data-stu-id="33c8e-762">String</span></span> | | <span data-ttu-id="33c8e-p151">Используется, только если свойству `type` задано значение `item`. Идентификатор вложения EWS. Это строка длиной до 100 символов.</span><span class="sxs-lookup"><span data-stu-id="33c8e-p151">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="33c8e-766">function</span><span class="sxs-lookup"><span data-stu-id="33c8e-766">function</span></span> | <span data-ttu-id="33c8e-767">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="33c8e-767">&lt;optional&gt;</span></span> | <span data-ttu-id="33c8e-768">По завершении работы метода функция, переданная параметру `callback`, вызывается с помощью одного параметра `asyncResult`, представляющего собой объект [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="33c8e-768">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="33c8e-769">Требования</span><span class="sxs-lookup"><span data-stu-id="33c8e-769">Requirements</span></span>

|<span data-ttu-id="33c8e-770">Требование</span><span class="sxs-lookup"><span data-stu-id="33c8e-770">Requirement</span></span>| <span data-ttu-id="33c8e-771">Значение</span><span class="sxs-lookup"><span data-stu-id="33c8e-771">Value</span></span>|
|---|---|
|[<span data-ttu-id="33c8e-772">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="33c8e-772">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="33c8e-773">1.0</span><span class="sxs-lookup"><span data-stu-id="33c8e-773">1.0</span></span>|
|[<span data-ttu-id="33c8e-774">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="33c8e-774">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="33c8e-775">ReadItem</span><span class="sxs-lookup"><span data-stu-id="33c8e-775">ReadItem</span></span>|
|[<span data-ttu-id="33c8e-776">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="33c8e-776">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="33c8e-777">Чтение</span><span class="sxs-lookup"><span data-stu-id="33c8e-777">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="33c8e-778">Примеры</span><span class="sxs-lookup"><span data-stu-id="33c8e-778">Examples</span></span>

<span data-ttu-id="33c8e-779">Приведенный ниже код передает строку в функцию `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="33c8e-779">The following code passes a string to the `displayReplyForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="33c8e-780">Ответ с пустым текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="33c8e-780">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="33c8e-781">Ответ только с текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="33c8e-781">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="33c8e-782">Ответ с текстом сообщения и вложенным файлом.</span><span class="sxs-lookup"><span data-stu-id="33c8e-782">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="33c8e-783">Ответ с текстом сообщения и вложенным элементом.</span><span class="sxs-lookup"><span data-stu-id="33c8e-783">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="33c8e-784">Ответ с текстом сообщения, вложенным файлом, вложенным элементом и обратным вызовом.</span><span class="sxs-lookup"><span data-stu-id="33c8e-784">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlookofficeentitiesviewoutlook-js-12"></a><span data-ttu-id="33c8e-785">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.2)}</span><span class="sxs-lookup"><span data-stu-id="33c8e-785">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.2)}</span></span>

<span data-ttu-id="33c8e-786">Получает сущности, обнаруженные в теле выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="33c8e-786">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="33c8e-787">Этот метод не поддерживается в Outlook на iOS или Android.</span><span class="sxs-lookup"><span data-stu-id="33c8e-787">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="33c8e-788">Требования</span><span class="sxs-lookup"><span data-stu-id="33c8e-788">Requirements</span></span>

|<span data-ttu-id="33c8e-789">Требование</span><span class="sxs-lookup"><span data-stu-id="33c8e-789">Requirement</span></span>| <span data-ttu-id="33c8e-790">Значение</span><span class="sxs-lookup"><span data-stu-id="33c8e-790">Value</span></span>|
|---|---|
|[<span data-ttu-id="33c8e-791">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="33c8e-791">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="33c8e-792">1.0</span><span class="sxs-lookup"><span data-stu-id="33c8e-792">1.0</span></span>|
|[<span data-ttu-id="33c8e-793">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="33c8e-793">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="33c8e-794">ReadItem</span><span class="sxs-lookup"><span data-stu-id="33c8e-794">ReadItem</span></span>|
|[<span data-ttu-id="33c8e-795">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="33c8e-795">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="33c8e-796">Чтение</span><span class="sxs-lookup"><span data-stu-id="33c8e-796">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="33c8e-797">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="33c8e-797">Returns:</span></span>

<span data-ttu-id="33c8e-798">Тип: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="33c8e-798">Type: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.2)</span></span>

##### <a name="example"></a><span data-ttu-id="33c8e-799">Пример</span><span class="sxs-lookup"><span data-stu-id="33c8e-799">Example</span></span>

<span data-ttu-id="33c8e-800">Ниже приведен пример получения доступа к сущностям контактов в теле текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="33c8e-800">The following example accesses the contacts entities in the current item's body.</span></span>

```js
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

<br>

---
---

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-12meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-12phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-12tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-12"></a><span data-ttu-id="33c8e-801">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.2)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.2)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.2)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.2))>}</span><span class="sxs-lookup"><span data-stu-id="33c8e-801">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.2)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.2)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.2)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.2))>}</span></span>

<span data-ttu-id="33c8e-802">Получает массив всех сущностей указанного типа, обнаруженных в теле выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="33c8e-802">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="33c8e-803">Этот метод не поддерживается в Outlook на iOS или Android.</span><span class="sxs-lookup"><span data-stu-id="33c8e-803">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="33c8e-804">Параметры</span><span class="sxs-lookup"><span data-stu-id="33c8e-804">Parameters</span></span>

|<span data-ttu-id="33c8e-805">Имя</span><span class="sxs-lookup"><span data-stu-id="33c8e-805">Name</span></span>| <span data-ttu-id="33c8e-806">Тип</span><span class="sxs-lookup"><span data-stu-id="33c8e-806">Type</span></span>| <span data-ttu-id="33c8e-807">Описание</span><span class="sxs-lookup"><span data-stu-id="33c8e-807">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="33c8e-808">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="33c8e-808">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook/office.mailboxenums.entitytype?view=outlook-js-1.2)|<span data-ttu-id="33c8e-809">Одно из значений перечисления EntityType.</span><span class="sxs-lookup"><span data-stu-id="33c8e-809">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="33c8e-810">Требования</span><span class="sxs-lookup"><span data-stu-id="33c8e-810">Requirements</span></span>

|<span data-ttu-id="33c8e-811">Требование</span><span class="sxs-lookup"><span data-stu-id="33c8e-811">Requirement</span></span>| <span data-ttu-id="33c8e-812">Значение</span><span class="sxs-lookup"><span data-stu-id="33c8e-812">Value</span></span>|
|---|---|
|[<span data-ttu-id="33c8e-813">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="33c8e-813">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="33c8e-814">1.0</span><span class="sxs-lookup"><span data-stu-id="33c8e-814">1.0</span></span>|
|[<span data-ttu-id="33c8e-815">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="33c8e-815">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="33c8e-816">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="33c8e-816">Restricted</span></span>|
|[<span data-ttu-id="33c8e-817">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="33c8e-817">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="33c8e-818">Чтение</span><span class="sxs-lookup"><span data-stu-id="33c8e-818">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="33c8e-819">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="33c8e-819">Returns:</span></span>

<span data-ttu-id="33c8e-820">Если значение, переданное в `entityType`, не является допустимым членом перечисления `EntityType`, метод возвращает значение NULL.</span><span class="sxs-lookup"><span data-stu-id="33c8e-820">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="33c8e-821">Если в теле элемента отсутствуют сущности указанного типа, метод возвращает пустой массив.</span><span class="sxs-lookup"><span data-stu-id="33c8e-821">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="33c8e-822">В противном случае тип объектов в возвращаемом массиве зависит от типа сущности, запрошенной в параметре `entityType`.</span><span class="sxs-lookup"><span data-stu-id="33c8e-822">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="33c8e-823">Хотя минимальный уровень разрешений для использования этого метода — **Restricted**, для некоторых типов сущностей требуется доступ на уровне **ReadItem**, как указано в приведенной ниже таблице.</span><span class="sxs-lookup"><span data-stu-id="33c8e-823">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="33c8e-824">Значение параметра `entityType`</span><span class="sxs-lookup"><span data-stu-id="33c8e-824">Value of `entityType`</span></span> | <span data-ttu-id="33c8e-825">Тип объектов в возвращаемом массиве</span><span class="sxs-lookup"><span data-stu-id="33c8e-825">Type of objects in returned array</span></span> | <span data-ttu-id="33c8e-826">Необходимый уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="33c8e-826">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="33c8e-827">String</span><span class="sxs-lookup"><span data-stu-id="33c8e-827">String</span></span> | <span data-ttu-id="33c8e-828">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="33c8e-828">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="33c8e-829">Contact</span><span class="sxs-lookup"><span data-stu-id="33c8e-829">Contact</span></span> | <span data-ttu-id="33c8e-830">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="33c8e-830">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="33c8e-831">String</span><span class="sxs-lookup"><span data-stu-id="33c8e-831">String</span></span> | <span data-ttu-id="33c8e-832">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="33c8e-832">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="33c8e-833">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="33c8e-833">MeetingSuggestion</span></span> | <span data-ttu-id="33c8e-834">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="33c8e-834">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="33c8e-835">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="33c8e-835">PhoneNumber</span></span> | <span data-ttu-id="33c8e-836">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="33c8e-836">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="33c8e-837">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="33c8e-837">TaskSuggestion</span></span> | <span data-ttu-id="33c8e-838">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="33c8e-838">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="33c8e-839">String</span><span class="sxs-lookup"><span data-stu-id="33c8e-839">String</span></span> | <span data-ttu-id="33c8e-840">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="33c8e-840">**Restricted**</span></span> |

<span data-ttu-id="33c8e-841">Тип: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.2)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.2)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.2)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.2))></span><span class="sxs-lookup"><span data-stu-id="33c8e-841">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.2)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.2)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.2)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.2))></span></span>

##### <a name="example"></a><span data-ttu-id="33c8e-842">Пример</span><span class="sxs-lookup"><span data-stu-id="33c8e-842">Example</span></span>

<span data-ttu-id="33c8e-843">В примере ниже показано, как получить доступ к массиву строк, которые представляют собой почтовые адреса в теле текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="33c8e-843">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-12meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-12phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-12tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-12"></a><span data-ttu-id="33c8e-844">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.2)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.2)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.2)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.2))>}</span><span class="sxs-lookup"><span data-stu-id="33c8e-844">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.2)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.2)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.2)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.2))>}</span></span>

<span data-ttu-id="33c8e-845">Возвращает известные сущности в выбранном элементе, которые проходят через именованный фильтр, определяемый в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="33c8e-845">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="33c8e-846">Этот метод не поддерживается в Outlook на iOS или Android.</span><span class="sxs-lookup"><span data-stu-id="33c8e-846">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="33c8e-847">Метод `getFilteredEntitiesByName` возвращает сущности, соответствующие регулярному выражению, которое определяется в элементе правила [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) в XML-файле манифеста, с использованием указанного значения элемента `FilterName`.</span><span class="sxs-lookup"><span data-stu-id="33c8e-847">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="33c8e-848">Параметры</span><span class="sxs-lookup"><span data-stu-id="33c8e-848">Parameters</span></span>

|<span data-ttu-id="33c8e-849">Имя</span><span class="sxs-lookup"><span data-stu-id="33c8e-849">Name</span></span>| <span data-ttu-id="33c8e-850">Тип</span><span class="sxs-lookup"><span data-stu-id="33c8e-850">Type</span></span>| <span data-ttu-id="33c8e-851">Описание</span><span class="sxs-lookup"><span data-stu-id="33c8e-851">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="33c8e-852">String</span><span class="sxs-lookup"><span data-stu-id="33c8e-852">String</span></span>|<span data-ttu-id="33c8e-853">Имя элемента правила `ItemHasKnownEntity`, определяющее соответствующий фильтр.</span><span class="sxs-lookup"><span data-stu-id="33c8e-853">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="33c8e-854">Требования</span><span class="sxs-lookup"><span data-stu-id="33c8e-854">Requirements</span></span>

|<span data-ttu-id="33c8e-855">Требование</span><span class="sxs-lookup"><span data-stu-id="33c8e-855">Requirement</span></span>| <span data-ttu-id="33c8e-856">Значение</span><span class="sxs-lookup"><span data-stu-id="33c8e-856">Value</span></span>|
|---|---|
|[<span data-ttu-id="33c8e-857">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="33c8e-857">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="33c8e-858">1.0</span><span class="sxs-lookup"><span data-stu-id="33c8e-858">1.0</span></span>|
|[<span data-ttu-id="33c8e-859">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="33c8e-859">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="33c8e-860">ReadItem</span><span class="sxs-lookup"><span data-stu-id="33c8e-860">ReadItem</span></span>|
|[<span data-ttu-id="33c8e-861">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="33c8e-861">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="33c8e-862">Чтение</span><span class="sxs-lookup"><span data-stu-id="33c8e-862">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="33c8e-863">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="33c8e-863">Returns:</span></span>

<span data-ttu-id="33c8e-p153">Если в манифесте нет элемента `ItemHasKnownEntity` со значением `FilterName`, соответствующим параметру `name`, метод возвращает `null`. Если параметр `name` соответствует элементу `ItemHasKnownEntity` в манифесте, но при этом в текущем элементе нет соответствующих сущностей, метод возвращает пустой массив.</span><span class="sxs-lookup"><span data-stu-id="33c8e-p153">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="33c8e-866">Тип: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.2)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.2)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.2)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.2))></span><span class="sxs-lookup"><span data-stu-id="33c8e-866">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.2)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.2)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.2)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.2))></span></span>

<br>

---
---

#### <a name="getregexmatches--object"></a><span data-ttu-id="33c8e-867">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="33c8e-867">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="33c8e-868">Возвращает строковые значения в выбранном элементе, которые соответствуют регулярным выражениям, определенным в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="33c8e-868">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="33c8e-869">Этот метод не поддерживается в Outlook на iOS или Android.</span><span class="sxs-lookup"><span data-stu-id="33c8e-869">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="33c8e-p154">Метод `getRegExMatches` возвращает строки, соответствующие регулярному выражению, которое определяется в каждом элементе правила `ItemHasRegularExpressionMatch` или `ItemHasKnownEntity` в XML-файле манифеста. Для правила `ItemHasRegularExpressionMatch` соответствующую строку должно содержать свойство элемента, указанного этим правилом. Простой тип `PropertyName` определяет поддерживаемые свойства.</span><span class="sxs-lookup"><span data-stu-id="33c8e-p154">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="33c8e-873">Например, рассмотрим манифест надстройки, который содержит указанный ниже элемент `Rule`.</span><span class="sxs-lookup"><span data-stu-id="33c8e-873">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="33c8e-874">Объект, возвращаемый методом `getRegExMatches`, будет содержать два свойства: `fruits` и `veggies`.</span><span class="sxs-lookup"><span data-stu-id="33c8e-874">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

> [!NOTE]
> <span data-ttu-id="33c8e-p155">Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты.</span><span class="sxs-lookup"><span data-stu-id="33c8e-p155">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="requirements"></a><span data-ttu-id="33c8e-877">Требования</span><span class="sxs-lookup"><span data-stu-id="33c8e-877">Requirements</span></span>

|<span data-ttu-id="33c8e-878">Требование</span><span class="sxs-lookup"><span data-stu-id="33c8e-878">Requirement</span></span>| <span data-ttu-id="33c8e-879">Значение</span><span class="sxs-lookup"><span data-stu-id="33c8e-879">Value</span></span>|
|---|---|
|[<span data-ttu-id="33c8e-880">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="33c8e-880">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="33c8e-881">1.0</span><span class="sxs-lookup"><span data-stu-id="33c8e-881">1.0</span></span>|
|[<span data-ttu-id="33c8e-882">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="33c8e-882">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="33c8e-883">ReadItem</span><span class="sxs-lookup"><span data-stu-id="33c8e-883">ReadItem</span></span>|
|[<span data-ttu-id="33c8e-884">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="33c8e-884">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="33c8e-885">Чтение</span><span class="sxs-lookup"><span data-stu-id="33c8e-885">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="33c8e-886">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="33c8e-886">Returns:</span></span>

<span data-ttu-id="33c8e-p156">Объект, содержащий массив строк, которые соответствуют регулярным выражениям, определяемым в XML-файле манифеста. Имя каждого массива равно соответствующему значению атрибута `RegExName` подходящего правила `ItemHasRegularExpressionMatch` или атрибута `FilterName` соответствующего правила `ItemHasKnownEntity`.</span><span class="sxs-lookup"><span data-stu-id="33c8e-p156">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<span data-ttu-id="33c8e-889">Тип: Object</span><span class="sxs-lookup"><span data-stu-id="33c8e-889">Type: Object</span></span>

##### <a name="example"></a><span data-ttu-id="33c8e-890">Пример</span><span class="sxs-lookup"><span data-stu-id="33c8e-890">Example</span></span>

<span data-ttu-id="33c8e-891">В примере ниже показано, как получить доступ к массиву совпадений для <rule>элементов регулярного выражения `fruits` и `veggies`, которые указаны в манифесте</rule>.</span><span class="sxs-lookup"><span data-stu-id="33c8e-891">The following example shows how to access the array of matches for the regular expression <rule>elements `fruits` and `veggies`, which are specified in the manifest.</rule></span></span>

```js
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

<br>

---
---

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="33c8e-892">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="33c8e-892">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="33c8e-893">Возвращает строковые значения в выбранном элементе, которые соответствуют именованному регулярному выражению, определенному в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="33c8e-893">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="33c8e-894">Этот метод не поддерживается в Outlook на iOS или Android.</span><span class="sxs-lookup"><span data-stu-id="33c8e-894">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="33c8e-895">Метод `getRegExMatchesByName` возвращает строки, соответствующие регулярному выражению, которое определяется в элементе правила `ItemHasRegularExpressionMatch` в XML-файле манифеста, с использованием указанного значения элемента `RegExName`.</span><span class="sxs-lookup"><span data-stu-id="33c8e-895">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="33c8e-p157">Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты.</span><span class="sxs-lookup"><span data-stu-id="33c8e-p157">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="33c8e-898">Параметры</span><span class="sxs-lookup"><span data-stu-id="33c8e-898">Parameters</span></span>

|<span data-ttu-id="33c8e-899">Имя</span><span class="sxs-lookup"><span data-stu-id="33c8e-899">Name</span></span>| <span data-ttu-id="33c8e-900">Тип</span><span class="sxs-lookup"><span data-stu-id="33c8e-900">Type</span></span>| <span data-ttu-id="33c8e-901">Описание</span><span class="sxs-lookup"><span data-stu-id="33c8e-901">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="33c8e-902">String</span><span class="sxs-lookup"><span data-stu-id="33c8e-902">String</span></span>|<span data-ttu-id="33c8e-903">Имя элемента правила `ItemHasRegularExpressionMatch`, определяющее соответствующий фильтр.</span><span class="sxs-lookup"><span data-stu-id="33c8e-903">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="33c8e-904">Требования</span><span class="sxs-lookup"><span data-stu-id="33c8e-904">Requirements</span></span>

|<span data-ttu-id="33c8e-905">Требование</span><span class="sxs-lookup"><span data-stu-id="33c8e-905">Requirement</span></span>| <span data-ttu-id="33c8e-906">Значение</span><span class="sxs-lookup"><span data-stu-id="33c8e-906">Value</span></span>|
|---|---|
|[<span data-ttu-id="33c8e-907">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="33c8e-907">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="33c8e-908">1.0</span><span class="sxs-lookup"><span data-stu-id="33c8e-908">1.0</span></span>|
|[<span data-ttu-id="33c8e-909">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="33c8e-909">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="33c8e-910">ReadItem</span><span class="sxs-lookup"><span data-stu-id="33c8e-910">ReadItem</span></span>|
|[<span data-ttu-id="33c8e-911">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="33c8e-911">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="33c8e-912">Чтение</span><span class="sxs-lookup"><span data-stu-id="33c8e-912">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="33c8e-913">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="33c8e-913">Returns:</span></span>

<span data-ttu-id="33c8e-914">Массив строк, соответствующих регулярному выражению, определяемому в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="33c8e-914">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<span data-ttu-id="33c8e-915">Тип: Array. < String ></span><span class="sxs-lookup"><span data-stu-id="33c8e-915">Type: Array.< String ></span></span>

##### <a name="example"></a><span data-ttu-id="33c8e-916">Пример</span><span class="sxs-lookup"><span data-stu-id="33c8e-916">Example</span></span>

```js
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

<br>

---
---

#### <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="33c8e-917">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="33c8e-917">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="33c8e-918">Асинхронно возвращает данные, выбранные в теме или тексте сообщения.</span><span class="sxs-lookup"><span data-stu-id="33c8e-918">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="33c8e-p158">Если выделенный фрагмент отсутствует, но курсор находится в тексте или теме, метод возвращает значение NULL для выбранных данных. Если выбраны не текст и не тема, метод возвращает ошибку `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="33c8e-p158">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="33c8e-921">Параметры</span><span class="sxs-lookup"><span data-stu-id="33c8e-921">Parameters</span></span>

|<span data-ttu-id="33c8e-922">Имя</span><span class="sxs-lookup"><span data-stu-id="33c8e-922">Name</span></span>| <span data-ttu-id="33c8e-923">Тип</span><span class="sxs-lookup"><span data-stu-id="33c8e-923">Type</span></span>| <span data-ttu-id="33c8e-924">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="33c8e-924">Attributes</span></span>| <span data-ttu-id="33c8e-925">Описание</span><span class="sxs-lookup"><span data-stu-id="33c8e-925">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="33c8e-926">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="33c8e-926">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="33c8e-p159">Запрашивает формат данных. Если задано значение Text, метод возвращает обычный текст как строку, удаляя все имеющиеся HTML-теги. Если задано значение HTML, метод возвращает выделенный текст (обычный текст или HTML).</span><span class="sxs-lookup"><span data-stu-id="33c8e-p159">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="33c8e-930">Объект</span><span class="sxs-lookup"><span data-stu-id="33c8e-930">Object</span></span>| <span data-ttu-id="33c8e-931">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="33c8e-931">&lt;optional&gt;</span></span>|<span data-ttu-id="33c8e-932">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="33c8e-932">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="33c8e-933">Объект</span><span class="sxs-lookup"><span data-stu-id="33c8e-933">Object</span></span>| <span data-ttu-id="33c8e-934">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="33c8e-934">&lt;optional&gt;</span></span>|<span data-ttu-id="33c8e-935">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="33c8e-935">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="33c8e-936">функция</span><span class="sxs-lookup"><span data-stu-id="33c8e-936">function</span></span>||<span data-ttu-id="33c8e-937">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="33c8e-937">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="33c8e-938">Чтобы получить доступ к выбранным данным из метода обратного вызова, вызовите `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="33c8e-938">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="33c8e-939">Для доступа к исходному свойству, представляющему собой источник выбранных данных, вызовите параметр `asyncResult.value.sourceProperty`, который может иметь значение `body` или `subject`.</span><span class="sxs-lookup"><span data-stu-id="33c8e-939">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="33c8e-940">Требования</span><span class="sxs-lookup"><span data-stu-id="33c8e-940">Requirements</span></span>

|<span data-ttu-id="33c8e-941">Требование</span><span class="sxs-lookup"><span data-stu-id="33c8e-941">Requirement</span></span>| <span data-ttu-id="33c8e-942">Значение</span><span class="sxs-lookup"><span data-stu-id="33c8e-942">Value</span></span>|
|---|---|
|[<span data-ttu-id="33c8e-943">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="33c8e-943">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="33c8e-944">1.2</span><span class="sxs-lookup"><span data-stu-id="33c8e-944">1.2</span></span>|
|[<span data-ttu-id="33c8e-945">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="33c8e-945">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="33c8e-946">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="33c8e-946">ReadWriteItem</span></span>|
|[<span data-ttu-id="33c8e-947">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="33c8e-947">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="33c8e-948">Создание</span><span class="sxs-lookup"><span data-stu-id="33c8e-948">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="33c8e-949">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="33c8e-949">Returns:</span></span>

<span data-ttu-id="33c8e-950">Выбранные данные в виде строки с форматом, определенным в параметре `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="33c8e-950">The selected data as a string with format determined by `coercionType`.</span></span>

<span data-ttu-id="33c8e-951">Тип: String</span><span class="sxs-lookup"><span data-stu-id="33c8e-951">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="33c8e-952">Пример</span><span class="sxs-lookup"><span data-stu-id="33c8e-952">Example</span></span>

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

#### <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="33c8e-953">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="33c8e-953">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="33c8e-954">Асинхронно загружает настраиваемые свойства для надстройки для выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="33c8e-954">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="33c8e-p161">Настраиваемые свойства сохраняются в виде пар "ключ-значение" для каждого приложения и каждого элемента. Этот метод возвращает объект `CustomProperties` при обратном вызове, который предоставляет методы для доступа к настраиваемым свойствам, характерным для текущего элемента и текущей надстройки. Настраиваемые свойства не шифруются для элемента, поэтому этот способ хранения не является безопасным.</span><span class="sxs-lookup"><span data-stu-id="33c8e-p161">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="33c8e-958">Параметры</span><span class="sxs-lookup"><span data-stu-id="33c8e-958">Parameters</span></span>

|<span data-ttu-id="33c8e-959">Имя</span><span class="sxs-lookup"><span data-stu-id="33c8e-959">Name</span></span>| <span data-ttu-id="33c8e-960">Тип</span><span class="sxs-lookup"><span data-stu-id="33c8e-960">Type</span></span>| <span data-ttu-id="33c8e-961">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="33c8e-961">Attributes</span></span>| <span data-ttu-id="33c8e-962">Описание</span><span class="sxs-lookup"><span data-stu-id="33c8e-962">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="33c8e-963">function</span><span class="sxs-lookup"><span data-stu-id="33c8e-963">function</span></span>||<span data-ttu-id="33c8e-964">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="33c8e-964">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="33c8e-965">Настраиваемые свойства предоставляются в виде объекта [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.2) в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="33c8e-965">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.2) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="33c8e-966">Этот объект позволяет получить, задать и удалить настраиваемые свойства для элемента, а также сохранить изменения, внесенные в набор настраиваемых свойств, на сервере.</span><span class="sxs-lookup"><span data-stu-id="33c8e-966">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="33c8e-967">Объект</span><span class="sxs-lookup"><span data-stu-id="33c8e-967">Object</span></span>| <span data-ttu-id="33c8e-968">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="33c8e-968">&lt;optional&gt;</span></span>|<span data-ttu-id="33c8e-969">Разработчики могут указать любой объект, к которому необходимо получить доступ, в функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="33c8e-969">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="33c8e-970">Доступ к этому объекту можно получить с помощью свойства `asyncResult.asyncContext` в функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="33c8e-970">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="33c8e-971">Требования</span><span class="sxs-lookup"><span data-stu-id="33c8e-971">Requirements</span></span>

|<span data-ttu-id="33c8e-972">Требование</span><span class="sxs-lookup"><span data-stu-id="33c8e-972">Requirement</span></span>| <span data-ttu-id="33c8e-973">Значение</span><span class="sxs-lookup"><span data-stu-id="33c8e-973">Value</span></span>|
|---|---|
|[<span data-ttu-id="33c8e-974">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="33c8e-974">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="33c8e-975">1.0</span><span class="sxs-lookup"><span data-stu-id="33c8e-975">1.0</span></span>|
|[<span data-ttu-id="33c8e-976">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="33c8e-976">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="33c8e-977">ReadItem</span><span class="sxs-lookup"><span data-stu-id="33c8e-977">ReadItem</span></span>|
|[<span data-ttu-id="33c8e-978">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="33c8e-978">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="33c8e-979">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="33c8e-979">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="33c8e-980">Пример</span><span class="sxs-lookup"><span data-stu-id="33c8e-980">Example</span></span>

<span data-ttu-id="33c8e-p164">Приведенный ниже пример кода показывает, как асинхронно загружать настраиваемые свойства, характерные для текущего элемента, с помощью метода `loadCustomPropertiesAsync`. Этот пример также показывает, как сохранять эти свойства на сервере с помощью метода `CustomProperties.saveAsync`. После загрузки настраиваемых свойств в этом примере кода метод `CustomProperties.get` используется для считывания настраиваемого свойства `myProp`, метод `CustomProperties.set` — для записи настраиваемого свойства `otherProp`, а метод `saveAsync` — для сохранения настраиваемых свойств.</span><span class="sxs-lookup"><span data-stu-id="33c8e-p164">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

#### <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="33c8e-984">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="33c8e-984">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="33c8e-985">Удаляет вложение из сообщения или встречи.</span><span class="sxs-lookup"><span data-stu-id="33c8e-985">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="33c8e-986">Метод `removeAttachmentAsync` удаляет из элемента вложение с указанным идентификатором.</span><span class="sxs-lookup"><span data-stu-id="33c8e-986">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="33c8e-987">Идентификатор вложения рекомендуется использовать для удаления вложения, только если оно добавлено тем же почтовым приложением в ходе текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="33c8e-987">As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session.</span></span> <span data-ttu-id="33c8e-988">В Outlook в Интернете и мобильных устройствах идентификатор вложения действителен только в рамках одного сеанса.</span><span class="sxs-lookup"><span data-stu-id="33c8e-988">In Outlook on the web and mobile devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="33c8e-989">Сеанс завершается, когда пользователь закрывает приложение или начинает создавать элемент во встроенной форме, а затем переходит из формы в отдельное окно.</span><span class="sxs-lookup"><span data-stu-id="33c8e-989">A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="33c8e-990">Параметры</span><span class="sxs-lookup"><span data-stu-id="33c8e-990">Parameters</span></span>

|<span data-ttu-id="33c8e-991">Имя</span><span class="sxs-lookup"><span data-stu-id="33c8e-991">Name</span></span>| <span data-ttu-id="33c8e-992">Тип</span><span class="sxs-lookup"><span data-stu-id="33c8e-992">Type</span></span>| <span data-ttu-id="33c8e-993">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="33c8e-993">Attributes</span></span>| <span data-ttu-id="33c8e-994">Описание</span><span class="sxs-lookup"><span data-stu-id="33c8e-994">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="33c8e-995">String</span><span class="sxs-lookup"><span data-stu-id="33c8e-995">String</span></span>||<span data-ttu-id="33c8e-996">Идентификатор удаляемого вложения.</span><span class="sxs-lookup"><span data-stu-id="33c8e-996">The identifier of the attachment to remove.</span></span>|
|`options`| <span data-ttu-id="33c8e-997">Object</span><span class="sxs-lookup"><span data-stu-id="33c8e-997">Object</span></span>| <span data-ttu-id="33c8e-998">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="33c8e-998">&lt;optional&gt;</span></span>|<span data-ttu-id="33c8e-999">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="33c8e-999">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="33c8e-1000">Объект</span><span class="sxs-lookup"><span data-stu-id="33c8e-1000">Object</span></span>| <span data-ttu-id="33c8e-1001">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="33c8e-1001">&lt;optional&gt;</span></span>|<span data-ttu-id="33c8e-1002">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="33c8e-1002">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="33c8e-1003">функция</span><span class="sxs-lookup"><span data-stu-id="33c8e-1003">function</span></span>| <span data-ttu-id="33c8e-1004">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="33c8e-1004">&lt;optional&gt;</span></span>|<span data-ttu-id="33c8e-1005">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="33c8e-1005">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="33c8e-1006">Если удалить вложение не удается, свойство `asyncResult.error` содержит код ошибки с указанием ее причины.</span><span class="sxs-lookup"><span data-stu-id="33c8e-1006">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="33c8e-1007">Ошибки</span><span class="sxs-lookup"><span data-stu-id="33c8e-1007">Errors</span></span>

| <span data-ttu-id="33c8e-1008">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="33c8e-1008">Error code</span></span> | <span data-ttu-id="33c8e-1009">Описание</span><span class="sxs-lookup"><span data-stu-id="33c8e-1009">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="33c8e-1010">Идентификатор вложения не существует.</span><span class="sxs-lookup"><span data-stu-id="33c8e-1010">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="33c8e-1011">Требования</span><span class="sxs-lookup"><span data-stu-id="33c8e-1011">Requirements</span></span>

|<span data-ttu-id="33c8e-1012">Требование</span><span class="sxs-lookup"><span data-stu-id="33c8e-1012">Requirement</span></span>| <span data-ttu-id="33c8e-1013">Значение</span><span class="sxs-lookup"><span data-stu-id="33c8e-1013">Value</span></span>|
|---|---|
|[<span data-ttu-id="33c8e-1014">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="33c8e-1014">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="33c8e-1015">1.1</span><span class="sxs-lookup"><span data-stu-id="33c8e-1015">1.1</span></span>|
|[<span data-ttu-id="33c8e-1016">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="33c8e-1016">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="33c8e-1017">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="33c8e-1017">ReadWriteItem</span></span>|
|[<span data-ttu-id="33c8e-1018">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="33c8e-1018">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="33c8e-1019">Создание</span><span class="sxs-lookup"><span data-stu-id="33c8e-1019">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="33c8e-1020">Пример</span><span class="sxs-lookup"><span data-stu-id="33c8e-1020">Example</span></span>

<span data-ttu-id="33c8e-1021">Указанный ниже код удаляет вложение с идентификатором "0".</span><span class="sxs-lookup"><span data-stu-id="33c8e-1021">The following code removes an attachment with an identifier of '0'.</span></span>

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

#### <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="33c8e-1022">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="33c8e-1022">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="33c8e-1023">Асинхронно вставляет данные в текст или тему сообщения.</span><span class="sxs-lookup"><span data-stu-id="33c8e-1023">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="33c8e-p166">Метод `setSelectedDataAsync` вставляет указанную строку в местоположение курсора в теме или тексте элемента либо, если текст выделен в редакторе, он заменяет выделенный текст. Если курсор находится вне текста или темы элемента, возвращается ошибка. После вставки курсор помещается в конец вставленного содержимого.</span><span class="sxs-lookup"><span data-stu-id="33c8e-p166">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="33c8e-1027">Параметры</span><span class="sxs-lookup"><span data-stu-id="33c8e-1027">Parameters</span></span>

|<span data-ttu-id="33c8e-1028">Имя</span><span class="sxs-lookup"><span data-stu-id="33c8e-1028">Name</span></span>| <span data-ttu-id="33c8e-1029">Тип</span><span class="sxs-lookup"><span data-stu-id="33c8e-1029">Type</span></span>| <span data-ttu-id="33c8e-1030">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="33c8e-1030">Attributes</span></span>| <span data-ttu-id="33c8e-1031">Описание</span><span class="sxs-lookup"><span data-stu-id="33c8e-1031">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="33c8e-1032">String</span><span class="sxs-lookup"><span data-stu-id="33c8e-1032">String</span></span>||<span data-ttu-id="33c8e-p167">Вставляемые данные. Объем данных не должен превышать 1 000 000 символов. Если передано больше 1 000 000 символов, возвращается исключение `ArgumentOutOfRange`.</span><span class="sxs-lookup"><span data-stu-id="33c8e-p167">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="33c8e-1036">Object</span><span class="sxs-lookup"><span data-stu-id="33c8e-1036">Object</span></span>| <span data-ttu-id="33c8e-1037">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="33c8e-1037">&lt;optional&gt;</span></span>|<span data-ttu-id="33c8e-1038">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="33c8e-1038">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="33c8e-1039">Объект</span><span class="sxs-lookup"><span data-stu-id="33c8e-1039">Object</span></span>| <span data-ttu-id="33c8e-1040">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="33c8e-1040">&lt;optional&gt;</span></span>|<span data-ttu-id="33c8e-1041">В методе обратного вызова разработчики могут указать любой объект, к которому необходимо получить доступ.</span><span class="sxs-lookup"><span data-stu-id="33c8e-1041">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="33c8e-1042">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="33c8e-1042">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="33c8e-1043">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="33c8e-1043">&lt;optional&gt;</span></span>|<span data-ttu-id="33c8e-1044">Если `text`текущий стиль применяется в Outlook для веб-клиентов и клиентов для настольных ПК.</span><span class="sxs-lookup"><span data-stu-id="33c8e-1044">If `text`, the current style is applied in Outlook on the web and desktop clients.</span></span> <span data-ttu-id="33c8e-1045">Если поле представляет собой редактор HTML, вставляются только текстовые данные, даже если они имеют формат HTML.</span><span class="sxs-lookup"><span data-stu-id="33c8e-1045">If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="33c8e-1046">Если `html` и поле поддерживает HTML (тема не используется), текущий стиль применяется в Outlook в Интернете, а в настольных клиентах Outlook применяется стиль по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="33c8e-1046">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook on the web and the default style is applied in Outlook desktop clients.</span></span> <span data-ttu-id="33c8e-1047">Если поле является текстовым, возвращается ошибка `InvalidDataFormat`.</span><span class="sxs-lookup"><span data-stu-id="33c8e-1047">If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="33c8e-1048">Если свойство `coercionType` не задано, результат зависит от поля: если поле имеет формат HTML, используется текст в формате HTML, а если поле текстовое, применяется обычный текст.</span><span class="sxs-lookup"><span data-stu-id="33c8e-1048">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="33c8e-1049">функция</span><span class="sxs-lookup"><span data-stu-id="33c8e-1049">function</span></span>||<span data-ttu-id="33c8e-1050">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="33c8e-1050">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="33c8e-1051">Требования</span><span class="sxs-lookup"><span data-stu-id="33c8e-1051">Requirements</span></span>

|<span data-ttu-id="33c8e-1052">Требование</span><span class="sxs-lookup"><span data-stu-id="33c8e-1052">Requirement</span></span>| <span data-ttu-id="33c8e-1053">Значение</span><span class="sxs-lookup"><span data-stu-id="33c8e-1053">Value</span></span>|
|---|---|
|[<span data-ttu-id="33c8e-1054">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="33c8e-1054">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="33c8e-1055">1.2</span><span class="sxs-lookup"><span data-stu-id="33c8e-1055">1.2</span></span>|
|[<span data-ttu-id="33c8e-1056">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="33c8e-1056">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="33c8e-1057">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="33c8e-1057">ReadWriteItem</span></span>|
|[<span data-ttu-id="33c8e-1058">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="33c8e-1058">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="33c8e-1059">Создание</span><span class="sxs-lookup"><span data-stu-id="33c8e-1059">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="33c8e-1060">Пример</span><span class="sxs-lookup"><span data-stu-id="33c8e-1060">Example</span></span>

```js
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
