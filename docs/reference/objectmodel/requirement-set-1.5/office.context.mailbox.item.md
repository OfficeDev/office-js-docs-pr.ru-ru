---
title: Office.context.mailbox.item — набор обязательных элементов 1.5
description: ''
ms.date: 09/23/2019
localization_priority: Priority
ms.openlocfilehash: a76440c2353483d0754aa0ba2fed98f051bc1c6f
ms.sourcegitcommit: 3c84fe6302341668c3f9f6dd64e636a97d03023c
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/26/2019
ms.locfileid: "37167370"
---
# <a name="item"></a><span data-ttu-id="99922-102">item</span><span class="sxs-lookup"><span data-stu-id="99922-102">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="99922-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="99922-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="99922-p101">Пространство имен `item` используется для доступа к выбранному в данный момент сообщению, приглашению на собрание или описанию встречи. Вы можете определить тип пространства имен `item` с помощью свойства [itemType](#itemtype-officemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="99922-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="99922-106">Requirements</span><span class="sxs-lookup"><span data-stu-id="99922-106">Requirements</span></span>

|<span data-ttu-id="99922-107">Требование</span><span class="sxs-lookup"><span data-stu-id="99922-107">Requirement</span></span>| <span data-ttu-id="99922-108">Значение</span><span class="sxs-lookup"><span data-stu-id="99922-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="99922-109">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="99922-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="99922-110">1.0</span><span class="sxs-lookup"><span data-stu-id="99922-110">1.0</span></span>|
|[<span data-ttu-id="99922-111">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="99922-111">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="99922-112">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="99922-112">Restricted</span></span>|
|[<span data-ttu-id="99922-113">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="99922-113">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="99922-114">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="99922-114">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="99922-115">Элементы и методы</span><span class="sxs-lookup"><span data-stu-id="99922-115">Members and methods</span></span>

| <span data-ttu-id="99922-116">Элемент</span><span class="sxs-lookup"><span data-stu-id="99922-116">Member</span></span> | <span data-ttu-id="99922-117">Тип</span><span class="sxs-lookup"><span data-stu-id="99922-117">Type</span></span> |
|--------|------|
| [<span data-ttu-id="99922-118">attachments</span><span class="sxs-lookup"><span data-stu-id="99922-118">attachments</span></span>](#attachments-arrayattachmentdetails) | <span data-ttu-id="99922-119">Элемент</span><span class="sxs-lookup"><span data-stu-id="99922-119">Member</span></span> |
| [<span data-ttu-id="99922-120">bcc</span><span class="sxs-lookup"><span data-stu-id="99922-120">bcc</span></span>](#bcc-recipients) | <span data-ttu-id="99922-121">Элемент</span><span class="sxs-lookup"><span data-stu-id="99922-121">Member</span></span> |
| [<span data-ttu-id="99922-122">body</span><span class="sxs-lookup"><span data-stu-id="99922-122">body</span></span>](#body-body) | <span data-ttu-id="99922-123">Элемент</span><span class="sxs-lookup"><span data-stu-id="99922-123">Member</span></span> |
| [<span data-ttu-id="99922-124">cc</span><span class="sxs-lookup"><span data-stu-id="99922-124">cc</span></span>](#cc-arrayemailaddressdetailsrecipients) | <span data-ttu-id="99922-125">Элемент</span><span class="sxs-lookup"><span data-stu-id="99922-125">Member</span></span> |
| [<span data-ttu-id="99922-126">conversationId</span><span class="sxs-lookup"><span data-stu-id="99922-126">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="99922-127">Элемент</span><span class="sxs-lookup"><span data-stu-id="99922-127">Member</span></span> |
| [<span data-ttu-id="99922-128">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="99922-128">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="99922-129">Элемент</span><span class="sxs-lookup"><span data-stu-id="99922-129">Member</span></span> |
| [<span data-ttu-id="99922-130">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="99922-130">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="99922-131">Элемент</span><span class="sxs-lookup"><span data-stu-id="99922-131">Member</span></span> |
| [<span data-ttu-id="99922-132">end</span><span class="sxs-lookup"><span data-stu-id="99922-132">end</span></span>](#end-datetime) | <span data-ttu-id="99922-133">Элемент</span><span class="sxs-lookup"><span data-stu-id="99922-133">Member</span></span> |
| [<span data-ttu-id="99922-134">from</span><span class="sxs-lookup"><span data-stu-id="99922-134">from</span></span>](#from-emailaddressdetails) | <span data-ttu-id="99922-135">Элемент</span><span class="sxs-lookup"><span data-stu-id="99922-135">Member</span></span> |
| [<span data-ttu-id="99922-136">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="99922-136">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="99922-137">Элемент</span><span class="sxs-lookup"><span data-stu-id="99922-137">Member</span></span> |
| [<span data-ttu-id="99922-138">itemClass</span><span class="sxs-lookup"><span data-stu-id="99922-138">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="99922-139">Элемент</span><span class="sxs-lookup"><span data-stu-id="99922-139">Member</span></span> |
| [<span data-ttu-id="99922-140">itemId</span><span class="sxs-lookup"><span data-stu-id="99922-140">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="99922-141">Элемент</span><span class="sxs-lookup"><span data-stu-id="99922-141">Member</span></span> |
| [<span data-ttu-id="99922-142">itemType</span><span class="sxs-lookup"><span data-stu-id="99922-142">itemType</span></span>](#itemtype-officemailboxenumsitemtype) | <span data-ttu-id="99922-143">Элемент</span><span class="sxs-lookup"><span data-stu-id="99922-143">Member</span></span> |
| [<span data-ttu-id="99922-144">location</span><span class="sxs-lookup"><span data-stu-id="99922-144">location</span></span>](#location-stringlocation) | <span data-ttu-id="99922-145">Элемент</span><span class="sxs-lookup"><span data-stu-id="99922-145">Member</span></span> |
| [<span data-ttu-id="99922-146">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="99922-146">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="99922-147">Элемент</span><span class="sxs-lookup"><span data-stu-id="99922-147">Member</span></span> |
| [<span data-ttu-id="99922-148">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="99922-148">notificationMessages</span></span>](#notificationmessages-notificationmessages) | <span data-ttu-id="99922-149">Элемент</span><span class="sxs-lookup"><span data-stu-id="99922-149">Member</span></span> |
| [<span data-ttu-id="99922-150">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="99922-150">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="99922-151">Элемент</span><span class="sxs-lookup"><span data-stu-id="99922-151">Member</span></span> |
| [<span data-ttu-id="99922-152">organizer</span><span class="sxs-lookup"><span data-stu-id="99922-152">organizer</span></span>](#organizer-emailaddressdetails) | <span data-ttu-id="99922-153">Элемент</span><span class="sxs-lookup"><span data-stu-id="99922-153">Member</span></span> |
| [<span data-ttu-id="99922-154">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="99922-154">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="99922-155">Member</span><span class="sxs-lookup"><span data-stu-id="99922-155">Member</span></span> |
| [<span data-ttu-id="99922-156">sender</span><span class="sxs-lookup"><span data-stu-id="99922-156">sender</span></span>](#sender-emailaddressdetails) | <span data-ttu-id="99922-157">Элемент</span><span class="sxs-lookup"><span data-stu-id="99922-157">Member</span></span> |
| [<span data-ttu-id="99922-158">start</span><span class="sxs-lookup"><span data-stu-id="99922-158">start</span></span>](#start-datetime) | <span data-ttu-id="99922-159">Элемент</span><span class="sxs-lookup"><span data-stu-id="99922-159">Member</span></span> |
| [<span data-ttu-id="99922-160">subject</span><span class="sxs-lookup"><span data-stu-id="99922-160">subject</span></span>](#subject-stringsubject) | <span data-ttu-id="99922-161">Элемент</span><span class="sxs-lookup"><span data-stu-id="99922-161">Member</span></span> |
| [<span data-ttu-id="99922-162">to</span><span class="sxs-lookup"><span data-stu-id="99922-162">to</span></span>](#to-arrayemailaddressdetailsrecipients) | <span data-ttu-id="99922-163">Элемент</span><span class="sxs-lookup"><span data-stu-id="99922-163">Member</span></span> |
| [<span data-ttu-id="99922-164">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="99922-164">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="99922-165">Метод</span><span class="sxs-lookup"><span data-stu-id="99922-165">Method</span></span> |
| [<span data-ttu-id="99922-166">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="99922-166">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="99922-167">Метод</span><span class="sxs-lookup"><span data-stu-id="99922-167">Method</span></span> |
| [<span data-ttu-id="99922-168">close</span><span class="sxs-lookup"><span data-stu-id="99922-168">close</span></span>](#close) | <span data-ttu-id="99922-169">Метод</span><span class="sxs-lookup"><span data-stu-id="99922-169">Method</span></span> |
| [<span data-ttu-id="99922-170">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="99922-170">displayReplyAllForm</span></span>](#displayreplyallformformdata-callback) | <span data-ttu-id="99922-171">Метод</span><span class="sxs-lookup"><span data-stu-id="99922-171">Method</span></span> |
| [<span data-ttu-id="99922-172">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="99922-172">displayReplyForm</span></span>](#displayreplyformformdata-callback) | <span data-ttu-id="99922-173">Метод</span><span class="sxs-lookup"><span data-stu-id="99922-173">Method</span></span> |
| [<span data-ttu-id="99922-174">getEntities</span><span class="sxs-lookup"><span data-stu-id="99922-174">getEntities</span></span>](#getentities--entities) | <span data-ttu-id="99922-175">Метод</span><span class="sxs-lookup"><span data-stu-id="99922-175">Method</span></span> |
| [<span data-ttu-id="99922-176">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="99922-176">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="99922-177">Метод</span><span class="sxs-lookup"><span data-stu-id="99922-177">Method</span></span> |
| [<span data-ttu-id="99922-178">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="99922-178">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="99922-179">Метод</span><span class="sxs-lookup"><span data-stu-id="99922-179">Method</span></span> |
| [<span data-ttu-id="99922-180">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="99922-180">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="99922-181">Метод</span><span class="sxs-lookup"><span data-stu-id="99922-181">Method</span></span> |
| [<span data-ttu-id="99922-182">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="99922-182">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="99922-183">Метод</span><span class="sxs-lookup"><span data-stu-id="99922-183">Method</span></span> |
| [<span data-ttu-id="99922-184">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="99922-184">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="99922-185">Метод</span><span class="sxs-lookup"><span data-stu-id="99922-185">Method</span></span> |
| [<span data-ttu-id="99922-186">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="99922-186">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="99922-187">Метод</span><span class="sxs-lookup"><span data-stu-id="99922-187">Method</span></span> |
| [<span data-ttu-id="99922-188">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="99922-188">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="99922-189">Метод</span><span class="sxs-lookup"><span data-stu-id="99922-189">Method</span></span> |
| [<span data-ttu-id="99922-190">saveAsync</span><span class="sxs-lookup"><span data-stu-id="99922-190">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="99922-191">Метод</span><span class="sxs-lookup"><span data-stu-id="99922-191">Method</span></span> |
| [<span data-ttu-id="99922-192">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="99922-192">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="99922-193">Метод</span><span class="sxs-lookup"><span data-stu-id="99922-193">Method</span></span> |

### <a name="example"></a><span data-ttu-id="99922-194">Пример</span><span class="sxs-lookup"><span data-stu-id="99922-194">Example</span></span>

<span data-ttu-id="99922-195">В примере кода JavaScript, приведенном ниже, показано, как получить доступ к свойству `subject` текущего элемента в Outlook.</span><span class="sxs-lookup"><span data-stu-id="99922-195">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="99922-196">Members</span><span class="sxs-lookup"><span data-stu-id="99922-196">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetailsviewoutlook-js-15"></a><span data-ttu-id="99922-197">attachments: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.5)></span><span class="sxs-lookup"><span data-stu-id="99922-197">attachments :Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.5)></span></span>

<span data-ttu-id="99922-p102">Получает массив вложений для элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="99922-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="99922-200">Outlook блокирует определенные типы файлов из-за потенциальных проблем с безопасностью, поэтому они не возвращаются.</span><span class="sxs-lookup"><span data-stu-id="99922-200">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="99922-201">Дополнительные сведения см. в статье [Блокировка вложений в Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="99922-201">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="99922-202">Тип</span><span class="sxs-lookup"><span data-stu-id="99922-202">Type</span></span>

*   <span data-ttu-id="99922-203">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.5)></span><span class="sxs-lookup"><span data-stu-id="99922-203">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.5)></span></span>

##### <a name="requirements"></a><span data-ttu-id="99922-204">Требования</span><span class="sxs-lookup"><span data-stu-id="99922-204">Requirements</span></span>

|<span data-ttu-id="99922-205">Требование</span><span class="sxs-lookup"><span data-stu-id="99922-205">Requirement</span></span>| <span data-ttu-id="99922-206">Значение</span><span class="sxs-lookup"><span data-stu-id="99922-206">Value</span></span>|
|---|---|
|[<span data-ttu-id="99922-207">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="99922-207">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="99922-208">1.0</span><span class="sxs-lookup"><span data-stu-id="99922-208">1.0</span></span>|
|[<span data-ttu-id="99922-209">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="99922-209">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="99922-210">ReadItem</span><span class="sxs-lookup"><span data-stu-id="99922-210">ReadItem</span></span>|
|[<span data-ttu-id="99922-211">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="99922-211">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="99922-212">Чтение</span><span class="sxs-lookup"><span data-stu-id="99922-212">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="99922-213">Пример</span><span class="sxs-lookup"><span data-stu-id="99922-213">Example</span></span>

<span data-ttu-id="99922-214">С помощью приведенного ниже кода можно создать HTML-строку с подробными сведениями обо всех вложениях для текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="99922-214">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

#### <a name="bcc-recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-15"></a><span data-ttu-id="99922-215">bcc: [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="99922-215">bcc :[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.5)</span></span>

<span data-ttu-id="99922-216">Получает объект, который предоставляет методы для получения или обновления получателей скрытой копии сообщения.</span><span class="sxs-lookup"><span data-stu-id="99922-216">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="99922-217">Только в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="99922-217">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="99922-218">Тип</span><span class="sxs-lookup"><span data-stu-id="99922-218">Type</span></span>

*   [<span data-ttu-id="99922-219">Получатели</span><span class="sxs-lookup"><span data-stu-id="99922-219">Recipients</span></span>](/javascript/api/outlook/office.recipients?view=outlook-js-1.5)

##### <a name="requirements"></a><span data-ttu-id="99922-220">Требования</span><span class="sxs-lookup"><span data-stu-id="99922-220">Requirements</span></span>

|<span data-ttu-id="99922-221">Требование</span><span class="sxs-lookup"><span data-stu-id="99922-221">Requirement</span></span>| <span data-ttu-id="99922-222">Значение</span><span class="sxs-lookup"><span data-stu-id="99922-222">Value</span></span>|
|---|---|
|[<span data-ttu-id="99922-223">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="99922-223">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="99922-224">1.1</span><span class="sxs-lookup"><span data-stu-id="99922-224">1.1</span></span>|
|[<span data-ttu-id="99922-225">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="99922-225">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="99922-226">ReadItem</span><span class="sxs-lookup"><span data-stu-id="99922-226">ReadItem</span></span>|
|[<span data-ttu-id="99922-227">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="99922-227">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="99922-228">Создание</span><span class="sxs-lookup"><span data-stu-id="99922-228">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="99922-229">Пример</span><span class="sxs-lookup"><span data-stu-id="99922-229">Example</span></span>

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

#### <a name="body-bodyjavascriptapioutlookofficebodyviewoutlook-js-15"></a><span data-ttu-id="99922-230">body: [Body](/javascript/api/outlook/office.body?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="99922-230">body :[Body](/javascript/api/outlook/office.body?view=outlook-js-1.5)</span></span>

<span data-ttu-id="99922-231">Получает объект, предоставляющий методы для работы с основным текстом элемента.</span><span class="sxs-lookup"><span data-stu-id="99922-231">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="99922-232">Тип</span><span class="sxs-lookup"><span data-stu-id="99922-232">Type</span></span>

*   [<span data-ttu-id="99922-233">Body</span><span class="sxs-lookup"><span data-stu-id="99922-233">Body</span></span>](/javascript/api/outlook/office.body?view=outlook-js-1.5)

##### <a name="requirements"></a><span data-ttu-id="99922-234">Требования</span><span class="sxs-lookup"><span data-stu-id="99922-234">Requirements</span></span>

|<span data-ttu-id="99922-235">Требование</span><span class="sxs-lookup"><span data-stu-id="99922-235">Requirement</span></span>| <span data-ttu-id="99922-236">Значение</span><span class="sxs-lookup"><span data-stu-id="99922-236">Value</span></span>|
|---|---|
|[<span data-ttu-id="99922-237">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="99922-237">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="99922-238">1.1</span><span class="sxs-lookup"><span data-stu-id="99922-238">1.1</span></span>|
|[<span data-ttu-id="99922-239">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="99922-239">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="99922-240">ReadItem</span><span class="sxs-lookup"><span data-stu-id="99922-240">ReadItem</span></span>|
|[<span data-ttu-id="99922-241">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="99922-241">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="99922-242">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="99922-242">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="99922-243">Пример</span><span class="sxs-lookup"><span data-stu-id="99922-243">Example</span></span>

<span data-ttu-id="99922-244">В этом примере возвращается текст сообщения в формате обычного текста.</span><span class="sxs-lookup"><span data-stu-id="99922-244">This example gets the body of the message in plain text.</span></span>

```js
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="99922-245">Ниже приведен пример итогового параметра, переданного функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="99922-245">The following is an example of the result parameter passed to the callback function.</span></span>

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

#### <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-15recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-15"></a><span data-ttu-id="99922-246">cc: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="99922-246">cc :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.5)</span></span>

<span data-ttu-id="99922-247">Предоставляет доступ к получателям копии сообщения.</span><span class="sxs-lookup"><span data-stu-id="99922-247">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="99922-248">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="99922-248">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="99922-249">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="99922-249">Read mode</span></span>

<span data-ttu-id="99922-p106">Свойство `cc` возвращает массив, который содержит объект `EmailAddressDetails` для каждого получателя, указанного в строке **Копия** сообщения. Коллекция может включать не более 100 элементов.</span><span class="sxs-lookup"><span data-stu-id="99922-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="99922-252">Режим создания</span><span class="sxs-lookup"><span data-stu-id="99922-252">Compose mode</span></span>

<span data-ttu-id="99922-253">Свойство `cc` возвращает объект `Recipients`, предоставляющий методы для получения или обновления получателей, которые указаны в строке **Копия** сообщения.</span><span class="sxs-lookup"><span data-stu-id="99922-253">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

```js
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="99922-254">Тип</span><span class="sxs-lookup"><span data-stu-id="99922-254">Type</span></span>

*   <span data-ttu-id="99922-255">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="99922-255">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.5)</span></span>

##### <a name="requirements"></a><span data-ttu-id="99922-256">Требования</span><span class="sxs-lookup"><span data-stu-id="99922-256">Requirements</span></span>

|<span data-ttu-id="99922-257">Требование</span><span class="sxs-lookup"><span data-stu-id="99922-257">Requirement</span></span>| <span data-ttu-id="99922-258">Значение</span><span class="sxs-lookup"><span data-stu-id="99922-258">Value</span></span>|
|---|---|
|[<span data-ttu-id="99922-259">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="99922-259">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="99922-260">1.0</span><span class="sxs-lookup"><span data-stu-id="99922-260">1.0</span></span>|
|[<span data-ttu-id="99922-261">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="99922-261">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="99922-262">ReadItem</span><span class="sxs-lookup"><span data-stu-id="99922-262">ReadItem</span></span>|
|[<span data-ttu-id="99922-263">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="99922-263">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="99922-264">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="99922-264">Compose or Read</span></span>|

<br>

---
---

#### <a name="nullable-conversationid-string"></a><span data-ttu-id="99922-265">(nullable) conversationId: String</span><span class="sxs-lookup"><span data-stu-id="99922-265">(nullable) conversationId :String</span></span>

<span data-ttu-id="99922-266">Получает идентификатор разговора по электронной почте, содержащего конкретное сообщение.</span><span class="sxs-lookup"><span data-stu-id="99922-266">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="99922-p107">Вы можете получить целочисленное значение этого свойства, если ваше почтовое приложение активируется в формах просмотра или формах создания ответов. Если пользователь изменит тему ответа, после его отправки идентификатор беседы будет изменен, и полученное ранее значение будет недействительным.</span><span class="sxs-lookup"><span data-stu-id="99922-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="99922-p108">Это свойство имеет значение NULL для нового элемента в форме создания. Свойство `conversationId` вернет значение, если пользователь задаст тему и сохранит элемент.</span><span class="sxs-lookup"><span data-stu-id="99922-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="99922-271">Тип</span><span class="sxs-lookup"><span data-stu-id="99922-271">Type</span></span>

*   <span data-ttu-id="99922-272">String</span><span class="sxs-lookup"><span data-stu-id="99922-272">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="99922-273">Требования</span><span class="sxs-lookup"><span data-stu-id="99922-273">Requirements</span></span>

|<span data-ttu-id="99922-274">Требование</span><span class="sxs-lookup"><span data-stu-id="99922-274">Requirement</span></span>| <span data-ttu-id="99922-275">Значение</span><span class="sxs-lookup"><span data-stu-id="99922-275">Value</span></span>|
|---|---|
|[<span data-ttu-id="99922-276">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="99922-276">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="99922-277">1.0</span><span class="sxs-lookup"><span data-stu-id="99922-277">1.0</span></span>|
|[<span data-ttu-id="99922-278">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="99922-278">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="99922-279">ReadItem</span><span class="sxs-lookup"><span data-stu-id="99922-279">ReadItem</span></span>|
|[<span data-ttu-id="99922-280">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="99922-280">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="99922-281">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="99922-281">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="99922-282">Пример</span><span class="sxs-lookup"><span data-stu-id="99922-282">Example</span></span>

```js
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

<br>

---
---

#### <a name="datetimecreated-date"></a><span data-ttu-id="99922-283">dateTimeCreated: Date</span><span class="sxs-lookup"><span data-stu-id="99922-283">dateTimeCreated :Date</span></span>

<span data-ttu-id="99922-p109">Получает дату и время создания элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="99922-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="99922-286">Тип</span><span class="sxs-lookup"><span data-stu-id="99922-286">Type</span></span>

*   <span data-ttu-id="99922-287">Дата</span><span class="sxs-lookup"><span data-stu-id="99922-287">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="99922-288">Требования</span><span class="sxs-lookup"><span data-stu-id="99922-288">Requirements</span></span>

|<span data-ttu-id="99922-289">Требование</span><span class="sxs-lookup"><span data-stu-id="99922-289">Requirement</span></span>| <span data-ttu-id="99922-290">Значение</span><span class="sxs-lookup"><span data-stu-id="99922-290">Value</span></span>|
|---|---|
|[<span data-ttu-id="99922-291">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="99922-291">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="99922-292">1.0</span><span class="sxs-lookup"><span data-stu-id="99922-292">1.0</span></span>|
|[<span data-ttu-id="99922-293">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="99922-293">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="99922-294">ReadItem</span><span class="sxs-lookup"><span data-stu-id="99922-294">ReadItem</span></span>|
|[<span data-ttu-id="99922-295">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="99922-295">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="99922-296">Чтение</span><span class="sxs-lookup"><span data-stu-id="99922-296">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="99922-297">Пример</span><span class="sxs-lookup"><span data-stu-id="99922-297">Example</span></span>

```js
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

<br>

---
---

#### <a name="datetimemodified-date"></a><span data-ttu-id="99922-298">dateTimeModified: Date</span><span class="sxs-lookup"><span data-stu-id="99922-298">dateTimeModified :Date</span></span>

<span data-ttu-id="99922-p110">Получает дату и время последнего изменения элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="99922-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="99922-301">Этот элемент не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="99922-301">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="99922-302">Тип</span><span class="sxs-lookup"><span data-stu-id="99922-302">Type</span></span>

*   <span data-ttu-id="99922-303">Дата</span><span class="sxs-lookup"><span data-stu-id="99922-303">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="99922-304">Требования</span><span class="sxs-lookup"><span data-stu-id="99922-304">Requirements</span></span>

|<span data-ttu-id="99922-305">Требование</span><span class="sxs-lookup"><span data-stu-id="99922-305">Requirement</span></span>| <span data-ttu-id="99922-306">Значение</span><span class="sxs-lookup"><span data-stu-id="99922-306">Value</span></span>|
|---|---|
|[<span data-ttu-id="99922-307">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="99922-307">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="99922-308">1.0</span><span class="sxs-lookup"><span data-stu-id="99922-308">1.0</span></span>|
|[<span data-ttu-id="99922-309">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="99922-309">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="99922-310">ReadItem</span><span class="sxs-lookup"><span data-stu-id="99922-310">ReadItem</span></span>|
|[<span data-ttu-id="99922-311">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="99922-311">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="99922-312">Чтение</span><span class="sxs-lookup"><span data-stu-id="99922-312">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="99922-313">Пример</span><span class="sxs-lookup"><span data-stu-id="99922-313">Example</span></span>

```js
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

<br>

---
---

#### <a name="end-datetimejavascriptapioutlookofficetimeviewoutlook-js-15"></a><span data-ttu-id="99922-314">end: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="99922-314">end :Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.5)</span></span>

<span data-ttu-id="99922-315">Получает или задает дату и время окончания встречи.</span><span class="sxs-lookup"><span data-stu-id="99922-315">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="99922-p111">Свойство `end` представлено в виде значения даты и времени в формате UTC. Преобразовать значение свойства end в местные значения даты и времени клиента можно с помощью метода [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime).</span><span class="sxs-lookup"><span data-stu-id="99922-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="99922-318">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="99922-318">Read mode</span></span>

<span data-ttu-id="99922-319">Свойство `end` возвращает объект `Date`.</span><span class="sxs-lookup"><span data-stu-id="99922-319">The `end` property returns a `Date` object.</span></span>

```js
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="99922-320">Режим создания</span><span class="sxs-lookup"><span data-stu-id="99922-320">Compose mode</span></span>

<span data-ttu-id="99922-321">Свойство `end` возвращает объект `Time`.</span><span class="sxs-lookup"><span data-stu-id="99922-321">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="99922-322">Если вы задаете время окончания с помощью метода [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.5#setasync-datetime--options--callback-), необходимо использовать метод [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) для преобразования местного времени на клиенте в формат UTC для сервера.</span><span class="sxs-lookup"><span data-stu-id="99922-322">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.5#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="99922-323">В примере ниже показано, как с помощью метода [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.5#setasync-datetime--options--callback-) объекта `Time` задать время окончания встречи.</span><span class="sxs-lookup"><span data-stu-id="99922-323">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.5#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="99922-324">Тип</span><span class="sxs-lookup"><span data-stu-id="99922-324">Type</span></span>

*   <span data-ttu-id="99922-325">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="99922-325">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.5)</span></span>

##### <a name="requirements"></a><span data-ttu-id="99922-326">Требования</span><span class="sxs-lookup"><span data-stu-id="99922-326">Requirements</span></span>

|<span data-ttu-id="99922-327">Требование</span><span class="sxs-lookup"><span data-stu-id="99922-327">Requirement</span></span>| <span data-ttu-id="99922-328">Значение</span><span class="sxs-lookup"><span data-stu-id="99922-328">Value</span></span>|
|---|---|
|[<span data-ttu-id="99922-329">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="99922-329">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="99922-330">1.0</span><span class="sxs-lookup"><span data-stu-id="99922-330">1.0</span></span>|
|[<span data-ttu-id="99922-331">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="99922-331">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="99922-332">ReadItem</span><span class="sxs-lookup"><span data-stu-id="99922-332">ReadItem</span></span>|
|[<span data-ttu-id="99922-333">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="99922-333">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="99922-334">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="99922-334">Compose or Read</span></span>|

<br>

---
---

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-15"></a><span data-ttu-id="99922-335">from: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="99922-335">from :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)</span></span>

<span data-ttu-id="99922-p112">Получает электронный адрес отправителя сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="99922-p112">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="99922-p113">Свойства `from` и [`sender`](#sender-emailaddressdetails) представляют одно лицо, если сообщение не отправлено представителем. В противном случае свойство `from` представляет лицо, делегировавшее полномочия, а свойство sender — представителя.</span><span class="sxs-lookup"><span data-stu-id="99922-p113">The `from` and [`sender`](#sender-emailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="99922-340">Свойству `recipientType`, принадлежащему объекту `EmailAddressDetails` в свойстве `from`, задано значение `undefined`.</span><span class="sxs-lookup"><span data-stu-id="99922-340">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="99922-341">Тип</span><span class="sxs-lookup"><span data-stu-id="99922-341">Type</span></span>

*   [<span data-ttu-id="99922-342">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="99922-342">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)

##### <a name="requirements"></a><span data-ttu-id="99922-343">Требования</span><span class="sxs-lookup"><span data-stu-id="99922-343">Requirements</span></span>

|<span data-ttu-id="99922-344">Требование</span><span class="sxs-lookup"><span data-stu-id="99922-344">Requirement</span></span>| <span data-ttu-id="99922-345">Значение</span><span class="sxs-lookup"><span data-stu-id="99922-345">Value</span></span>|
|---|---|
|[<span data-ttu-id="99922-346">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="99922-346">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="99922-347">1.0</span><span class="sxs-lookup"><span data-stu-id="99922-347">1.0</span></span>|
|[<span data-ttu-id="99922-348">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="99922-348">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="99922-349">ReadItem</span><span class="sxs-lookup"><span data-stu-id="99922-349">ReadItem</span></span>|
|[<span data-ttu-id="99922-350">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="99922-350">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="99922-351">Чтение</span><span class="sxs-lookup"><span data-stu-id="99922-351">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="99922-352">Пример</span><span class="sxs-lookup"><span data-stu-id="99922-352">Example</span></span>

```js
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

<br>

---
---

#### <a name="internetmessageid-string"></a><span data-ttu-id="99922-353">internetMessageId: String</span><span class="sxs-lookup"><span data-stu-id="99922-353">internetMessageId :String</span></span>

<span data-ttu-id="99922-p114">Получает идентификатор интернет-сообщения для электронного сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="99922-p114">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="99922-356">Тип</span><span class="sxs-lookup"><span data-stu-id="99922-356">Type</span></span>

*   <span data-ttu-id="99922-357">String</span><span class="sxs-lookup"><span data-stu-id="99922-357">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="99922-358">Требования</span><span class="sxs-lookup"><span data-stu-id="99922-358">Requirements</span></span>

|<span data-ttu-id="99922-359">Требование</span><span class="sxs-lookup"><span data-stu-id="99922-359">Requirement</span></span>| <span data-ttu-id="99922-360">Значение</span><span class="sxs-lookup"><span data-stu-id="99922-360">Value</span></span>|
|---|---|
|[<span data-ttu-id="99922-361">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="99922-361">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="99922-362">1.0</span><span class="sxs-lookup"><span data-stu-id="99922-362">1.0</span></span>|
|[<span data-ttu-id="99922-363">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="99922-363">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="99922-364">ReadItem</span><span class="sxs-lookup"><span data-stu-id="99922-364">ReadItem</span></span>|
|[<span data-ttu-id="99922-365">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="99922-365">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="99922-366">Чтение</span><span class="sxs-lookup"><span data-stu-id="99922-366">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="99922-367">Пример</span><span class="sxs-lookup"><span data-stu-id="99922-367">Example</span></span>

```js
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

<br>

---
---

#### <a name="itemclass-string"></a><span data-ttu-id="99922-368">itemClass: String</span><span class="sxs-lookup"><span data-stu-id="99922-368">itemClass :String</span></span>

<span data-ttu-id="99922-p115">Получает класс элемента веб-служб Exchange для выбранного элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="99922-p115">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="99922-p116">Свойство `itemClass` указывает класс сообщения выбранного элемента. Ниже приводятся классы сообщения по умолчанию для элемента сообщения или встречи.</span><span class="sxs-lookup"><span data-stu-id="99922-p116">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="99922-373">Тип</span><span class="sxs-lookup"><span data-stu-id="99922-373">Type</span></span> | <span data-ttu-id="99922-374">Описание</span><span class="sxs-lookup"><span data-stu-id="99922-374">Description</span></span> | <span data-ttu-id="99922-375">Класс элемента</span><span class="sxs-lookup"><span data-stu-id="99922-375">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="99922-376">Элементы встречи</span><span class="sxs-lookup"><span data-stu-id="99922-376">Appointment items</span></span> | <span data-ttu-id="99922-377">Это элементы календаря для класса элемента `IPM.Appointment` или `IPM.Appointment.Occurrence`.</span><span class="sxs-lookup"><span data-stu-id="99922-377">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurrence` |
| <span data-ttu-id="99922-378">Элементы сообщения</span><span class="sxs-lookup"><span data-stu-id="99922-378">Message items</span></span> | <span data-ttu-id="99922-379">Сюда входят электронные сообщения, для которых по умолчанию задан класс сообщения `IPM.Note`, а также приглашения на собрания, ответы на них и уведомления об их отмене, использующие `IPM.Schedule.Meeting` в качестве базового класса сообщения.</span><span class="sxs-lookup"><span data-stu-id="99922-379">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="99922-380">Можно создавать настраиваемые классы сообщения, расширяющие классы сообщения по умолчанию, например настраиваемый класс сообщения о встрече `IPM.Appointment.Contoso`.</span><span class="sxs-lookup"><span data-stu-id="99922-380">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="99922-381">Тип</span><span class="sxs-lookup"><span data-stu-id="99922-381">Type</span></span>

*   <span data-ttu-id="99922-382">String</span><span class="sxs-lookup"><span data-stu-id="99922-382">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="99922-383">Требования</span><span class="sxs-lookup"><span data-stu-id="99922-383">Requirements</span></span>

|<span data-ttu-id="99922-384">Требование</span><span class="sxs-lookup"><span data-stu-id="99922-384">Requirement</span></span>| <span data-ttu-id="99922-385">Значение</span><span class="sxs-lookup"><span data-stu-id="99922-385">Value</span></span>|
|---|---|
|[<span data-ttu-id="99922-386">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="99922-386">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="99922-387">1.0</span><span class="sxs-lookup"><span data-stu-id="99922-387">1.0</span></span>|
|[<span data-ttu-id="99922-388">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="99922-388">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="99922-389">ReadItem</span><span class="sxs-lookup"><span data-stu-id="99922-389">ReadItem</span></span>|
|[<span data-ttu-id="99922-390">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="99922-390">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="99922-391">Чтение</span><span class="sxs-lookup"><span data-stu-id="99922-391">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="99922-392">Пример</span><span class="sxs-lookup"><span data-stu-id="99922-392">Example</span></span>

```js
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

<br>

---
---

#### <a name="nullable-itemid-string"></a><span data-ttu-id="99922-393">(nullable) itemId: String</span><span class="sxs-lookup"><span data-stu-id="99922-393">(nullable) itemId :String</span></span>

<span data-ttu-id="99922-p117">Получает идентификатор элемента веб-служб Exchange для текущего элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="99922-p117">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="99922-396">Идентификатор, возвращаемый свойством `itemId`, совпадает с идентификатором элемента веб-служб Exchange.</span><span class="sxs-lookup"><span data-stu-id="99922-396">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="99922-397">Свойство `itemId` не совпадает с идентификатором записи Outlook, а также идентификатором, который используется REST API Outlook.</span><span class="sxs-lookup"><span data-stu-id="99922-397">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="99922-398">Прежде чем совершать вызовы REST API, используя это значение, его необходимо преобразовать с помощью [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="99922-398">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="99922-399">Дополнительные сведения см. в статье [Использование REST API Outlook из надстройки Outlook](/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="99922-399">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="99922-p119">Свойство `itemId` недоступно в режиме создания. Если требуется идентификатор элемента, с помощью метода [`saveAsync`](#saveasyncoptions-callback) можно сохранить элемент в хранилище. При этом в параметре [`AsyncResult.value`](/javascript/api/office/office.asyncresult) функции обратного вызова возвращается идентификатор элемента.</span><span class="sxs-lookup"><span data-stu-id="99922-p119">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="99922-402">Тип</span><span class="sxs-lookup"><span data-stu-id="99922-402">Type</span></span>

*   <span data-ttu-id="99922-403">String</span><span class="sxs-lookup"><span data-stu-id="99922-403">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="99922-404">Требования</span><span class="sxs-lookup"><span data-stu-id="99922-404">Requirements</span></span>

|<span data-ttu-id="99922-405">Требование</span><span class="sxs-lookup"><span data-stu-id="99922-405">Requirement</span></span>| <span data-ttu-id="99922-406">Значение</span><span class="sxs-lookup"><span data-stu-id="99922-406">Value</span></span>|
|---|---|
|[<span data-ttu-id="99922-407">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="99922-407">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="99922-408">1.0</span><span class="sxs-lookup"><span data-stu-id="99922-408">1.0</span></span>|
|[<span data-ttu-id="99922-409">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="99922-409">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="99922-410">ReadItem</span><span class="sxs-lookup"><span data-stu-id="99922-410">ReadItem</span></span>|
|[<span data-ttu-id="99922-411">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="99922-411">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="99922-412">Чтение</span><span class="sxs-lookup"><span data-stu-id="99922-412">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="99922-413">Пример</span><span class="sxs-lookup"><span data-stu-id="99922-413">Example</span></span>

<span data-ttu-id="99922-p120">Указанный ниже код проверяет наличие идентификатора элемента. Если свойство `itemId` возвращает значение `null` или `undefined`, элемент будет сохранен в хранилище, а из асинхронного результата будет получен идентификатор элемента.</span><span class="sxs-lookup"><span data-stu-id="99922-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

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

#### <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtypeviewoutlook-js-15"></a><span data-ttu-id="99922-416">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="99922-416">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.5)</span></span>

<span data-ttu-id="99922-417">Получает тип элемента, который представляет экземпляр.</span><span class="sxs-lookup"><span data-stu-id="99922-417">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="99922-418">Свойство `itemType` возвращает одно из значений перечисления `ItemType`, которое указывает, является ли экземпляр объекта `item` сообщением или встречей.</span><span class="sxs-lookup"><span data-stu-id="99922-418">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="99922-419">Тип</span><span class="sxs-lookup"><span data-stu-id="99922-419">Type</span></span>

*   [<span data-ttu-id="99922-420">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="99922-420">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.5)

##### <a name="requirements"></a><span data-ttu-id="99922-421">Требования</span><span class="sxs-lookup"><span data-stu-id="99922-421">Requirements</span></span>

|<span data-ttu-id="99922-422">Требование</span><span class="sxs-lookup"><span data-stu-id="99922-422">Requirement</span></span>| <span data-ttu-id="99922-423">Значение</span><span class="sxs-lookup"><span data-stu-id="99922-423">Value</span></span>|
|---|---|
|[<span data-ttu-id="99922-424">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="99922-424">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="99922-425">1.0</span><span class="sxs-lookup"><span data-stu-id="99922-425">1.0</span></span>|
|[<span data-ttu-id="99922-426">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="99922-426">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="99922-427">ReadItem</span><span class="sxs-lookup"><span data-stu-id="99922-427">ReadItem</span></span>|
|[<span data-ttu-id="99922-428">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="99922-428">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="99922-429">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="99922-429">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="99922-430">Пример</span><span class="sxs-lookup"><span data-stu-id="99922-430">Example</span></span>

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

#### <a name="location-stringlocationjavascriptapioutlookofficelocationviewoutlook-js-15"></a><span data-ttu-id="99922-431">location: String|[Location](/javascript/api/outlook/office.location?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="99922-431">location :String|[Location](/javascript/api/outlook/office.location?view=outlook-js-1.5)</span></span>

<span data-ttu-id="99922-432">Получает или задает место встречи.</span><span class="sxs-lookup"><span data-stu-id="99922-432">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="99922-433">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="99922-433">Read mode</span></span>

<span data-ttu-id="99922-434">Свойство `location` возвращает строку, содержащую сведения о месте встречи.</span><span class="sxs-lookup"><span data-stu-id="99922-434">The `location` property returns a string that contains the location of the appointment.</span></span>

```js
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="99922-435">Режим создания</span><span class="sxs-lookup"><span data-stu-id="99922-435">Compose mode</span></span>

<span data-ttu-id="99922-436">Свойство `location` возвращает объект `Location`, предоставляющий методы, которые используются для получения и задания места встречи.</span><span class="sxs-lookup"><span data-stu-id="99922-436">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```js
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="99922-437">Тип</span><span class="sxs-lookup"><span data-stu-id="99922-437">Type</span></span>

*   <span data-ttu-id="99922-438">String | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="99922-438">String | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.5)</span></span>

##### <a name="requirements"></a><span data-ttu-id="99922-439">Требования</span><span class="sxs-lookup"><span data-stu-id="99922-439">Requirements</span></span>

|<span data-ttu-id="99922-440">Требование</span><span class="sxs-lookup"><span data-stu-id="99922-440">Requirement</span></span>| <span data-ttu-id="99922-441">Значение</span><span class="sxs-lookup"><span data-stu-id="99922-441">Value</span></span>|
|---|---|
|[<span data-ttu-id="99922-442">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="99922-442">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="99922-443">1.0</span><span class="sxs-lookup"><span data-stu-id="99922-443">1.0</span></span>|
|[<span data-ttu-id="99922-444">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="99922-444">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="99922-445">ReadItem</span><span class="sxs-lookup"><span data-stu-id="99922-445">ReadItem</span></span>|
|[<span data-ttu-id="99922-446">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="99922-446">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="99922-447">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="99922-447">Compose or Read</span></span>|

<br>

---
---

#### <a name="normalizedsubject-string"></a><span data-ttu-id="99922-448">normalizedSubject: String</span><span class="sxs-lookup"><span data-stu-id="99922-448">normalizedSubject :String</span></span>

<span data-ttu-id="99922-p121">Получает тему элемента со всеми удаленными префиксами (включая `RE:` и `FWD:`). Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="99922-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="99922-p122">Свойство normalizedSubject получает тему элемента со стандартными префиксами (такими как `RE:` и `FW:`), добавляемыми почтовыми программами. Для получения темы элемента с неизмененными префиксами используйте свойство [`subject`](#subject-stringsubject).</span><span class="sxs-lookup"><span data-stu-id="99922-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="99922-453">Тип</span><span class="sxs-lookup"><span data-stu-id="99922-453">Type</span></span>

*   <span data-ttu-id="99922-454">String</span><span class="sxs-lookup"><span data-stu-id="99922-454">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="99922-455">Требования</span><span class="sxs-lookup"><span data-stu-id="99922-455">Requirements</span></span>

|<span data-ttu-id="99922-456">Требование</span><span class="sxs-lookup"><span data-stu-id="99922-456">Requirement</span></span>| <span data-ttu-id="99922-457">Значение</span><span class="sxs-lookup"><span data-stu-id="99922-457">Value</span></span>|
|---|---|
|[<span data-ttu-id="99922-458">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="99922-458">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="99922-459">1.0</span><span class="sxs-lookup"><span data-stu-id="99922-459">1.0</span></span>|
|[<span data-ttu-id="99922-460">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="99922-460">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="99922-461">ReadItem</span><span class="sxs-lookup"><span data-stu-id="99922-461">ReadItem</span></span>|
|[<span data-ttu-id="99922-462">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="99922-462">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="99922-463">Чтение</span><span class="sxs-lookup"><span data-stu-id="99922-463">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="99922-464">Пример</span><span class="sxs-lookup"><span data-stu-id="99922-464">Example</span></span>

```js
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

<br>

---
---

#### <a name="notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessagesviewoutlook-js-15"></a><span data-ttu-id="99922-465">notificationMessages: [NotificationMessages](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="99922-465">notificationMessages :[NotificationMessages](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.5)</span></span>

<span data-ttu-id="99922-466">Получает сообщения уведомления для элемента.</span><span class="sxs-lookup"><span data-stu-id="99922-466">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="99922-467">Тип</span><span class="sxs-lookup"><span data-stu-id="99922-467">Type</span></span>

*   [<span data-ttu-id="99922-468">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="99922-468">NotificationMessages</span></span>](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.5)

##### <a name="requirements"></a><span data-ttu-id="99922-469">Требования</span><span class="sxs-lookup"><span data-stu-id="99922-469">Requirements</span></span>

|<span data-ttu-id="99922-470">Требование</span><span class="sxs-lookup"><span data-stu-id="99922-470">Requirement</span></span>| <span data-ttu-id="99922-471">Значение</span><span class="sxs-lookup"><span data-stu-id="99922-471">Value</span></span>|
|---|---|
|[<span data-ttu-id="99922-472">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="99922-472">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="99922-473">1.3</span><span class="sxs-lookup"><span data-stu-id="99922-473">1.3</span></span>|
|[<span data-ttu-id="99922-474">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="99922-474">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="99922-475">ReadItem</span><span class="sxs-lookup"><span data-stu-id="99922-475">ReadItem</span></span>|
|[<span data-ttu-id="99922-476">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="99922-476">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="99922-477">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="99922-477">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="99922-478">Пример</span><span class="sxs-lookup"><span data-stu-id="99922-478">Example</span></span>

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

#### <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-15recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-15"></a><span data-ttu-id="99922-479">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="99922-479">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.5)</span></span>

<span data-ttu-id="99922-480">Предоставляет доступ к необязательным участникам события.</span><span class="sxs-lookup"><span data-stu-id="99922-480">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="99922-481">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="99922-481">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="99922-482">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="99922-482">Read mode</span></span>

<span data-ttu-id="99922-483">Свойство `optionalAttendees` возвращает массив, содержащий объект `EmailAddressDetails` для каждого необязательного участника собрания.</span><span class="sxs-lookup"><span data-stu-id="99922-483">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

```js
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="99922-484">Режим создания</span><span class="sxs-lookup"><span data-stu-id="99922-484">Compose mode</span></span>

<span data-ttu-id="99922-485">Свойство `optionalAttendees` возвращает объект `Recipients`, который предоставляет методы для получения или обновления необязательных участников собрания.</span><span class="sxs-lookup"><span data-stu-id="99922-485">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

```js
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="99922-486">Тип</span><span class="sxs-lookup"><span data-stu-id="99922-486">Type</span></span>

*   <span data-ttu-id="99922-487">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="99922-487">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.5)</span></span>

##### <a name="requirements"></a><span data-ttu-id="99922-488">Требования</span><span class="sxs-lookup"><span data-stu-id="99922-488">Requirements</span></span>

|<span data-ttu-id="99922-489">Требование</span><span class="sxs-lookup"><span data-stu-id="99922-489">Requirement</span></span>| <span data-ttu-id="99922-490">Значение</span><span class="sxs-lookup"><span data-stu-id="99922-490">Value</span></span>|
|---|---|
|[<span data-ttu-id="99922-491">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="99922-491">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="99922-492">1.0</span><span class="sxs-lookup"><span data-stu-id="99922-492">1.0</span></span>|
|[<span data-ttu-id="99922-493">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="99922-493">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="99922-494">ReadItem</span><span class="sxs-lookup"><span data-stu-id="99922-494">ReadItem</span></span>|
|[<span data-ttu-id="99922-495">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="99922-495">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="99922-496">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="99922-496">Compose or Read</span></span>|

<br>

---
---

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-15"></a><span data-ttu-id="99922-497">organizer: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="99922-497">organizer :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)</span></span>

<span data-ttu-id="99922-p124">Получает электронный адрес организатора указанного собрания. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="99922-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="99922-500">Тип</span><span class="sxs-lookup"><span data-stu-id="99922-500">Type</span></span>

*   [<span data-ttu-id="99922-501">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="99922-501">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)

##### <a name="requirements"></a><span data-ttu-id="99922-502">Требования</span><span class="sxs-lookup"><span data-stu-id="99922-502">Requirements</span></span>

|<span data-ttu-id="99922-503">Требование</span><span class="sxs-lookup"><span data-stu-id="99922-503">Requirement</span></span>| <span data-ttu-id="99922-504">Значение</span><span class="sxs-lookup"><span data-stu-id="99922-504">Value</span></span>|
|---|---|
|[<span data-ttu-id="99922-505">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="99922-505">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="99922-506">1.0</span><span class="sxs-lookup"><span data-stu-id="99922-506">1.0</span></span>|
|[<span data-ttu-id="99922-507">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="99922-507">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="99922-508">ReadItem</span><span class="sxs-lookup"><span data-stu-id="99922-508">ReadItem</span></span>|
|[<span data-ttu-id="99922-509">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="99922-509">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="99922-510">Чтение</span><span class="sxs-lookup"><span data-stu-id="99922-510">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="99922-511">Пример</span><span class="sxs-lookup"><span data-stu-id="99922-511">Example</span></span>

```js
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

<br>

---
---

#### <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-15recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-15"></a><span data-ttu-id="99922-512">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="99922-512">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.5)</span></span>

<span data-ttu-id="99922-513">Предоставляет доступ к обязательным участникам события.</span><span class="sxs-lookup"><span data-stu-id="99922-513">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="99922-514">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="99922-514">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="99922-515">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="99922-515">Read mode</span></span>

<span data-ttu-id="99922-516">Свойство `requiredAttendees` возвращает массив, содержащий объект `EmailAddressDetails` для каждого обязательного участника собрания.</span><span class="sxs-lookup"><span data-stu-id="99922-516">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

```js
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="99922-517">Режим создания</span><span class="sxs-lookup"><span data-stu-id="99922-517">Compose mode</span></span>

<span data-ttu-id="99922-518">Свойство `requiredAttendees` возвращает объект `Recipients`, предоставляющий методы, с помощью которых можно получить или обновить сведения об обязательных участниках собрания.</span><span class="sxs-lookup"><span data-stu-id="99922-518">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

```js
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="99922-519">Тип</span><span class="sxs-lookup"><span data-stu-id="99922-519">Type</span></span>

*   <span data-ttu-id="99922-520">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="99922-520">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.5)</span></span>

##### <a name="requirements"></a><span data-ttu-id="99922-521">Требования</span><span class="sxs-lookup"><span data-stu-id="99922-521">Requirements</span></span>

|<span data-ttu-id="99922-522">Требование</span><span class="sxs-lookup"><span data-stu-id="99922-522">Requirement</span></span>| <span data-ttu-id="99922-523">Значение</span><span class="sxs-lookup"><span data-stu-id="99922-523">Value</span></span>|
|---|---|
|[<span data-ttu-id="99922-524">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="99922-524">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="99922-525">1.0</span><span class="sxs-lookup"><span data-stu-id="99922-525">1.0</span></span>|
|[<span data-ttu-id="99922-526">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="99922-526">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="99922-527">ReadItem</span><span class="sxs-lookup"><span data-stu-id="99922-527">ReadItem</span></span>|
|[<span data-ttu-id="99922-528">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="99922-528">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="99922-529">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="99922-529">Compose or Read</span></span>|

<br>

---
---

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-15"></a><span data-ttu-id="99922-530">sender: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="99922-530">sender :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)</span></span>

<span data-ttu-id="99922-p126">Получает электронный адрес отправителя электронного сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="99922-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="99922-p127">Свойства [`from`](#from-emailaddressdetails) и `sender` представляют одно лицо, если сообщение не отправлено представителем. В противном случае свойство `from` представляет лицо, делегировавшее полномочия, а свойство sender — представителя.</span><span class="sxs-lookup"><span data-stu-id="99922-p127">The [`from`](#from-emailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="99922-535">Свойству `recipientType`, принадлежащему объекту `EmailAddressDetails` в свойстве `sender`, задано значение `undefined`.</span><span class="sxs-lookup"><span data-stu-id="99922-535">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="99922-536">Тип</span><span class="sxs-lookup"><span data-stu-id="99922-536">Type</span></span>

*   [<span data-ttu-id="99922-537">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="99922-537">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)

##### <a name="requirements"></a><span data-ttu-id="99922-538">Требования</span><span class="sxs-lookup"><span data-stu-id="99922-538">Requirements</span></span>

|<span data-ttu-id="99922-539">Требование</span><span class="sxs-lookup"><span data-stu-id="99922-539">Requirement</span></span>| <span data-ttu-id="99922-540">Значение</span><span class="sxs-lookup"><span data-stu-id="99922-540">Value</span></span>|
|---|---|
|[<span data-ttu-id="99922-541">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="99922-541">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="99922-542">1.0</span><span class="sxs-lookup"><span data-stu-id="99922-542">1.0</span></span>|
|[<span data-ttu-id="99922-543">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="99922-543">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="99922-544">ReadItem</span><span class="sxs-lookup"><span data-stu-id="99922-544">ReadItem</span></span>|
|[<span data-ttu-id="99922-545">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="99922-545">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="99922-546">Чтение</span><span class="sxs-lookup"><span data-stu-id="99922-546">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="99922-547">Пример</span><span class="sxs-lookup"><span data-stu-id="99922-547">Example</span></span>

```js
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

<br>

---
---

#### <a name="start-datetimejavascriptapioutlookofficetimeviewoutlook-js-15"></a><span data-ttu-id="99922-548">start: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="99922-548">start :Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.5)</span></span>

<span data-ttu-id="99922-549">Получает или задает дату и время начала встречи.</span><span class="sxs-lookup"><span data-stu-id="99922-549">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="99922-p128">Свойство `start` представлено в виде значения даты и времени в формате UTC. Это значение можно преобразовать в местные значения даты и времени клиента с помощью метода [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime).</span><span class="sxs-lookup"><span data-stu-id="99922-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="99922-552">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="99922-552">Read mode</span></span>

<span data-ttu-id="99922-553">Свойство `start` возвращает объект `Date`.</span><span class="sxs-lookup"><span data-stu-id="99922-553">The `start` property returns a `Date` object.</span></span>

```js
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="99922-554">Режим создания</span><span class="sxs-lookup"><span data-stu-id="99922-554">Compose mode</span></span>

<span data-ttu-id="99922-555">Свойство `start` возвращает объект `Time`.</span><span class="sxs-lookup"><span data-stu-id="99922-555">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="99922-556">Если вы задаете время начала с помощью метода [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.5#setasync-datetime--options--callback-), необходимо использовать метод [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) для преобразования местного времени на клиенте в формат UTC для сервера.</span><span class="sxs-lookup"><span data-stu-id="99922-556">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.5#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="99922-557">В примере ниже с помощью метода [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.5#setasync-datetime--options--callback-) объекта `Time` задается время начала встречи в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="99922-557">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.5#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="99922-558">Тип</span><span class="sxs-lookup"><span data-stu-id="99922-558">Type</span></span>

*   <span data-ttu-id="99922-559">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="99922-559">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.5)</span></span>

##### <a name="requirements"></a><span data-ttu-id="99922-560">Требования</span><span class="sxs-lookup"><span data-stu-id="99922-560">Requirements</span></span>

|<span data-ttu-id="99922-561">Требование</span><span class="sxs-lookup"><span data-stu-id="99922-561">Requirement</span></span>| <span data-ttu-id="99922-562">Значение</span><span class="sxs-lookup"><span data-stu-id="99922-562">Value</span></span>|
|---|---|
|[<span data-ttu-id="99922-563">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="99922-563">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="99922-564">1.0</span><span class="sxs-lookup"><span data-stu-id="99922-564">1.0</span></span>|
|[<span data-ttu-id="99922-565">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="99922-565">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="99922-566">ReadItem</span><span class="sxs-lookup"><span data-stu-id="99922-566">ReadItem</span></span>|
|[<span data-ttu-id="99922-567">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="99922-567">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="99922-568">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="99922-568">Compose or Read</span></span>|

<br>

---
---

#### <a name="subject-stringsubjectjavascriptapioutlookofficesubjectviewoutlook-js-15"></a><span data-ttu-id="99922-569">subject: String|[Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="99922-569">subject :String|[Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.5)</span></span>

<span data-ttu-id="99922-570">Получает или задает описание, которое отображается в поле темы элемента.</span><span class="sxs-lookup"><span data-stu-id="99922-570">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="99922-571">Свойство `subject` получает или задает всю тему элемента для отправки с почтового сервера.</span><span class="sxs-lookup"><span data-stu-id="99922-571">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="99922-572">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="99922-572">Read mode</span></span>

<span data-ttu-id="99922-p129">Свойство `subject` возвращает строку. С помощью свойства [`normalizedSubject`](#normalizedsubject-string) можно получить тему без начальных префиксов, таких как `RE:` и `FW:`.</span><span class="sxs-lookup"><span data-stu-id="99922-p129">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```js
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="99922-575">Режим создания</span><span class="sxs-lookup"><span data-stu-id="99922-575">Compose mode</span></span>

<span data-ttu-id="99922-576">Свойство `subject` возвращает объект `Subject`, который предоставляет методы для получения и задания темы.</span><span class="sxs-lookup"><span data-stu-id="99922-576">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```js
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="99922-577">Тип</span><span class="sxs-lookup"><span data-stu-id="99922-577">Type</span></span>

*   <span data-ttu-id="99922-578">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="99922-578">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.5)</span></span>

##### <a name="requirements"></a><span data-ttu-id="99922-579">Требования</span><span class="sxs-lookup"><span data-stu-id="99922-579">Requirements</span></span>

|<span data-ttu-id="99922-580">Требование</span><span class="sxs-lookup"><span data-stu-id="99922-580">Requirement</span></span>| <span data-ttu-id="99922-581">Значение</span><span class="sxs-lookup"><span data-stu-id="99922-581">Value</span></span>|
|---|---|
|[<span data-ttu-id="99922-582">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="99922-582">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="99922-583">1.0</span><span class="sxs-lookup"><span data-stu-id="99922-583">1.0</span></span>|
|[<span data-ttu-id="99922-584">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="99922-584">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="99922-585">ReadItem</span><span class="sxs-lookup"><span data-stu-id="99922-585">ReadItem</span></span>|
|[<span data-ttu-id="99922-586">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="99922-586">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="99922-587">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="99922-587">Compose or Read</span></span>|

<br>

---
---

#### <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-15recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-15"></a><span data-ttu-id="99922-588">to: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="99922-588">to :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.5)</span></span>

<span data-ttu-id="99922-589">Предоставляет доступ к получателям, указанным в строке **Кому** сообщения.</span><span class="sxs-lookup"><span data-stu-id="99922-589">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="99922-590">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="99922-590">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="99922-591">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="99922-591">Read mode</span></span>

<span data-ttu-id="99922-p131">Свойство `to` возвращает массив, содержащий объект `EmailAddressDetails` для каждого получателя в строке **Кому** сообщения. Коллекция может включать не более 100 элементов.</span><span class="sxs-lookup"><span data-stu-id="99922-p131">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="99922-594">Режим создания</span><span class="sxs-lookup"><span data-stu-id="99922-594">Compose mode</span></span>

<span data-ttu-id="99922-595">Свойство `to` возвращает объект `Recipients`, предоставляющий методы для получения или обновления получателей, которые указаны в строке **Кому** сообщения.</span><span class="sxs-lookup"><span data-stu-id="99922-595">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

```js
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="99922-596">Тип</span><span class="sxs-lookup"><span data-stu-id="99922-596">Type</span></span>

*   <span data-ttu-id="99922-597">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="99922-597">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.5)</span></span>

##### <a name="requirements"></a><span data-ttu-id="99922-598">Требования</span><span class="sxs-lookup"><span data-stu-id="99922-598">Requirements</span></span>

|<span data-ttu-id="99922-599">Требование</span><span class="sxs-lookup"><span data-stu-id="99922-599">Requirement</span></span>| <span data-ttu-id="99922-600">Значение</span><span class="sxs-lookup"><span data-stu-id="99922-600">Value</span></span>|
|---|---|
|[<span data-ttu-id="99922-601">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="99922-601">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="99922-602">1.0</span><span class="sxs-lookup"><span data-stu-id="99922-602">1.0</span></span>|
|[<span data-ttu-id="99922-603">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="99922-603">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="99922-604">ReadItem</span><span class="sxs-lookup"><span data-stu-id="99922-604">ReadItem</span></span>|
|[<span data-ttu-id="99922-605">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="99922-605">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="99922-606">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="99922-606">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="99922-607">Методы</span><span class="sxs-lookup"><span data-stu-id="99922-607">Methods</span></span>

#### <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="99922-608">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="99922-608">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="99922-609">Добавляет файл в сообщение или встречу в качестве вложения.</span><span class="sxs-lookup"><span data-stu-id="99922-609">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="99922-610">Метод `addFileAttachmentAsync` передает файл по указанному универсальному коду ресурса (URI) и вкладывает его в элемент в форме создания.</span><span class="sxs-lookup"><span data-stu-id="99922-610">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="99922-611">Идентификатор можно использовать с методом [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback), чтобы удалить вложение, добавленное во время текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="99922-611">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="99922-612">Параметры</span><span class="sxs-lookup"><span data-stu-id="99922-612">Parameters</span></span>

|<span data-ttu-id="99922-613">Имя</span><span class="sxs-lookup"><span data-stu-id="99922-613">Name</span></span>| <span data-ttu-id="99922-614">Тип</span><span class="sxs-lookup"><span data-stu-id="99922-614">Type</span></span>| <span data-ttu-id="99922-615">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="99922-615">Attributes</span></span>| <span data-ttu-id="99922-616">Описание</span><span class="sxs-lookup"><span data-stu-id="99922-616">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="99922-617">String</span><span class="sxs-lookup"><span data-stu-id="99922-617">String</span></span>||<span data-ttu-id="99922-p132">Универсальный код ресурса (URI), представляющий расположение файла, который нужно вложить в сообщение или встречу. Максимальная длина — 2048 символов.</span><span class="sxs-lookup"><span data-stu-id="99922-p132">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="99922-620">String</span><span class="sxs-lookup"><span data-stu-id="99922-620">String</span></span>||<span data-ttu-id="99922-p133">Имя вложения, которое отображается при передаче вложения. Максимальная длина — 255 символов.</span><span class="sxs-lookup"><span data-stu-id="99922-p133">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="99922-623">Object</span><span class="sxs-lookup"><span data-stu-id="99922-623">Object</span></span>| <span data-ttu-id="99922-624">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="99922-624">&lt;optional&gt;</span></span>|<span data-ttu-id="99922-625">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="99922-625">An object literal that contains one or more of the following properties.</span></span>|
| `options.asyncContext` | <span data-ttu-id="99922-626">Object</span><span class="sxs-lookup"><span data-stu-id="99922-626">Object</span></span> | <span data-ttu-id="99922-627">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="99922-627">&lt;optional&gt;</span></span> | <span data-ttu-id="99922-628">В методе обратного вызова разработчики могут указать любой объект, к которому необходимо получить доступ.</span><span class="sxs-lookup"><span data-stu-id="99922-628">Developers can provide any object they wish to access in the callback method.</span></span> |
| `options.isInline` | <span data-ttu-id="99922-629">Boolean</span><span class="sxs-lookup"><span data-stu-id="99922-629">Boolean</span></span> | <span data-ttu-id="99922-630">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="99922-630">&lt;optional&gt;</span></span> | <span data-ttu-id="99922-631">Значение `true` указывает на то, что вложение будет встроено в текст сообщения и не должно отображаться в списке вложений.</span><span class="sxs-lookup"><span data-stu-id="99922-631">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
|`callback`| <span data-ttu-id="99922-632">function</span><span class="sxs-lookup"><span data-stu-id="99922-632">function</span></span>| <span data-ttu-id="99922-633">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="99922-633">&lt;optional&gt;</span></span>|<span data-ttu-id="99922-634">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="99922-634">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="99922-635">После успешного выполнения идентификатор вложения будет представлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="99922-635">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="99922-636">Если добавить вложение не удастся, объект `asyncResult` будет содержать объект `Error` с описанием ошибки.</span><span class="sxs-lookup"><span data-stu-id="99922-636">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="99922-637">Ошибки</span><span class="sxs-lookup"><span data-stu-id="99922-637">Errors</span></span>

| <span data-ttu-id="99922-638">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="99922-638">Error code</span></span> | <span data-ttu-id="99922-639">Описание</span><span class="sxs-lookup"><span data-stu-id="99922-639">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="99922-640">Вложение превышает максимальный размер.</span><span class="sxs-lookup"><span data-stu-id="99922-640">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="99922-641">Расширение вложения не поддерживается.</span><span class="sxs-lookup"><span data-stu-id="99922-641">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="99922-642">Сообщение или встреча содержат слишком много вложений.</span><span class="sxs-lookup"><span data-stu-id="99922-642">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="99922-643">Требования</span><span class="sxs-lookup"><span data-stu-id="99922-643">Requirements</span></span>

|<span data-ttu-id="99922-644">Требование</span><span class="sxs-lookup"><span data-stu-id="99922-644">Requirement</span></span>| <span data-ttu-id="99922-645">Значение</span><span class="sxs-lookup"><span data-stu-id="99922-645">Value</span></span>|
|---|---|
|[<span data-ttu-id="99922-646">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="99922-646">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="99922-647">1.1</span><span class="sxs-lookup"><span data-stu-id="99922-647">1.1</span></span>|
|[<span data-ttu-id="99922-648">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="99922-648">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="99922-649">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="99922-649">ReadWriteItem</span></span>|
|[<span data-ttu-id="99922-650">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="99922-650">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="99922-651">Создание</span><span class="sxs-lookup"><span data-stu-id="99922-651">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="99922-652">Примеры</span><span class="sxs-lookup"><span data-stu-id="99922-652">Examples</span></span>

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

<span data-ttu-id="99922-653">В приведенном ниже примере файл изображения добавляется в качестве встроенного вложения, а ссылка на вложение добавляется в текст сообщения.</span><span class="sxs-lookup"><span data-stu-id="99922-653">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

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

#### <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="99922-654">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="99922-654">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="99922-655">Добавляет к сообщению элемент Exchange, например сообщение, в виде вложения.</span><span class="sxs-lookup"><span data-stu-id="99922-655">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="99922-p134">С помощью метода `addItemAttachmentAsync` можно в элемент формы создания вложить элемент с указанным идентификатором Exchange. Если указать метод обратного вызова, то этот метод вызывается с помощью параметра `asyncResult`, который содержит идентификатор вложения или код, указывающий на ошибки, которые произошли при вложении элемента. При необходимости можно использовать параметр `options` для передачи сведений о состоянии методу обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="99922-p134">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="99922-659">Идентификатор можно использовать с методом [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback), чтобы удалить вложение, добавленное во время текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="99922-659">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="99922-660">Если ваша надстройка Office выполняется в Outlook в Интернете, метод `addItemAttachmentAsync` обеспечивает вложение элементов в элементы, отличные от редактируемого. Однако это действие не рекомендуется выполнять, так как оно не поддерживается.</span><span class="sxs-lookup"><span data-stu-id="99922-660">If your Office Add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="99922-661">Параметры</span><span class="sxs-lookup"><span data-stu-id="99922-661">Parameters</span></span>

|<span data-ttu-id="99922-662">Имя</span><span class="sxs-lookup"><span data-stu-id="99922-662">Name</span></span>| <span data-ttu-id="99922-663">Тип</span><span class="sxs-lookup"><span data-stu-id="99922-663">Type</span></span>| <span data-ttu-id="99922-664">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="99922-664">Attributes</span></span>| <span data-ttu-id="99922-665">Описание</span><span class="sxs-lookup"><span data-stu-id="99922-665">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="99922-666">String</span><span class="sxs-lookup"><span data-stu-id="99922-666">String</span></span>||<span data-ttu-id="99922-p135">Идентификатор Exchange для вкладываемого элемента. Максимальная длина — 100 символов.</span><span class="sxs-lookup"><span data-stu-id="99922-p135">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="99922-669">String</span><span class="sxs-lookup"><span data-stu-id="99922-669">String</span></span>||<span data-ttu-id="99922-670">Тема вкладываемого элемента.</span><span class="sxs-lookup"><span data-stu-id="99922-670">The subject of the item to be attached.</span></span> <span data-ttu-id="99922-671">Максимальная длина: 255 символов.</span><span class="sxs-lookup"><span data-stu-id="99922-671">The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="99922-672">Object</span><span class="sxs-lookup"><span data-stu-id="99922-672">Object</span></span>| <span data-ttu-id="99922-673">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="99922-673">&lt;optional&gt;</span></span>|<span data-ttu-id="99922-674">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="99922-674">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="99922-675">Object</span><span class="sxs-lookup"><span data-stu-id="99922-675">Object</span></span>| <span data-ttu-id="99922-676">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="99922-676">&lt;optional&gt;</span></span>|<span data-ttu-id="99922-677">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="99922-677">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="99922-678">функция</span><span class="sxs-lookup"><span data-stu-id="99922-678">function</span></span>| <span data-ttu-id="99922-679">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="99922-679">&lt;optional&gt;</span></span>|<span data-ttu-id="99922-680">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="99922-680">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="99922-681">После успешного выполнения идентификатор вложения будет представлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="99922-681">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="99922-682">Если добавить вложение не удастся, объект `asyncResult` будет содержать объект `Error` с описанием ошибки.</span><span class="sxs-lookup"><span data-stu-id="99922-682">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="99922-683">Ошибки</span><span class="sxs-lookup"><span data-stu-id="99922-683">Errors</span></span>

| <span data-ttu-id="99922-684">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="99922-684">Error code</span></span> | <span data-ttu-id="99922-685">Описание</span><span class="sxs-lookup"><span data-stu-id="99922-685">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="99922-686">Сообщение или встреча содержат слишком много вложений.</span><span class="sxs-lookup"><span data-stu-id="99922-686">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="99922-687">Требования</span><span class="sxs-lookup"><span data-stu-id="99922-687">Requirements</span></span>

|<span data-ttu-id="99922-688">Требование</span><span class="sxs-lookup"><span data-stu-id="99922-688">Requirement</span></span>| <span data-ttu-id="99922-689">Значение</span><span class="sxs-lookup"><span data-stu-id="99922-689">Value</span></span>|
|---|---|
|[<span data-ttu-id="99922-690">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="99922-690">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="99922-691">1.1</span><span class="sxs-lookup"><span data-stu-id="99922-691">1.1</span></span>|
|[<span data-ttu-id="99922-692">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="99922-692">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="99922-693">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="99922-693">ReadWriteItem</span></span>|
|[<span data-ttu-id="99922-694">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="99922-694">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="99922-695">Создание</span><span class="sxs-lookup"><span data-stu-id="99922-695">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="99922-696">Пример</span><span class="sxs-lookup"><span data-stu-id="99922-696">Example</span></span>

<span data-ttu-id="99922-697">В следующем примере существующий элемент Outlook добавляется в виде вложения с именем `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="99922-697">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

#### <a name="close"></a><span data-ttu-id="99922-698">close()</span><span class="sxs-lookup"><span data-stu-id="99922-698">close()</span></span>

<span data-ttu-id="99922-699">Закрывает текущий создаваемый элемент.</span><span class="sxs-lookup"><span data-stu-id="99922-699">Closes the current item that is being composed.</span></span>

<span data-ttu-id="99922-p137">Работа метода `close` зависит от текущего состояния создаваемого элемента. Если элемент содержит несохраненные изменения, клиент предложит пользователю сохранить или отклонить их либо отменить действие закрытия.</span><span class="sxs-lookup"><span data-stu-id="99922-p137">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="99922-702">Если в Outlook в Интернете элемент представляет собой встречу, ранее сохраненную с помощью метода `saveAsync`, пользователю предлагается сохранить, отклонить или отменить действие, даже если с момента последнего сохранения элемента не вносились какие-либо изменения.</span><span class="sxs-lookup"><span data-stu-id="99922-702">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="99922-703">Если в клиенте Outlook для настольных ПК сообщение представляет собой ответ в тексте, метод `close` не работает.</span><span class="sxs-lookup"><span data-stu-id="99922-703">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="99922-704">Требования</span><span class="sxs-lookup"><span data-stu-id="99922-704">Requirements</span></span>

|<span data-ttu-id="99922-705">Требование</span><span class="sxs-lookup"><span data-stu-id="99922-705">Requirement</span></span>| <span data-ttu-id="99922-706">Значение</span><span class="sxs-lookup"><span data-stu-id="99922-706">Value</span></span>|
|---|---|
|[<span data-ttu-id="99922-707">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="99922-707">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="99922-708">1.3</span><span class="sxs-lookup"><span data-stu-id="99922-708">1.3</span></span>|
|[<span data-ttu-id="99922-709">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="99922-709">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="99922-710">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="99922-710">Restricted</span></span>|
|[<span data-ttu-id="99922-711">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="99922-711">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="99922-712">Создание</span><span class="sxs-lookup"><span data-stu-id="99922-712">Compose</span></span>|

<br>

---
---

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="99922-713">displayReplyAllForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="99922-713">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="99922-714">Отображает форму ответа, включающую отправителя и всех получателей выбранного сообщения или организатора и всех участников выбранной встречи.</span><span class="sxs-lookup"><span data-stu-id="99922-714">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="99922-715">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="99922-715">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="99922-716">В Outlook в Интернете форма ответа отображается в виде всплывающей формы в представлении с 3 либо 1 или 2 столбцами.</span><span class="sxs-lookup"><span data-stu-id="99922-716">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="99922-717">Если любой строковый параметр превышает указанные для него ограничения, `displayReplyAllForm` возвращает исключение.</span><span class="sxs-lookup"><span data-stu-id="99922-717">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="99922-p138">Если в параметре `formData.attachments` указаны вложения, Outlook в Интернете и классические клиенты пытаются скачать их и вложить в форму ответа. Если какие-либо вложения добавить не удается, в форме отображается сообщение об ошибке. Если сообщения об ошибках не предусмотрены, то они не отображаются.</span><span class="sxs-lookup"><span data-stu-id="99922-p138">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="99922-721">Параметры</span><span class="sxs-lookup"><span data-stu-id="99922-721">Parameters</span></span>

| <span data-ttu-id="99922-722">Имя</span><span class="sxs-lookup"><span data-stu-id="99922-722">Name</span></span> | <span data-ttu-id="99922-723">Тип</span><span class="sxs-lookup"><span data-stu-id="99922-723">Type</span></span> | <span data-ttu-id="99922-724">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="99922-724">Attributes</span></span> | <span data-ttu-id="99922-725">Описание</span><span class="sxs-lookup"><span data-stu-id="99922-725">Description</span></span> |
|---|---|---|---|
|`formData`| <span data-ttu-id="99922-726">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="99922-726">String &#124; Object</span></span>| |<span data-ttu-id="99922-p139">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="99922-p139">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="99922-729">**ИЛИ**</span><span class="sxs-lookup"><span data-stu-id="99922-729">**OR**</span></span><br/><span data-ttu-id="99922-p140">Объект, который содержит текст или данные вложения и функцию обратного вызова. Ниже представлено определение этого объекта.</span><span class="sxs-lookup"><span data-stu-id="99922-p140">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="99922-732">String</span><span class="sxs-lookup"><span data-stu-id="99922-732">String</span></span> | <span data-ttu-id="99922-733">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="99922-733">&lt;optional&gt;</span></span> | <span data-ttu-id="99922-p141">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="99922-p141">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="99922-736">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="99922-736">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="99922-737">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="99922-737">&lt;optional&gt;</span></span> | <span data-ttu-id="99922-738">Массив объектов JSON, представляющих собой вложенные файлы или элементы.</span><span class="sxs-lookup"><span data-stu-id="99922-738">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="99922-739">String</span><span class="sxs-lookup"><span data-stu-id="99922-739">String</span></span> | | <span data-ttu-id="99922-p142">Указывает тип вложения. Допустимые значения: `file` для вложенного файла и `item` для вложенного элемента.</span><span class="sxs-lookup"><span data-stu-id="99922-p142">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="99922-742">String</span><span class="sxs-lookup"><span data-stu-id="99922-742">String</span></span> | | <span data-ttu-id="99922-743">Строка, содержащая имя вложения, длиной до 255 символов.</span><span class="sxs-lookup"><span data-stu-id="99922-743">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="99922-744">String</span><span class="sxs-lookup"><span data-stu-id="99922-744">String</span></span> | | <span data-ttu-id="99922-p143">Используется, только если свойству `type` задано значение `file`. URI расположения файла.</span><span class="sxs-lookup"><span data-stu-id="99922-p143">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.isInline` | <span data-ttu-id="99922-747">Логический</span><span class="sxs-lookup"><span data-stu-id="99922-747">Boolean</span></span> | | <span data-ttu-id="99922-p144">Используется, только если свойству `type` задано значение `file`. Значение `true` указывает на то, что вложение будет встроено в текст сообщения и не должно отображаться в списке вложений.</span><span class="sxs-lookup"><span data-stu-id="99922-p144">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="99922-750">String</span><span class="sxs-lookup"><span data-stu-id="99922-750">String</span></span> | | <span data-ttu-id="99922-p145">Используется, только если свойству `type` задано значение `item`. Идентификатор вложения EWS. Это строка длиной до 100 символов.</span><span class="sxs-lookup"><span data-stu-id="99922-p145">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="99922-754">function</span><span class="sxs-lookup"><span data-stu-id="99922-754">function</span></span> | <span data-ttu-id="99922-755">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="99922-755">&lt;optional&gt;</span></span> | <span data-ttu-id="99922-756">По завершении работы метода функция, переданная параметру `callback`, вызывается с помощью одного параметра `asyncResult`, представляющего собой объект [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="99922-756">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="99922-757">Требования</span><span class="sxs-lookup"><span data-stu-id="99922-757">Requirements</span></span>

|<span data-ttu-id="99922-758">Требование</span><span class="sxs-lookup"><span data-stu-id="99922-758">Requirement</span></span>| <span data-ttu-id="99922-759">Значение</span><span class="sxs-lookup"><span data-stu-id="99922-759">Value</span></span>|
|---|---|
|[<span data-ttu-id="99922-760">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="99922-760">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="99922-761">1.0</span><span class="sxs-lookup"><span data-stu-id="99922-761">1.0</span></span>|
|[<span data-ttu-id="99922-762">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="99922-762">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="99922-763">ReadItem</span><span class="sxs-lookup"><span data-stu-id="99922-763">ReadItem</span></span>|
|[<span data-ttu-id="99922-764">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="99922-764">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="99922-765">Чтение</span><span class="sxs-lookup"><span data-stu-id="99922-765">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="99922-766">Примеры</span><span class="sxs-lookup"><span data-stu-id="99922-766">Examples</span></span>

<span data-ttu-id="99922-767">Приведенный ниже код передает строку в функцию `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="99922-767">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="99922-768">Ответ с пустым текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="99922-768">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="99922-769">Ответ только с текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="99922-769">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="99922-770">Ответ с текстом сообщения и вложенным файлом.</span><span class="sxs-lookup"><span data-stu-id="99922-770">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="99922-771">Ответ с текстом сообщения и вложенным элементом.</span><span class="sxs-lookup"><span data-stu-id="99922-771">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="99922-772">Ответ с текстом сообщения, вложенным файлом, вложенным элементом и обратным вызовом.</span><span class="sxs-lookup"><span data-stu-id="99922-772">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="99922-773">displayReplyForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="99922-773">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="99922-774">Отображает форму ответа, включающую только отправителя выбранного сообщения или организатора выбранной встречи.</span><span class="sxs-lookup"><span data-stu-id="99922-774">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="99922-775">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="99922-775">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="99922-776">В Outlook в Интернете форма ответа отображается в виде всплывающей формы в представлении с 3 либо 1 или 2 столбцами.</span><span class="sxs-lookup"><span data-stu-id="99922-776">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="99922-777">Если любой строковый параметр превышает указанные для него ограничения, `displayReplyForm` возвращает исключение.</span><span class="sxs-lookup"><span data-stu-id="99922-777">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="99922-p146">Если в параметре `formData.attachments` указаны вложения, Outlook в Интернете и классические клиенты пытаются скачать их и вложить в форму ответа. Если какие-либо вложения добавить не удается, в форме отображается сообщение об ошибке. Если сообщения об ошибках не предусмотрены, то они не отображаются.</span><span class="sxs-lookup"><span data-stu-id="99922-p146">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="99922-781">Параметры</span><span class="sxs-lookup"><span data-stu-id="99922-781">Parameters</span></span>

| <span data-ttu-id="99922-782">Имя</span><span class="sxs-lookup"><span data-stu-id="99922-782">Name</span></span> | <span data-ttu-id="99922-783">Тип</span><span class="sxs-lookup"><span data-stu-id="99922-783">Type</span></span> | <span data-ttu-id="99922-784">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="99922-784">Attributes</span></span> | <span data-ttu-id="99922-785">Описание</span><span class="sxs-lookup"><span data-stu-id="99922-785">Description</span></span> |
|---|---|---|---|
|`formData`| <span data-ttu-id="99922-786">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="99922-786">String &#124; Object</span></span>| | <span data-ttu-id="99922-p147">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="99922-p147">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="99922-789">**ИЛИ**</span><span class="sxs-lookup"><span data-stu-id="99922-789">**OR**</span></span><br/><span data-ttu-id="99922-p148">Объект, который содержит текст или данные вложения и функцию обратного вызова. Ниже представлено определение этого объекта.</span><span class="sxs-lookup"><span data-stu-id="99922-p148">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="99922-792">String</span><span class="sxs-lookup"><span data-stu-id="99922-792">String</span></span> | <span data-ttu-id="99922-793">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="99922-793">&lt;optional&gt;</span></span> | <span data-ttu-id="99922-p149">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="99922-p149">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="99922-796">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="99922-796">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="99922-797">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="99922-797">&lt;optional&gt;</span></span> | <span data-ttu-id="99922-798">Массив объектов JSON, представляющих собой вложенные файлы или элементы.</span><span class="sxs-lookup"><span data-stu-id="99922-798">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="99922-799">String</span><span class="sxs-lookup"><span data-stu-id="99922-799">String</span></span> | | <span data-ttu-id="99922-p150">Указывает тип вложения. Допустимые значения: `file` для вложенного файла и `item` для вложенного элемента.</span><span class="sxs-lookup"><span data-stu-id="99922-p150">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="99922-802">String</span><span class="sxs-lookup"><span data-stu-id="99922-802">String</span></span> | | <span data-ttu-id="99922-803">Строка, содержащая имя вложения, длиной до 255 символов.</span><span class="sxs-lookup"><span data-stu-id="99922-803">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="99922-804">String</span><span class="sxs-lookup"><span data-stu-id="99922-804">String</span></span> | | <span data-ttu-id="99922-p151">Используется, только если свойству `type` задано значение `file`. URI расположения файла.</span><span class="sxs-lookup"><span data-stu-id="99922-p151">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.isInline` | <span data-ttu-id="99922-807">Логический</span><span class="sxs-lookup"><span data-stu-id="99922-807">Boolean</span></span> | | <span data-ttu-id="99922-p152">Используется, только если свойству `type` задано значение `file`. Значение `true` указывает на то, что вложение будет встроено в текст сообщения и не должно отображаться в списке вложений.</span><span class="sxs-lookup"><span data-stu-id="99922-p152">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="99922-810">String</span><span class="sxs-lookup"><span data-stu-id="99922-810">String</span></span> | | <span data-ttu-id="99922-p153">Используется, только если свойству `type` задано значение `item`. Идентификатор вложения EWS. Это строка длиной до 100 символов.</span><span class="sxs-lookup"><span data-stu-id="99922-p153">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="99922-814">function</span><span class="sxs-lookup"><span data-stu-id="99922-814">function</span></span> | <span data-ttu-id="99922-815">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="99922-815">&lt;optional&gt;</span></span> | <span data-ttu-id="99922-816">По завершении работы метода функция, переданная параметру `callback`, вызывается с помощью одного параметра `asyncResult`, представляющего собой объект [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="99922-816">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="99922-817">Требования</span><span class="sxs-lookup"><span data-stu-id="99922-817">Requirements</span></span>

|<span data-ttu-id="99922-818">Требование</span><span class="sxs-lookup"><span data-stu-id="99922-818">Requirement</span></span>| <span data-ttu-id="99922-819">Значение</span><span class="sxs-lookup"><span data-stu-id="99922-819">Value</span></span>|
|---|---|
|[<span data-ttu-id="99922-820">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="99922-820">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="99922-821">1.0</span><span class="sxs-lookup"><span data-stu-id="99922-821">1.0</span></span>|
|[<span data-ttu-id="99922-822">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="99922-822">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="99922-823">ReadItem</span><span class="sxs-lookup"><span data-stu-id="99922-823">ReadItem</span></span>|
|[<span data-ttu-id="99922-824">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="99922-824">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="99922-825">Чтение</span><span class="sxs-lookup"><span data-stu-id="99922-825">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="99922-826">Примеры</span><span class="sxs-lookup"><span data-stu-id="99922-826">Examples</span></span>

<span data-ttu-id="99922-827">Приведенный ниже код передает строку в функцию `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="99922-827">The following code passes a string to the `displayReplyForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="99922-828">Ответ с пустым текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="99922-828">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="99922-829">Ответ только с текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="99922-829">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="99922-830">Ответ с текстом сообщения и вложенным файлом.</span><span class="sxs-lookup"><span data-stu-id="99922-830">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="99922-831">Ответ с текстом сообщения и вложенным элементом.</span><span class="sxs-lookup"><span data-stu-id="99922-831">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="99922-832">Ответ с текстом сообщения, вложенным файлом, вложенным элементом и обратным вызовом.</span><span class="sxs-lookup"><span data-stu-id="99922-832">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlookofficeentitiesviewoutlook-js-15"></a><span data-ttu-id="99922-833">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.5)}</span><span class="sxs-lookup"><span data-stu-id="99922-833">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.5)}</span></span>

<span data-ttu-id="99922-834">Получает сущности, обнаруженные в теле выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="99922-834">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="99922-835">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="99922-835">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="99922-836">Требования</span><span class="sxs-lookup"><span data-stu-id="99922-836">Requirements</span></span>

|<span data-ttu-id="99922-837">Требование</span><span class="sxs-lookup"><span data-stu-id="99922-837">Requirement</span></span>| <span data-ttu-id="99922-838">Значение</span><span class="sxs-lookup"><span data-stu-id="99922-838">Value</span></span>|
|---|---|
|[<span data-ttu-id="99922-839">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="99922-839">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="99922-840">1.0</span><span class="sxs-lookup"><span data-stu-id="99922-840">1.0</span></span>|
|[<span data-ttu-id="99922-841">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="99922-841">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="99922-842">ReadItem</span><span class="sxs-lookup"><span data-stu-id="99922-842">ReadItem</span></span>|
|[<span data-ttu-id="99922-843">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="99922-843">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="99922-844">Чтение</span><span class="sxs-lookup"><span data-stu-id="99922-844">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="99922-845">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="99922-845">Returns:</span></span>

<span data-ttu-id="99922-846">Тип: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="99922-846">Type: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.5)</span></span>

##### <a name="example"></a><span data-ttu-id="99922-847">Пример</span><span class="sxs-lookup"><span data-stu-id="99922-847">Example</span></span>

<span data-ttu-id="99922-848">Ниже приведен пример получения доступа к сущностям контактов в теле текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="99922-848">The following example accesses the contacts entities in the current item's body.</span></span>

```js
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

<br>

---
---

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-15meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-15phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-15tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-15"></a><span data-ttu-id="99922-849">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.5)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.5)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.5)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.5))>}</span><span class="sxs-lookup"><span data-stu-id="99922-849">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.5)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.5)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.5)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.5))>}</span></span>

<span data-ttu-id="99922-850">Получает массив всех сущностей указанного типа, обнаруженных в теле выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="99922-850">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="99922-851">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="99922-851">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="99922-852">Параметры</span><span class="sxs-lookup"><span data-stu-id="99922-852">Parameters</span></span>

|<span data-ttu-id="99922-853">Имя</span><span class="sxs-lookup"><span data-stu-id="99922-853">Name</span></span>| <span data-ttu-id="99922-854">Тип</span><span class="sxs-lookup"><span data-stu-id="99922-854">Type</span></span>| <span data-ttu-id="99922-855">Описание</span><span class="sxs-lookup"><span data-stu-id="99922-855">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="99922-856">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="99922-856">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook/office.mailboxenums.entitytype?view=outlook-js-1.5)|<span data-ttu-id="99922-857">Одно из значений перечисления EntityType.</span><span class="sxs-lookup"><span data-stu-id="99922-857">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="99922-858">Требования</span><span class="sxs-lookup"><span data-stu-id="99922-858">Requirements</span></span>

|<span data-ttu-id="99922-859">Требование</span><span class="sxs-lookup"><span data-stu-id="99922-859">Requirement</span></span>| <span data-ttu-id="99922-860">Значение</span><span class="sxs-lookup"><span data-stu-id="99922-860">Value</span></span>|
|---|---|
|[<span data-ttu-id="99922-861">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="99922-861">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="99922-862">1.0</span><span class="sxs-lookup"><span data-stu-id="99922-862">1.0</span></span>|
|[<span data-ttu-id="99922-863">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="99922-863">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="99922-864">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="99922-864">Restricted</span></span>|
|[<span data-ttu-id="99922-865">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="99922-865">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="99922-866">Чтение</span><span class="sxs-lookup"><span data-stu-id="99922-866">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="99922-867">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="99922-867">Returns:</span></span>

<span data-ttu-id="99922-868">Если значение, переданное в `entityType`, не является допустимым членом перечисления `EntityType`, метод возвращает значение NULL.</span><span class="sxs-lookup"><span data-stu-id="99922-868">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="99922-869">Если в теле элемента отсутствуют сущности указанного типа, метод возвращает пустой массив.</span><span class="sxs-lookup"><span data-stu-id="99922-869">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="99922-870">В противном случае тип объектов в возвращаемом массиве зависит от типа сущности, запрошенной в параметре `entityType`.</span><span class="sxs-lookup"><span data-stu-id="99922-870">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="99922-871">Хотя минимальный уровень разрешений для использования этого метода — **Restricted**, для некоторых типов сущностей требуется доступ на уровне **ReadItem**, как указано в приведенной ниже таблице.</span><span class="sxs-lookup"><span data-stu-id="99922-871">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="99922-872">Значение параметра `entityType`</span><span class="sxs-lookup"><span data-stu-id="99922-872">Value of `entityType`</span></span> | <span data-ttu-id="99922-873">Тип объектов в возвращаемом массиве</span><span class="sxs-lookup"><span data-stu-id="99922-873">Type of objects in returned array</span></span> | <span data-ttu-id="99922-874">Необходимый уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="99922-874">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="99922-875">String</span><span class="sxs-lookup"><span data-stu-id="99922-875">String</span></span> | <span data-ttu-id="99922-876">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="99922-876">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="99922-877">Contact</span><span class="sxs-lookup"><span data-stu-id="99922-877">Contact</span></span> | <span data-ttu-id="99922-878">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="99922-878">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="99922-879">String</span><span class="sxs-lookup"><span data-stu-id="99922-879">String</span></span> | <span data-ttu-id="99922-880">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="99922-880">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="99922-881">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="99922-881">MeetingSuggestion</span></span> | <span data-ttu-id="99922-882">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="99922-882">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="99922-883">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="99922-883">PhoneNumber</span></span> | <span data-ttu-id="99922-884">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="99922-884">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="99922-885">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="99922-885">TaskSuggestion</span></span> | <span data-ttu-id="99922-886">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="99922-886">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="99922-887">String</span><span class="sxs-lookup"><span data-stu-id="99922-887">String</span></span> | <span data-ttu-id="99922-888">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="99922-888">**Restricted**</span></span> |

<span data-ttu-id="99922-889">Тип: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.5)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.5)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.5)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.5))></span><span class="sxs-lookup"><span data-stu-id="99922-889">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.5)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.5)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.5)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.5))></span></span>

##### <a name="example"></a><span data-ttu-id="99922-890">Пример</span><span class="sxs-lookup"><span data-stu-id="99922-890">Example</span></span>

<span data-ttu-id="99922-891">В примере ниже показано, как получить доступ к массиву строк, которые представляют собой почтовые адреса в теле текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="99922-891">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-15meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-15phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-15tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-15"></a><span data-ttu-id="99922-892">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.5)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.5)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.5)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.5))>}</span><span class="sxs-lookup"><span data-stu-id="99922-892">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.5)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.5)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.5)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.5))>}</span></span>

<span data-ttu-id="99922-893">Возвращает известные сущности в выбранном элементе, которые проходят через именованный фильтр, определяемый в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="99922-893">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="99922-894">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="99922-894">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="99922-895">Метод `getFilteredEntitiesByName` возвращает сущности, соответствующие регулярному выражению, которое определяется в элементе правила [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) в XML-файле манифеста, с использованием указанного значения элемента `FilterName`.</span><span class="sxs-lookup"><span data-stu-id="99922-895">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="99922-896">Параметры</span><span class="sxs-lookup"><span data-stu-id="99922-896">Parameters</span></span>

|<span data-ttu-id="99922-897">Имя</span><span class="sxs-lookup"><span data-stu-id="99922-897">Name</span></span>| <span data-ttu-id="99922-898">Тип</span><span class="sxs-lookup"><span data-stu-id="99922-898">Type</span></span>| <span data-ttu-id="99922-899">Описание</span><span class="sxs-lookup"><span data-stu-id="99922-899">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="99922-900">String</span><span class="sxs-lookup"><span data-stu-id="99922-900">String</span></span>|<span data-ttu-id="99922-901">Имя элемента правила `ItemHasKnownEntity`, определяющее соответствующий фильтр.</span><span class="sxs-lookup"><span data-stu-id="99922-901">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="99922-902">Требования</span><span class="sxs-lookup"><span data-stu-id="99922-902">Requirements</span></span>

|<span data-ttu-id="99922-903">Требование</span><span class="sxs-lookup"><span data-stu-id="99922-903">Requirement</span></span>| <span data-ttu-id="99922-904">Значение</span><span class="sxs-lookup"><span data-stu-id="99922-904">Value</span></span>|
|---|---|
|[<span data-ttu-id="99922-905">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="99922-905">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="99922-906">1.0</span><span class="sxs-lookup"><span data-stu-id="99922-906">1.0</span></span>|
|[<span data-ttu-id="99922-907">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="99922-907">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="99922-908">ReadItem</span><span class="sxs-lookup"><span data-stu-id="99922-908">ReadItem</span></span>|
|[<span data-ttu-id="99922-909">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="99922-909">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="99922-910">Чтение</span><span class="sxs-lookup"><span data-stu-id="99922-910">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="99922-911">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="99922-911">Returns:</span></span>

<span data-ttu-id="99922-p155">Если в манифесте нет элемента `ItemHasKnownEntity` со значением `FilterName`, соответствующим параметру `name`, метод возвращает `null`. Если параметр `name` соответствует элементу `ItemHasKnownEntity` в манифесте, но при этом в текущем элементе нет соответствующих сущностей, метод возвращает пустой массив.</span><span class="sxs-lookup"><span data-stu-id="99922-p155">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="99922-914">Тип: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.5)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.5)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.5)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.5))></span><span class="sxs-lookup"><span data-stu-id="99922-914">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.5)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.5)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.5)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.5))></span></span>

<br>

---
---

#### <a name="getregexmatches--object"></a><span data-ttu-id="99922-915">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="99922-915">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="99922-916">Возвращает строковые значения в выбранном элементе, которые соответствуют регулярным выражениям, определенным в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="99922-916">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="99922-917">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="99922-917">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="99922-p156">Метод `getRegExMatches` возвращает строки, соответствующие регулярному выражению, которое определяется в каждом элементе правила `ItemHasRegularExpressionMatch` или `ItemHasKnownEntity` в XML-файле манифеста. Для правила `ItemHasRegularExpressionMatch` соответствующую строку должно содержать свойство элемента, указанного этим правилом. Простой тип `PropertyName` определяет поддерживаемые свойства.</span><span class="sxs-lookup"><span data-stu-id="99922-p156">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="99922-921">Например, рассмотрим манифест надстройки, который содержит указанный ниже элемент `Rule`.</span><span class="sxs-lookup"><span data-stu-id="99922-921">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="99922-922">Объект, возвращаемый методом `getRegExMatches`, будет содержать два свойства: `fruits` и `veggies`.</span><span class="sxs-lookup"><span data-stu-id="99922-922">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="99922-p157">Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты. Лучше используйте метод [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.5#getasync-coerciontype--options--callback-) для этого.</span><span class="sxs-lookup"><span data-stu-id="99922-p157">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.5#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="99922-926">Requirements</span><span class="sxs-lookup"><span data-stu-id="99922-926">Requirements</span></span>

|<span data-ttu-id="99922-927">Требование</span><span class="sxs-lookup"><span data-stu-id="99922-927">Requirement</span></span>| <span data-ttu-id="99922-928">Значение</span><span class="sxs-lookup"><span data-stu-id="99922-928">Value</span></span>|
|---|---|
|[<span data-ttu-id="99922-929">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="99922-929">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="99922-930">1.0</span><span class="sxs-lookup"><span data-stu-id="99922-930">1.0</span></span>|
|[<span data-ttu-id="99922-931">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="99922-931">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="99922-932">ReadItem</span><span class="sxs-lookup"><span data-stu-id="99922-932">ReadItem</span></span>|
|[<span data-ttu-id="99922-933">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="99922-933">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="99922-934">Чтение</span><span class="sxs-lookup"><span data-stu-id="99922-934">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="99922-935">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="99922-935">Returns:</span></span>

<span data-ttu-id="99922-p158">Объект, содержащий массив строк, которые соответствуют регулярным выражениям, определяемым в XML-файле манифеста. Имя каждого массива равно соответствующему значению атрибута `RegExName` подходящего правила `ItemHasRegularExpressionMatch` или атрибута `FilterName` соответствующего правила `ItemHasKnownEntity`.</span><span class="sxs-lookup"><span data-stu-id="99922-p158">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<span data-ttu-id="99922-938">Тип: Object</span><span class="sxs-lookup"><span data-stu-id="99922-938">Type:  object</span></span>

##### <a name="example"></a><span data-ttu-id="99922-939">Пример</span><span class="sxs-lookup"><span data-stu-id="99922-939">Example</span></span>

<span data-ttu-id="99922-940">В примере ниже показано, как получить доступ к массиву совпадений для <rule>элементов регулярного выражения `fruits` и `veggies`, которые указаны в манифесте</rule>.</span><span class="sxs-lookup"><span data-stu-id="99922-940">The following example shows how to access the array of matches for the regular expression <rule>elements `fruits` and `veggies`, which are specified in the manifest.</rule></span></span>

```js
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

<br>

---
---

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="99922-941">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="99922-941">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="99922-942">Возвращает строковые значения в выбранном элементе, которые соответствуют именованному регулярному выражению, определенному в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="99922-942">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="99922-943">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="99922-943">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="99922-944">Метод `getRegExMatchesByName` возвращает строки, соответствующие регулярному выражению, которое определяется в элементе правила `ItemHasRegularExpressionMatch` в XML-файле манифеста, с использованием указанного значения элемента `RegExName`.</span><span class="sxs-lookup"><span data-stu-id="99922-944">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="99922-p159">Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты.</span><span class="sxs-lookup"><span data-stu-id="99922-p159">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="99922-947">Параметры</span><span class="sxs-lookup"><span data-stu-id="99922-947">Parameters</span></span>

|<span data-ttu-id="99922-948">Имя</span><span class="sxs-lookup"><span data-stu-id="99922-948">Name</span></span>| <span data-ttu-id="99922-949">Тип</span><span class="sxs-lookup"><span data-stu-id="99922-949">Type</span></span>| <span data-ttu-id="99922-950">Описание</span><span class="sxs-lookup"><span data-stu-id="99922-950">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="99922-951">String</span><span class="sxs-lookup"><span data-stu-id="99922-951">String</span></span>|<span data-ttu-id="99922-952">Имя элемента правила `ItemHasRegularExpressionMatch`, определяющее соответствующий фильтр.</span><span class="sxs-lookup"><span data-stu-id="99922-952">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="99922-953">Требования</span><span class="sxs-lookup"><span data-stu-id="99922-953">Requirements</span></span>

|<span data-ttu-id="99922-954">Требование</span><span class="sxs-lookup"><span data-stu-id="99922-954">Requirement</span></span>| <span data-ttu-id="99922-955">Значение</span><span class="sxs-lookup"><span data-stu-id="99922-955">Value</span></span>|
|---|---|
|[<span data-ttu-id="99922-956">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="99922-956">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="99922-957">1.0</span><span class="sxs-lookup"><span data-stu-id="99922-957">1.0</span></span>|
|[<span data-ttu-id="99922-958">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="99922-958">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="99922-959">ReadItem</span><span class="sxs-lookup"><span data-stu-id="99922-959">ReadItem</span></span>|
|[<span data-ttu-id="99922-960">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="99922-960">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="99922-961">Чтение</span><span class="sxs-lookup"><span data-stu-id="99922-961">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="99922-962">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="99922-962">Returns:</span></span>

<span data-ttu-id="99922-963">Массив строк, соответствующих регулярному выражению, определяемому в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="99922-963">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<span data-ttu-id="99922-964">Тип: Array.< String ></span><span class="sxs-lookup"><span data-stu-id="99922-964">Array.< String ></span></span>

##### <a name="example"></a><span data-ttu-id="99922-965">Пример</span><span class="sxs-lookup"><span data-stu-id="99922-965">Example</span></span>

```js
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

<br>

---
---

#### <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="99922-966">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="99922-966">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="99922-967">Асинхронно возвращает данные, выбранные в теме или тексте сообщения.</span><span class="sxs-lookup"><span data-stu-id="99922-967">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="99922-p160">Если выделенный фрагмент отсутствует, но курсор находится в тексте или теме, метод возвращает значение NULL для выбранных данных. Если выбраны не текст и не тема, метод возвращает ошибку `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="99922-p160">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="99922-970">Параметры</span><span class="sxs-lookup"><span data-stu-id="99922-970">Parameters</span></span>

|<span data-ttu-id="99922-971">Имя</span><span class="sxs-lookup"><span data-stu-id="99922-971">Name</span></span>| <span data-ttu-id="99922-972">Тип</span><span class="sxs-lookup"><span data-stu-id="99922-972">Type</span></span>| <span data-ttu-id="99922-973">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="99922-973">Attributes</span></span>| <span data-ttu-id="99922-974">Описание</span><span class="sxs-lookup"><span data-stu-id="99922-974">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="99922-975">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="99922-975">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="99922-p161">Запрашивает формат данных. Если задано значение Text, метод возвращает обычный текст как строку, удаляя все имеющиеся HTML-теги. Если задано значение HTML, метод возвращает выделенный текст (обычный текст или HTML).</span><span class="sxs-lookup"><span data-stu-id="99922-p161">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="99922-979">Object</span><span class="sxs-lookup"><span data-stu-id="99922-979">Object</span></span>| <span data-ttu-id="99922-980">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="99922-980">&lt;optional&gt;</span></span>|<span data-ttu-id="99922-981">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="99922-981">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="99922-982">Object</span><span class="sxs-lookup"><span data-stu-id="99922-982">Object</span></span>| <span data-ttu-id="99922-983">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="99922-983">&lt;optional&gt;</span></span>|<span data-ttu-id="99922-984">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="99922-984">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="99922-985">функция</span><span class="sxs-lookup"><span data-stu-id="99922-985">function</span></span>||<span data-ttu-id="99922-986">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="99922-986">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="99922-987">Чтобы получить доступ к выбранным данным из метода обратного вызова, вызовите `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="99922-987">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="99922-988">Для доступа к исходному свойству, представляющему собой источник выбранных данных, вызовите параметр `asyncResult.value.sourceProperty`, который может иметь значение `body` или `subject`.</span><span class="sxs-lookup"><span data-stu-id="99922-988">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="99922-989">Требования</span><span class="sxs-lookup"><span data-stu-id="99922-989">Requirements</span></span>

|<span data-ttu-id="99922-990">Требование</span><span class="sxs-lookup"><span data-stu-id="99922-990">Requirement</span></span>| <span data-ttu-id="99922-991">Значение</span><span class="sxs-lookup"><span data-stu-id="99922-991">Value</span></span>|
|---|---|
|[<span data-ttu-id="99922-992">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="99922-992">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="99922-993">1.2</span><span class="sxs-lookup"><span data-stu-id="99922-993">1.2</span></span>|
|[<span data-ttu-id="99922-994">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="99922-994">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="99922-995">ReadItem</span><span class="sxs-lookup"><span data-stu-id="99922-995">ReadItem</span></span>|
|[<span data-ttu-id="99922-996">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="99922-996">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="99922-997">Создание</span><span class="sxs-lookup"><span data-stu-id="99922-997">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="99922-998">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="99922-998">Returns:</span></span>

<span data-ttu-id="99922-999">Выбранные данные в виде строки с форматом, определенным в параметре `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="99922-999">The selected data as a string with format determined by `coercionType`.</span></span>

<span data-ttu-id="99922-1000">Тип: строка</span><span class="sxs-lookup"><span data-stu-id="99922-1000">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="99922-1001">Пример</span><span class="sxs-lookup"><span data-stu-id="99922-1001">Example</span></span>

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

#### <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="99922-1002">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="99922-1002">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="99922-1003">Асинхронно загружает настраиваемые свойства для надстройки для выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="99922-1003">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="99922-p163">Настраиваемые свойства сохраняются в виде пар "ключ-значение" для каждого приложения и каждого элемента. Этот метод возвращает объект `CustomProperties` при обратном вызове, который предоставляет методы для доступа к настраиваемым свойствам, характерным для текущего элемента и текущей надстройки. Настраиваемые свойства не шифруются для элемента, поэтому этот способ хранения не является безопасным.</span><span class="sxs-lookup"><span data-stu-id="99922-p163">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="99922-1007">Параметры</span><span class="sxs-lookup"><span data-stu-id="99922-1007">Parameters</span></span>

|<span data-ttu-id="99922-1008">Имя</span><span class="sxs-lookup"><span data-stu-id="99922-1008">Name</span></span>| <span data-ttu-id="99922-1009">Тип</span><span class="sxs-lookup"><span data-stu-id="99922-1009">Type</span></span>| <span data-ttu-id="99922-1010">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="99922-1010">Attributes</span></span>| <span data-ttu-id="99922-1011">Описание</span><span class="sxs-lookup"><span data-stu-id="99922-1011">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="99922-1012">function</span><span class="sxs-lookup"><span data-stu-id="99922-1012">function</span></span>||<span data-ttu-id="99922-1013">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="99922-1013">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="99922-1014">Настраиваемые свойства предоставляются в виде объекта [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.5) в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="99922-1014">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.5) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="99922-1015">Этот объект позволяет получить, задать и удалить настраиваемые свойства для элемента, а также сохранить изменения, внесенные в набор настраиваемых свойств, на сервере.</span><span class="sxs-lookup"><span data-stu-id="99922-1015">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="99922-1016">Объект</span><span class="sxs-lookup"><span data-stu-id="99922-1016">Object</span></span>| <span data-ttu-id="99922-1017">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="99922-1017">&lt;optional&gt;</span></span>|<span data-ttu-id="99922-1018">Разработчики могут указать любой объект, к которому необходимо получить доступ, в функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="99922-1018">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="99922-1019">Доступ к этому объекту можно получить с помощью свойства `asyncResult.asyncContext` в функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="99922-1019">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="99922-1020">Требования</span><span class="sxs-lookup"><span data-stu-id="99922-1020">Requirements</span></span>

|<span data-ttu-id="99922-1021">Требование</span><span class="sxs-lookup"><span data-stu-id="99922-1021">Requirement</span></span>| <span data-ttu-id="99922-1022">Значение</span><span class="sxs-lookup"><span data-stu-id="99922-1022">Value</span></span>|
|---|---|
|[<span data-ttu-id="99922-1023">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="99922-1023">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="99922-1024">1.0</span><span class="sxs-lookup"><span data-stu-id="99922-1024">1.0</span></span>|
|[<span data-ttu-id="99922-1025">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="99922-1025">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="99922-1026">ReadItem</span><span class="sxs-lookup"><span data-stu-id="99922-1026">ReadItem</span></span>|
|[<span data-ttu-id="99922-1027">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="99922-1027">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="99922-1028">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="99922-1028">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="99922-1029">Пример</span><span class="sxs-lookup"><span data-stu-id="99922-1029">Example</span></span>

<span data-ttu-id="99922-p166">Приведенный ниже пример кода показывает, как асинхронно загружать настраиваемые свойства, характерные для текущего элемента, с помощью метода `loadCustomPropertiesAsync`. Этот пример также показывает, как сохранять эти свойства на сервере с помощью метода `CustomProperties.saveAsync`. После загрузки настраиваемых свойств в этом примере кода метод `CustomProperties.get` используется для считывания настраиваемого свойства `myProp`, метод `CustomProperties.set` — для записи настраиваемого свойства `otherProp`, а метод `saveAsync` — для сохранения настраиваемых свойств.</span><span class="sxs-lookup"><span data-stu-id="99922-p166">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

#### <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="99922-1033">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="99922-1033">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="99922-1034">Удаляет вложение из сообщения или встречи.</span><span class="sxs-lookup"><span data-stu-id="99922-1034">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="99922-1035">Метод `removeAttachmentAsync` удаляет из элемента вложение с указанным идентификатором.</span><span class="sxs-lookup"><span data-stu-id="99922-1035">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="99922-1036">Идентификатор вложения рекомендуется использовать для удаления вложения, только если оно добавлено тем же почтовым приложением в ходе текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="99922-1036">As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session.</span></span> <span data-ttu-id="99922-1037">В Outlook в Интернете и на мобильных устройствах идентификатор вложения действителен только в течение одного сеанса.</span><span class="sxs-lookup"><span data-stu-id="99922-1037">In Outlook on the web and OWA for Devices, the attachment ID is valid only within the same session.</span></span> <span data-ttu-id="99922-1038">Сеанс завершается, когда пользователь закрывает приложение или начинает создавать элемент во встроенной форме, а затем переходит из формы в отдельное окно.</span><span class="sxs-lookup"><span data-stu-id="99922-1038">A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="99922-1039">Параметры</span><span class="sxs-lookup"><span data-stu-id="99922-1039">Parameters</span></span>

|<span data-ttu-id="99922-1040">Имя</span><span class="sxs-lookup"><span data-stu-id="99922-1040">Name</span></span>| <span data-ttu-id="99922-1041">Тип</span><span class="sxs-lookup"><span data-stu-id="99922-1041">Type</span></span>| <span data-ttu-id="99922-1042">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="99922-1042">Attributes</span></span>| <span data-ttu-id="99922-1043">Описание</span><span class="sxs-lookup"><span data-stu-id="99922-1043">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="99922-1044">String</span><span class="sxs-lookup"><span data-stu-id="99922-1044">String</span></span>||<span data-ttu-id="99922-1045">Идентификатор удаляемого вложения.</span><span class="sxs-lookup"><span data-stu-id="99922-1045">The identifier of the attachment to remove.</span></span>|
|`options`| <span data-ttu-id="99922-1046">Object</span><span class="sxs-lookup"><span data-stu-id="99922-1046">Object</span></span>| <span data-ttu-id="99922-1047">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="99922-1047">&lt;optional&gt;</span></span>|<span data-ttu-id="99922-1048">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="99922-1048">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="99922-1049">Object</span><span class="sxs-lookup"><span data-stu-id="99922-1049">Object</span></span>| <span data-ttu-id="99922-1050">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="99922-1050">&lt;optional&gt;</span></span>|<span data-ttu-id="99922-1051">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="99922-1051">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="99922-1052">функция</span><span class="sxs-lookup"><span data-stu-id="99922-1052">function</span></span>| <span data-ttu-id="99922-1053">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="99922-1053">&lt;optional&gt;</span></span>|<span data-ttu-id="99922-1054">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="99922-1054">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="99922-1055">Если удалить вложение не удается, свойство `asyncResult.error` содержит код ошибки с указанием ее причины.</span><span class="sxs-lookup"><span data-stu-id="99922-1055">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="99922-1056">Ошибки</span><span class="sxs-lookup"><span data-stu-id="99922-1056">Errors</span></span>

| <span data-ttu-id="99922-1057">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="99922-1057">Error code</span></span> | <span data-ttu-id="99922-1058">Описание</span><span class="sxs-lookup"><span data-stu-id="99922-1058">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="99922-1059">Идентификатор вложения не существует.</span><span class="sxs-lookup"><span data-stu-id="99922-1059">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="99922-1060">Требования</span><span class="sxs-lookup"><span data-stu-id="99922-1060">Requirements</span></span>

|<span data-ttu-id="99922-1061">Требование</span><span class="sxs-lookup"><span data-stu-id="99922-1061">Requirement</span></span>| <span data-ttu-id="99922-1062">Значение</span><span class="sxs-lookup"><span data-stu-id="99922-1062">Value</span></span>|
|---|---|
|[<span data-ttu-id="99922-1063">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="99922-1063">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="99922-1064">1.1</span><span class="sxs-lookup"><span data-stu-id="99922-1064">1.1</span></span>|
|[<span data-ttu-id="99922-1065">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="99922-1065">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="99922-1066">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="99922-1066">ReadWriteItem</span></span>|
|[<span data-ttu-id="99922-1067">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="99922-1067">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="99922-1068">Создание</span><span class="sxs-lookup"><span data-stu-id="99922-1068">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="99922-1069">Пример</span><span class="sxs-lookup"><span data-stu-id="99922-1069">Example</span></span>

<span data-ttu-id="99922-1070">Указанный ниже код удаляет вложение с идентификатором "0".</span><span class="sxs-lookup"><span data-stu-id="99922-1070">The following code removes an attachment with an identifier of '0'.</span></span>

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

#### <a name="saveasyncoptions-callback"></a><span data-ttu-id="99922-1071">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="99922-1071">saveAsync([options], callback)</span></span>

<span data-ttu-id="99922-1072">Асинхронно сохраняет элемент.</span><span class="sxs-lookup"><span data-stu-id="99922-1072">Asynchronously saves an item.</span></span>

<span data-ttu-id="99922-1073">При вызове этот метод сохраняет текущее сообщение в виде черновика и возвращает идентификатор элемента с помощью метода обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="99922-1073">When invoked, this method saves the current message as a draft and returns the item id via the callback method.</span></span> <span data-ttu-id="99922-1074">В Outlook в Интернете или интерактивном режиме Outlook этот элемент сохраняется на сервере.</span><span class="sxs-lookup"><span data-stu-id="99922-1074">In Outlook Web App or Outlook in online mode, the item is saved to the server.</span></span> <span data-ttu-id="99922-1075">В Outlook в режиме кэширования этот элемент сохраняется в локальном кэше.</span><span class="sxs-lookup"><span data-stu-id="99922-1075">In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="99922-1076">Если в режиме создания надстройка вызывает для элемента метод `saveAsync`, чтобы получить параметр `itemId` для использования с EWS или REST API, необходимо помнить, что синхронизация элемента с сервером может занять много времени, если Outlook работает в режиме кэширования данных.</span><span class="sxs-lookup"><span data-stu-id="99922-1076">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="99922-1077">До окончания синхронизации элемента применение параметра `itemId` будет приводить к ошибке.</span><span class="sxs-lookup"><span data-stu-id="99922-1077">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="99922-p170">Если метод `saveAsync` вызывается для встречи в режиме создания, она сохраняется как обычная встреча в календаре пользователя, а не как черновик. При сохранении новой встречи приглашения не отправляются. При сохранении существующей встречи уведомления отправляются добавленным или удаленным участникам.</span><span class="sxs-lookup"><span data-stu-id="99922-p170">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="99922-1081">Следующие клиенты отличаются другим поведением `saveAsync` в отношении встреч в режиме создания:</span><span class="sxs-lookup"><span data-stu-id="99922-1081">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="99922-1082">Outlook для Mac не поддерживает сохранение собрания.</span><span class="sxs-lookup"><span data-stu-id="99922-1082">Outlook on Mac does not support saving a meeting.</span></span> <span data-ttu-id="99922-1083">Метод `saveAsync` не работает при вызове из собрания в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="99922-1083">The `saveAsync` method fails when called from a meeting in compose mode.</span></span> <span data-ttu-id="99922-1084">Временное решение представлено в статье [Не удается сохранить встречу как черновик в Outlook для Mac с помощью API JS для Office](https://support.microsoft.com/help/4505745).</span><span class="sxs-lookup"><span data-stu-id="99922-1084">See [Cannot save a meeting as a draft in Outlook for Mac by using Office JS API](https://support.microsoft.com/help/4505745) for a workaround.</span></span>
> - <span data-ttu-id="99922-1085">Outlook в Интернете всегда отправляет приглашение или обновление при вызове `saveAsync` для встречи в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="99922-1085">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="99922-1086">Параметры</span><span class="sxs-lookup"><span data-stu-id="99922-1086">Parameters</span></span>

|<span data-ttu-id="99922-1087">Имя</span><span class="sxs-lookup"><span data-stu-id="99922-1087">Name</span></span>| <span data-ttu-id="99922-1088">Тип</span><span class="sxs-lookup"><span data-stu-id="99922-1088">Type</span></span>| <span data-ttu-id="99922-1089">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="99922-1089">Attributes</span></span>| <span data-ttu-id="99922-1090">Описание</span><span class="sxs-lookup"><span data-stu-id="99922-1090">Description</span></span>|
|---|---|---|---|
|`options`| <span data-ttu-id="99922-1091">Object</span><span class="sxs-lookup"><span data-stu-id="99922-1091">Object</span></span>| <span data-ttu-id="99922-1092">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="99922-1092">&lt;optional&gt;</span></span>|<span data-ttu-id="99922-1093">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="99922-1093">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="99922-1094">Object</span><span class="sxs-lookup"><span data-stu-id="99922-1094">Object</span></span>| <span data-ttu-id="99922-1095">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="99922-1095">&lt;optional&gt;</span></span>|<span data-ttu-id="99922-1096">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="99922-1096">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="99922-1097">функция</span><span class="sxs-lookup"><span data-stu-id="99922-1097">function</span></span>||<span data-ttu-id="99922-1098">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="99922-1098">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="99922-1099">После успешного выполнения идентификатор элемента будет предоставлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="99922-1099">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="99922-1100">Требования</span><span class="sxs-lookup"><span data-stu-id="99922-1100">Requirements</span></span>

|<span data-ttu-id="99922-1101">Требование</span><span class="sxs-lookup"><span data-stu-id="99922-1101">Requirement</span></span>| <span data-ttu-id="99922-1102">Значение</span><span class="sxs-lookup"><span data-stu-id="99922-1102">Value</span></span>|
|---|---|
|[<span data-ttu-id="99922-1103">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="99922-1103">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="99922-1104">1.3</span><span class="sxs-lookup"><span data-stu-id="99922-1104">1.3</span></span>|
|[<span data-ttu-id="99922-1105">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="99922-1105">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="99922-1106">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="99922-1106">ReadWriteItem</span></span>|
|[<span data-ttu-id="99922-1107">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="99922-1107">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="99922-1108">Создание</span><span class="sxs-lookup"><span data-stu-id="99922-1108">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="99922-1109">Примеры</span><span class="sxs-lookup"><span data-stu-id="99922-1109">Examples</span></span>

```js
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result.
  });
```

<span data-ttu-id="99922-p172">Ниже приведен пример параметра `result`, переданного функции обратного вызова. Свойство `value` содержит идентификатор элемента.</span><span class="sxs-lookup"><span data-stu-id="99922-p172">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

<br>

---
---

#### <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="99922-1112">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="99922-1112">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="99922-1113">Асинхронно вставляет данные в текст или тему сообщения.</span><span class="sxs-lookup"><span data-stu-id="99922-1113">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="99922-p173">Метод `setSelectedDataAsync` вставляет указанную строку в местоположение курсора в теме или тексте элемента либо, если текст выделен в редакторе, он заменяет выделенный текст. Если курсор находится вне текста или темы элемента, возвращается ошибка. После вставки курсор помещается в конец вставленного содержимого.</span><span class="sxs-lookup"><span data-stu-id="99922-p173">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="99922-1117">Параметры</span><span class="sxs-lookup"><span data-stu-id="99922-1117">Parameters</span></span>

|<span data-ttu-id="99922-1118">Имя</span><span class="sxs-lookup"><span data-stu-id="99922-1118">Name</span></span>| <span data-ttu-id="99922-1119">Тип</span><span class="sxs-lookup"><span data-stu-id="99922-1119">Type</span></span>| <span data-ttu-id="99922-1120">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="99922-1120">Attributes</span></span>| <span data-ttu-id="99922-1121">Описание</span><span class="sxs-lookup"><span data-stu-id="99922-1121">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="99922-1122">String</span><span class="sxs-lookup"><span data-stu-id="99922-1122">String</span></span>||<span data-ttu-id="99922-p174">Вставляемые данные. Объем данных не должен превышать 1 000 000 символов. Если передано больше 1 000 000 символов, возвращается исключение `ArgumentOutOfRange`.</span><span class="sxs-lookup"><span data-stu-id="99922-p174">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="99922-1126">Object</span><span class="sxs-lookup"><span data-stu-id="99922-1126">Object</span></span>| <span data-ttu-id="99922-1127">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="99922-1127">&lt;optional&gt;</span></span>|<span data-ttu-id="99922-1128">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="99922-1128">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="99922-1129">Object</span><span class="sxs-lookup"><span data-stu-id="99922-1129">Object</span></span>| <span data-ttu-id="99922-1130">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="99922-1130">&lt;optional&gt;</span></span>|<span data-ttu-id="99922-1131">В методе обратного вызова разработчики могут указать любой объект, к которому необходимо получить доступ.</span><span class="sxs-lookup"><span data-stu-id="99922-1131">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="99922-1132">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="99922-1132">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="99922-1133">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="99922-1133">&lt;optional&gt;</span></span>|<span data-ttu-id="99922-1134">Если задано значение `text`, текущий стиль применяется в Outlook в Интернете и классических клиентах.</span><span class="sxs-lookup"><span data-stu-id="99922-1134">If `text`, the current style is applied in Outlook on the web and desktop clients.</span></span> <span data-ttu-id="99922-1135">Если поле представляет собой редактор HTML, вставляются только текстовые данные, даже если они имеют формат HTML.</span><span class="sxs-lookup"><span data-stu-id="99922-1135">If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="99922-1136">Если задано значение `html` и поле (не тема) поддерживает HTML, в Outlook в Интернете применяется текущий стиль, а в классических клиентах Outlook — стиль по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="99922-1136">If html and the field supports HTML (the subject doesn't), the current style is applied in olwebshort and the default style is applied in Outlook.</span></span> <span data-ttu-id="99922-1137">Если поле является текстовым, возвращается ошибка `InvalidDataFormat`.</span><span class="sxs-lookup"><span data-stu-id="99922-1137">If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="99922-1138">Если свойство `coercionType` не задано, результат зависит от поля: если поле имеет формат HTML, используется текст в формате HTML, а если поле текстовое, применяется обычный текст.</span><span class="sxs-lookup"><span data-stu-id="99922-1138">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="99922-1139">функция</span><span class="sxs-lookup"><span data-stu-id="99922-1139">function</span></span>||<span data-ttu-id="99922-1140">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="99922-1140">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="99922-1141">Требования</span><span class="sxs-lookup"><span data-stu-id="99922-1141">Requirements</span></span>

|<span data-ttu-id="99922-1142">Требование</span><span class="sxs-lookup"><span data-stu-id="99922-1142">Requirement</span></span>| <span data-ttu-id="99922-1143">Значение</span><span class="sxs-lookup"><span data-stu-id="99922-1143">Value</span></span>|
|---|---|
|[<span data-ttu-id="99922-1144">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="99922-1144">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="99922-1145">1.2</span><span class="sxs-lookup"><span data-stu-id="99922-1145">1.2</span></span>|
|[<span data-ttu-id="99922-1146">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="99922-1146">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="99922-1147">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="99922-1147">ReadWriteItem</span></span>|
|[<span data-ttu-id="99922-1148">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="99922-1148">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="99922-1149">Создание</span><span class="sxs-lookup"><span data-stu-id="99922-1149">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="99922-1150">Пример</span><span class="sxs-lookup"><span data-stu-id="99922-1150">Example</span></span>

```js
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
