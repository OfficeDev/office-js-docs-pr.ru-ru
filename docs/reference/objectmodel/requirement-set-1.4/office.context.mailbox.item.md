---
title: Office. Context. Mailbox. Item — набор требований 1,4
description: ''
ms.date: 09/23/2019
localization_priority: Normal
ms.openlocfilehash: 4a8a97403c43e4af4ee5d840d7a4fd843d5bd398
ms.sourcegitcommit: 3c84fe6302341668c3f9f6dd64e636a97d03023c
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/26/2019
ms.locfileid: "37167342"
---
# <a name="item"></a><span data-ttu-id="7facb-102">item</span><span class="sxs-lookup"><span data-stu-id="7facb-102">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="7facb-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="7facb-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="7facb-p101">Пространство имен `item` используется для доступа к выбранному в данный момент сообщению, приглашению на собрание или описанию встречи. Вы можете определить тип пространства имен `item` с помощью свойства [itemType](#itemtype-officemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="7facb-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="7facb-106">Requirements</span><span class="sxs-lookup"><span data-stu-id="7facb-106">Requirements</span></span>

|<span data-ttu-id="7facb-107">Требование</span><span class="sxs-lookup"><span data-stu-id="7facb-107">Requirement</span></span>| <span data-ttu-id="7facb-108">Значение</span><span class="sxs-lookup"><span data-stu-id="7facb-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="7facb-109">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="7facb-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7facb-110">1.0</span><span class="sxs-lookup"><span data-stu-id="7facb-110">1.0</span></span>|
|[<span data-ttu-id="7facb-111">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="7facb-111">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7facb-112">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="7facb-112">Restricted</span></span>|
|[<span data-ttu-id="7facb-113">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="7facb-113">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7facb-114">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="7facb-114">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="7facb-115">Элементы и методы</span><span class="sxs-lookup"><span data-stu-id="7facb-115">Members and methods</span></span>

| <span data-ttu-id="7facb-116">Элемент</span><span class="sxs-lookup"><span data-stu-id="7facb-116">Member</span></span> | <span data-ttu-id="7facb-117">Тип</span><span class="sxs-lookup"><span data-stu-id="7facb-117">Type</span></span> |
|--------|------|
| [<span data-ttu-id="7facb-118">attachments</span><span class="sxs-lookup"><span data-stu-id="7facb-118">attachments</span></span>](#attachments-arrayattachmentdetails) | <span data-ttu-id="7facb-119">Элемент</span><span class="sxs-lookup"><span data-stu-id="7facb-119">Member</span></span> |
| [<span data-ttu-id="7facb-120">bcc</span><span class="sxs-lookup"><span data-stu-id="7facb-120">bcc</span></span>](#bcc-recipients) | <span data-ttu-id="7facb-121">Элемент</span><span class="sxs-lookup"><span data-stu-id="7facb-121">Member</span></span> |
| [<span data-ttu-id="7facb-122">body</span><span class="sxs-lookup"><span data-stu-id="7facb-122">body</span></span>](#body-body) | <span data-ttu-id="7facb-123">Элемент</span><span class="sxs-lookup"><span data-stu-id="7facb-123">Member</span></span> |
| [<span data-ttu-id="7facb-124">cc</span><span class="sxs-lookup"><span data-stu-id="7facb-124">cc</span></span>](#cc-arrayemailaddressdetailsrecipients) | <span data-ttu-id="7facb-125">Элемент</span><span class="sxs-lookup"><span data-stu-id="7facb-125">Member</span></span> |
| [<span data-ttu-id="7facb-126">conversationId</span><span class="sxs-lookup"><span data-stu-id="7facb-126">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="7facb-127">Элемент</span><span class="sxs-lookup"><span data-stu-id="7facb-127">Member</span></span> |
| [<span data-ttu-id="7facb-128">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="7facb-128">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="7facb-129">Элемент</span><span class="sxs-lookup"><span data-stu-id="7facb-129">Member</span></span> |
| [<span data-ttu-id="7facb-130">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="7facb-130">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="7facb-131">Элемент</span><span class="sxs-lookup"><span data-stu-id="7facb-131">Member</span></span> |
| [<span data-ttu-id="7facb-132">end</span><span class="sxs-lookup"><span data-stu-id="7facb-132">end</span></span>](#end-datetime) | <span data-ttu-id="7facb-133">Элемент</span><span class="sxs-lookup"><span data-stu-id="7facb-133">Member</span></span> |
| [<span data-ttu-id="7facb-134">from</span><span class="sxs-lookup"><span data-stu-id="7facb-134">from</span></span>](#from-emailaddressdetails) | <span data-ttu-id="7facb-135">Элемент</span><span class="sxs-lookup"><span data-stu-id="7facb-135">Member</span></span> |
| [<span data-ttu-id="7facb-136">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="7facb-136">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="7facb-137">Элемент</span><span class="sxs-lookup"><span data-stu-id="7facb-137">Member</span></span> |
| [<span data-ttu-id="7facb-138">itemClass</span><span class="sxs-lookup"><span data-stu-id="7facb-138">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="7facb-139">Элемент</span><span class="sxs-lookup"><span data-stu-id="7facb-139">Member</span></span> |
| [<span data-ttu-id="7facb-140">itemId</span><span class="sxs-lookup"><span data-stu-id="7facb-140">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="7facb-141">Элемент</span><span class="sxs-lookup"><span data-stu-id="7facb-141">Member</span></span> |
| [<span data-ttu-id="7facb-142">itemType</span><span class="sxs-lookup"><span data-stu-id="7facb-142">itemType</span></span>](#itemtype-officemailboxenumsitemtype) | <span data-ttu-id="7facb-143">Элемент</span><span class="sxs-lookup"><span data-stu-id="7facb-143">Member</span></span> |
| [<span data-ttu-id="7facb-144">location</span><span class="sxs-lookup"><span data-stu-id="7facb-144">location</span></span>](#location-stringlocation) | <span data-ttu-id="7facb-145">Элемент</span><span class="sxs-lookup"><span data-stu-id="7facb-145">Member</span></span> |
| [<span data-ttu-id="7facb-146">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="7facb-146">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="7facb-147">Элемент</span><span class="sxs-lookup"><span data-stu-id="7facb-147">Member</span></span> |
| [<span data-ttu-id="7facb-148">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="7facb-148">notificationMessages</span></span>](#notificationmessages-notificationmessages) | <span data-ttu-id="7facb-149">Элемент</span><span class="sxs-lookup"><span data-stu-id="7facb-149">Member</span></span> |
| [<span data-ttu-id="7facb-150">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="7facb-150">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="7facb-151">Элемент</span><span class="sxs-lookup"><span data-stu-id="7facb-151">Member</span></span> |
| [<span data-ttu-id="7facb-152">organizer</span><span class="sxs-lookup"><span data-stu-id="7facb-152">organizer</span></span>](#organizer-emailaddressdetails) | <span data-ttu-id="7facb-153">Элемент</span><span class="sxs-lookup"><span data-stu-id="7facb-153">Member</span></span> |
| [<span data-ttu-id="7facb-154">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="7facb-154">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="7facb-155">Member</span><span class="sxs-lookup"><span data-stu-id="7facb-155">Member</span></span> |
| [<span data-ttu-id="7facb-156">sender</span><span class="sxs-lookup"><span data-stu-id="7facb-156">sender</span></span>](#sender-emailaddressdetails) | <span data-ttu-id="7facb-157">Элемент</span><span class="sxs-lookup"><span data-stu-id="7facb-157">Member</span></span> |
| [<span data-ttu-id="7facb-158">start</span><span class="sxs-lookup"><span data-stu-id="7facb-158">start</span></span>](#start-datetime) | <span data-ttu-id="7facb-159">Элемент</span><span class="sxs-lookup"><span data-stu-id="7facb-159">Member</span></span> |
| [<span data-ttu-id="7facb-160">subject</span><span class="sxs-lookup"><span data-stu-id="7facb-160">subject</span></span>](#subject-stringsubject) | <span data-ttu-id="7facb-161">Элемент</span><span class="sxs-lookup"><span data-stu-id="7facb-161">Member</span></span> |
| [<span data-ttu-id="7facb-162">to</span><span class="sxs-lookup"><span data-stu-id="7facb-162">to</span></span>](#to-arrayemailaddressdetailsrecipients) | <span data-ttu-id="7facb-163">Элемент</span><span class="sxs-lookup"><span data-stu-id="7facb-163">Member</span></span> |
| [<span data-ttu-id="7facb-164">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="7facb-164">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="7facb-165">Метод</span><span class="sxs-lookup"><span data-stu-id="7facb-165">Method</span></span> |
| [<span data-ttu-id="7facb-166">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="7facb-166">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="7facb-167">Метод</span><span class="sxs-lookup"><span data-stu-id="7facb-167">Method</span></span> |
| [<span data-ttu-id="7facb-168">close</span><span class="sxs-lookup"><span data-stu-id="7facb-168">close</span></span>](#close) | <span data-ttu-id="7facb-169">Метод</span><span class="sxs-lookup"><span data-stu-id="7facb-169">Method</span></span> |
| [<span data-ttu-id="7facb-170">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="7facb-170">displayReplyAllForm</span></span>](#displayreplyallformformdata-callback) | <span data-ttu-id="7facb-171">Метод</span><span class="sxs-lookup"><span data-stu-id="7facb-171">Method</span></span> |
| [<span data-ttu-id="7facb-172">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="7facb-172">displayReplyForm</span></span>](#displayreplyformformdata-callback) | <span data-ttu-id="7facb-173">Метод</span><span class="sxs-lookup"><span data-stu-id="7facb-173">Method</span></span> |
| [<span data-ttu-id="7facb-174">getEntities</span><span class="sxs-lookup"><span data-stu-id="7facb-174">getEntities</span></span>](#getentities--entities) | <span data-ttu-id="7facb-175">Метод</span><span class="sxs-lookup"><span data-stu-id="7facb-175">Method</span></span> |
| [<span data-ttu-id="7facb-176">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="7facb-176">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="7facb-177">Метод</span><span class="sxs-lookup"><span data-stu-id="7facb-177">Method</span></span> |
| [<span data-ttu-id="7facb-178">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="7facb-178">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="7facb-179">Метод</span><span class="sxs-lookup"><span data-stu-id="7facb-179">Method</span></span> |
| [<span data-ttu-id="7facb-180">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="7facb-180">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="7facb-181">Метод</span><span class="sxs-lookup"><span data-stu-id="7facb-181">Method</span></span> |
| [<span data-ttu-id="7facb-182">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="7facb-182">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="7facb-183">Метод</span><span class="sxs-lookup"><span data-stu-id="7facb-183">Method</span></span> |
| [<span data-ttu-id="7facb-184">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="7facb-184">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="7facb-185">Метод</span><span class="sxs-lookup"><span data-stu-id="7facb-185">Method</span></span> |
| [<span data-ttu-id="7facb-186">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="7facb-186">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="7facb-187">Метод</span><span class="sxs-lookup"><span data-stu-id="7facb-187">Method</span></span> |
| [<span data-ttu-id="7facb-188">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="7facb-188">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="7facb-189">Метод</span><span class="sxs-lookup"><span data-stu-id="7facb-189">Method</span></span> |
| [<span data-ttu-id="7facb-190">saveAsync</span><span class="sxs-lookup"><span data-stu-id="7facb-190">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="7facb-191">Метод</span><span class="sxs-lookup"><span data-stu-id="7facb-191">Method</span></span> |
| [<span data-ttu-id="7facb-192">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="7facb-192">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="7facb-193">Метод</span><span class="sxs-lookup"><span data-stu-id="7facb-193">Method</span></span> |

### <a name="example"></a><span data-ttu-id="7facb-194">Пример</span><span class="sxs-lookup"><span data-stu-id="7facb-194">Example</span></span>

<span data-ttu-id="7facb-195">В примере кода JavaScript, приведенном ниже, показано, как получить доступ к свойству `subject` текущего элемента в Outlook.</span><span class="sxs-lookup"><span data-stu-id="7facb-195">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="7facb-196">Элементы</span><span class="sxs-lookup"><span data-stu-id="7facb-196">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetailsviewoutlook-js-14"></a><span data-ttu-id="7facb-197">вложения: Array. <[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.4)></span><span class="sxs-lookup"><span data-stu-id="7facb-197">attachments: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.4)></span></span>

<span data-ttu-id="7facb-p102">Получает массив вложений для элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="7facb-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="7facb-200">Outlook блокирует определенные типы файлов из-за потенциальных проблем с безопасностью, поэтому они не возвращаются.</span><span class="sxs-lookup"><span data-stu-id="7facb-200">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="7facb-201">Дополнительные сведения см. в статье [Блокировка вложений в Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="7facb-201">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="7facb-202">Тип</span><span class="sxs-lookup"><span data-stu-id="7facb-202">Type</span></span>

*   <span data-ttu-id="7facb-203">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.4)></span><span class="sxs-lookup"><span data-stu-id="7facb-203">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.4)></span></span>

##### <a name="requirements"></a><span data-ttu-id="7facb-204">Требования</span><span class="sxs-lookup"><span data-stu-id="7facb-204">Requirements</span></span>

|<span data-ttu-id="7facb-205">Требование</span><span class="sxs-lookup"><span data-stu-id="7facb-205">Requirement</span></span>| <span data-ttu-id="7facb-206">Значение</span><span class="sxs-lookup"><span data-stu-id="7facb-206">Value</span></span>|
|---|---|
|[<span data-ttu-id="7facb-207">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="7facb-207">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7facb-208">1.0</span><span class="sxs-lookup"><span data-stu-id="7facb-208">1.0</span></span>|
|[<span data-ttu-id="7facb-209">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="7facb-209">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7facb-210">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7facb-210">ReadItem</span></span>|
|[<span data-ttu-id="7facb-211">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="7facb-211">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7facb-212">Чтение</span><span class="sxs-lookup"><span data-stu-id="7facb-212">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="7facb-213">Пример</span><span class="sxs-lookup"><span data-stu-id="7facb-213">Example</span></span>

<span data-ttu-id="7facb-214">С помощью приведенного ниже кода можно создать HTML-строку с подробными сведениями обо всех вложениях для текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="7facb-214">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

#### <a name="bcc-recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-14"></a><span data-ttu-id="7facb-215">СК: [получатели](/javascript/api/outlook/office.recipients?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="7facb-215">bcc: [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.4)</span></span>

<span data-ttu-id="7facb-216">Получает объект, который предоставляет методы для получения или обновления строки "СК" (Скрытая копия) сообщения.</span><span class="sxs-lookup"><span data-stu-id="7facb-216">Gets an object that provides methods to get or update the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="7facb-217">Только в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="7facb-217">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="7facb-218">Тип</span><span class="sxs-lookup"><span data-stu-id="7facb-218">Type</span></span>

*   [<span data-ttu-id="7facb-219">Получатели</span><span class="sxs-lookup"><span data-stu-id="7facb-219">Recipients</span></span>](/javascript/api/outlook/office.recipients?view=outlook-js-1.4)

##### <a name="requirements"></a><span data-ttu-id="7facb-220">Требования</span><span class="sxs-lookup"><span data-stu-id="7facb-220">Requirements</span></span>

|<span data-ttu-id="7facb-221">Требование</span><span class="sxs-lookup"><span data-stu-id="7facb-221">Requirement</span></span>| <span data-ttu-id="7facb-222">Значение</span><span class="sxs-lookup"><span data-stu-id="7facb-222">Value</span></span>|
|---|---|
|[<span data-ttu-id="7facb-223">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="7facb-223">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7facb-224">1.1</span><span class="sxs-lookup"><span data-stu-id="7facb-224">1.1</span></span>|
|[<span data-ttu-id="7facb-225">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="7facb-225">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7facb-226">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7facb-226">ReadItem</span></span>|
|[<span data-ttu-id="7facb-227">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="7facb-227">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7facb-228">Создание</span><span class="sxs-lookup"><span data-stu-id="7facb-228">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="7facb-229">Пример</span><span class="sxs-lookup"><span data-stu-id="7facb-229">Example</span></span>

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

#### <a name="body-bodyjavascriptapioutlookofficebodyviewoutlook-js-14"></a><span data-ttu-id="7facb-230">основной текст: [Body](/javascript/api/outlook/office.body?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="7facb-230">body: [Body](/javascript/api/outlook/office.body?view=outlook-js-1.4)</span></span>

<span data-ttu-id="7facb-231">Получает объект, предоставляющий методы для работы с основным текстом элемента.</span><span class="sxs-lookup"><span data-stu-id="7facb-231">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="7facb-232">Тип</span><span class="sxs-lookup"><span data-stu-id="7facb-232">Type</span></span>

*   [<span data-ttu-id="7facb-233">Body</span><span class="sxs-lookup"><span data-stu-id="7facb-233">Body</span></span>](/javascript/api/outlook/office.body?view=outlook-js-1.4)

##### <a name="requirements"></a><span data-ttu-id="7facb-234">Требования</span><span class="sxs-lookup"><span data-stu-id="7facb-234">Requirements</span></span>

|<span data-ttu-id="7facb-235">Требование</span><span class="sxs-lookup"><span data-stu-id="7facb-235">Requirement</span></span>| <span data-ttu-id="7facb-236">Значение</span><span class="sxs-lookup"><span data-stu-id="7facb-236">Value</span></span>|
|---|---|
|[<span data-ttu-id="7facb-237">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="7facb-237">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7facb-238">1.1</span><span class="sxs-lookup"><span data-stu-id="7facb-238">1.1</span></span>|
|[<span data-ttu-id="7facb-239">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="7facb-239">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7facb-240">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7facb-240">ReadItem</span></span>|
|[<span data-ttu-id="7facb-241">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="7facb-241">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7facb-242">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="7facb-242">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="7facb-243">Пример</span><span class="sxs-lookup"><span data-stu-id="7facb-243">Example</span></span>

<span data-ttu-id="7facb-244">В этом примере возвращается текст сообщения в формате обычного текста.</span><span class="sxs-lookup"><span data-stu-id="7facb-244">This example gets the body of the message in plain text.</span></span>

```js
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="7facb-245">Ниже приведен пример итогового параметра, переданного функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="7facb-245">The following is an example of the result parameter passed to the callback function.</span></span>

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

#### <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-14recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-14"></a><span data-ttu-id="7facb-246">CC: Array. <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)>|[получатели](/javascript/api/outlook/office.recipients?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="7facb-246">cc: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.4)</span></span>

<span data-ttu-id="7facb-247">Предоставляет доступ к получателям копии сообщения.</span><span class="sxs-lookup"><span data-stu-id="7facb-247">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="7facb-248">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="7facb-248">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="7facb-249">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="7facb-249">Read mode</span></span>

<span data-ttu-id="7facb-p106">Свойство `cc` возвращает массив, который содержит объект `EmailAddressDetails` для каждого получателя, указанного в строке **Копия** сообщения. Коллекция может включать не более 100 элементов.</span><span class="sxs-lookup"><span data-stu-id="7facb-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="7facb-252">Режим создания</span><span class="sxs-lookup"><span data-stu-id="7facb-252">Compose mode</span></span>

<span data-ttu-id="7facb-253">Свойство `cc` возвращает объект `Recipients`, предоставляющий методы для получения или обновления получателей, которые указаны в строке **Копия** сообщения.</span><span class="sxs-lookup"><span data-stu-id="7facb-253">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

```js
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="7facb-254">Тип</span><span class="sxs-lookup"><span data-stu-id="7facb-254">Type</span></span>

*   <span data-ttu-id="7facb-255">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="7facb-255">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.4)</span></span>

##### <a name="requirements"></a><span data-ttu-id="7facb-256">Требования</span><span class="sxs-lookup"><span data-stu-id="7facb-256">Requirements</span></span>

|<span data-ttu-id="7facb-257">Требование</span><span class="sxs-lookup"><span data-stu-id="7facb-257">Requirement</span></span>| <span data-ttu-id="7facb-258">Значение</span><span class="sxs-lookup"><span data-stu-id="7facb-258">Value</span></span>|
|---|---|
|[<span data-ttu-id="7facb-259">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="7facb-259">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7facb-260">1.0</span><span class="sxs-lookup"><span data-stu-id="7facb-260">1.0</span></span>|
|[<span data-ttu-id="7facb-261">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="7facb-261">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7facb-262">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7facb-262">ReadItem</span></span>|
|[<span data-ttu-id="7facb-263">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="7facb-263">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7facb-264">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="7facb-264">Compose or Read</span></span>|

<br>

---
---

#### <a name="nullable-conversationid-string"></a><span data-ttu-id="7facb-265">(Nullable) conversationId: строка</span><span class="sxs-lookup"><span data-stu-id="7facb-265">(nullable) conversationId: String</span></span>

<span data-ttu-id="7facb-266">Получает идентификатор разговора по электронной почте, содержащего конкретное сообщение.</span><span class="sxs-lookup"><span data-stu-id="7facb-266">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="7facb-p107">Вы можете получить целочисленное значение этого свойства, если ваше почтовое приложение активируется в формах просмотра или формах создания ответов. Если пользователь изменит тему ответа, после его отправки идентификатор беседы будет изменен, и полученное ранее значение будет недействительным.</span><span class="sxs-lookup"><span data-stu-id="7facb-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="7facb-p108">Это свойство имеет значение NULL для нового элемента в форме создания. Свойство `conversationId` вернет значение, если пользователь задаст тему и сохранит элемент.</span><span class="sxs-lookup"><span data-stu-id="7facb-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="7facb-271">Тип</span><span class="sxs-lookup"><span data-stu-id="7facb-271">Type</span></span>

*   <span data-ttu-id="7facb-272">String</span><span class="sxs-lookup"><span data-stu-id="7facb-272">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="7facb-273">Требования</span><span class="sxs-lookup"><span data-stu-id="7facb-273">Requirements</span></span>

|<span data-ttu-id="7facb-274">Требование</span><span class="sxs-lookup"><span data-stu-id="7facb-274">Requirement</span></span>| <span data-ttu-id="7facb-275">Значение</span><span class="sxs-lookup"><span data-stu-id="7facb-275">Value</span></span>|
|---|---|
|[<span data-ttu-id="7facb-276">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="7facb-276">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7facb-277">1.0</span><span class="sxs-lookup"><span data-stu-id="7facb-277">1.0</span></span>|
|[<span data-ttu-id="7facb-278">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="7facb-278">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7facb-279">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7facb-279">ReadItem</span></span>|
|[<span data-ttu-id="7facb-280">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="7facb-280">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7facb-281">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="7facb-281">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="7facb-282">Пример</span><span class="sxs-lookup"><span data-stu-id="7facb-282">Example</span></span>

```js
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

<br>

---
---

#### <a name="datetimecreated-date"></a><span data-ttu-id="7facb-283">dateTimeCreated: Дата</span><span class="sxs-lookup"><span data-stu-id="7facb-283">dateTimeCreated: Date</span></span>

<span data-ttu-id="7facb-p109">Получает дату и время создания элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="7facb-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="7facb-286">Тип</span><span class="sxs-lookup"><span data-stu-id="7facb-286">Type</span></span>

*   <span data-ttu-id="7facb-287">Дата</span><span class="sxs-lookup"><span data-stu-id="7facb-287">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="7facb-288">Требования</span><span class="sxs-lookup"><span data-stu-id="7facb-288">Requirements</span></span>

|<span data-ttu-id="7facb-289">Требование</span><span class="sxs-lookup"><span data-stu-id="7facb-289">Requirement</span></span>| <span data-ttu-id="7facb-290">Значение</span><span class="sxs-lookup"><span data-stu-id="7facb-290">Value</span></span>|
|---|---|
|[<span data-ttu-id="7facb-291">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="7facb-291">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7facb-292">1.0</span><span class="sxs-lookup"><span data-stu-id="7facb-292">1.0</span></span>|
|[<span data-ttu-id="7facb-293">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="7facb-293">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7facb-294">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7facb-294">ReadItem</span></span>|
|[<span data-ttu-id="7facb-295">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="7facb-295">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7facb-296">Чтение</span><span class="sxs-lookup"><span data-stu-id="7facb-296">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="7facb-297">Пример</span><span class="sxs-lookup"><span data-stu-id="7facb-297">Example</span></span>

```js
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

<br>

---
---

#### <a name="datetimemodified-date"></a><span data-ttu-id="7facb-298">dateTimeModified: Дата</span><span class="sxs-lookup"><span data-stu-id="7facb-298">dateTimeModified: Date</span></span>

<span data-ttu-id="7facb-p110">Получает дату и время последнего изменения элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="7facb-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="7facb-301">Этот элемент не поддерживается в Outlook на iOS или Android.</span><span class="sxs-lookup"><span data-stu-id="7facb-301">This member is not supported in Outlook on iOS or Android.</span></span>

##### <a name="type"></a><span data-ttu-id="7facb-302">Тип</span><span class="sxs-lookup"><span data-stu-id="7facb-302">Type</span></span>

*   <span data-ttu-id="7facb-303">Дата</span><span class="sxs-lookup"><span data-stu-id="7facb-303">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="7facb-304">Требования</span><span class="sxs-lookup"><span data-stu-id="7facb-304">Requirements</span></span>

|<span data-ttu-id="7facb-305">Требование</span><span class="sxs-lookup"><span data-stu-id="7facb-305">Requirement</span></span>| <span data-ttu-id="7facb-306">Значение</span><span class="sxs-lookup"><span data-stu-id="7facb-306">Value</span></span>|
|---|---|
|[<span data-ttu-id="7facb-307">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="7facb-307">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7facb-308">1.0</span><span class="sxs-lookup"><span data-stu-id="7facb-308">1.0</span></span>|
|[<span data-ttu-id="7facb-309">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="7facb-309">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7facb-310">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7facb-310">ReadItem</span></span>|
|[<span data-ttu-id="7facb-311">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="7facb-311">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7facb-312">Чтение</span><span class="sxs-lookup"><span data-stu-id="7facb-312">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="7facb-313">Пример</span><span class="sxs-lookup"><span data-stu-id="7facb-313">Example</span></span>

```js
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

<br>

---
---

#### <a name="end-datetimejavascriptapioutlookofficetimeviewoutlook-js-14"></a><span data-ttu-id="7facb-314">конец: Дата | [Time (время](/javascript/api/outlook/office.time?view=outlook-js-1.4) )</span><span class="sxs-lookup"><span data-stu-id="7facb-314">end: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.4)</span></span>

<span data-ttu-id="7facb-315">Получает или задает дату и время окончания встречи.</span><span class="sxs-lookup"><span data-stu-id="7facb-315">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="7facb-p111">Свойство `end` представлено в виде значения даты и времени в формате UTC. Преобразовать значение свойства end в местные значения даты и времени клиента можно с помощью метода [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime).</span><span class="sxs-lookup"><span data-stu-id="7facb-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="7facb-318">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="7facb-318">Read mode</span></span>

<span data-ttu-id="7facb-319">Свойство `end` возвращает объект `Date`.</span><span class="sxs-lookup"><span data-stu-id="7facb-319">The `end` property returns a `Date` object.</span></span>

```js
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="7facb-320">Режим создания</span><span class="sxs-lookup"><span data-stu-id="7facb-320">Compose mode</span></span>

<span data-ttu-id="7facb-321">Свойство `end` возвращает объект `Time`.</span><span class="sxs-lookup"><span data-stu-id="7facb-321">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="7facb-322">Если вы задаете время окончания с помощью метода [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.4#setasync-datetime--options--callback-), необходимо использовать метод [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) для преобразования местного времени на клиенте в формат UTC для сервера.</span><span class="sxs-lookup"><span data-stu-id="7facb-322">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.4#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="7facb-323">В примере ниже показано, как с помощью метода [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.4#setasync-datetime--options--callback-) объекта `Time` задать время окончания встречи.</span><span class="sxs-lookup"><span data-stu-id="7facb-323">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.4#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="7facb-324">Тип</span><span class="sxs-lookup"><span data-stu-id="7facb-324">Type</span></span>

*   <span data-ttu-id="7facb-325">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="7facb-325">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.4)</span></span>

##### <a name="requirements"></a><span data-ttu-id="7facb-326">Требования</span><span class="sxs-lookup"><span data-stu-id="7facb-326">Requirements</span></span>

|<span data-ttu-id="7facb-327">Требование</span><span class="sxs-lookup"><span data-stu-id="7facb-327">Requirement</span></span>| <span data-ttu-id="7facb-328">Значение</span><span class="sxs-lookup"><span data-stu-id="7facb-328">Value</span></span>|
|---|---|
|[<span data-ttu-id="7facb-329">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="7facb-329">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7facb-330">1.0</span><span class="sxs-lookup"><span data-stu-id="7facb-330">1.0</span></span>|
|[<span data-ttu-id="7facb-331">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="7facb-331">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7facb-332">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7facb-332">ReadItem</span></span>|
|[<span data-ttu-id="7facb-333">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="7facb-333">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7facb-334">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="7facb-334">Compose or Read</span></span>|

<br>

---
---

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-14"></a><span data-ttu-id="7facb-335">от: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="7facb-335">from: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)</span></span>

<span data-ttu-id="7facb-p112">Получает электронный адрес отправителя сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="7facb-p112">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="7facb-p113">Свойства `from` и [`sender`](#sender-emailaddressdetails) представляют одно лицо, если сообщение не отправлено представителем. В противном случае свойство `from` представляет лицо, делегировавшее полномочия, а свойство sender — представителя.</span><span class="sxs-lookup"><span data-stu-id="7facb-p113">The `from` and [`sender`](#sender-emailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="7facb-340">Свойству `recipientType`, принадлежащему объекту `EmailAddressDetails` в свойстве `from`, задано значение `undefined`.</span><span class="sxs-lookup"><span data-stu-id="7facb-340">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="7facb-341">Тип</span><span class="sxs-lookup"><span data-stu-id="7facb-341">Type</span></span>

*   [<span data-ttu-id="7facb-342">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="7facb-342">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)

##### <a name="requirements"></a><span data-ttu-id="7facb-343">Требования</span><span class="sxs-lookup"><span data-stu-id="7facb-343">Requirements</span></span>

|<span data-ttu-id="7facb-344">Требование</span><span class="sxs-lookup"><span data-stu-id="7facb-344">Requirement</span></span>| <span data-ttu-id="7facb-345">Значение</span><span class="sxs-lookup"><span data-stu-id="7facb-345">Value</span></span>|
|---|---|
|[<span data-ttu-id="7facb-346">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="7facb-346">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7facb-347">1.0</span><span class="sxs-lookup"><span data-stu-id="7facb-347">1.0</span></span>|
|[<span data-ttu-id="7facb-348">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="7facb-348">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7facb-349">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7facb-349">ReadItem</span></span>|
|[<span data-ttu-id="7facb-350">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="7facb-350">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7facb-351">Чтение</span><span class="sxs-lookup"><span data-stu-id="7facb-351">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="7facb-352">Пример</span><span class="sxs-lookup"><span data-stu-id="7facb-352">Example</span></span>

```js
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

<br>

---
---

#### <a name="internetmessageid-string"></a><span data-ttu-id="7facb-353">internetMessageId: строка</span><span class="sxs-lookup"><span data-stu-id="7facb-353">internetMessageId: String</span></span>

<span data-ttu-id="7facb-p114">Получает идентификатор интернет-сообщения для электронного сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="7facb-p114">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="7facb-356">Тип</span><span class="sxs-lookup"><span data-stu-id="7facb-356">Type</span></span>

*   <span data-ttu-id="7facb-357">String</span><span class="sxs-lookup"><span data-stu-id="7facb-357">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="7facb-358">Требования</span><span class="sxs-lookup"><span data-stu-id="7facb-358">Requirements</span></span>

|<span data-ttu-id="7facb-359">Требование</span><span class="sxs-lookup"><span data-stu-id="7facb-359">Requirement</span></span>| <span data-ttu-id="7facb-360">Значение</span><span class="sxs-lookup"><span data-stu-id="7facb-360">Value</span></span>|
|---|---|
|[<span data-ttu-id="7facb-361">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="7facb-361">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7facb-362">1.0</span><span class="sxs-lookup"><span data-stu-id="7facb-362">1.0</span></span>|
|[<span data-ttu-id="7facb-363">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="7facb-363">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7facb-364">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7facb-364">ReadItem</span></span>|
|[<span data-ttu-id="7facb-365">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="7facb-365">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7facb-366">Чтение</span><span class="sxs-lookup"><span data-stu-id="7facb-366">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="7facb-367">Пример</span><span class="sxs-lookup"><span data-stu-id="7facb-367">Example</span></span>

```js
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

<br>

---
---

#### <a name="itemclass-string"></a><span data-ttu-id="7facb-368">itemClass: строка</span><span class="sxs-lookup"><span data-stu-id="7facb-368">itemClass: String</span></span>

<span data-ttu-id="7facb-p115">Получает класс элемента веб-служб Exchange для выбранного элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="7facb-p115">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="7facb-p116">Свойство `itemClass` указывает класс сообщения выбранного элемента. Ниже приводятся классы сообщения по умолчанию для элемента сообщения или встречи.</span><span class="sxs-lookup"><span data-stu-id="7facb-p116">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="7facb-373">Тип</span><span class="sxs-lookup"><span data-stu-id="7facb-373">Type</span></span> | <span data-ttu-id="7facb-374">Описание</span><span class="sxs-lookup"><span data-stu-id="7facb-374">Description</span></span> | <span data-ttu-id="7facb-375">Класс элемента</span><span class="sxs-lookup"><span data-stu-id="7facb-375">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="7facb-376">Элементы встречи</span><span class="sxs-lookup"><span data-stu-id="7facb-376">Appointment items</span></span> | <span data-ttu-id="7facb-377">Это элементы календаря для класса элемента `IPM.Appointment` или `IPM.Appointment.Occurrence`.</span><span class="sxs-lookup"><span data-stu-id="7facb-377">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurrence` |
| <span data-ttu-id="7facb-378">Элементы сообщения</span><span class="sxs-lookup"><span data-stu-id="7facb-378">Message items</span></span> | <span data-ttu-id="7facb-379">Сюда входят электронные сообщения, для которых по умолчанию задан класс сообщения `IPM.Note`, а также приглашения на собрания, ответы на них и уведомления об их отмене, использующие `IPM.Schedule.Meeting` в качестве базового класса сообщения.</span><span class="sxs-lookup"><span data-stu-id="7facb-379">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="7facb-380">Можно создавать настраиваемые классы сообщения, расширяющие классы сообщения по умолчанию, например настраиваемый класс сообщения о встрече `IPM.Appointment.Contoso`.</span><span class="sxs-lookup"><span data-stu-id="7facb-380">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="7facb-381">Тип</span><span class="sxs-lookup"><span data-stu-id="7facb-381">Type</span></span>

*   <span data-ttu-id="7facb-382">String</span><span class="sxs-lookup"><span data-stu-id="7facb-382">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="7facb-383">Требования</span><span class="sxs-lookup"><span data-stu-id="7facb-383">Requirements</span></span>

|<span data-ttu-id="7facb-384">Требование</span><span class="sxs-lookup"><span data-stu-id="7facb-384">Requirement</span></span>| <span data-ttu-id="7facb-385">Значение</span><span class="sxs-lookup"><span data-stu-id="7facb-385">Value</span></span>|
|---|---|
|[<span data-ttu-id="7facb-386">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="7facb-386">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7facb-387">1.0</span><span class="sxs-lookup"><span data-stu-id="7facb-387">1.0</span></span>|
|[<span data-ttu-id="7facb-388">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="7facb-388">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7facb-389">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7facb-389">ReadItem</span></span>|
|[<span data-ttu-id="7facb-390">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="7facb-390">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7facb-391">Чтение</span><span class="sxs-lookup"><span data-stu-id="7facb-391">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="7facb-392">Пример</span><span class="sxs-lookup"><span data-stu-id="7facb-392">Example</span></span>

```js
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

<br>

---
---

#### <a name="nullable-itemid-string"></a><span data-ttu-id="7facb-393">(Nullable) itemId: строка</span><span class="sxs-lookup"><span data-stu-id="7facb-393">(nullable) itemId: String</span></span>

<span data-ttu-id="7facb-p117">Получает идентификатор элемента веб-служб Exchange для текущего элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="7facb-p117">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="7facb-396">Идентификатор, возвращаемый свойством `itemId`, совпадает с идентификатором элемента веб-служб Exchange.</span><span class="sxs-lookup"><span data-stu-id="7facb-396">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="7facb-397">Свойство `itemId` не совпадает с идентификатором записи Outlook, а также идентификатором, который используется REST API Outlook.</span><span class="sxs-lookup"><span data-stu-id="7facb-397">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="7facb-398">Прежде чем совершать вызовы REST API, используя это значение, его необходимо преобразовать с помощью [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="7facb-398">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="7facb-399">Дополнительные сведения см. в статье [Использование REST API Outlook из надстройки Outlook](/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="7facb-399">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="7facb-p119">Свойство `itemId` недоступно в режиме создания. Если требуется идентификатор элемента, с помощью метода [`saveAsync`](#saveasyncoptions-callback) можно сохранить элемент в хранилище. При этом в параметре [`AsyncResult.value`](/javascript/api/office/office.asyncresult) функции обратного вызова возвращается идентификатор элемента.</span><span class="sxs-lookup"><span data-stu-id="7facb-p119">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="7facb-402">Тип</span><span class="sxs-lookup"><span data-stu-id="7facb-402">Type</span></span>

*   <span data-ttu-id="7facb-403">String</span><span class="sxs-lookup"><span data-stu-id="7facb-403">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="7facb-404">Требования</span><span class="sxs-lookup"><span data-stu-id="7facb-404">Requirements</span></span>

|<span data-ttu-id="7facb-405">Требование</span><span class="sxs-lookup"><span data-stu-id="7facb-405">Requirement</span></span>| <span data-ttu-id="7facb-406">Значение</span><span class="sxs-lookup"><span data-stu-id="7facb-406">Value</span></span>|
|---|---|
|[<span data-ttu-id="7facb-407">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="7facb-407">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7facb-408">1.0</span><span class="sxs-lookup"><span data-stu-id="7facb-408">1.0</span></span>|
|[<span data-ttu-id="7facb-409">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="7facb-409">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7facb-410">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7facb-410">ReadItem</span></span>|
|[<span data-ttu-id="7facb-411">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="7facb-411">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7facb-412">Чтение</span><span class="sxs-lookup"><span data-stu-id="7facb-412">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="7facb-413">Пример</span><span class="sxs-lookup"><span data-stu-id="7facb-413">Example</span></span>

<span data-ttu-id="7facb-p120">Указанный ниже код проверяет наличие идентификатора элемента. Если свойство `itemId` возвращает значение `null` или `undefined`, элемент будет сохранен в хранилище, а из асинхронного результата будет получен идентификатор элемента.</span><span class="sxs-lookup"><span data-stu-id="7facb-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

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

#### <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtypeviewoutlook-js-14"></a><span data-ttu-id="7facb-416">itemType: [Office. MailboxEnums. ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="7facb-416">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.4)</span></span>

<span data-ttu-id="7facb-417">Получает тип элемента, который представляет экземпляр.</span><span class="sxs-lookup"><span data-stu-id="7facb-417">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="7facb-418">Свойство `itemType` возвращает одно из значений перечисления `ItemType`, которое указывает, является ли экземпляр объекта `item` сообщением или встречей.</span><span class="sxs-lookup"><span data-stu-id="7facb-418">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="7facb-419">Тип</span><span class="sxs-lookup"><span data-stu-id="7facb-419">Type</span></span>

*   [<span data-ttu-id="7facb-420">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="7facb-420">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.4)

##### <a name="requirements"></a><span data-ttu-id="7facb-421">Требования</span><span class="sxs-lookup"><span data-stu-id="7facb-421">Requirements</span></span>

|<span data-ttu-id="7facb-422">Требование</span><span class="sxs-lookup"><span data-stu-id="7facb-422">Requirement</span></span>| <span data-ttu-id="7facb-423">Значение</span><span class="sxs-lookup"><span data-stu-id="7facb-423">Value</span></span>|
|---|---|
|[<span data-ttu-id="7facb-424">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="7facb-424">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7facb-425">1.0</span><span class="sxs-lookup"><span data-stu-id="7facb-425">1.0</span></span>|
|[<span data-ttu-id="7facb-426">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="7facb-426">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7facb-427">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7facb-427">ReadItem</span></span>|
|[<span data-ttu-id="7facb-428">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="7facb-428">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7facb-429">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="7facb-429">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="7facb-430">Пример</span><span class="sxs-lookup"><span data-stu-id="7facb-430">Example</span></span>

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

#### <a name="location-stringlocationjavascriptapioutlookofficelocationviewoutlook-js-14"></a><span data-ttu-id="7facb-431">Местоположение: строка | [Location (расположение](/javascript/api/outlook/office.location?view=outlook-js-1.4) )</span><span class="sxs-lookup"><span data-stu-id="7facb-431">location: String|[Location](/javascript/api/outlook/office.location?view=outlook-js-1.4)</span></span>

<span data-ttu-id="7facb-432">Получает или задает место встречи.</span><span class="sxs-lookup"><span data-stu-id="7facb-432">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="7facb-433">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="7facb-433">Read mode</span></span>

<span data-ttu-id="7facb-434">Свойство `location` возвращает строку, содержащую сведения о месте встречи.</span><span class="sxs-lookup"><span data-stu-id="7facb-434">The `location` property returns a string that contains the location of the appointment.</span></span>

```js
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="7facb-435">Режим создания</span><span class="sxs-lookup"><span data-stu-id="7facb-435">Compose mode</span></span>

<span data-ttu-id="7facb-436">Свойство `location` возвращает объект `Location`, предоставляющий методы, которые используются для получения и задания места встречи.</span><span class="sxs-lookup"><span data-stu-id="7facb-436">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```js
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="7facb-437">Тип</span><span class="sxs-lookup"><span data-stu-id="7facb-437">Type</span></span>

*   <span data-ttu-id="7facb-438">String | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="7facb-438">String | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.4)</span></span>

##### <a name="requirements"></a><span data-ttu-id="7facb-439">Требования</span><span class="sxs-lookup"><span data-stu-id="7facb-439">Requirements</span></span>

|<span data-ttu-id="7facb-440">Требование</span><span class="sxs-lookup"><span data-stu-id="7facb-440">Requirement</span></span>| <span data-ttu-id="7facb-441">Значение</span><span class="sxs-lookup"><span data-stu-id="7facb-441">Value</span></span>|
|---|---|
|[<span data-ttu-id="7facb-442">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="7facb-442">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7facb-443">1.0</span><span class="sxs-lookup"><span data-stu-id="7facb-443">1.0</span></span>|
|[<span data-ttu-id="7facb-444">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="7facb-444">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7facb-445">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7facb-445">ReadItem</span></span>|
|[<span data-ttu-id="7facb-446">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="7facb-446">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7facb-447">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="7facb-447">Compose or Read</span></span>|

<br>

---
---

#### <a name="normalizedsubject-string"></a><span data-ttu-id="7facb-448">normalizedSubject: строка</span><span class="sxs-lookup"><span data-stu-id="7facb-448">normalizedSubject: String</span></span>

<span data-ttu-id="7facb-p121">Получает тему элемента со всеми удаленными префиксами (включая `RE:` и `FWD:`). Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="7facb-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="7facb-p122">Свойство normalizedSubject получает тему элемента со стандартными префиксами (такими как `RE:` и `FW:`), добавляемыми почтовыми программами. Для получения темы элемента с неизмененными префиксами используйте свойство [`subject`](#subject-stringsubject).</span><span class="sxs-lookup"><span data-stu-id="7facb-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="7facb-453">Тип</span><span class="sxs-lookup"><span data-stu-id="7facb-453">Type</span></span>

*   <span data-ttu-id="7facb-454">String</span><span class="sxs-lookup"><span data-stu-id="7facb-454">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="7facb-455">Требования</span><span class="sxs-lookup"><span data-stu-id="7facb-455">Requirements</span></span>

|<span data-ttu-id="7facb-456">Требование</span><span class="sxs-lookup"><span data-stu-id="7facb-456">Requirement</span></span>| <span data-ttu-id="7facb-457">Значение</span><span class="sxs-lookup"><span data-stu-id="7facb-457">Value</span></span>|
|---|---|
|[<span data-ttu-id="7facb-458">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="7facb-458">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7facb-459">1.0</span><span class="sxs-lookup"><span data-stu-id="7facb-459">1.0</span></span>|
|[<span data-ttu-id="7facb-460">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="7facb-460">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7facb-461">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7facb-461">ReadItem</span></span>|
|[<span data-ttu-id="7facb-462">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="7facb-462">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7facb-463">Чтение</span><span class="sxs-lookup"><span data-stu-id="7facb-463">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="7facb-464">Пример</span><span class="sxs-lookup"><span data-stu-id="7facb-464">Example</span></span>

```js
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

<br>

---
---

#### <a name="notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessagesviewoutlook-js-14"></a><span data-ttu-id="7facb-465">notificationMessages: [notificationMessages](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="7facb-465">notificationMessages: [NotificationMessages](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.4)</span></span>

<span data-ttu-id="7facb-466">Получает сообщения уведомления для элемента.</span><span class="sxs-lookup"><span data-stu-id="7facb-466">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="7facb-467">Тип</span><span class="sxs-lookup"><span data-stu-id="7facb-467">Type</span></span>

*   [<span data-ttu-id="7facb-468">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="7facb-468">NotificationMessages</span></span>](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.4)

##### <a name="requirements"></a><span data-ttu-id="7facb-469">Требования</span><span class="sxs-lookup"><span data-stu-id="7facb-469">Requirements</span></span>

|<span data-ttu-id="7facb-470">Требование</span><span class="sxs-lookup"><span data-stu-id="7facb-470">Requirement</span></span>| <span data-ttu-id="7facb-471">Значение</span><span class="sxs-lookup"><span data-stu-id="7facb-471">Value</span></span>|
|---|---|
|[<span data-ttu-id="7facb-472">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="7facb-472">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7facb-473">1.3</span><span class="sxs-lookup"><span data-stu-id="7facb-473">1.3</span></span>|
|[<span data-ttu-id="7facb-474">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="7facb-474">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7facb-475">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7facb-475">ReadItem</span></span>|
|[<span data-ttu-id="7facb-476">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="7facb-476">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7facb-477">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="7facb-477">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="7facb-478">Пример</span><span class="sxs-lookup"><span data-stu-id="7facb-478">Example</span></span>

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

#### <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-14recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-14"></a><span data-ttu-id="7facb-479">optionalAttendees: Array. <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)>|[получатели](/javascript/api/outlook/office.recipients?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="7facb-479">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.4)</span></span>

<span data-ttu-id="7facb-480">Предоставляет доступ к необязательным участникам события.</span><span class="sxs-lookup"><span data-stu-id="7facb-480">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="7facb-481">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="7facb-481">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="7facb-482">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="7facb-482">Read mode</span></span>

<span data-ttu-id="7facb-483">Свойство `optionalAttendees` возвращает массив, содержащий объект `EmailAddressDetails` для каждого необязательного участника собрания.</span><span class="sxs-lookup"><span data-stu-id="7facb-483">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

```js
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="7facb-484">Режим создания</span><span class="sxs-lookup"><span data-stu-id="7facb-484">Compose mode</span></span>

<span data-ttu-id="7facb-485">Свойство `optionalAttendees` возвращает объект `Recipients`, который предоставляет методы для получения или обновления необязательных участников собрания.</span><span class="sxs-lookup"><span data-stu-id="7facb-485">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

```js
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="7facb-486">Тип</span><span class="sxs-lookup"><span data-stu-id="7facb-486">Type</span></span>

*   <span data-ttu-id="7facb-487">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="7facb-487">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.4)</span></span>

##### <a name="requirements"></a><span data-ttu-id="7facb-488">Требования</span><span class="sxs-lookup"><span data-stu-id="7facb-488">Requirements</span></span>

|<span data-ttu-id="7facb-489">Требование</span><span class="sxs-lookup"><span data-stu-id="7facb-489">Requirement</span></span>| <span data-ttu-id="7facb-490">Значение</span><span class="sxs-lookup"><span data-stu-id="7facb-490">Value</span></span>|
|---|---|
|[<span data-ttu-id="7facb-491">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="7facb-491">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7facb-492">1.0</span><span class="sxs-lookup"><span data-stu-id="7facb-492">1.0</span></span>|
|[<span data-ttu-id="7facb-493">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="7facb-493">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7facb-494">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7facb-494">ReadItem</span></span>|
|[<span data-ttu-id="7facb-495">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="7facb-495">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7facb-496">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="7facb-496">Compose or Read</span></span>|

<br>

---
---

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-14"></a><span data-ttu-id="7facb-497">Организатор: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="7facb-497">organizer: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)</span></span>

<span data-ttu-id="7facb-p124">Получает электронный адрес организатора указанного собрания. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="7facb-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="7facb-500">Тип</span><span class="sxs-lookup"><span data-stu-id="7facb-500">Type</span></span>

*   [<span data-ttu-id="7facb-501">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="7facb-501">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)

##### <a name="requirements"></a><span data-ttu-id="7facb-502">Требования</span><span class="sxs-lookup"><span data-stu-id="7facb-502">Requirements</span></span>

|<span data-ttu-id="7facb-503">Требование</span><span class="sxs-lookup"><span data-stu-id="7facb-503">Requirement</span></span>| <span data-ttu-id="7facb-504">Значение</span><span class="sxs-lookup"><span data-stu-id="7facb-504">Value</span></span>|
|---|---|
|[<span data-ttu-id="7facb-505">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="7facb-505">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7facb-506">1.0</span><span class="sxs-lookup"><span data-stu-id="7facb-506">1.0</span></span>|
|[<span data-ttu-id="7facb-507">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="7facb-507">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7facb-508">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7facb-508">ReadItem</span></span>|
|[<span data-ttu-id="7facb-509">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="7facb-509">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7facb-510">Чтение</span><span class="sxs-lookup"><span data-stu-id="7facb-510">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="7facb-511">Пример</span><span class="sxs-lookup"><span data-stu-id="7facb-511">Example</span></span>

```js
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

<br>

---
---

#### <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-14recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-14"></a><span data-ttu-id="7facb-512">requiredAttendees: Array. <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)>|[получатели](/javascript/api/outlook/office.recipients?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="7facb-512">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.4)</span></span>

<span data-ttu-id="7facb-513">Предоставляет доступ к обязательным участникам события.</span><span class="sxs-lookup"><span data-stu-id="7facb-513">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="7facb-514">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="7facb-514">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="7facb-515">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="7facb-515">Read mode</span></span>

<span data-ttu-id="7facb-516">Свойство `requiredAttendees` возвращает массив, содержащий объект `EmailAddressDetails` для каждого обязательного участника собрания.</span><span class="sxs-lookup"><span data-stu-id="7facb-516">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

```js
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="7facb-517">Режим создания</span><span class="sxs-lookup"><span data-stu-id="7facb-517">Compose mode</span></span>

<span data-ttu-id="7facb-518">Свойство `requiredAttendees` возвращает объект `Recipients`, предоставляющий методы, с помощью которых можно получить или обновить сведения об обязательных участниках собрания.</span><span class="sxs-lookup"><span data-stu-id="7facb-518">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

```js
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="7facb-519">Тип</span><span class="sxs-lookup"><span data-stu-id="7facb-519">Type</span></span>

*   <span data-ttu-id="7facb-520">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="7facb-520">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.4)</span></span>

##### <a name="requirements"></a><span data-ttu-id="7facb-521">Требования</span><span class="sxs-lookup"><span data-stu-id="7facb-521">Requirements</span></span>

|<span data-ttu-id="7facb-522">Требование</span><span class="sxs-lookup"><span data-stu-id="7facb-522">Requirement</span></span>| <span data-ttu-id="7facb-523">Значение</span><span class="sxs-lookup"><span data-stu-id="7facb-523">Value</span></span>|
|---|---|
|[<span data-ttu-id="7facb-524">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="7facb-524">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7facb-525">1.0</span><span class="sxs-lookup"><span data-stu-id="7facb-525">1.0</span></span>|
|[<span data-ttu-id="7facb-526">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="7facb-526">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7facb-527">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7facb-527">ReadItem</span></span>|
|[<span data-ttu-id="7facb-528">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="7facb-528">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7facb-529">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="7facb-529">Compose or Read</span></span>|

<br>

---
---

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-14"></a><span data-ttu-id="7facb-530">Отправитель: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="7facb-530">sender: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)</span></span>

<span data-ttu-id="7facb-p126">Получает электронный адрес отправителя электронного сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="7facb-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="7facb-p127">Свойства [`from`](#from-emailaddressdetails) и `sender` представляют одно лицо, если сообщение не отправлено представителем. В противном случае свойство `from` представляет лицо, делегировавшее полномочия, а свойство sender — представителя.</span><span class="sxs-lookup"><span data-stu-id="7facb-p127">The [`from`](#from-emailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="7facb-535">Свойству `recipientType`, принадлежащему объекту `EmailAddressDetails` в свойстве `sender`, задано значение `undefined`.</span><span class="sxs-lookup"><span data-stu-id="7facb-535">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="7facb-536">Тип</span><span class="sxs-lookup"><span data-stu-id="7facb-536">Type</span></span>

*   [<span data-ttu-id="7facb-537">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="7facb-537">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)

##### <a name="requirements"></a><span data-ttu-id="7facb-538">Требования</span><span class="sxs-lookup"><span data-stu-id="7facb-538">Requirements</span></span>

|<span data-ttu-id="7facb-539">Требование</span><span class="sxs-lookup"><span data-stu-id="7facb-539">Requirement</span></span>| <span data-ttu-id="7facb-540">Значение</span><span class="sxs-lookup"><span data-stu-id="7facb-540">Value</span></span>|
|---|---|
|[<span data-ttu-id="7facb-541">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="7facb-541">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7facb-542">1.0</span><span class="sxs-lookup"><span data-stu-id="7facb-542">1.0</span></span>|
|[<span data-ttu-id="7facb-543">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="7facb-543">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7facb-544">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7facb-544">ReadItem</span></span>|
|[<span data-ttu-id="7facb-545">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="7facb-545">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7facb-546">Чтение</span><span class="sxs-lookup"><span data-stu-id="7facb-546">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="7facb-547">Пример</span><span class="sxs-lookup"><span data-stu-id="7facb-547">Example</span></span>

```js
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

<br>

---
---

#### <a name="start-datetimejavascriptapioutlookofficetimeviewoutlook-js-14"></a><span data-ttu-id="7facb-548">Начало: Дата | [Time (время](/javascript/api/outlook/office.time?view=outlook-js-1.4) )</span><span class="sxs-lookup"><span data-stu-id="7facb-548">start: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.4)</span></span>

<span data-ttu-id="7facb-549">Получает или задает дату и время начала встречи.</span><span class="sxs-lookup"><span data-stu-id="7facb-549">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="7facb-p128">Свойство `start` представлено в виде значения даты и времени в формате UTC. Это значение можно преобразовать в местные значения даты и времени клиента с помощью метода [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime).</span><span class="sxs-lookup"><span data-stu-id="7facb-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="7facb-552">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="7facb-552">Read mode</span></span>

<span data-ttu-id="7facb-553">Свойство `start` возвращает объект `Date`.</span><span class="sxs-lookup"><span data-stu-id="7facb-553">The `start` property returns a `Date` object.</span></span>

```js
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="7facb-554">Режим создания</span><span class="sxs-lookup"><span data-stu-id="7facb-554">Compose mode</span></span>

<span data-ttu-id="7facb-555">Свойство `start` возвращает объект `Time`.</span><span class="sxs-lookup"><span data-stu-id="7facb-555">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="7facb-556">Если вы задаете время начала с помощью метода [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.4#setasync-datetime--options--callback-), необходимо использовать метод [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) для преобразования местного времени на клиенте в формат UTC для сервера.</span><span class="sxs-lookup"><span data-stu-id="7facb-556">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.4#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="7facb-557">В примере ниже с помощью метода [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.4#setasync-datetime--options--callback-) объекта `Time` задается время начала встречи в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="7facb-557">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.4#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="7facb-558">Тип</span><span class="sxs-lookup"><span data-stu-id="7facb-558">Type</span></span>

*   <span data-ttu-id="7facb-559">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="7facb-559">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.4)</span></span>

##### <a name="requirements"></a><span data-ttu-id="7facb-560">Требования</span><span class="sxs-lookup"><span data-stu-id="7facb-560">Requirements</span></span>

|<span data-ttu-id="7facb-561">Требование</span><span class="sxs-lookup"><span data-stu-id="7facb-561">Requirement</span></span>| <span data-ttu-id="7facb-562">Значение</span><span class="sxs-lookup"><span data-stu-id="7facb-562">Value</span></span>|
|---|---|
|[<span data-ttu-id="7facb-563">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="7facb-563">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7facb-564">1.0</span><span class="sxs-lookup"><span data-stu-id="7facb-564">1.0</span></span>|
|[<span data-ttu-id="7facb-565">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="7facb-565">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7facb-566">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7facb-566">ReadItem</span></span>|
|[<span data-ttu-id="7facb-567">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="7facb-567">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7facb-568">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="7facb-568">Compose or Read</span></span>|

<br>

---
---

#### <a name="subject-stringsubjectjavascriptapioutlookofficesubjectviewoutlook-js-14"></a><span data-ttu-id="7facb-569">Тема: строка | [Subject (тема](/javascript/api/outlook/office.subject?view=outlook-js-1.4) )</span><span class="sxs-lookup"><span data-stu-id="7facb-569">subject: String|[Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.4)</span></span>

<span data-ttu-id="7facb-570">Получает или задает описание, которое отображается в поле темы элемента.</span><span class="sxs-lookup"><span data-stu-id="7facb-570">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="7facb-571">Свойство `subject` получает или задает всю тему элемента для отправки с почтового сервера.</span><span class="sxs-lookup"><span data-stu-id="7facb-571">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="7facb-572">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="7facb-572">Read mode</span></span>

<span data-ttu-id="7facb-p129">Свойство `subject` возвращает строку. С помощью свойства [`normalizedSubject`](#normalizedsubject-string) можно получить тему без начальных префиксов, таких как `RE:` и `FW:`.</span><span class="sxs-lookup"><span data-stu-id="7facb-p129">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```js
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="7facb-575">Режим создания</span><span class="sxs-lookup"><span data-stu-id="7facb-575">Compose mode</span></span>

<span data-ttu-id="7facb-576">Свойство `subject` возвращает объект `Subject`, который предоставляет методы для получения и задания темы.</span><span class="sxs-lookup"><span data-stu-id="7facb-576">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```js
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="7facb-577">Тип</span><span class="sxs-lookup"><span data-stu-id="7facb-577">Type</span></span>

*   <span data-ttu-id="7facb-578">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="7facb-578">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.4)</span></span>

##### <a name="requirements"></a><span data-ttu-id="7facb-579">Требования</span><span class="sxs-lookup"><span data-stu-id="7facb-579">Requirements</span></span>

|<span data-ttu-id="7facb-580">Требование</span><span class="sxs-lookup"><span data-stu-id="7facb-580">Requirement</span></span>| <span data-ttu-id="7facb-581">Значение</span><span class="sxs-lookup"><span data-stu-id="7facb-581">Value</span></span>|
|---|---|
|[<span data-ttu-id="7facb-582">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="7facb-582">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7facb-583">1.0</span><span class="sxs-lookup"><span data-stu-id="7facb-583">1.0</span></span>|
|[<span data-ttu-id="7facb-584">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="7facb-584">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7facb-585">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7facb-585">ReadItem</span></span>|
|[<span data-ttu-id="7facb-586">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="7facb-586">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7facb-587">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="7facb-587">Compose or Read</span></span>|

<br>

---
---

#### <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-14recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-14"></a><span data-ttu-id="7facb-588">Кому: Array. <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)>|[получатели](/javascript/api/outlook/office.recipients?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="7facb-588">to: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.4)</span></span>

<span data-ttu-id="7facb-589">Предоставляет доступ к получателям, указанным в строке **Кому** сообщения.</span><span class="sxs-lookup"><span data-stu-id="7facb-589">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="7facb-590">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="7facb-590">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="7facb-591">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="7facb-591">Read mode</span></span>

<span data-ttu-id="7facb-p131">Свойство `to` возвращает массив, содержащий объект `EmailAddressDetails` для каждого получателя в строке **Кому** сообщения. Коллекция может включать не более 100 элементов.</span><span class="sxs-lookup"><span data-stu-id="7facb-p131">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="7facb-594">Режим создания</span><span class="sxs-lookup"><span data-stu-id="7facb-594">Compose mode</span></span>

<span data-ttu-id="7facb-595">Свойство `to` возвращает объект `Recipients`, предоставляющий методы для получения или обновления получателей, которые указаны в строке **Кому** сообщения.</span><span class="sxs-lookup"><span data-stu-id="7facb-595">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

```js
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="7facb-596">Тип</span><span class="sxs-lookup"><span data-stu-id="7facb-596">Type</span></span>

*   <span data-ttu-id="7facb-597">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="7facb-597">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.4)</span></span>

##### <a name="requirements"></a><span data-ttu-id="7facb-598">Требования</span><span class="sxs-lookup"><span data-stu-id="7facb-598">Requirements</span></span>

|<span data-ttu-id="7facb-599">Требование</span><span class="sxs-lookup"><span data-stu-id="7facb-599">Requirement</span></span>| <span data-ttu-id="7facb-600">Значение</span><span class="sxs-lookup"><span data-stu-id="7facb-600">Value</span></span>|
|---|---|
|[<span data-ttu-id="7facb-601">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="7facb-601">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7facb-602">1.0</span><span class="sxs-lookup"><span data-stu-id="7facb-602">1.0</span></span>|
|[<span data-ttu-id="7facb-603">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="7facb-603">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7facb-604">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7facb-604">ReadItem</span></span>|
|[<span data-ttu-id="7facb-605">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="7facb-605">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7facb-606">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="7facb-606">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="7facb-607">Методы</span><span class="sxs-lookup"><span data-stu-id="7facb-607">Methods</span></span>

#### <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="7facb-608">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="7facb-608">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="7facb-609">Добавляет файл в сообщение или встречу в качестве вложения.</span><span class="sxs-lookup"><span data-stu-id="7facb-609">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="7facb-610">Метод `addFileAttachmentAsync` передает файл по указанному универсальному коду ресурса (URI) и вкладывает его в элемент в форме создания.</span><span class="sxs-lookup"><span data-stu-id="7facb-610">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="7facb-611">Идентификатор можно использовать с методом [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback), чтобы удалить вложение, добавленное во время текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="7facb-611">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="7facb-612">Параметры</span><span class="sxs-lookup"><span data-stu-id="7facb-612">Parameters</span></span>

|<span data-ttu-id="7facb-613">Имя</span><span class="sxs-lookup"><span data-stu-id="7facb-613">Name</span></span>| <span data-ttu-id="7facb-614">Тип</span><span class="sxs-lookup"><span data-stu-id="7facb-614">Type</span></span>| <span data-ttu-id="7facb-615">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="7facb-615">Attributes</span></span>| <span data-ttu-id="7facb-616">Описание</span><span class="sxs-lookup"><span data-stu-id="7facb-616">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="7facb-617">String</span><span class="sxs-lookup"><span data-stu-id="7facb-617">String</span></span>||<span data-ttu-id="7facb-p132">Универсальный код ресурса (URI), представляющий расположение файла, который нужно вложить в сообщение или встречу. Максимальная длина — 2048 символов.</span><span class="sxs-lookup"><span data-stu-id="7facb-p132">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="7facb-620">String</span><span class="sxs-lookup"><span data-stu-id="7facb-620">String</span></span>||<span data-ttu-id="7facb-p133">Имя вложения, которое отображается при передаче вложения. Максимальная длина — 255 символов.</span><span class="sxs-lookup"><span data-stu-id="7facb-p133">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="7facb-623">Object</span><span class="sxs-lookup"><span data-stu-id="7facb-623">Object</span></span>| <span data-ttu-id="7facb-624">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="7facb-624">&lt;optional&gt;</span></span>|<span data-ttu-id="7facb-625">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="7facb-625">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="7facb-626">Object</span><span class="sxs-lookup"><span data-stu-id="7facb-626">Object</span></span>| <span data-ttu-id="7facb-627">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="7facb-627">&lt;optional&gt;</span></span>|<span data-ttu-id="7facb-628">В методе обратного вызова разработчики могут указать любой объект, к которому необходимо получить доступ.</span><span class="sxs-lookup"><span data-stu-id="7facb-628">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="7facb-629">функция</span><span class="sxs-lookup"><span data-stu-id="7facb-629">function</span></span>| <span data-ttu-id="7facb-630">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="7facb-630">&lt;optional&gt;</span></span>|<span data-ttu-id="7facb-631">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="7facb-631">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="7facb-632">После успешного выполнения идентификатор вложения будет представлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="7facb-632">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="7facb-633">Если добавить вложение не удастся, объект `asyncResult` будет содержать объект `Error` с описанием ошибки.</span><span class="sxs-lookup"><span data-stu-id="7facb-633">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="7facb-634">Ошибки</span><span class="sxs-lookup"><span data-stu-id="7facb-634">Errors</span></span>

| <span data-ttu-id="7facb-635">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="7facb-635">Error code</span></span> | <span data-ttu-id="7facb-636">Описание</span><span class="sxs-lookup"><span data-stu-id="7facb-636">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="7facb-637">Вложение превышает максимальный размер.</span><span class="sxs-lookup"><span data-stu-id="7facb-637">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="7facb-638">Расширение вложения не поддерживается.</span><span class="sxs-lookup"><span data-stu-id="7facb-638">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="7facb-639">Сообщение или встреча содержат слишком много вложений.</span><span class="sxs-lookup"><span data-stu-id="7facb-639">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="7facb-640">Требования</span><span class="sxs-lookup"><span data-stu-id="7facb-640">Requirements</span></span>

|<span data-ttu-id="7facb-641">Требование</span><span class="sxs-lookup"><span data-stu-id="7facb-641">Requirement</span></span>| <span data-ttu-id="7facb-642">Значение</span><span class="sxs-lookup"><span data-stu-id="7facb-642">Value</span></span>|
|---|---|
|[<span data-ttu-id="7facb-643">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="7facb-643">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7facb-644">1.1</span><span class="sxs-lookup"><span data-stu-id="7facb-644">1.1</span></span>|
|[<span data-ttu-id="7facb-645">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="7facb-645">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7facb-646">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="7facb-646">ReadWriteItem</span></span>|
|[<span data-ttu-id="7facb-647">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="7facb-647">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7facb-648">Создание</span><span class="sxs-lookup"><span data-stu-id="7facb-648">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="7facb-649">Пример</span><span class="sxs-lookup"><span data-stu-id="7facb-649">Example</span></span>

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

#### <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="7facb-650">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="7facb-650">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="7facb-651">Добавляет к сообщению элемент Exchange, например сообщение, в виде вложения.</span><span class="sxs-lookup"><span data-stu-id="7facb-651">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="7facb-p134">С помощью метода `addItemAttachmentAsync` можно в элемент формы создания вложить элемент с указанным идентификатором Exchange. Если указать метод обратного вызова, то этот метод вызывается с помощью параметра `asyncResult`, который содержит идентификатор вложения или код, указывающий на ошибки, которые произошли при вложении элемента. При необходимости можно использовать параметр `options` для передачи сведений о состоянии методу обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="7facb-p134">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="7facb-655">Идентификатор можно использовать с методом [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback), чтобы удалить вложение, добавленное во время текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="7facb-655">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="7facb-656">Если ваша надстройка Office работает в Outlook в Интернете, `addItemAttachmentAsync` метод может присоединять элементы к элементам, отличным от редактируемого элемента; Однако это не поддерживается и не рекомендуется.</span><span class="sxs-lookup"><span data-stu-id="7facb-656">If your Office Add-in is running in Outlook on the web, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="7facb-657">Параметры</span><span class="sxs-lookup"><span data-stu-id="7facb-657">Parameters</span></span>

|<span data-ttu-id="7facb-658">Имя</span><span class="sxs-lookup"><span data-stu-id="7facb-658">Name</span></span>| <span data-ttu-id="7facb-659">Тип</span><span class="sxs-lookup"><span data-stu-id="7facb-659">Type</span></span>| <span data-ttu-id="7facb-660">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="7facb-660">Attributes</span></span>| <span data-ttu-id="7facb-661">Описание</span><span class="sxs-lookup"><span data-stu-id="7facb-661">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="7facb-662">String</span><span class="sxs-lookup"><span data-stu-id="7facb-662">String</span></span>||<span data-ttu-id="7facb-p135">Идентификатор Exchange для вкладываемого элемента. Максимальная длина — 100 символов.</span><span class="sxs-lookup"><span data-stu-id="7facb-p135">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="7facb-665">String</span><span class="sxs-lookup"><span data-stu-id="7facb-665">String</span></span>||<span data-ttu-id="7facb-666">Тема вкладываемого элемента.</span><span class="sxs-lookup"><span data-stu-id="7facb-666">The subject of the item to be attached.</span></span> <span data-ttu-id="7facb-667">Максимальная длина: 255 символов.</span><span class="sxs-lookup"><span data-stu-id="7facb-667">The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="7facb-668">Object</span><span class="sxs-lookup"><span data-stu-id="7facb-668">Object</span></span>| <span data-ttu-id="7facb-669">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="7facb-669">&lt;optional&gt;</span></span>|<span data-ttu-id="7facb-670">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="7facb-670">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="7facb-671">Объект</span><span class="sxs-lookup"><span data-stu-id="7facb-671">Object</span></span>| <span data-ttu-id="7facb-672">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="7facb-672">&lt;optional&gt;</span></span>|<span data-ttu-id="7facb-673">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="7facb-673">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="7facb-674">функция</span><span class="sxs-lookup"><span data-stu-id="7facb-674">function</span></span>| <span data-ttu-id="7facb-675">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="7facb-675">&lt;optional&gt;</span></span>|<span data-ttu-id="7facb-676">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="7facb-676">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="7facb-677">После успешного выполнения идентификатор вложения будет представлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="7facb-677">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="7facb-678">Если добавить вложение не удастся, объект `asyncResult` будет содержать объект `Error` с описанием ошибки.</span><span class="sxs-lookup"><span data-stu-id="7facb-678">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="7facb-679">Ошибки</span><span class="sxs-lookup"><span data-stu-id="7facb-679">Errors</span></span>

| <span data-ttu-id="7facb-680">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="7facb-680">Error code</span></span> | <span data-ttu-id="7facb-681">Описание</span><span class="sxs-lookup"><span data-stu-id="7facb-681">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="7facb-682">Сообщение или встреча содержат слишком много вложений.</span><span class="sxs-lookup"><span data-stu-id="7facb-682">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="7facb-683">Требования</span><span class="sxs-lookup"><span data-stu-id="7facb-683">Requirements</span></span>

|<span data-ttu-id="7facb-684">Требование</span><span class="sxs-lookup"><span data-stu-id="7facb-684">Requirement</span></span>| <span data-ttu-id="7facb-685">Значение</span><span class="sxs-lookup"><span data-stu-id="7facb-685">Value</span></span>|
|---|---|
|[<span data-ttu-id="7facb-686">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="7facb-686">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7facb-687">1.1</span><span class="sxs-lookup"><span data-stu-id="7facb-687">1.1</span></span>|
|[<span data-ttu-id="7facb-688">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="7facb-688">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7facb-689">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="7facb-689">ReadWriteItem</span></span>|
|[<span data-ttu-id="7facb-690">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="7facb-690">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7facb-691">Создание</span><span class="sxs-lookup"><span data-stu-id="7facb-691">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="7facb-692">Пример</span><span class="sxs-lookup"><span data-stu-id="7facb-692">Example</span></span>

<span data-ttu-id="7facb-693">В следующем примере существующий элемент Outlook добавляется в виде вложения с именем `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="7facb-693">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

#### <a name="close"></a><span data-ttu-id="7facb-694">close()</span><span class="sxs-lookup"><span data-stu-id="7facb-694">close()</span></span>

<span data-ttu-id="7facb-695">Закрывает текущий создаваемый элемент.</span><span class="sxs-lookup"><span data-stu-id="7facb-695">Closes the current item that is being composed.</span></span>

<span data-ttu-id="7facb-p137">Работа метода `close` зависит от текущего состояния создаваемого элемента. Если элемент содержит несохраненные изменения, клиент предложит пользователю сохранить или отклонить их либо отменить действие закрытия.</span><span class="sxs-lookup"><span data-stu-id="7facb-p137">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="7facb-698">Если в Outlook в Интернете элемент представляет собой встречу, ранее сохраненную с помощью метода `saveAsync`, пользователю предлагается сохранить, отклонить или отменить действие, даже если с момента последнего сохранения элемента не вносились какие-либо изменения.</span><span class="sxs-lookup"><span data-stu-id="7facb-698">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="7facb-699">Если в клиенте Outlook для настольных ПК сообщение представляет собой ответ в тексте, метод `close` не работает.</span><span class="sxs-lookup"><span data-stu-id="7facb-699">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="7facb-700">Требования</span><span class="sxs-lookup"><span data-stu-id="7facb-700">Requirements</span></span>

|<span data-ttu-id="7facb-701">Требование</span><span class="sxs-lookup"><span data-stu-id="7facb-701">Requirement</span></span>| <span data-ttu-id="7facb-702">Значение</span><span class="sxs-lookup"><span data-stu-id="7facb-702">Value</span></span>|
|---|---|
|[<span data-ttu-id="7facb-703">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="7facb-703">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7facb-704">1.3</span><span class="sxs-lookup"><span data-stu-id="7facb-704">1.3</span></span>|
|[<span data-ttu-id="7facb-705">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="7facb-705">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7facb-706">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="7facb-706">Restricted</span></span>|
|[<span data-ttu-id="7facb-707">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="7facb-707">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7facb-708">Создание</span><span class="sxs-lookup"><span data-stu-id="7facb-708">Compose</span></span>|

<br>

---
---

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="7facb-709">displayReplyAllForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="7facb-709">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="7facb-710">Отображает форму ответа, включающую отправителя и всех получателей выбранного сообщения или организатора и всех участников выбранной встречи.</span><span class="sxs-lookup"><span data-stu-id="7facb-710">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="7facb-711">Этот метод не поддерживается в Outlook на iOS или Android.</span><span class="sxs-lookup"><span data-stu-id="7facb-711">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="7facb-712">В Outlook в Интернете форма ответа отображается в виде всплывающей формы в представлении из трех столбцов и всплывающей формы в представлении с 2 или 1 столбца.</span><span class="sxs-lookup"><span data-stu-id="7facb-712">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="7facb-713">Если любой строковый параметр превышает указанные для него ограничения, `displayReplyAllForm` возвращает исключение.</span><span class="sxs-lookup"><span data-stu-id="7facb-713">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="7facb-714">Если в `formData.attachments` параметре указаны вложения, Outlook в Интернете и клиенте для настольных компьютеров пытаются скачать все вложения и присоединить их к форме ответа.</span><span class="sxs-lookup"><span data-stu-id="7facb-714">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form.</span></span> <span data-ttu-id="7facb-715">Если какие-либо вложения добавить не удается, в форме отображается сообщение об ошибке.</span><span class="sxs-lookup"><span data-stu-id="7facb-715">If any attachments fail to be added, an error is shown in the form UI.</span></span> <span data-ttu-id="7facb-716">Если сообщения об ошибках не предусмотрены, то они не отображаются.</span><span class="sxs-lookup"><span data-stu-id="7facb-716">If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="7facb-717">Параметры</span><span class="sxs-lookup"><span data-stu-id="7facb-717">Parameters</span></span>

|<span data-ttu-id="7facb-718">Имя</span><span class="sxs-lookup"><span data-stu-id="7facb-718">Name</span></span>| <span data-ttu-id="7facb-719">Тип</span><span class="sxs-lookup"><span data-stu-id="7facb-719">Type</span></span>| <span data-ttu-id="7facb-720">Описание</span><span class="sxs-lookup"><span data-stu-id="7facb-720">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="7facb-721">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="7facb-721">String &#124; Object</span></span>| |<span data-ttu-id="7facb-p139">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="7facb-p139">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="7facb-724">**ИЛИ**</span><span class="sxs-lookup"><span data-stu-id="7facb-724">**OR**</span></span><br/><span data-ttu-id="7facb-p140">Объект, который содержит текст или данные вложения и функцию обратного вызова. Ниже представлено определение этого объекта.</span><span class="sxs-lookup"><span data-stu-id="7facb-p140">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="7facb-727">String.</span><span class="sxs-lookup"><span data-stu-id="7facb-727">String</span></span> | <span data-ttu-id="7facb-728">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="7facb-728">&lt;optional&gt;</span></span> | <span data-ttu-id="7facb-p141">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="7facb-p141">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="7facb-731">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="7facb-731">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="7facb-732">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="7facb-732">&lt;optional&gt;</span></span> | <span data-ttu-id="7facb-733">Массив объектов JSON, представляющих собой вложенные файлы или элементы.</span><span class="sxs-lookup"><span data-stu-id="7facb-733">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="7facb-734">String.</span><span class="sxs-lookup"><span data-stu-id="7facb-734">String</span></span> | | <span data-ttu-id="7facb-p142">Указывает тип вложения. Допустимые значения: `file` для вложенного файла и `item` для вложенного элемента.</span><span class="sxs-lookup"><span data-stu-id="7facb-p142">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="7facb-737">Строка</span><span class="sxs-lookup"><span data-stu-id="7facb-737">String</span></span> | | <span data-ttu-id="7facb-738">Строка, содержащая имя вложения, длиной до 255 символов.</span><span class="sxs-lookup"><span data-stu-id="7facb-738">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="7facb-739">String</span><span class="sxs-lookup"><span data-stu-id="7facb-739">String</span></span> | | <span data-ttu-id="7facb-p143">Используется, только если свойству `type` задано значение `file`. URI расположения файла.</span><span class="sxs-lookup"><span data-stu-id="7facb-p143">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="7facb-742">String</span><span class="sxs-lookup"><span data-stu-id="7facb-742">String</span></span> | | <span data-ttu-id="7facb-p144">Используется, только если свойству `type` задано значение `item`. Идентификатор вложения EWS. Это строка длиной до 100 символов.</span><span class="sxs-lookup"><span data-stu-id="7facb-p144">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="7facb-746">function</span><span class="sxs-lookup"><span data-stu-id="7facb-746">function</span></span> | <span data-ttu-id="7facb-747">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="7facb-747">&lt;optional&gt;</span></span> | <span data-ttu-id="7facb-748">По завершении работы метода функция, переданная параметру `callback`, вызывается с помощью одного параметра `asyncResult`, представляющего собой объект [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="7facb-748">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="7facb-749">Требования</span><span class="sxs-lookup"><span data-stu-id="7facb-749">Requirements</span></span>

|<span data-ttu-id="7facb-750">Требование</span><span class="sxs-lookup"><span data-stu-id="7facb-750">Requirement</span></span>| <span data-ttu-id="7facb-751">Значение</span><span class="sxs-lookup"><span data-stu-id="7facb-751">Value</span></span>|
|---|---|
|[<span data-ttu-id="7facb-752">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="7facb-752">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7facb-753">1.0</span><span class="sxs-lookup"><span data-stu-id="7facb-753">1.0</span></span>|
|[<span data-ttu-id="7facb-754">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="7facb-754">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7facb-755">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7facb-755">ReadItem</span></span>|
|[<span data-ttu-id="7facb-756">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="7facb-756">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7facb-757">Чтение</span><span class="sxs-lookup"><span data-stu-id="7facb-757">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="7facb-758">Примеры</span><span class="sxs-lookup"><span data-stu-id="7facb-758">Examples</span></span>

<span data-ttu-id="7facb-759">Приведенный ниже код передает строку в функцию `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="7facb-759">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="7facb-760">Ответ с пустым текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="7facb-760">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="7facb-761">Ответ только с текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="7facb-761">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="7facb-762">Ответ с текстом сообщения и вложенным файлом.</span><span class="sxs-lookup"><span data-stu-id="7facb-762">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="7facb-763">Ответ с текстом сообщения и вложенным элементом.</span><span class="sxs-lookup"><span data-stu-id="7facb-763">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="7facb-764">Ответ с текстом сообщения, вложенным файлом, вложенным элементом и обратным вызовом.</span><span class="sxs-lookup"><span data-stu-id="7facb-764">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="7facb-765">displayReplyForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="7facb-765">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="7facb-766">Отображает форму ответа, включающую только отправителя выбранного сообщения или организатора выбранной встречи.</span><span class="sxs-lookup"><span data-stu-id="7facb-766">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="7facb-767">Этот метод не поддерживается в Outlook на iOS или Android.</span><span class="sxs-lookup"><span data-stu-id="7facb-767">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="7facb-768">В Outlook в Интернете форма ответа отображается в виде всплывающей формы в представлении из трех столбцов и всплывающей формы в представлении с 2 или 1 столбца.</span><span class="sxs-lookup"><span data-stu-id="7facb-768">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="7facb-769">Если любой строковый параметр превышает указанные для него ограничения, `displayReplyForm` возвращает исключение.</span><span class="sxs-lookup"><span data-stu-id="7facb-769">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="7facb-770">Если в `formData.attachments` параметре указаны вложения, Outlook в Интернете и клиенте для настольных компьютеров пытаются скачать все вложения и присоединить их к форме ответа.</span><span class="sxs-lookup"><span data-stu-id="7facb-770">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form.</span></span> <span data-ttu-id="7facb-771">Если какие-либо вложения добавить не удается, в форме отображается сообщение об ошибке.</span><span class="sxs-lookup"><span data-stu-id="7facb-771">If any attachments fail to be added, an error is shown in the form UI.</span></span> <span data-ttu-id="7facb-772">Если сообщения об ошибках не предусмотрены, то они не отображаются.</span><span class="sxs-lookup"><span data-stu-id="7facb-772">If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="7facb-773">Параметры</span><span class="sxs-lookup"><span data-stu-id="7facb-773">Parameters</span></span>

|<span data-ttu-id="7facb-774">Имя</span><span class="sxs-lookup"><span data-stu-id="7facb-774">Name</span></span>| <span data-ttu-id="7facb-775">Тип</span><span class="sxs-lookup"><span data-stu-id="7facb-775">Type</span></span>| <span data-ttu-id="7facb-776">Описание</span><span class="sxs-lookup"><span data-stu-id="7facb-776">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="7facb-777">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="7facb-777">String &#124; Object</span></span>| | <span data-ttu-id="7facb-p146">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="7facb-p146">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="7facb-780">**ИЛИ**</span><span class="sxs-lookup"><span data-stu-id="7facb-780">**OR**</span></span><br/><span data-ttu-id="7facb-p147">Объект, который содержит текст или данные вложения и функцию обратного вызова. Ниже представлено определение этого объекта.</span><span class="sxs-lookup"><span data-stu-id="7facb-p147">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="7facb-783">String.</span><span class="sxs-lookup"><span data-stu-id="7facb-783">String</span></span> | <span data-ttu-id="7facb-784">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="7facb-784">&lt;optional&gt;</span></span> | <span data-ttu-id="7facb-p148">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="7facb-p148">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="7facb-787">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="7facb-787">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="7facb-788">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="7facb-788">&lt;optional&gt;</span></span> | <span data-ttu-id="7facb-789">Массив объектов JSON, представляющих собой вложенные файлы или элементы.</span><span class="sxs-lookup"><span data-stu-id="7facb-789">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="7facb-790">String.</span><span class="sxs-lookup"><span data-stu-id="7facb-790">String</span></span> | | <span data-ttu-id="7facb-p149">Указывает тип вложения. Допустимые значения: `file` для вложенного файла и `item` для вложенного элемента.</span><span class="sxs-lookup"><span data-stu-id="7facb-p149">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="7facb-793">Строка</span><span class="sxs-lookup"><span data-stu-id="7facb-793">String</span></span> | | <span data-ttu-id="7facb-794">Строка, содержащая имя вложения, длиной до 255 символов.</span><span class="sxs-lookup"><span data-stu-id="7facb-794">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="7facb-795">String</span><span class="sxs-lookup"><span data-stu-id="7facb-795">String</span></span> | | <span data-ttu-id="7facb-p150">Используется, только если свойству `type` задано значение `file`. URI расположения файла.</span><span class="sxs-lookup"><span data-stu-id="7facb-p150">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="7facb-798">String</span><span class="sxs-lookup"><span data-stu-id="7facb-798">String</span></span> | | <span data-ttu-id="7facb-p151">Используется, только если свойству `type` задано значение `item`. Идентификатор вложения EWS. Это строка длиной до 100 символов.</span><span class="sxs-lookup"><span data-stu-id="7facb-p151">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="7facb-802">function</span><span class="sxs-lookup"><span data-stu-id="7facb-802">function</span></span> | <span data-ttu-id="7facb-803">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="7facb-803">&lt;optional&gt;</span></span> | <span data-ttu-id="7facb-804">По завершении работы метода функция, переданная параметру `callback`, вызывается с помощью одного параметра `asyncResult`, представляющего собой объект [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="7facb-804">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="7facb-805">Требования</span><span class="sxs-lookup"><span data-stu-id="7facb-805">Requirements</span></span>

|<span data-ttu-id="7facb-806">Требование</span><span class="sxs-lookup"><span data-stu-id="7facb-806">Requirement</span></span>| <span data-ttu-id="7facb-807">Значение</span><span class="sxs-lookup"><span data-stu-id="7facb-807">Value</span></span>|
|---|---|
|[<span data-ttu-id="7facb-808">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="7facb-808">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7facb-809">1.0</span><span class="sxs-lookup"><span data-stu-id="7facb-809">1.0</span></span>|
|[<span data-ttu-id="7facb-810">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="7facb-810">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7facb-811">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7facb-811">ReadItem</span></span>|
|[<span data-ttu-id="7facb-812">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="7facb-812">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7facb-813">Чтение</span><span class="sxs-lookup"><span data-stu-id="7facb-813">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="7facb-814">Примеры</span><span class="sxs-lookup"><span data-stu-id="7facb-814">Examples</span></span>

<span data-ttu-id="7facb-815">Приведенный ниже код передает строку в функцию `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="7facb-815">The following code passes a string to the `displayReplyForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="7facb-816">Ответ с пустым текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="7facb-816">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="7facb-817">Ответ только с текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="7facb-817">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="7facb-818">Ответ с текстом сообщения и вложенным файлом.</span><span class="sxs-lookup"><span data-stu-id="7facb-818">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="7facb-819">Ответ с текстом сообщения и вложенным элементом.</span><span class="sxs-lookup"><span data-stu-id="7facb-819">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="7facb-820">Ответ с текстом сообщения, вложенным файлом, вложенным элементом и обратным вызовом.</span><span class="sxs-lookup"><span data-stu-id="7facb-820">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlookofficeentitiesviewoutlook-js-14"></a><span data-ttu-id="7facb-821">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.4)}</span><span class="sxs-lookup"><span data-stu-id="7facb-821">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.4)}</span></span>

<span data-ttu-id="7facb-822">Получает сущности, обнаруженные в теле выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="7facb-822">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="7facb-823">Этот метод не поддерживается в Outlook на iOS или Android.</span><span class="sxs-lookup"><span data-stu-id="7facb-823">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="7facb-824">Требования</span><span class="sxs-lookup"><span data-stu-id="7facb-824">Requirements</span></span>

|<span data-ttu-id="7facb-825">Требование</span><span class="sxs-lookup"><span data-stu-id="7facb-825">Requirement</span></span>| <span data-ttu-id="7facb-826">Значение</span><span class="sxs-lookup"><span data-stu-id="7facb-826">Value</span></span>|
|---|---|
|[<span data-ttu-id="7facb-827">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="7facb-827">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7facb-828">1.0</span><span class="sxs-lookup"><span data-stu-id="7facb-828">1.0</span></span>|
|[<span data-ttu-id="7facb-829">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="7facb-829">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7facb-830">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7facb-830">ReadItem</span></span>|
|[<span data-ttu-id="7facb-831">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="7facb-831">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7facb-832">Чтение</span><span class="sxs-lookup"><span data-stu-id="7facb-832">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="7facb-833">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="7facb-833">Returns:</span></span>

<span data-ttu-id="7facb-834">Тип: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="7facb-834">Type: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.4)</span></span>

##### <a name="example"></a><span data-ttu-id="7facb-835">Пример</span><span class="sxs-lookup"><span data-stu-id="7facb-835">Example</span></span>

<span data-ttu-id="7facb-836">Ниже приведен пример получения доступа к сущностям контактов в теле текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="7facb-836">The following example accesses the contacts entities in the current item's body.</span></span>

```js
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

<br>

---
---

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-14meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-14phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-14tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-14"></a><span data-ttu-id="7facb-837">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.4)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.4)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.4)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.4))>}</span><span class="sxs-lookup"><span data-stu-id="7facb-837">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.4)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.4)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.4)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.4))>}</span></span>

<span data-ttu-id="7facb-838">Получает массив всех сущностей указанного типа, обнаруженных в теле выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="7facb-838">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="7facb-839">Этот метод не поддерживается в Outlook на iOS или Android.</span><span class="sxs-lookup"><span data-stu-id="7facb-839">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="7facb-840">Параметры</span><span class="sxs-lookup"><span data-stu-id="7facb-840">Parameters</span></span>

|<span data-ttu-id="7facb-841">Имя</span><span class="sxs-lookup"><span data-stu-id="7facb-841">Name</span></span>| <span data-ttu-id="7facb-842">Тип</span><span class="sxs-lookup"><span data-stu-id="7facb-842">Type</span></span>| <span data-ttu-id="7facb-843">Описание</span><span class="sxs-lookup"><span data-stu-id="7facb-843">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="7facb-844">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="7facb-844">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook/office.mailboxenums.entitytype?view=outlook-js-1.4)|<span data-ttu-id="7facb-845">Одно из значений перечисления EntityType.</span><span class="sxs-lookup"><span data-stu-id="7facb-845">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="7facb-846">Требования</span><span class="sxs-lookup"><span data-stu-id="7facb-846">Requirements</span></span>

|<span data-ttu-id="7facb-847">Требование</span><span class="sxs-lookup"><span data-stu-id="7facb-847">Requirement</span></span>| <span data-ttu-id="7facb-848">Значение</span><span class="sxs-lookup"><span data-stu-id="7facb-848">Value</span></span>|
|---|---|
|[<span data-ttu-id="7facb-849">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="7facb-849">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7facb-850">1.0</span><span class="sxs-lookup"><span data-stu-id="7facb-850">1.0</span></span>|
|[<span data-ttu-id="7facb-851">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="7facb-851">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7facb-852">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="7facb-852">Restricted</span></span>|
|[<span data-ttu-id="7facb-853">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="7facb-853">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7facb-854">Чтение</span><span class="sxs-lookup"><span data-stu-id="7facb-854">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="7facb-855">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="7facb-855">Returns:</span></span>

<span data-ttu-id="7facb-856">Если значение, переданное в `entityType`, не является допустимым членом перечисления `EntityType`, метод возвращает значение NULL.</span><span class="sxs-lookup"><span data-stu-id="7facb-856">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="7facb-857">Если в теле элемента отсутствуют сущности указанного типа, метод возвращает пустой массив.</span><span class="sxs-lookup"><span data-stu-id="7facb-857">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="7facb-858">В противном случае тип объектов в возвращаемом массиве зависит от типа сущности, запрошенной в параметре `entityType`.</span><span class="sxs-lookup"><span data-stu-id="7facb-858">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="7facb-859">Хотя минимальный уровень разрешений для использования этого метода — **Restricted**, для некоторых типов сущностей требуется доступ на уровне **ReadItem**, как указано в приведенной ниже таблице.</span><span class="sxs-lookup"><span data-stu-id="7facb-859">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="7facb-860">Значение параметра `entityType`</span><span class="sxs-lookup"><span data-stu-id="7facb-860">Value of `entityType`</span></span> | <span data-ttu-id="7facb-861">Тип объектов в возвращаемом массиве</span><span class="sxs-lookup"><span data-stu-id="7facb-861">Type of objects in returned array</span></span> | <span data-ttu-id="7facb-862">Необходимый уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="7facb-862">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="7facb-863">String</span><span class="sxs-lookup"><span data-stu-id="7facb-863">String</span></span> | <span data-ttu-id="7facb-864">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="7facb-864">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="7facb-865">Contact</span><span class="sxs-lookup"><span data-stu-id="7facb-865">Contact</span></span> | <span data-ttu-id="7facb-866">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="7facb-866">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="7facb-867">String</span><span class="sxs-lookup"><span data-stu-id="7facb-867">String</span></span> | <span data-ttu-id="7facb-868">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="7facb-868">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="7facb-869">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="7facb-869">MeetingSuggestion</span></span> | <span data-ttu-id="7facb-870">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="7facb-870">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="7facb-871">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="7facb-871">PhoneNumber</span></span> | <span data-ttu-id="7facb-872">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="7facb-872">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="7facb-873">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="7facb-873">TaskSuggestion</span></span> | <span data-ttu-id="7facb-874">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="7facb-874">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="7facb-875">String</span><span class="sxs-lookup"><span data-stu-id="7facb-875">String</span></span> | <span data-ttu-id="7facb-876">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="7facb-876">**Restricted**</span></span> |

<span data-ttu-id="7facb-877">Тип: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.4)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.4)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.4)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.4))></span><span class="sxs-lookup"><span data-stu-id="7facb-877">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.4)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.4)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.4)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.4))></span></span>

##### <a name="example"></a><span data-ttu-id="7facb-878">Пример</span><span class="sxs-lookup"><span data-stu-id="7facb-878">Example</span></span>

<span data-ttu-id="7facb-879">В примере ниже показано, как получить доступ к массиву строк, которые представляют собой почтовые адреса в теле текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="7facb-879">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-14meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-14phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-14tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-14"></a><span data-ttu-id="7facb-880">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.4)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.4)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.4)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.4))>}</span><span class="sxs-lookup"><span data-stu-id="7facb-880">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.4)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.4)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.4)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.4))>}</span></span>

<span data-ttu-id="7facb-881">Возвращает известные сущности в выбранном элементе, которые проходят через именованный фильтр, определяемый в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="7facb-881">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="7facb-882">Этот метод не поддерживается в Outlook на iOS или Android.</span><span class="sxs-lookup"><span data-stu-id="7facb-882">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="7facb-883">Метод `getFilteredEntitiesByName` возвращает сущности, соответствующие регулярному выражению, которое определяется в элементе правила [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) в XML-файле манифеста, с использованием указанного значения элемента `FilterName`.</span><span class="sxs-lookup"><span data-stu-id="7facb-883">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="7facb-884">Параметры</span><span class="sxs-lookup"><span data-stu-id="7facb-884">Parameters</span></span>

|<span data-ttu-id="7facb-885">Имя</span><span class="sxs-lookup"><span data-stu-id="7facb-885">Name</span></span>| <span data-ttu-id="7facb-886">Тип</span><span class="sxs-lookup"><span data-stu-id="7facb-886">Type</span></span>| <span data-ttu-id="7facb-887">Описание</span><span class="sxs-lookup"><span data-stu-id="7facb-887">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="7facb-888">String</span><span class="sxs-lookup"><span data-stu-id="7facb-888">String</span></span>|<span data-ttu-id="7facb-889">Имя элемента правила `ItemHasKnownEntity`, определяющее соответствующий фильтр.</span><span class="sxs-lookup"><span data-stu-id="7facb-889">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="7facb-890">Требования</span><span class="sxs-lookup"><span data-stu-id="7facb-890">Requirements</span></span>

|<span data-ttu-id="7facb-891">Требование</span><span class="sxs-lookup"><span data-stu-id="7facb-891">Requirement</span></span>| <span data-ttu-id="7facb-892">Значение</span><span class="sxs-lookup"><span data-stu-id="7facb-892">Value</span></span>|
|---|---|
|[<span data-ttu-id="7facb-893">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="7facb-893">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7facb-894">1.0</span><span class="sxs-lookup"><span data-stu-id="7facb-894">1.0</span></span>|
|[<span data-ttu-id="7facb-895">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="7facb-895">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7facb-896">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7facb-896">ReadItem</span></span>|
|[<span data-ttu-id="7facb-897">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="7facb-897">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7facb-898">Чтение</span><span class="sxs-lookup"><span data-stu-id="7facb-898">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="7facb-899">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="7facb-899">Returns:</span></span>

<span data-ttu-id="7facb-p153">Если в манифесте нет элемента `ItemHasKnownEntity` со значением `FilterName`, соответствующим параметру `name`, метод возвращает `null`. Если параметр `name` соответствует элементу `ItemHasKnownEntity` в манифесте, но при этом в текущем элементе нет соответствующих сущностей, метод возвращает пустой массив.</span><span class="sxs-lookup"><span data-stu-id="7facb-p153">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="7facb-902">Тип: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.4)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.4)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.4)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.4))></span><span class="sxs-lookup"><span data-stu-id="7facb-902">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.4)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.4)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.4)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.4))></span></span>

<br>

---
---

#### <a name="getregexmatches--object"></a><span data-ttu-id="7facb-903">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="7facb-903">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="7facb-904">Возвращает строковые значения в выбранном элементе, которые соответствуют регулярным выражениям, определенным в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="7facb-904">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="7facb-905">Этот метод не поддерживается в Outlook на iOS или Android.</span><span class="sxs-lookup"><span data-stu-id="7facb-905">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="7facb-p154">Метод `getRegExMatches` возвращает строки, соответствующие регулярному выражению, которое определяется в каждом элементе правила `ItemHasRegularExpressionMatch` или `ItemHasKnownEntity` в XML-файле манифеста. Для правила `ItemHasRegularExpressionMatch` соответствующую строку должно содержать свойство элемента, указанного этим правилом. Простой тип `PropertyName` определяет поддерживаемые свойства.</span><span class="sxs-lookup"><span data-stu-id="7facb-p154">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="7facb-909">Например, рассмотрим манифест надстройки, который содержит указанный ниже элемент `Rule`.</span><span class="sxs-lookup"><span data-stu-id="7facb-909">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="7facb-910">Объект, возвращаемый методом `getRegExMatches`, будет содержать два свойства: `fruits` и `veggies`.</span><span class="sxs-lookup"><span data-stu-id="7facb-910">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="7facb-p155">Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты. Лучше используйте метод [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.4#getasync-coerciontype--options--callback-) для этого.</span><span class="sxs-lookup"><span data-stu-id="7facb-p155">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.4#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="7facb-914">Requirements</span><span class="sxs-lookup"><span data-stu-id="7facb-914">Requirements</span></span>

|<span data-ttu-id="7facb-915">Требование</span><span class="sxs-lookup"><span data-stu-id="7facb-915">Requirement</span></span>| <span data-ttu-id="7facb-916">Значение</span><span class="sxs-lookup"><span data-stu-id="7facb-916">Value</span></span>|
|---|---|
|[<span data-ttu-id="7facb-917">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="7facb-917">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7facb-918">1.0</span><span class="sxs-lookup"><span data-stu-id="7facb-918">1.0</span></span>|
|[<span data-ttu-id="7facb-919">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="7facb-919">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7facb-920">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7facb-920">ReadItem</span></span>|
|[<span data-ttu-id="7facb-921">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="7facb-921">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7facb-922">Чтение</span><span class="sxs-lookup"><span data-stu-id="7facb-922">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="7facb-923">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="7facb-923">Returns:</span></span>

<span data-ttu-id="7facb-p156">Объект, содержащий массив строк, которые соответствуют регулярным выражениям, определяемым в XML-файле манифеста. Имя каждого массива равно соответствующему значению атрибута `RegExName` подходящего правила `ItemHasRegularExpressionMatch` или атрибута `FilterName` соответствующего правила `ItemHasKnownEntity`.</span><span class="sxs-lookup"><span data-stu-id="7facb-p156">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<span data-ttu-id="7facb-926">Тип: Object</span><span class="sxs-lookup"><span data-stu-id="7facb-926">Type: Object</span></span>

##### <a name="example"></a><span data-ttu-id="7facb-927">Пример</span><span class="sxs-lookup"><span data-stu-id="7facb-927">Example</span></span>

<span data-ttu-id="7facb-928">В примере ниже показано, как получить доступ к массиву совпадений для <rule>элементов регулярного выражения `fruits` и `veggies`, которые указаны в манифесте</rule>.</span><span class="sxs-lookup"><span data-stu-id="7facb-928">The following example shows how to access the array of matches for the regular expression <rule>elements `fruits` and `veggies`, which are specified in the manifest.</rule></span></span>

```js
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

<br>

---
---

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="7facb-929">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="7facb-929">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="7facb-930">Возвращает строковые значения в выбранном элементе, которые соответствуют именованному регулярному выражению, определенному в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="7facb-930">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="7facb-931">Этот метод не поддерживается в Outlook на iOS или Android.</span><span class="sxs-lookup"><span data-stu-id="7facb-931">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="7facb-932">Метод `getRegExMatchesByName` возвращает строки, соответствующие регулярному выражению, которое определяется в элементе правила `ItemHasRegularExpressionMatch` в XML-файле манифеста, с использованием указанного значения элемента `RegExName`.</span><span class="sxs-lookup"><span data-stu-id="7facb-932">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="7facb-p157">Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты.</span><span class="sxs-lookup"><span data-stu-id="7facb-p157">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="7facb-935">Параметры</span><span class="sxs-lookup"><span data-stu-id="7facb-935">Parameters</span></span>

|<span data-ttu-id="7facb-936">Имя</span><span class="sxs-lookup"><span data-stu-id="7facb-936">Name</span></span>| <span data-ttu-id="7facb-937">Тип</span><span class="sxs-lookup"><span data-stu-id="7facb-937">Type</span></span>| <span data-ttu-id="7facb-938">Описание</span><span class="sxs-lookup"><span data-stu-id="7facb-938">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="7facb-939">String</span><span class="sxs-lookup"><span data-stu-id="7facb-939">String</span></span>|<span data-ttu-id="7facb-940">Имя элемента правила `ItemHasRegularExpressionMatch`, определяющее соответствующий фильтр.</span><span class="sxs-lookup"><span data-stu-id="7facb-940">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="7facb-941">Требования</span><span class="sxs-lookup"><span data-stu-id="7facb-941">Requirements</span></span>

|<span data-ttu-id="7facb-942">Требование</span><span class="sxs-lookup"><span data-stu-id="7facb-942">Requirement</span></span>| <span data-ttu-id="7facb-943">Значение</span><span class="sxs-lookup"><span data-stu-id="7facb-943">Value</span></span>|
|---|---|
|[<span data-ttu-id="7facb-944">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="7facb-944">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7facb-945">1.0</span><span class="sxs-lookup"><span data-stu-id="7facb-945">1.0</span></span>|
|[<span data-ttu-id="7facb-946">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="7facb-946">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7facb-947">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7facb-947">ReadItem</span></span>|
|[<span data-ttu-id="7facb-948">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="7facb-948">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7facb-949">Чтение</span><span class="sxs-lookup"><span data-stu-id="7facb-949">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="7facb-950">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="7facb-950">Returns:</span></span>

<span data-ttu-id="7facb-951">Массив строк, соответствующих регулярному выражению, определяемому в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="7facb-951">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<span data-ttu-id="7facb-952">Тип: Array. < String ></span><span class="sxs-lookup"><span data-stu-id="7facb-952">Type: Array.< String ></span></span>

##### <a name="example"></a><span data-ttu-id="7facb-953">Пример</span><span class="sxs-lookup"><span data-stu-id="7facb-953">Example</span></span>

```js
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

<br>

---
---

#### <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="7facb-954">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="7facb-954">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="7facb-955">Асинхронно возвращает данные, выбранные в теме или тексте сообщения.</span><span class="sxs-lookup"><span data-stu-id="7facb-955">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="7facb-p158">Если выделенный фрагмент отсутствует, но курсор находится в тексте или теме, метод возвращает значение NULL для выбранных данных. Если выбраны не текст и не тема, метод возвращает ошибку `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="7facb-p158">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="7facb-958">Параметры</span><span class="sxs-lookup"><span data-stu-id="7facb-958">Parameters</span></span>

|<span data-ttu-id="7facb-959">Имя</span><span class="sxs-lookup"><span data-stu-id="7facb-959">Name</span></span>| <span data-ttu-id="7facb-960">Тип</span><span class="sxs-lookup"><span data-stu-id="7facb-960">Type</span></span>| <span data-ttu-id="7facb-961">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="7facb-961">Attributes</span></span>| <span data-ttu-id="7facb-962">Описание</span><span class="sxs-lookup"><span data-stu-id="7facb-962">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="7facb-963">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="7facb-963">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="7facb-p159">Запрашивает формат данных. Если задано значение Text, метод возвращает обычный текст как строку, удаляя все имеющиеся HTML-теги. Если задано значение HTML, метод возвращает выделенный текст (обычный текст или HTML).</span><span class="sxs-lookup"><span data-stu-id="7facb-p159">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="7facb-967">Объект</span><span class="sxs-lookup"><span data-stu-id="7facb-967">Object</span></span>| <span data-ttu-id="7facb-968">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="7facb-968">&lt;optional&gt;</span></span>|<span data-ttu-id="7facb-969">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="7facb-969">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="7facb-970">Объект</span><span class="sxs-lookup"><span data-stu-id="7facb-970">Object</span></span>| <span data-ttu-id="7facb-971">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="7facb-971">&lt;optional&gt;</span></span>|<span data-ttu-id="7facb-972">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="7facb-972">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="7facb-973">функция</span><span class="sxs-lookup"><span data-stu-id="7facb-973">function</span></span>||<span data-ttu-id="7facb-974">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="7facb-974">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="7facb-975">Чтобы получить доступ к выбранным данным из метода обратного вызова, вызовите `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="7facb-975">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="7facb-976">Для доступа к исходному свойству, представляющему собой источник выбранных данных, вызовите параметр `asyncResult.value.sourceProperty`, который может иметь значение `body` или `subject`.</span><span class="sxs-lookup"><span data-stu-id="7facb-976">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="7facb-977">Требования</span><span class="sxs-lookup"><span data-stu-id="7facb-977">Requirements</span></span>

|<span data-ttu-id="7facb-978">Требование</span><span class="sxs-lookup"><span data-stu-id="7facb-978">Requirement</span></span>| <span data-ttu-id="7facb-979">Значение</span><span class="sxs-lookup"><span data-stu-id="7facb-979">Value</span></span>|
|---|---|
|[<span data-ttu-id="7facb-980">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="7facb-980">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7facb-981">1.2</span><span class="sxs-lookup"><span data-stu-id="7facb-981">1.2</span></span>|
|[<span data-ttu-id="7facb-982">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="7facb-982">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7facb-983">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7facb-983">ReadItem</span></span>|
|[<span data-ttu-id="7facb-984">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="7facb-984">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7facb-985">Создание</span><span class="sxs-lookup"><span data-stu-id="7facb-985">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="7facb-986">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="7facb-986">Returns:</span></span>

<span data-ttu-id="7facb-987">Выбранные данные в виде строки с форматом, определенным в параметре `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="7facb-987">The selected data as a string with format determined by `coercionType`.</span></span>

<span data-ttu-id="7facb-988">Тип: String</span><span class="sxs-lookup"><span data-stu-id="7facb-988">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="7facb-989">Пример</span><span class="sxs-lookup"><span data-stu-id="7facb-989">Example</span></span>

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

#### <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="7facb-990">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="7facb-990">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="7facb-991">Асинхронно загружает настраиваемые свойства для надстройки для выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="7facb-991">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="7facb-p161">Настраиваемые свойства сохраняются в виде пар "ключ-значение" для каждого приложения и каждого элемента. Этот метод возвращает объект `CustomProperties` при обратном вызове, который предоставляет методы для доступа к настраиваемым свойствам, характерным для текущего элемента и текущей надстройки. Настраиваемые свойства не шифруются для элемента, поэтому этот способ хранения не является безопасным.</span><span class="sxs-lookup"><span data-stu-id="7facb-p161">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="7facb-995">Параметры</span><span class="sxs-lookup"><span data-stu-id="7facb-995">Parameters</span></span>

|<span data-ttu-id="7facb-996">Имя</span><span class="sxs-lookup"><span data-stu-id="7facb-996">Name</span></span>| <span data-ttu-id="7facb-997">Тип</span><span class="sxs-lookup"><span data-stu-id="7facb-997">Type</span></span>| <span data-ttu-id="7facb-998">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="7facb-998">Attributes</span></span>| <span data-ttu-id="7facb-999">Описание</span><span class="sxs-lookup"><span data-stu-id="7facb-999">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="7facb-1000">function</span><span class="sxs-lookup"><span data-stu-id="7facb-1000">function</span></span>||<span data-ttu-id="7facb-1001">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="7facb-1001">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="7facb-1002">Настраиваемые свойства предоставляются в виде объекта [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.4) в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="7facb-1002">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.4) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="7facb-1003">Этот объект позволяет получить, задать и удалить настраиваемые свойства для элемента, а также сохранить изменения, внесенные в набор настраиваемых свойств, на сервере.</span><span class="sxs-lookup"><span data-stu-id="7facb-1003">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="7facb-1004">Объект</span><span class="sxs-lookup"><span data-stu-id="7facb-1004">Object</span></span>| <span data-ttu-id="7facb-1005">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="7facb-1005">&lt;optional&gt;</span></span>|<span data-ttu-id="7facb-1006">Разработчики могут указать любой объект, к которому необходимо получить доступ, в функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="7facb-1006">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="7facb-1007">Доступ к этому объекту можно получить с помощью свойства `asyncResult.asyncContext` в функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="7facb-1007">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="7facb-1008">Требования</span><span class="sxs-lookup"><span data-stu-id="7facb-1008">Requirements</span></span>

|<span data-ttu-id="7facb-1009">Требование</span><span class="sxs-lookup"><span data-stu-id="7facb-1009">Requirement</span></span>| <span data-ttu-id="7facb-1010">Значение</span><span class="sxs-lookup"><span data-stu-id="7facb-1010">Value</span></span>|
|---|---|
|[<span data-ttu-id="7facb-1011">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="7facb-1011">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7facb-1012">1.0</span><span class="sxs-lookup"><span data-stu-id="7facb-1012">1.0</span></span>|
|[<span data-ttu-id="7facb-1013">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="7facb-1013">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7facb-1014">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7facb-1014">ReadItem</span></span>|
|[<span data-ttu-id="7facb-1015">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="7facb-1015">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7facb-1016">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="7facb-1016">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="7facb-1017">Пример</span><span class="sxs-lookup"><span data-stu-id="7facb-1017">Example</span></span>

<span data-ttu-id="7facb-p164">Приведенный ниже пример кода показывает, как асинхронно загружать настраиваемые свойства, характерные для текущего элемента, с помощью метода `loadCustomPropertiesAsync`. Этот пример также показывает, как сохранять эти свойства на сервере с помощью метода `CustomProperties.saveAsync`. После загрузки настраиваемых свойств в этом примере кода метод `CustomProperties.get` используется для считывания настраиваемого свойства `myProp`, метод `CustomProperties.set` — для записи настраиваемого свойства `otherProp`, а метод `saveAsync` — для сохранения настраиваемых свойств.</span><span class="sxs-lookup"><span data-stu-id="7facb-p164">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

#### <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="7facb-1021">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="7facb-1021">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="7facb-1022">Удаляет вложение из сообщения или встречи.</span><span class="sxs-lookup"><span data-stu-id="7facb-1022">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="7facb-1023">Метод `removeAttachmentAsync` удаляет из элемента вложение с указанным идентификатором.</span><span class="sxs-lookup"><span data-stu-id="7facb-1023">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="7facb-1024">Идентификатор вложения рекомендуется использовать для удаления вложения, только если оно добавлено тем же почтовым приложением в ходе текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="7facb-1024">As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session.</span></span> <span data-ttu-id="7facb-1025">В Outlook в Интернете и мобильных устройствах идентификатор вложения действителен только в рамках одного сеанса.</span><span class="sxs-lookup"><span data-stu-id="7facb-1025">In Outlook on the web and mobile devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="7facb-1026">Сеанс завершается, когда пользователь закрывает приложение или начинает создавать элемент во встроенной форме, а затем переходит из формы в отдельное окно.</span><span class="sxs-lookup"><span data-stu-id="7facb-1026">A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="7facb-1027">Параметры</span><span class="sxs-lookup"><span data-stu-id="7facb-1027">Parameters</span></span>

|<span data-ttu-id="7facb-1028">Имя</span><span class="sxs-lookup"><span data-stu-id="7facb-1028">Name</span></span>| <span data-ttu-id="7facb-1029">Тип</span><span class="sxs-lookup"><span data-stu-id="7facb-1029">Type</span></span>| <span data-ttu-id="7facb-1030">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="7facb-1030">Attributes</span></span>| <span data-ttu-id="7facb-1031">Описание</span><span class="sxs-lookup"><span data-stu-id="7facb-1031">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="7facb-1032">String</span><span class="sxs-lookup"><span data-stu-id="7facb-1032">String</span></span>||<span data-ttu-id="7facb-1033">Идентификатор удаляемого вложения.</span><span class="sxs-lookup"><span data-stu-id="7facb-1033">The identifier of the attachment to remove.</span></span>|
|`options`| <span data-ttu-id="7facb-1034">Object</span><span class="sxs-lookup"><span data-stu-id="7facb-1034">Object</span></span>| <span data-ttu-id="7facb-1035">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="7facb-1035">&lt;optional&gt;</span></span>|<span data-ttu-id="7facb-1036">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="7facb-1036">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="7facb-1037">Объект</span><span class="sxs-lookup"><span data-stu-id="7facb-1037">Object</span></span>| <span data-ttu-id="7facb-1038">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="7facb-1038">&lt;optional&gt;</span></span>|<span data-ttu-id="7facb-1039">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="7facb-1039">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="7facb-1040">функция</span><span class="sxs-lookup"><span data-stu-id="7facb-1040">function</span></span>| <span data-ttu-id="7facb-1041">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="7facb-1041">&lt;optional&gt;</span></span>|<span data-ttu-id="7facb-1042">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="7facb-1042">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="7facb-1043">Если удалить вложение не удается, свойство `asyncResult.error` содержит код ошибки с указанием ее причины.</span><span class="sxs-lookup"><span data-stu-id="7facb-1043">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="7facb-1044">Ошибки</span><span class="sxs-lookup"><span data-stu-id="7facb-1044">Errors</span></span>

| <span data-ttu-id="7facb-1045">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="7facb-1045">Error code</span></span> | <span data-ttu-id="7facb-1046">Описание</span><span class="sxs-lookup"><span data-stu-id="7facb-1046">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="7facb-1047">Идентификатор вложения не существует.</span><span class="sxs-lookup"><span data-stu-id="7facb-1047">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="7facb-1048">Требования</span><span class="sxs-lookup"><span data-stu-id="7facb-1048">Requirements</span></span>

|<span data-ttu-id="7facb-1049">Требование</span><span class="sxs-lookup"><span data-stu-id="7facb-1049">Requirement</span></span>| <span data-ttu-id="7facb-1050">Значение</span><span class="sxs-lookup"><span data-stu-id="7facb-1050">Value</span></span>|
|---|---|
|[<span data-ttu-id="7facb-1051">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="7facb-1051">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7facb-1052">1.1</span><span class="sxs-lookup"><span data-stu-id="7facb-1052">1.1</span></span>|
|[<span data-ttu-id="7facb-1053">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="7facb-1053">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7facb-1054">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="7facb-1054">ReadWriteItem</span></span>|
|[<span data-ttu-id="7facb-1055">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="7facb-1055">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7facb-1056">Создание</span><span class="sxs-lookup"><span data-stu-id="7facb-1056">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="7facb-1057">Пример</span><span class="sxs-lookup"><span data-stu-id="7facb-1057">Example</span></span>

<span data-ttu-id="7facb-1058">Указанный ниже код удаляет вложение с идентификатором "0".</span><span class="sxs-lookup"><span data-stu-id="7facb-1058">The following code removes an attachment with an identifier of '0'.</span></span>

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

#### <a name="saveasyncoptions-callback"></a><span data-ttu-id="7facb-1059">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="7facb-1059">saveAsync([options], callback)</span></span>

<span data-ttu-id="7facb-1060">Асинхронно сохраняет элемент.</span><span class="sxs-lookup"><span data-stu-id="7facb-1060">Asynchronously saves an item.</span></span>

<span data-ttu-id="7facb-1061">При вызове этот метод сохраняет текущее сообщение в виде черновика и возвращает идентификатор элемента с помощью метода обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="7facb-1061">When invoked, this method saves the current message as a draft and returns the item id via the callback method.</span></span> <span data-ttu-id="7facb-1062">В Outlook в Интернете или Outlook в интерактивном режиме элемент сохраняется на сервере.</span><span class="sxs-lookup"><span data-stu-id="7facb-1062">In Outlook on the web or Outlook in online mode, the item is saved to the server.</span></span> <span data-ttu-id="7facb-1063">В Outlook в режиме кэширования этот элемент сохраняется в локальном кэше.</span><span class="sxs-lookup"><span data-stu-id="7facb-1063">In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="7facb-1064">Если в режиме создания надстройка вызывает для элемента метод `saveAsync`, чтобы получить параметр `itemId` для использования с EWS или REST API, необходимо помнить, что синхронизация элемента с сервером может занять много времени, если Outlook работает в режиме кэширования данных.</span><span class="sxs-lookup"><span data-stu-id="7facb-1064">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="7facb-1065">До окончания синхронизации элемента применение параметра `itemId` будет приводить к ошибке.</span><span class="sxs-lookup"><span data-stu-id="7facb-1065">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="7facb-p168">Если метод `saveAsync` вызывается для встречи в режиме создания, она сохраняется как обычная встреча в календаре пользователя, а не как черновик. При сохранении новой встречи приглашения не отправляются. При сохранении существующей встречи уведомления отправляются добавленным или удаленным участникам.</span><span class="sxs-lookup"><span data-stu-id="7facb-p168">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="7facb-1069">Следующие клиенты отличаются другим поведением `saveAsync` в отношении встреч в режиме создания:</span><span class="sxs-lookup"><span data-stu-id="7facb-1069">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="7facb-1070">Outlook в Mac не поддерживает сохранение собраний.</span><span class="sxs-lookup"><span data-stu-id="7facb-1070">Outlook on Mac does not support saving a meeting.</span></span> <span data-ttu-id="7facb-1071">`saveAsync` Метод завершается с ошибкой при вызове из собрания в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="7facb-1071">The `saveAsync` method fails when called from a meeting in compose mode.</span></span> <span data-ttu-id="7facb-1072">Просмотреть [не удается сохранить собрание в виде черновика в Outlook для Mac с помощью API Office JS](https://support.microsoft.com/help/4505745) для обхода.</span><span class="sxs-lookup"><span data-stu-id="7facb-1072">See [Cannot save a meeting as a draft in Outlook for Mac by using Office JS API](https://support.microsoft.com/help/4505745) for a workaround.</span></span>
> - <span data-ttu-id="7facb-1073">Outlook в Интернете всегда отправляет приглашение или обновление при вызове `saveAsync` для встречи в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="7facb-1073">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="7facb-1074">Параметры</span><span class="sxs-lookup"><span data-stu-id="7facb-1074">Parameters</span></span>

|<span data-ttu-id="7facb-1075">Имя</span><span class="sxs-lookup"><span data-stu-id="7facb-1075">Name</span></span>| <span data-ttu-id="7facb-1076">Тип</span><span class="sxs-lookup"><span data-stu-id="7facb-1076">Type</span></span>| <span data-ttu-id="7facb-1077">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="7facb-1077">Attributes</span></span>| <span data-ttu-id="7facb-1078">Описание</span><span class="sxs-lookup"><span data-stu-id="7facb-1078">Description</span></span>|
|---|---|---|---|
|`options`| <span data-ttu-id="7facb-1079">Объект</span><span class="sxs-lookup"><span data-stu-id="7facb-1079">Object</span></span>| <span data-ttu-id="7facb-1080">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="7facb-1080">&lt;optional&gt;</span></span>|<span data-ttu-id="7facb-1081">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="7facb-1081">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="7facb-1082">Объект</span><span class="sxs-lookup"><span data-stu-id="7facb-1082">Object</span></span>| <span data-ttu-id="7facb-1083">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="7facb-1083">&lt;optional&gt;</span></span>|<span data-ttu-id="7facb-1084">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="7facb-1084">Developers can provide any object they wish to access in the callback method.</span></span>||
|`callback`| <span data-ttu-id="7facb-1085">функция</span><span class="sxs-lookup"><span data-stu-id="7facb-1085">function</span></span>||<span data-ttu-id="7facb-1086">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="7facb-1086">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="7facb-1087">После успешного выполнения идентификатор элемента будет предоставлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="7facb-1087">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="7facb-1088">Требования</span><span class="sxs-lookup"><span data-stu-id="7facb-1088">Requirements</span></span>

|<span data-ttu-id="7facb-1089">Требование</span><span class="sxs-lookup"><span data-stu-id="7facb-1089">Requirement</span></span>| <span data-ttu-id="7facb-1090">Значение</span><span class="sxs-lookup"><span data-stu-id="7facb-1090">Value</span></span>|
|---|---|
|[<span data-ttu-id="7facb-1091">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="7facb-1091">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7facb-1092">1.3</span><span class="sxs-lookup"><span data-stu-id="7facb-1092">1.3</span></span>|
|[<span data-ttu-id="7facb-1093">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="7facb-1093">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7facb-1094">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="7facb-1094">ReadWriteItem</span></span>|
|[<span data-ttu-id="7facb-1095">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="7facb-1095">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7facb-1096">Создание</span><span class="sxs-lookup"><span data-stu-id="7facb-1096">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="7facb-1097">Примеры</span><span class="sxs-lookup"><span data-stu-id="7facb-1097">Examples</span></span>

```js
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result.
  });
```

<span data-ttu-id="7facb-p170">Ниже приведен пример параметра `result`, переданного функции обратного вызова. Свойство `value` содержит идентификатор элемента.</span><span class="sxs-lookup"><span data-stu-id="7facb-p170">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

<br>

---
---

#### <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="7facb-1100">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="7facb-1100">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="7facb-1101">Асинхронно вставляет данные в текст или тему сообщения.</span><span class="sxs-lookup"><span data-stu-id="7facb-1101">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="7facb-p171">Метод `setSelectedDataAsync` вставляет указанную строку в местоположение курсора в теме или тексте элемента либо, если текст выделен в редакторе, он заменяет выделенный текст. Если курсор находится вне текста или темы элемента, возвращается ошибка. После вставки курсор помещается в конец вставленного содержимого.</span><span class="sxs-lookup"><span data-stu-id="7facb-p171">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="7facb-1105">Параметры</span><span class="sxs-lookup"><span data-stu-id="7facb-1105">Parameters</span></span>

|<span data-ttu-id="7facb-1106">Имя</span><span class="sxs-lookup"><span data-stu-id="7facb-1106">Name</span></span>| <span data-ttu-id="7facb-1107">Тип</span><span class="sxs-lookup"><span data-stu-id="7facb-1107">Type</span></span>| <span data-ttu-id="7facb-1108">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="7facb-1108">Attributes</span></span>| <span data-ttu-id="7facb-1109">Описание</span><span class="sxs-lookup"><span data-stu-id="7facb-1109">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="7facb-1110">String</span><span class="sxs-lookup"><span data-stu-id="7facb-1110">String</span></span>||<span data-ttu-id="7facb-p172">Вставляемые данные. Объем данных не должен превышать 1 000 000 символов. Если передано больше 1 000 000 символов, возвращается исключение `ArgumentOutOfRange`.</span><span class="sxs-lookup"><span data-stu-id="7facb-p172">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="7facb-1114">Object</span><span class="sxs-lookup"><span data-stu-id="7facb-1114">Object</span></span>| <span data-ttu-id="7facb-1115">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="7facb-1115">&lt;optional&gt;</span></span>|<span data-ttu-id="7facb-1116">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="7facb-1116">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="7facb-1117">Объект</span><span class="sxs-lookup"><span data-stu-id="7facb-1117">Object</span></span>| <span data-ttu-id="7facb-1118">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="7facb-1118">&lt;optional&gt;</span></span>|<span data-ttu-id="7facb-1119">В методе обратного вызова разработчики могут указать любой объект, к которому необходимо получить доступ.</span><span class="sxs-lookup"><span data-stu-id="7facb-1119">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="7facb-1120">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="7facb-1120">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="7facb-1121">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="7facb-1121">&lt;optional&gt;</span></span>|<span data-ttu-id="7facb-1122">Если `text`текущий стиль применяется в Outlook для веб-клиентов и клиентов для настольных ПК.</span><span class="sxs-lookup"><span data-stu-id="7facb-1122">If `text`, the current style is applied in Outlook on the web and desktop clients.</span></span> <span data-ttu-id="7facb-1123">Если поле представляет собой редактор HTML, вставляются только текстовые данные, даже если они имеют формат HTML.</span><span class="sxs-lookup"><span data-stu-id="7facb-1123">If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="7facb-1124">Если `html` и поле поддерживает HTML (тема не используется), текущий стиль применяется в Outlook в Интернете, а в настольных клиентах Outlook применяется стиль по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="7facb-1124">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook on the web and the default style is applied in Outlook desktop clients.</span></span> <span data-ttu-id="7facb-1125">Если поле является текстовым, возвращается ошибка `InvalidDataFormat`.</span><span class="sxs-lookup"><span data-stu-id="7facb-1125">If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="7facb-1126">Если свойство `coercionType` не задано, результат зависит от поля: если поле имеет формат HTML, используется текст в формате HTML, а если поле текстовое, применяется обычный текст.</span><span class="sxs-lookup"><span data-stu-id="7facb-1126">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="7facb-1127">функция</span><span class="sxs-lookup"><span data-stu-id="7facb-1127">function</span></span>||<span data-ttu-id="7facb-1128">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="7facb-1128">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="7facb-1129">Требования</span><span class="sxs-lookup"><span data-stu-id="7facb-1129">Requirements</span></span>

|<span data-ttu-id="7facb-1130">Требование</span><span class="sxs-lookup"><span data-stu-id="7facb-1130">Requirement</span></span>| <span data-ttu-id="7facb-1131">Значение</span><span class="sxs-lookup"><span data-stu-id="7facb-1131">Value</span></span>|
|---|---|
|[<span data-ttu-id="7facb-1132">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="7facb-1132">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7facb-1133">1.2</span><span class="sxs-lookup"><span data-stu-id="7facb-1133">1.2</span></span>|
|[<span data-ttu-id="7facb-1134">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="7facb-1134">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7facb-1135">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="7facb-1135">ReadWriteItem</span></span>|
|[<span data-ttu-id="7facb-1136">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="7facb-1136">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7facb-1137">Создание</span><span class="sxs-lookup"><span data-stu-id="7facb-1137">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="7facb-1138">Пример</span><span class="sxs-lookup"><span data-stu-id="7facb-1138">Example</span></span>

```js
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
