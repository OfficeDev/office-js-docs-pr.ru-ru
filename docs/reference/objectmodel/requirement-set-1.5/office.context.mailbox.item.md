---
title: Office.context.mailbox.item — набор обязательных элементов 1.5
description: ''
ms.date: 05/30/2019
localization_priority: Priority
ms.openlocfilehash: aba1590d570a55ed512f1eb223b7d927c953dbaf
ms.sourcegitcommit: 567aa05d6ee6b3639f65c50188df2331b7685857
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/04/2019
ms.locfileid: "34706339"
---
# <a name="item"></a><span data-ttu-id="00378-102">item</span><span class="sxs-lookup"><span data-stu-id="00378-102">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="00378-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="00378-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="00378-p101">Пространство имен `item` используется для доступа к выбранному в данный момент сообщению, приглашению на собрание или описанию встречи. Вы можете определить тип пространства имен `item` с помощью свойства [itemType](#itemtype-officemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="00378-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="00378-106">Requirements</span><span class="sxs-lookup"><span data-stu-id="00378-106">Requirements</span></span>

|<span data-ttu-id="00378-107">Требование</span><span class="sxs-lookup"><span data-stu-id="00378-107">Requirement</span></span>| <span data-ttu-id="00378-108">Значение</span><span class="sxs-lookup"><span data-stu-id="00378-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="00378-109">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="00378-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="00378-110">1.0</span><span class="sxs-lookup"><span data-stu-id="00378-110">1.0</span></span>|
|[<span data-ttu-id="00378-111">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="00378-111">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="00378-112">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="00378-112">Restricted</span></span>|
|[<span data-ttu-id="00378-113">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="00378-113">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="00378-114">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="00378-114">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="00378-115">Элементы и методы</span><span class="sxs-lookup"><span data-stu-id="00378-115">Members and methods</span></span>

| <span data-ttu-id="00378-116">Элемент</span><span class="sxs-lookup"><span data-stu-id="00378-116">Member</span></span> | <span data-ttu-id="00378-117">Тип</span><span class="sxs-lookup"><span data-stu-id="00378-117">Type</span></span> |
|--------|------|
| [<span data-ttu-id="00378-118">attachments</span><span class="sxs-lookup"><span data-stu-id="00378-118">attachments</span></span>](#attachments-arrayattachmentdetails) | <span data-ttu-id="00378-119">Элемент</span><span class="sxs-lookup"><span data-stu-id="00378-119">Member</span></span> |
| [<span data-ttu-id="00378-120">bcc</span><span class="sxs-lookup"><span data-stu-id="00378-120">bcc</span></span>](#bcc-recipients) | <span data-ttu-id="00378-121">Элемент</span><span class="sxs-lookup"><span data-stu-id="00378-121">Member</span></span> |
| [<span data-ttu-id="00378-122">body</span><span class="sxs-lookup"><span data-stu-id="00378-122">body</span></span>](#body-body) | <span data-ttu-id="00378-123">Элемент</span><span class="sxs-lookup"><span data-stu-id="00378-123">Member</span></span> |
| [<span data-ttu-id="00378-124">cc</span><span class="sxs-lookup"><span data-stu-id="00378-124">cc</span></span>](#cc-arrayemailaddressdetailsrecipients) | <span data-ttu-id="00378-125">Элемент</span><span class="sxs-lookup"><span data-stu-id="00378-125">Member</span></span> |
| [<span data-ttu-id="00378-126">conversationId</span><span class="sxs-lookup"><span data-stu-id="00378-126">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="00378-127">Элемент</span><span class="sxs-lookup"><span data-stu-id="00378-127">Member</span></span> |
| [<span data-ttu-id="00378-128">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="00378-128">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="00378-129">Элемент</span><span class="sxs-lookup"><span data-stu-id="00378-129">Member</span></span> |
| [<span data-ttu-id="00378-130">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="00378-130">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="00378-131">Элемент</span><span class="sxs-lookup"><span data-stu-id="00378-131">Member</span></span> |
| [<span data-ttu-id="00378-132">end</span><span class="sxs-lookup"><span data-stu-id="00378-132">end</span></span>](#end-datetime) | <span data-ttu-id="00378-133">Элемент</span><span class="sxs-lookup"><span data-stu-id="00378-133">Member</span></span> |
| [<span data-ttu-id="00378-134">from</span><span class="sxs-lookup"><span data-stu-id="00378-134">from</span></span>](#from-emailaddressdetails) | <span data-ttu-id="00378-135">Элемент</span><span class="sxs-lookup"><span data-stu-id="00378-135">Member</span></span> |
| [<span data-ttu-id="00378-136">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="00378-136">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="00378-137">Элемент</span><span class="sxs-lookup"><span data-stu-id="00378-137">Member</span></span> |
| [<span data-ttu-id="00378-138">itemClass</span><span class="sxs-lookup"><span data-stu-id="00378-138">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="00378-139">Элемент</span><span class="sxs-lookup"><span data-stu-id="00378-139">Member</span></span> |
| [<span data-ttu-id="00378-140">itemId</span><span class="sxs-lookup"><span data-stu-id="00378-140">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="00378-141">Элемент</span><span class="sxs-lookup"><span data-stu-id="00378-141">Member</span></span> |
| [<span data-ttu-id="00378-142">itemType</span><span class="sxs-lookup"><span data-stu-id="00378-142">itemType</span></span>](#itemtype-officemailboxenumsitemtype) | <span data-ttu-id="00378-143">Элемент</span><span class="sxs-lookup"><span data-stu-id="00378-143">Member</span></span> |
| [<span data-ttu-id="00378-144">location</span><span class="sxs-lookup"><span data-stu-id="00378-144">location</span></span>](#location-stringlocation) | <span data-ttu-id="00378-145">Элемент</span><span class="sxs-lookup"><span data-stu-id="00378-145">Member</span></span> |
| [<span data-ttu-id="00378-146">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="00378-146">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="00378-147">Элемент</span><span class="sxs-lookup"><span data-stu-id="00378-147">Member</span></span> |
| [<span data-ttu-id="00378-148">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="00378-148">notificationMessages</span></span>](#notificationmessages-notificationmessages) | <span data-ttu-id="00378-149">Элемент</span><span class="sxs-lookup"><span data-stu-id="00378-149">Member</span></span> |
| [<span data-ttu-id="00378-150">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="00378-150">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="00378-151">Элемент</span><span class="sxs-lookup"><span data-stu-id="00378-151">Member</span></span> |
| [<span data-ttu-id="00378-152">organizer</span><span class="sxs-lookup"><span data-stu-id="00378-152">organizer</span></span>](#organizer-emailaddressdetails) | <span data-ttu-id="00378-153">Элемент</span><span class="sxs-lookup"><span data-stu-id="00378-153">Member</span></span> |
| [<span data-ttu-id="00378-154">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="00378-154">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="00378-155">Member</span><span class="sxs-lookup"><span data-stu-id="00378-155">Member</span></span> |
| [<span data-ttu-id="00378-156">sender</span><span class="sxs-lookup"><span data-stu-id="00378-156">sender</span></span>](#sender-emailaddressdetails) | <span data-ttu-id="00378-157">Элемент</span><span class="sxs-lookup"><span data-stu-id="00378-157">Member</span></span> |
| [<span data-ttu-id="00378-158">start</span><span class="sxs-lookup"><span data-stu-id="00378-158">start</span></span>](#start-datetime) | <span data-ttu-id="00378-159">Элемент</span><span class="sxs-lookup"><span data-stu-id="00378-159">Member</span></span> |
| [<span data-ttu-id="00378-160">subject</span><span class="sxs-lookup"><span data-stu-id="00378-160">subject</span></span>](#subject-stringsubject) | <span data-ttu-id="00378-161">Элемент</span><span class="sxs-lookup"><span data-stu-id="00378-161">Member</span></span> |
| [<span data-ttu-id="00378-162">to</span><span class="sxs-lookup"><span data-stu-id="00378-162">to</span></span>](#to-arrayemailaddressdetailsrecipients) | <span data-ttu-id="00378-163">Элемент</span><span class="sxs-lookup"><span data-stu-id="00378-163">Member</span></span> |
| [<span data-ttu-id="00378-164">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="00378-164">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="00378-165">Метод</span><span class="sxs-lookup"><span data-stu-id="00378-165">Method</span></span> |
| [<span data-ttu-id="00378-166">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="00378-166">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="00378-167">Метод</span><span class="sxs-lookup"><span data-stu-id="00378-167">Method</span></span> |
| [<span data-ttu-id="00378-168">close</span><span class="sxs-lookup"><span data-stu-id="00378-168">close</span></span>](#close) | <span data-ttu-id="00378-169">Метод</span><span class="sxs-lookup"><span data-stu-id="00378-169">Method</span></span> |
| [<span data-ttu-id="00378-170">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="00378-170">displayReplyAllForm</span></span>](#displayreplyallformformdata-callback) | <span data-ttu-id="00378-171">Метод</span><span class="sxs-lookup"><span data-stu-id="00378-171">Method</span></span> |
| [<span data-ttu-id="00378-172">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="00378-172">displayReplyForm</span></span>](#displayreplyformformdata-callback) | <span data-ttu-id="00378-173">Метод</span><span class="sxs-lookup"><span data-stu-id="00378-173">Method</span></span> |
| [<span data-ttu-id="00378-174">getEntities</span><span class="sxs-lookup"><span data-stu-id="00378-174">getEntities</span></span>](#getentities--entities) | <span data-ttu-id="00378-175">Метод</span><span class="sxs-lookup"><span data-stu-id="00378-175">Method</span></span> |
| [<span data-ttu-id="00378-176">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="00378-176">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="00378-177">Метод</span><span class="sxs-lookup"><span data-stu-id="00378-177">Method</span></span> |
| [<span data-ttu-id="00378-178">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="00378-178">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="00378-179">Метод</span><span class="sxs-lookup"><span data-stu-id="00378-179">Method</span></span> |
| [<span data-ttu-id="00378-180">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="00378-180">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="00378-181">Метод</span><span class="sxs-lookup"><span data-stu-id="00378-181">Method</span></span> |
| [<span data-ttu-id="00378-182">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="00378-182">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="00378-183">Метод</span><span class="sxs-lookup"><span data-stu-id="00378-183">Method</span></span> |
| [<span data-ttu-id="00378-184">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="00378-184">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="00378-185">Метод</span><span class="sxs-lookup"><span data-stu-id="00378-185">Method</span></span> |
| [<span data-ttu-id="00378-186">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="00378-186">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="00378-187">Метод</span><span class="sxs-lookup"><span data-stu-id="00378-187">Method</span></span> |
| [<span data-ttu-id="00378-188">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="00378-188">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="00378-189">Метод</span><span class="sxs-lookup"><span data-stu-id="00378-189">Method</span></span> |
| [<span data-ttu-id="00378-190">saveAsync</span><span class="sxs-lookup"><span data-stu-id="00378-190">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="00378-191">Метод</span><span class="sxs-lookup"><span data-stu-id="00378-191">Method</span></span> |
| [<span data-ttu-id="00378-192">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="00378-192">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="00378-193">Метод</span><span class="sxs-lookup"><span data-stu-id="00378-193">Method</span></span> |

### <a name="example"></a><span data-ttu-id="00378-194">Пример</span><span class="sxs-lookup"><span data-stu-id="00378-194">Example</span></span>

<span data-ttu-id="00378-195">В примере кода JavaScript, приведенном ниже, показано, как получить доступ к свойству `subject` текущего элемента в Outlook.</span><span class="sxs-lookup"><span data-stu-id="00378-195">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="00378-196">Members</span><span class="sxs-lookup"><span data-stu-id="00378-196">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook15officeattachmentdetails"></a><span data-ttu-id="00378-197">attachments: Array.<[AttachmentDetails](/javascript/api/outlook_1_5/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="00378-197">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_5/office.attachmentdetails)></span></span>

<span data-ttu-id="00378-p102">Получает массив вложений для элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="00378-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="00378-200">Outlook блокирует определенные типы файлов из-за потенциальных проблем с безопасностью, поэтому они не возвращаются.</span><span class="sxs-lookup"><span data-stu-id="00378-200">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="00378-201">Дополнительные сведения см. в статье [Блокировка вложений в Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="00378-201">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="00378-202">Тип</span><span class="sxs-lookup"><span data-stu-id="00378-202">Type</span></span>

*   <span data-ttu-id="00378-203">Array.<[AttachmentDetails](/javascript/api/outlook_1_5/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="00378-203">Array.<[AttachmentDetails](/javascript/api/outlook_1_5/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="00378-204">Требования</span><span class="sxs-lookup"><span data-stu-id="00378-204">Requirements</span></span>

|<span data-ttu-id="00378-205">Требование</span><span class="sxs-lookup"><span data-stu-id="00378-205">Requirement</span></span>| <span data-ttu-id="00378-206">Значение</span><span class="sxs-lookup"><span data-stu-id="00378-206">Value</span></span>|
|---|---|
|[<span data-ttu-id="00378-207">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="00378-207">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="00378-208">1.0</span><span class="sxs-lookup"><span data-stu-id="00378-208">1.0</span></span>|
|[<span data-ttu-id="00378-209">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="00378-209">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="00378-210">ReadItem</span><span class="sxs-lookup"><span data-stu-id="00378-210">ReadItem</span></span>|
|[<span data-ttu-id="00378-211">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="00378-211">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="00378-212">Чтение</span><span class="sxs-lookup"><span data-stu-id="00378-212">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="00378-213">Пример</span><span class="sxs-lookup"><span data-stu-id="00378-213">Example</span></span>

<span data-ttu-id="00378-214">С помощью приведенного ниже кода можно создать HTML-строку с подробными сведениями обо всех вложениях для текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="00378-214">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

#### <a name="bcc-recipientsjavascriptapioutlook15officerecipients"></a><span data-ttu-id="00378-215">bcc: [Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="00378-215">bcc :[Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

<span data-ttu-id="00378-216">Получает объект, который предоставляет методы для получения или обновления получателей скрытой копии сообщения.</span><span class="sxs-lookup"><span data-stu-id="00378-216">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="00378-217">Только в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="00378-217">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="00378-218">Тип</span><span class="sxs-lookup"><span data-stu-id="00378-218">Type</span></span>

*   [<span data-ttu-id="00378-219">Получатели</span><span class="sxs-lookup"><span data-stu-id="00378-219">Recipients</span></span>](/javascript/api/outlook_1_5/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="00378-220">Требования</span><span class="sxs-lookup"><span data-stu-id="00378-220">Requirements</span></span>

|<span data-ttu-id="00378-221">Требование</span><span class="sxs-lookup"><span data-stu-id="00378-221">Requirement</span></span>| <span data-ttu-id="00378-222">Значение</span><span class="sxs-lookup"><span data-stu-id="00378-222">Value</span></span>|
|---|---|
|[<span data-ttu-id="00378-223">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="00378-223">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="00378-224">1.1</span><span class="sxs-lookup"><span data-stu-id="00378-224">1.1</span></span>|
|[<span data-ttu-id="00378-225">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="00378-225">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="00378-226">ReadItem</span><span class="sxs-lookup"><span data-stu-id="00378-226">ReadItem</span></span>|
|[<span data-ttu-id="00378-227">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="00378-227">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="00378-228">Создание</span><span class="sxs-lookup"><span data-stu-id="00378-228">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="00378-229">Пример</span><span class="sxs-lookup"><span data-stu-id="00378-229">Example</span></span>

```javascript
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

#### <a name="body-bodyjavascriptapioutlook15officebody"></a><span data-ttu-id="00378-230">body: [Body](/javascript/api/outlook_1_5/office.body)</span><span class="sxs-lookup"><span data-stu-id="00378-230">body :[Body](/javascript/api/outlook_1_5/office.body)</span></span>

<span data-ttu-id="00378-231">Получает объект, предоставляющий методы для работы с основным текстом элемента.</span><span class="sxs-lookup"><span data-stu-id="00378-231">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="00378-232">Тип</span><span class="sxs-lookup"><span data-stu-id="00378-232">Type</span></span>

*   [<span data-ttu-id="00378-233">Body</span><span class="sxs-lookup"><span data-stu-id="00378-233">Body</span></span>](/javascript/api/outlook_1_5/office.body)

##### <a name="requirements"></a><span data-ttu-id="00378-234">Требования</span><span class="sxs-lookup"><span data-stu-id="00378-234">Requirements</span></span>

|<span data-ttu-id="00378-235">Требование</span><span class="sxs-lookup"><span data-stu-id="00378-235">Requirement</span></span>| <span data-ttu-id="00378-236">Значение</span><span class="sxs-lookup"><span data-stu-id="00378-236">Value</span></span>|
|---|---|
|[<span data-ttu-id="00378-237">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="00378-237">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="00378-238">1.1</span><span class="sxs-lookup"><span data-stu-id="00378-238">1.1</span></span>|
|[<span data-ttu-id="00378-239">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="00378-239">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="00378-240">ReadItem</span><span class="sxs-lookup"><span data-stu-id="00378-240">ReadItem</span></span>|
|[<span data-ttu-id="00378-241">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="00378-241">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="00378-242">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="00378-242">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="00378-243">Пример</span><span class="sxs-lookup"><span data-stu-id="00378-243">Example</span></span>

<span data-ttu-id="00378-244">В этом примере возвращается текст сообщения в формате обычного текста.</span><span class="sxs-lookup"><span data-stu-id="00378-244">This example gets the body of the message in plain text.</span></span>

```javascript
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="00378-245">Ниже приведен пример итогового параметра, переданного функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="00378-245">The following is an example of the result parameter passed to the callback function.</span></span>

```json
{
  "value": "TEXT of whole body (including threads below)",
  "status": "succeeded",
  "asyncContext": "This is passed to the callback"
}
```

#### <a name="cc-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients"></a><span data-ttu-id="00378-246">cc: Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="00378-246">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

<span data-ttu-id="00378-247">Предоставляет доступ к получателям копии сообщения.</span><span class="sxs-lookup"><span data-stu-id="00378-247">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="00378-248">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="00378-248">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="00378-249">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="00378-249">Read mode</span></span>

<span data-ttu-id="00378-p106">Свойство `cc` возвращает массив, который содержит объект `EmailAddressDetails` для каждого получателя, указанного в строке **Копия** сообщения. Коллекция может включать не более 100 элементов.</span><span class="sxs-lookup"><span data-stu-id="00378-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="00378-252">Режим создания</span><span class="sxs-lookup"><span data-stu-id="00378-252">Compose mode</span></span>

<span data-ttu-id="00378-253">Свойство `cc` возвращает объект `Recipients`, предоставляющий методы для получения или обновления получателей, которые указаны в строке **Копия** сообщения.</span><span class="sxs-lookup"><span data-stu-id="00378-253">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

```javascript
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="00378-254">Тип</span><span class="sxs-lookup"><span data-stu-id="00378-254">Type</span></span>

*   <span data-ttu-id="00378-255">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="00378-255">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="00378-256">Требования</span><span class="sxs-lookup"><span data-stu-id="00378-256">Requirements</span></span>

|<span data-ttu-id="00378-257">Требование</span><span class="sxs-lookup"><span data-stu-id="00378-257">Requirement</span></span>| <span data-ttu-id="00378-258">Значение</span><span class="sxs-lookup"><span data-stu-id="00378-258">Value</span></span>|
|---|---|
|[<span data-ttu-id="00378-259">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="00378-259">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="00378-260">1.0</span><span class="sxs-lookup"><span data-stu-id="00378-260">1.0</span></span>|
|[<span data-ttu-id="00378-261">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="00378-261">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="00378-262">ReadItem</span><span class="sxs-lookup"><span data-stu-id="00378-262">ReadItem</span></span>|
|[<span data-ttu-id="00378-263">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="00378-263">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="00378-264">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="00378-264">Compose or Read</span></span>|

#### <a name="nullable-conversationid-string"></a><span data-ttu-id="00378-265">(nullable) conversationId: String</span><span class="sxs-lookup"><span data-stu-id="00378-265">(nullable) conversationId :String</span></span>

<span data-ttu-id="00378-266">Получает идентификатор разговора по электронной почте, содержащего конкретное сообщение.</span><span class="sxs-lookup"><span data-stu-id="00378-266">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="00378-p107">Вы можете получить целочисленное значение этого свойства, если ваше почтовое приложение активируется в формах просмотра или формах создания ответов. Если пользователь изменит тему ответа, после его отправки идентификатор беседы будет изменен, и полученное ранее значение будет недействительным.</span><span class="sxs-lookup"><span data-stu-id="00378-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="00378-p108">Это свойство имеет значение NULL для нового элемента в форме создания. Свойство `conversationId` вернет значение, если пользователь задаст тему и сохранит элемент.</span><span class="sxs-lookup"><span data-stu-id="00378-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="00378-271">Тип</span><span class="sxs-lookup"><span data-stu-id="00378-271">Type</span></span>

*   <span data-ttu-id="00378-272">String</span><span class="sxs-lookup"><span data-stu-id="00378-272">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="00378-273">Требования</span><span class="sxs-lookup"><span data-stu-id="00378-273">Requirements</span></span>

|<span data-ttu-id="00378-274">Требование</span><span class="sxs-lookup"><span data-stu-id="00378-274">Requirement</span></span>| <span data-ttu-id="00378-275">Значение</span><span class="sxs-lookup"><span data-stu-id="00378-275">Value</span></span>|
|---|---|
|[<span data-ttu-id="00378-276">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="00378-276">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="00378-277">1.0</span><span class="sxs-lookup"><span data-stu-id="00378-277">1.0</span></span>|
|[<span data-ttu-id="00378-278">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="00378-278">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="00378-279">ReadItem</span><span class="sxs-lookup"><span data-stu-id="00378-279">ReadItem</span></span>|
|[<span data-ttu-id="00378-280">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="00378-280">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="00378-281">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="00378-281">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="00378-282">Пример</span><span class="sxs-lookup"><span data-stu-id="00378-282">Example</span></span>

```javascript
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

#### <a name="datetimecreated-date"></a><span data-ttu-id="00378-283">dateTimeCreated: Date</span><span class="sxs-lookup"><span data-stu-id="00378-283">dateTimeCreated :Date</span></span>

<span data-ttu-id="00378-p109">Получает дату и время создания элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="00378-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="00378-286">Тип</span><span class="sxs-lookup"><span data-stu-id="00378-286">Type</span></span>

*   <span data-ttu-id="00378-287">Дата</span><span class="sxs-lookup"><span data-stu-id="00378-287">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="00378-288">Требования</span><span class="sxs-lookup"><span data-stu-id="00378-288">Requirements</span></span>

|<span data-ttu-id="00378-289">Требование</span><span class="sxs-lookup"><span data-stu-id="00378-289">Requirement</span></span>| <span data-ttu-id="00378-290">Значение</span><span class="sxs-lookup"><span data-stu-id="00378-290">Value</span></span>|
|---|---|
|[<span data-ttu-id="00378-291">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="00378-291">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="00378-292">1.0</span><span class="sxs-lookup"><span data-stu-id="00378-292">1.0</span></span>|
|[<span data-ttu-id="00378-293">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="00378-293">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="00378-294">ReadItem</span><span class="sxs-lookup"><span data-stu-id="00378-294">ReadItem</span></span>|
|[<span data-ttu-id="00378-295">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="00378-295">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="00378-296">Чтение</span><span class="sxs-lookup"><span data-stu-id="00378-296">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="00378-297">Пример</span><span class="sxs-lookup"><span data-stu-id="00378-297">Example</span></span>

```javascript
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="00378-298">dateTimeModified: Date</span><span class="sxs-lookup"><span data-stu-id="00378-298">dateTimeModified :Date</span></span>

<span data-ttu-id="00378-p110">Получает дату и время последнего изменения элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="00378-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="00378-301">Этот элемент не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="00378-301">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="00378-302">Тип</span><span class="sxs-lookup"><span data-stu-id="00378-302">Type</span></span>

*   <span data-ttu-id="00378-303">Дата</span><span class="sxs-lookup"><span data-stu-id="00378-303">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="00378-304">Требования</span><span class="sxs-lookup"><span data-stu-id="00378-304">Requirements</span></span>

|<span data-ttu-id="00378-305">Требование</span><span class="sxs-lookup"><span data-stu-id="00378-305">Requirement</span></span>| <span data-ttu-id="00378-306">Значение</span><span class="sxs-lookup"><span data-stu-id="00378-306">Value</span></span>|
|---|---|
|[<span data-ttu-id="00378-307">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="00378-307">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="00378-308">1.0</span><span class="sxs-lookup"><span data-stu-id="00378-308">1.0</span></span>|
|[<span data-ttu-id="00378-309">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="00378-309">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="00378-310">ReadItem</span><span class="sxs-lookup"><span data-stu-id="00378-310">ReadItem</span></span>|
|[<span data-ttu-id="00378-311">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="00378-311">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="00378-312">Чтение</span><span class="sxs-lookup"><span data-stu-id="00378-312">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="00378-313">Пример</span><span class="sxs-lookup"><span data-stu-id="00378-313">Example</span></span>

```javascript
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

#### <a name="end-datetimejavascriptapioutlook15officetime"></a><span data-ttu-id="00378-314">end: Date|[Time](/javascript/api/outlook_1_5/office.time)</span><span class="sxs-lookup"><span data-stu-id="00378-314">end :Date|[Time](/javascript/api/outlook_1_5/office.time)</span></span>

<span data-ttu-id="00378-315">Получает или задает дату и время окончания встречи.</span><span class="sxs-lookup"><span data-stu-id="00378-315">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="00378-p111">Свойство `end` представлено в виде значения даты и времени в формате UTC. Преобразовать значение свойства end в местные значения даты и времени клиента можно с помощью метода [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime).</span><span class="sxs-lookup"><span data-stu-id="00378-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="00378-318">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="00378-318">Read mode</span></span>

<span data-ttu-id="00378-319">Свойство `end` возвращает объект `Date`.</span><span class="sxs-lookup"><span data-stu-id="00378-319">The `end` property returns a `Date` object.</span></span>

```javascript
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="00378-320">Режим создания</span><span class="sxs-lookup"><span data-stu-id="00378-320">Compose mode</span></span>

<span data-ttu-id="00378-321">Свойство `end` возвращает объект `Time`.</span><span class="sxs-lookup"><span data-stu-id="00378-321">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="00378-322">Если вы задаете время окончания с помощью метода [`Time.setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-), необходимо использовать метод [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) для преобразования местного времени на клиенте в формат UTC для сервера.</span><span class="sxs-lookup"><span data-stu-id="00378-322">When you use the [`Time.setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="00378-323">В примере ниже показано, как с помощью метода [`setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) объекта `Time` задать время окончания встречи.</span><span class="sxs-lookup"><span data-stu-id="00378-323">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="00378-324">Тип</span><span class="sxs-lookup"><span data-stu-id="00378-324">Type</span></span>

*   <span data-ttu-id="00378-325">Date | [Time](/javascript/api/outlook_1_5/office.time)</span><span class="sxs-lookup"><span data-stu-id="00378-325">Date | [Time](/javascript/api/outlook_1_5/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="00378-326">Требования</span><span class="sxs-lookup"><span data-stu-id="00378-326">Requirements</span></span>

|<span data-ttu-id="00378-327">Требование</span><span class="sxs-lookup"><span data-stu-id="00378-327">Requirement</span></span>| <span data-ttu-id="00378-328">Значение</span><span class="sxs-lookup"><span data-stu-id="00378-328">Value</span></span>|
|---|---|
|[<span data-ttu-id="00378-329">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="00378-329">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="00378-330">1.0</span><span class="sxs-lookup"><span data-stu-id="00378-330">1.0</span></span>|
|[<span data-ttu-id="00378-331">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="00378-331">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="00378-332">ReadItem</span><span class="sxs-lookup"><span data-stu-id="00378-332">ReadItem</span></span>|
|[<span data-ttu-id="00378-333">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="00378-333">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="00378-334">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="00378-334">Compose or Read</span></span>|

#### <a name="from-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails"></a><span data-ttu-id="00378-335">from: [EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="00378-335">from :[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)</span></span>

<span data-ttu-id="00378-p112">Получает электронный адрес отправителя сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="00378-p112">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="00378-p113">Свойства `from` и [`sender`](#sender-emailaddressdetails) представляют одно лицо, если сообщение не отправлено представителем. В противном случае свойство `from` представляет лицо, делегировавшее полномочия, а свойство sender — представителя.</span><span class="sxs-lookup"><span data-stu-id="00378-p113">The `from` and [`sender`](#sender-emailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="00378-340">Свойству `recipientType`, принадлежащему объекту `EmailAddressDetails` в свойстве `from`, задано значение `undefined`.</span><span class="sxs-lookup"><span data-stu-id="00378-340">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="00378-341">Тип</span><span class="sxs-lookup"><span data-stu-id="00378-341">Type</span></span>

*   [<span data-ttu-id="00378-342">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="00378-342">EmailAddressDetails</span></span>](/javascript/api/outlook_1_5/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="00378-343">Требования</span><span class="sxs-lookup"><span data-stu-id="00378-343">Requirements</span></span>

|<span data-ttu-id="00378-344">Требование</span><span class="sxs-lookup"><span data-stu-id="00378-344">Requirement</span></span>| <span data-ttu-id="00378-345">Значение</span><span class="sxs-lookup"><span data-stu-id="00378-345">Value</span></span>|
|---|---|
|[<span data-ttu-id="00378-346">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="00378-346">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="00378-347">1.0</span><span class="sxs-lookup"><span data-stu-id="00378-347">1.0</span></span>|
|[<span data-ttu-id="00378-348">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="00378-348">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="00378-349">ReadItem</span><span class="sxs-lookup"><span data-stu-id="00378-349">ReadItem</span></span>|
|[<span data-ttu-id="00378-350">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="00378-350">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="00378-351">Чтение</span><span class="sxs-lookup"><span data-stu-id="00378-351">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="00378-352">Пример</span><span class="sxs-lookup"><span data-stu-id="00378-352">Example</span></span>

```javascript
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

#### <a name="internetmessageid-string"></a><span data-ttu-id="00378-353">internetMessageId: String</span><span class="sxs-lookup"><span data-stu-id="00378-353">internetMessageId :String</span></span>

<span data-ttu-id="00378-p114">Получает идентификатор интернет-сообщения для электронного сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="00378-p114">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="00378-356">Тип</span><span class="sxs-lookup"><span data-stu-id="00378-356">Type</span></span>

*   <span data-ttu-id="00378-357">String</span><span class="sxs-lookup"><span data-stu-id="00378-357">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="00378-358">Требования</span><span class="sxs-lookup"><span data-stu-id="00378-358">Requirements</span></span>

|<span data-ttu-id="00378-359">Требование</span><span class="sxs-lookup"><span data-stu-id="00378-359">Requirement</span></span>| <span data-ttu-id="00378-360">Значение</span><span class="sxs-lookup"><span data-stu-id="00378-360">Value</span></span>|
|---|---|
|[<span data-ttu-id="00378-361">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="00378-361">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="00378-362">1.0</span><span class="sxs-lookup"><span data-stu-id="00378-362">1.0</span></span>|
|[<span data-ttu-id="00378-363">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="00378-363">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="00378-364">ReadItem</span><span class="sxs-lookup"><span data-stu-id="00378-364">ReadItem</span></span>|
|[<span data-ttu-id="00378-365">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="00378-365">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="00378-366">Чтение</span><span class="sxs-lookup"><span data-stu-id="00378-366">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="00378-367">Пример</span><span class="sxs-lookup"><span data-stu-id="00378-367">Example</span></span>

```javascript
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="00378-368">itemClass: String</span><span class="sxs-lookup"><span data-stu-id="00378-368">itemClass :String</span></span>

<span data-ttu-id="00378-p115">Получает класс элемента веб-служб Exchange для выбранного элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="00378-p115">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="00378-p116">Свойство `itemClass` указывает класс сообщения выбранного элемента. Ниже приводятся классы сообщения по умолчанию для элемента сообщения или встречи.</span><span class="sxs-lookup"><span data-stu-id="00378-p116">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="00378-373">Тип</span><span class="sxs-lookup"><span data-stu-id="00378-373">Type</span></span> | <span data-ttu-id="00378-374">Описание</span><span class="sxs-lookup"><span data-stu-id="00378-374">Description</span></span> | <span data-ttu-id="00378-375">Класс элемента</span><span class="sxs-lookup"><span data-stu-id="00378-375">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="00378-376">Элементы встречи</span><span class="sxs-lookup"><span data-stu-id="00378-376">Appointment items</span></span> | <span data-ttu-id="00378-377">Это элементы календаря для класса элемента `IPM.Appointment` или `IPM.Appointment.Occurrence`.</span><span class="sxs-lookup"><span data-stu-id="00378-377">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurrence` |
| <span data-ttu-id="00378-378">Элементы сообщения</span><span class="sxs-lookup"><span data-stu-id="00378-378">Message items</span></span> | <span data-ttu-id="00378-379">Сюда входят электронные сообщения, для которых по умолчанию задан класс сообщения `IPM.Note`, а также приглашения на собрания, ответы на них и уведомления об их отмене, использующие `IPM.Schedule.Meeting` в качестве базового класса сообщения.</span><span class="sxs-lookup"><span data-stu-id="00378-379">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="00378-380">Можно создавать настраиваемые классы сообщения, расширяющие классы сообщения по умолчанию, например настраиваемый класс сообщения о встрече `IPM.Appointment.Contoso`.</span><span class="sxs-lookup"><span data-stu-id="00378-380">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="00378-381">Тип</span><span class="sxs-lookup"><span data-stu-id="00378-381">Type</span></span>

*   <span data-ttu-id="00378-382">String</span><span class="sxs-lookup"><span data-stu-id="00378-382">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="00378-383">Требования</span><span class="sxs-lookup"><span data-stu-id="00378-383">Requirements</span></span>

|<span data-ttu-id="00378-384">Требование</span><span class="sxs-lookup"><span data-stu-id="00378-384">Requirement</span></span>| <span data-ttu-id="00378-385">Значение</span><span class="sxs-lookup"><span data-stu-id="00378-385">Value</span></span>|
|---|---|
|[<span data-ttu-id="00378-386">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="00378-386">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="00378-387">1.0</span><span class="sxs-lookup"><span data-stu-id="00378-387">1.0</span></span>|
|[<span data-ttu-id="00378-388">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="00378-388">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="00378-389">ReadItem</span><span class="sxs-lookup"><span data-stu-id="00378-389">ReadItem</span></span>|
|[<span data-ttu-id="00378-390">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="00378-390">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="00378-391">Чтение</span><span class="sxs-lookup"><span data-stu-id="00378-391">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="00378-392">Пример</span><span class="sxs-lookup"><span data-stu-id="00378-392">Example</span></span>

```javascript
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="00378-393">(nullable) itemId: String</span><span class="sxs-lookup"><span data-stu-id="00378-393">(nullable) itemId :String</span></span>

<span data-ttu-id="00378-p117">Получает идентификатор элемента веб-служб Exchange для текущего элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="00378-p117">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="00378-396">Идентификатор, возвращаемый свойством `itemId`, совпадает с идентификатором элемента веб-служб Exchange.</span><span class="sxs-lookup"><span data-stu-id="00378-396">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="00378-397">Свойство `itemId` не совпадает с идентификатором записи Outlook, а также идентификатором, который используется REST API Outlook.</span><span class="sxs-lookup"><span data-stu-id="00378-397">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="00378-398">Прежде чем совершать вызовы REST API, используя это значение, его необходимо преобразовать с помощью [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="00378-398">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="00378-399">Дополнительные сведения см. в статье [Использование REST API Outlook из надстройки Outlook](/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="00378-399">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="00378-p119">Свойство `itemId` недоступно в режиме создания. Если требуется идентификатор элемента, с помощью метода [`saveAsync`](#saveasyncoptions-callback) можно сохранить элемент в хранилище. При этом в параметре [`AsyncResult.value`](/javascript/api/office/office.asyncresult) функции обратного вызова возвращается идентификатор элемента.</span><span class="sxs-lookup"><span data-stu-id="00378-p119">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="00378-402">Тип</span><span class="sxs-lookup"><span data-stu-id="00378-402">Type</span></span>

*   <span data-ttu-id="00378-403">String</span><span class="sxs-lookup"><span data-stu-id="00378-403">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="00378-404">Требования</span><span class="sxs-lookup"><span data-stu-id="00378-404">Requirements</span></span>

|<span data-ttu-id="00378-405">Требование</span><span class="sxs-lookup"><span data-stu-id="00378-405">Requirement</span></span>| <span data-ttu-id="00378-406">Значение</span><span class="sxs-lookup"><span data-stu-id="00378-406">Value</span></span>|
|---|---|
|[<span data-ttu-id="00378-407">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="00378-407">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="00378-408">1.0</span><span class="sxs-lookup"><span data-stu-id="00378-408">1.0</span></span>|
|[<span data-ttu-id="00378-409">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="00378-409">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="00378-410">ReadItem</span><span class="sxs-lookup"><span data-stu-id="00378-410">ReadItem</span></span>|
|[<span data-ttu-id="00378-411">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="00378-411">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="00378-412">Чтение</span><span class="sxs-lookup"><span data-stu-id="00378-412">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="00378-413">Пример</span><span class="sxs-lookup"><span data-stu-id="00378-413">Example</span></span>

<span data-ttu-id="00378-p120">Указанный ниже код проверяет наличие идентификатора элемента. Если свойство `itemId` возвращает значение `null` или `undefined`, элемент будет сохранен в хранилище, а из асинхронного результата будет получен идентификатор элемента.</span><span class="sxs-lookup"><span data-stu-id="00378-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```javascript
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result) {
    itemId = result.value;
  });
}
```

#### <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook15officemailboxenumsitemtype"></a><span data-ttu-id="00378-416">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook_1_5/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="00378-416">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_5/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="00378-417">Получает тип элемента, который представляет экземпляр.</span><span class="sxs-lookup"><span data-stu-id="00378-417">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="00378-418">Свойство `itemType` возвращает одно из значений перечисления `ItemType`, которое указывает, является ли экземпляр объекта `item` сообщением или встречей.</span><span class="sxs-lookup"><span data-stu-id="00378-418">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="00378-419">Тип</span><span class="sxs-lookup"><span data-stu-id="00378-419">Type</span></span>

*   [<span data-ttu-id="00378-420">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="00378-420">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook_1_5/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="00378-421">Требования</span><span class="sxs-lookup"><span data-stu-id="00378-421">Requirements</span></span>

|<span data-ttu-id="00378-422">Требование</span><span class="sxs-lookup"><span data-stu-id="00378-422">Requirement</span></span>| <span data-ttu-id="00378-423">Значение</span><span class="sxs-lookup"><span data-stu-id="00378-423">Value</span></span>|
|---|---|
|[<span data-ttu-id="00378-424">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="00378-424">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="00378-425">1.0</span><span class="sxs-lookup"><span data-stu-id="00378-425">1.0</span></span>|
|[<span data-ttu-id="00378-426">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="00378-426">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="00378-427">ReadItem</span><span class="sxs-lookup"><span data-stu-id="00378-427">ReadItem</span></span>|
|[<span data-ttu-id="00378-428">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="00378-428">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="00378-429">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="00378-429">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="00378-430">Пример</span><span class="sxs-lookup"><span data-stu-id="00378-430">Example</span></span>

```javascript
if (Office.context.mailbox.item.itemType === Office.MailboxEnums.ItemType.Message) {
  // Do something.
} else {
  // Do something else.
}
```

#### <a name="location-stringlocationjavascriptapioutlook15officelocation"></a><span data-ttu-id="00378-431">location: String|[Location](/javascript/api/outlook_1_5/office.location)</span><span class="sxs-lookup"><span data-stu-id="00378-431">location :String|[Location](/javascript/api/outlook_1_5/office.location)</span></span>

<span data-ttu-id="00378-432">Получает или задает место встречи.</span><span class="sxs-lookup"><span data-stu-id="00378-432">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="00378-433">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="00378-433">Read mode</span></span>

<span data-ttu-id="00378-434">Свойство `location` возвращает строку, содержащую сведения о месте встречи.</span><span class="sxs-lookup"><span data-stu-id="00378-434">The `location` property returns a string that contains the location of the appointment.</span></span>

```javascript
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="00378-435">Режим создания</span><span class="sxs-lookup"><span data-stu-id="00378-435">Compose mode</span></span>

<span data-ttu-id="00378-436">Свойство `location` возвращает объект `Location`, предоставляющий методы, которые используются для получения и задания места встречи.</span><span class="sxs-lookup"><span data-stu-id="00378-436">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```javascript
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="00378-437">Тип</span><span class="sxs-lookup"><span data-stu-id="00378-437">Type</span></span>

*   <span data-ttu-id="00378-438">String | [Location](/javascript/api/outlook_1_5/office.location)</span><span class="sxs-lookup"><span data-stu-id="00378-438">String | [Location](/javascript/api/outlook_1_5/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="00378-439">Требования</span><span class="sxs-lookup"><span data-stu-id="00378-439">Requirements</span></span>

|<span data-ttu-id="00378-440">Требование</span><span class="sxs-lookup"><span data-stu-id="00378-440">Requirement</span></span>| <span data-ttu-id="00378-441">Значение</span><span class="sxs-lookup"><span data-stu-id="00378-441">Value</span></span>|
|---|---|
|[<span data-ttu-id="00378-442">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="00378-442">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="00378-443">1.0</span><span class="sxs-lookup"><span data-stu-id="00378-443">1.0</span></span>|
|[<span data-ttu-id="00378-444">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="00378-444">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="00378-445">ReadItem</span><span class="sxs-lookup"><span data-stu-id="00378-445">ReadItem</span></span>|
|[<span data-ttu-id="00378-446">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="00378-446">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="00378-447">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="00378-447">Compose or Read</span></span>|

#### <a name="normalizedsubject-string"></a><span data-ttu-id="00378-448">normalizedSubject: String</span><span class="sxs-lookup"><span data-stu-id="00378-448">normalizedSubject :String</span></span>

<span data-ttu-id="00378-p121">Получает тему элемента со всеми удаленными префиксами (включая `RE:` и `FWD:`). Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="00378-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="00378-p122">Свойство normalizedSubject получает тему элемента со стандартными префиксами (такими как `RE:` и `FW:`), добавляемыми почтовыми программами. Для получения темы элемента с неизмененными префиксами используйте свойство [`subject`](#subject-stringsubject).</span><span class="sxs-lookup"><span data-stu-id="00378-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="00378-453">Тип</span><span class="sxs-lookup"><span data-stu-id="00378-453">Type</span></span>

*   <span data-ttu-id="00378-454">String</span><span class="sxs-lookup"><span data-stu-id="00378-454">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="00378-455">Требования</span><span class="sxs-lookup"><span data-stu-id="00378-455">Requirements</span></span>

|<span data-ttu-id="00378-456">Требование</span><span class="sxs-lookup"><span data-stu-id="00378-456">Requirement</span></span>| <span data-ttu-id="00378-457">Значение</span><span class="sxs-lookup"><span data-stu-id="00378-457">Value</span></span>|
|---|---|
|[<span data-ttu-id="00378-458">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="00378-458">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="00378-459">1.0</span><span class="sxs-lookup"><span data-stu-id="00378-459">1.0</span></span>|
|[<span data-ttu-id="00378-460">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="00378-460">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="00378-461">ReadItem</span><span class="sxs-lookup"><span data-stu-id="00378-461">ReadItem</span></span>|
|[<span data-ttu-id="00378-462">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="00378-462">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="00378-463">Чтение</span><span class="sxs-lookup"><span data-stu-id="00378-463">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="00378-464">Пример</span><span class="sxs-lookup"><span data-stu-id="00378-464">Example</span></span>

```javascript
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

#### <a name="notificationmessages-notificationmessagesjavascriptapioutlook15officenotificationmessages"></a><span data-ttu-id="00378-465">notificationMessages: [NotificationMessages](/javascript/api/outlook_1_5/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="00378-465">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_5/office.notificationmessages)</span></span>

<span data-ttu-id="00378-466">Получает сообщения уведомления для элемента.</span><span class="sxs-lookup"><span data-stu-id="00378-466">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="00378-467">Тип</span><span class="sxs-lookup"><span data-stu-id="00378-467">Type</span></span>

*   [<span data-ttu-id="00378-468">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="00378-468">NotificationMessages</span></span>](/javascript/api/outlook_1_5/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="00378-469">Требования</span><span class="sxs-lookup"><span data-stu-id="00378-469">Requirements</span></span>

|<span data-ttu-id="00378-470">Требование</span><span class="sxs-lookup"><span data-stu-id="00378-470">Requirement</span></span>| <span data-ttu-id="00378-471">Значение</span><span class="sxs-lookup"><span data-stu-id="00378-471">Value</span></span>|
|---|---|
|[<span data-ttu-id="00378-472">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="00378-472">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="00378-473">1.3</span><span class="sxs-lookup"><span data-stu-id="00378-473">1.3</span></span>|
|[<span data-ttu-id="00378-474">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="00378-474">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="00378-475">ReadItem</span><span class="sxs-lookup"><span data-stu-id="00378-475">ReadItem</span></span>|
|[<span data-ttu-id="00378-476">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="00378-476">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="00378-477">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="00378-477">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="00378-478">Пример</span><span class="sxs-lookup"><span data-stu-id="00378-478">Example</span></span>

```javascript
// Get all notifications.
Office.context.mailbox.item.notificationMessages.getAllAsync(
  function (asyncResult) {
    console.log(JSON.stringify(asyncResult));
  }
);
```

#### <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients"></a><span data-ttu-id="00378-479">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="00378-479">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

<span data-ttu-id="00378-480">Предоставляет доступ к необязательным участникам события.</span><span class="sxs-lookup"><span data-stu-id="00378-480">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="00378-481">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="00378-481">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="00378-482">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="00378-482">Read mode</span></span>

<span data-ttu-id="00378-483">Свойство `optionalAttendees` возвращает массив, содержащий объект `EmailAddressDetails` для каждого необязательного участника собрания.</span><span class="sxs-lookup"><span data-stu-id="00378-483">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

```javascript
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="00378-484">Режим создания</span><span class="sxs-lookup"><span data-stu-id="00378-484">Compose mode</span></span>

<span data-ttu-id="00378-485">Свойство `optionalAttendees` возвращает объект `Recipients`, который предоставляет методы для получения или обновления необязательных участников собрания.</span><span class="sxs-lookup"><span data-stu-id="00378-485">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="00378-486">Тип</span><span class="sxs-lookup"><span data-stu-id="00378-486">Type</span></span>

*   <span data-ttu-id="00378-487">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="00378-487">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="00378-488">Требования</span><span class="sxs-lookup"><span data-stu-id="00378-488">Requirements</span></span>

|<span data-ttu-id="00378-489">Требование</span><span class="sxs-lookup"><span data-stu-id="00378-489">Requirement</span></span>| <span data-ttu-id="00378-490">Значение</span><span class="sxs-lookup"><span data-stu-id="00378-490">Value</span></span>|
|---|---|
|[<span data-ttu-id="00378-491">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="00378-491">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="00378-492">1.0</span><span class="sxs-lookup"><span data-stu-id="00378-492">1.0</span></span>|
|[<span data-ttu-id="00378-493">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="00378-493">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="00378-494">ReadItem</span><span class="sxs-lookup"><span data-stu-id="00378-494">ReadItem</span></span>|
|[<span data-ttu-id="00378-495">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="00378-495">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="00378-496">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="00378-496">Compose or Read</span></span>|

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails"></a><span data-ttu-id="00378-497">organizer: [EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="00378-497">organizer :[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)</span></span>

<span data-ttu-id="00378-p124">Получает электронный адрес организатора указанного собрания. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="00378-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="00378-500">Тип</span><span class="sxs-lookup"><span data-stu-id="00378-500">Type</span></span>

*   [<span data-ttu-id="00378-501">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="00378-501">EmailAddressDetails</span></span>](/javascript/api/outlook_1_5/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="00378-502">Требования</span><span class="sxs-lookup"><span data-stu-id="00378-502">Requirements</span></span>

|<span data-ttu-id="00378-503">Требование</span><span class="sxs-lookup"><span data-stu-id="00378-503">Requirement</span></span>| <span data-ttu-id="00378-504">Значение</span><span class="sxs-lookup"><span data-stu-id="00378-504">Value</span></span>|
|---|---|
|[<span data-ttu-id="00378-505">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="00378-505">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="00378-506">1.0</span><span class="sxs-lookup"><span data-stu-id="00378-506">1.0</span></span>|
|[<span data-ttu-id="00378-507">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="00378-507">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="00378-508">ReadItem</span><span class="sxs-lookup"><span data-stu-id="00378-508">ReadItem</span></span>|
|[<span data-ttu-id="00378-509">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="00378-509">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="00378-510">Чтение</span><span class="sxs-lookup"><span data-stu-id="00378-510">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="00378-511">Пример</span><span class="sxs-lookup"><span data-stu-id="00378-511">Example</span></span>

```javascript
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

#### <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients"></a><span data-ttu-id="00378-512">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="00378-512">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

<span data-ttu-id="00378-513">Предоставляет доступ к обязательным участникам события.</span><span class="sxs-lookup"><span data-stu-id="00378-513">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="00378-514">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="00378-514">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="00378-515">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="00378-515">Read mode</span></span>

<span data-ttu-id="00378-516">Свойство `requiredAttendees` возвращает массив, содержащий объект `EmailAddressDetails` для каждого обязательного участника собрания.</span><span class="sxs-lookup"><span data-stu-id="00378-516">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

```javascript
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="00378-517">Режим создания</span><span class="sxs-lookup"><span data-stu-id="00378-517">Compose mode</span></span>

<span data-ttu-id="00378-518">Свойство `requiredAttendees` возвращает объект `Recipients`, предоставляющий методы, с помощью которых можно получить или обновить сведения об обязательных участниках собрания.</span><span class="sxs-lookup"><span data-stu-id="00378-518">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="00378-519">Тип</span><span class="sxs-lookup"><span data-stu-id="00378-519">Type</span></span>

*   <span data-ttu-id="00378-520">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="00378-520">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="00378-521">Требования</span><span class="sxs-lookup"><span data-stu-id="00378-521">Requirements</span></span>

|<span data-ttu-id="00378-522">Требование</span><span class="sxs-lookup"><span data-stu-id="00378-522">Requirement</span></span>| <span data-ttu-id="00378-523">Значение</span><span class="sxs-lookup"><span data-stu-id="00378-523">Value</span></span>|
|---|---|
|[<span data-ttu-id="00378-524">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="00378-524">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="00378-525">1.0</span><span class="sxs-lookup"><span data-stu-id="00378-525">1.0</span></span>|
|[<span data-ttu-id="00378-526">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="00378-526">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="00378-527">ReadItem</span><span class="sxs-lookup"><span data-stu-id="00378-527">ReadItem</span></span>|
|[<span data-ttu-id="00378-528">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="00378-528">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="00378-529">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="00378-529">Compose or Read</span></span>|

#### <a name="sender-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails"></a><span data-ttu-id="00378-530">sender: [EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="00378-530">sender :[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)</span></span>

<span data-ttu-id="00378-p126">Получает электронный адрес отправителя электронного сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="00378-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="00378-p127">Свойства [`from`](#from-emailaddressdetails) и `sender` представляют одно лицо, если сообщение не отправлено представителем. В противном случае свойство `from` представляет лицо, делегировавшее полномочия, а свойство sender — представителя.</span><span class="sxs-lookup"><span data-stu-id="00378-p127">The [`from`](#from-emailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="00378-535">Свойству `recipientType`, принадлежащему объекту `EmailAddressDetails` в свойстве `sender`, задано значение `undefined`.</span><span class="sxs-lookup"><span data-stu-id="00378-535">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="00378-536">Тип</span><span class="sxs-lookup"><span data-stu-id="00378-536">Type</span></span>

*   [<span data-ttu-id="00378-537">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="00378-537">EmailAddressDetails</span></span>](/javascript/api/outlook_1_5/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="00378-538">Требования</span><span class="sxs-lookup"><span data-stu-id="00378-538">Requirements</span></span>

|<span data-ttu-id="00378-539">Требование</span><span class="sxs-lookup"><span data-stu-id="00378-539">Requirement</span></span>| <span data-ttu-id="00378-540">Значение</span><span class="sxs-lookup"><span data-stu-id="00378-540">Value</span></span>|
|---|---|
|[<span data-ttu-id="00378-541">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="00378-541">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="00378-542">1.0</span><span class="sxs-lookup"><span data-stu-id="00378-542">1.0</span></span>|
|[<span data-ttu-id="00378-543">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="00378-543">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="00378-544">ReadItem</span><span class="sxs-lookup"><span data-stu-id="00378-544">ReadItem</span></span>|
|[<span data-ttu-id="00378-545">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="00378-545">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="00378-546">Чтение</span><span class="sxs-lookup"><span data-stu-id="00378-546">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="00378-547">Пример</span><span class="sxs-lookup"><span data-stu-id="00378-547">Example</span></span>

```javascript
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

#### <a name="start-datetimejavascriptapioutlook15officetime"></a><span data-ttu-id="00378-548">start: Date|[Time](/javascript/api/outlook_1_5/office.time)</span><span class="sxs-lookup"><span data-stu-id="00378-548">start :Date|[Time](/javascript/api/outlook_1_5/office.time)</span></span>

<span data-ttu-id="00378-549">Получает или задает дату и время начала встречи.</span><span class="sxs-lookup"><span data-stu-id="00378-549">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="00378-p128">Свойство `start` представлено в виде значения даты и времени в формате UTC. Это значение можно преобразовать в местные значения даты и времени клиента с помощью метода [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime).</span><span class="sxs-lookup"><span data-stu-id="00378-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="00378-552">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="00378-552">Read mode</span></span>

<span data-ttu-id="00378-553">Свойство `start` возвращает объект `Date`.</span><span class="sxs-lookup"><span data-stu-id="00378-553">The `start` property returns a `Date` object.</span></span>

```javascript
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="00378-554">Режим создания</span><span class="sxs-lookup"><span data-stu-id="00378-554">Compose mode</span></span>

<span data-ttu-id="00378-555">Свойство `start` возвращает объект `Time`.</span><span class="sxs-lookup"><span data-stu-id="00378-555">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="00378-556">Если вы задаете время начала с помощью метода [`Time.setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-), необходимо использовать метод [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) для преобразования местного времени на клиенте в формат UTC для сервера.</span><span class="sxs-lookup"><span data-stu-id="00378-556">When you use the [`Time.setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="00378-557">В примере ниже с помощью метода [`setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) объекта `Time` задается время начала встречи в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="00378-557">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="00378-558">Тип</span><span class="sxs-lookup"><span data-stu-id="00378-558">Type</span></span>

*   <span data-ttu-id="00378-559">Date | [Time](/javascript/api/outlook_1_5/office.time)</span><span class="sxs-lookup"><span data-stu-id="00378-559">Date | [Time](/javascript/api/outlook_1_5/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="00378-560">Требования</span><span class="sxs-lookup"><span data-stu-id="00378-560">Requirements</span></span>

|<span data-ttu-id="00378-561">Требование</span><span class="sxs-lookup"><span data-stu-id="00378-561">Requirement</span></span>| <span data-ttu-id="00378-562">Значение</span><span class="sxs-lookup"><span data-stu-id="00378-562">Value</span></span>|
|---|---|
|[<span data-ttu-id="00378-563">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="00378-563">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="00378-564">1.0</span><span class="sxs-lookup"><span data-stu-id="00378-564">1.0</span></span>|
|[<span data-ttu-id="00378-565">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="00378-565">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="00378-566">ReadItem</span><span class="sxs-lookup"><span data-stu-id="00378-566">ReadItem</span></span>|
|[<span data-ttu-id="00378-567">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="00378-567">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="00378-568">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="00378-568">Compose or Read</span></span>|

#### <a name="subject-stringsubjectjavascriptapioutlook15officesubject"></a><span data-ttu-id="00378-569">subject: String|[Subject](/javascript/api/outlook_1_5/office.subject)</span><span class="sxs-lookup"><span data-stu-id="00378-569">subject :String|[Subject](/javascript/api/outlook_1_5/office.subject)</span></span>

<span data-ttu-id="00378-570">Получает или задает описание, которое отображается в поле темы элемента.</span><span class="sxs-lookup"><span data-stu-id="00378-570">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="00378-571">Свойство `subject` получает или задает всю тему элемента для отправки с почтового сервера.</span><span class="sxs-lookup"><span data-stu-id="00378-571">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="00378-572">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="00378-572">Read mode</span></span>

<span data-ttu-id="00378-p129">Свойство `subject` возвращает строку. С помощью свойства [`normalizedSubject`](#normalizedsubject-string) можно получить тему без начальных префиксов, таких как `RE:` и `FW:`.</span><span class="sxs-lookup"><span data-stu-id="00378-p129">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```javascript
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="00378-575">Режим создания</span><span class="sxs-lookup"><span data-stu-id="00378-575">Compose mode</span></span>

<span data-ttu-id="00378-576">Свойство `subject` возвращает объект `Subject`, который предоставляет методы для получения и задания темы.</span><span class="sxs-lookup"><span data-stu-id="00378-576">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```javascript
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="00378-577">Тип</span><span class="sxs-lookup"><span data-stu-id="00378-577">Type</span></span>

*   <span data-ttu-id="00378-578">String | [Subject](/javascript/api/outlook_1_5/office.subject)</span><span class="sxs-lookup"><span data-stu-id="00378-578">String | [Subject](/javascript/api/outlook_1_5/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="00378-579">Требования</span><span class="sxs-lookup"><span data-stu-id="00378-579">Requirements</span></span>

|<span data-ttu-id="00378-580">Требование</span><span class="sxs-lookup"><span data-stu-id="00378-580">Requirement</span></span>| <span data-ttu-id="00378-581">Значение</span><span class="sxs-lookup"><span data-stu-id="00378-581">Value</span></span>|
|---|---|
|[<span data-ttu-id="00378-582">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="00378-582">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="00378-583">1.0</span><span class="sxs-lookup"><span data-stu-id="00378-583">1.0</span></span>|
|[<span data-ttu-id="00378-584">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="00378-584">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="00378-585">ReadItem</span><span class="sxs-lookup"><span data-stu-id="00378-585">ReadItem</span></span>|
|[<span data-ttu-id="00378-586">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="00378-586">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="00378-587">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="00378-587">Compose or Read</span></span>|

#### <a name="to-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients"></a><span data-ttu-id="00378-588">to: Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="00378-588">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

<span data-ttu-id="00378-589">Предоставляет доступ к получателям, указанным в строке **Кому** сообщения.</span><span class="sxs-lookup"><span data-stu-id="00378-589">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="00378-590">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="00378-590">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="00378-591">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="00378-591">Read mode</span></span>

<span data-ttu-id="00378-p131">Свойство `to` возвращает массив, содержащий объект `EmailAddressDetails` для каждого получателя в строке **Кому** сообщения. Коллекция может включать не более 100 элементов.</span><span class="sxs-lookup"><span data-stu-id="00378-p131">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="00378-594">Режим создания</span><span class="sxs-lookup"><span data-stu-id="00378-594">Compose mode</span></span>

<span data-ttu-id="00378-595">Свойство `to` возвращает объект `Recipients`, предоставляющий методы для получения или обновления получателей, которые указаны в строке **Кому** сообщения.</span><span class="sxs-lookup"><span data-stu-id="00378-595">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

```javascript
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="00378-596">Тип</span><span class="sxs-lookup"><span data-stu-id="00378-596">Type</span></span>

*   <span data-ttu-id="00378-597">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="00378-597">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="00378-598">Требования</span><span class="sxs-lookup"><span data-stu-id="00378-598">Requirements</span></span>

|<span data-ttu-id="00378-599">Требование</span><span class="sxs-lookup"><span data-stu-id="00378-599">Requirement</span></span>| <span data-ttu-id="00378-600">Значение</span><span class="sxs-lookup"><span data-stu-id="00378-600">Value</span></span>|
|---|---|
|[<span data-ttu-id="00378-601">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="00378-601">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="00378-602">1.0</span><span class="sxs-lookup"><span data-stu-id="00378-602">1.0</span></span>|
|[<span data-ttu-id="00378-603">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="00378-603">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="00378-604">ReadItem</span><span class="sxs-lookup"><span data-stu-id="00378-604">ReadItem</span></span>|
|[<span data-ttu-id="00378-605">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="00378-605">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="00378-606">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="00378-606">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="00378-607">Методы</span><span class="sxs-lookup"><span data-stu-id="00378-607">Methods</span></span>

#### <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="00378-608">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="00378-608">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="00378-609">Добавляет файл в сообщение или встречу в качестве вложения.</span><span class="sxs-lookup"><span data-stu-id="00378-609">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="00378-610">Метод `addFileAttachmentAsync` передает файл по указанному универсальному коду ресурса (URI) и вкладывает его в элемент в форме создания.</span><span class="sxs-lookup"><span data-stu-id="00378-610">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="00378-611">Идентификатор можно использовать с методом [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback), чтобы удалить вложение, добавленное во время текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="00378-611">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="00378-612">Параметры</span><span class="sxs-lookup"><span data-stu-id="00378-612">Parameters</span></span>

|<span data-ttu-id="00378-613">Имя</span><span class="sxs-lookup"><span data-stu-id="00378-613">Name</span></span>| <span data-ttu-id="00378-614">Тип</span><span class="sxs-lookup"><span data-stu-id="00378-614">Type</span></span>| <span data-ttu-id="00378-615">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="00378-615">Attributes</span></span>| <span data-ttu-id="00378-616">Описание</span><span class="sxs-lookup"><span data-stu-id="00378-616">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="00378-617">String</span><span class="sxs-lookup"><span data-stu-id="00378-617">String</span></span>||<span data-ttu-id="00378-p132">Универсальный код ресурса (URI), представляющий расположение файла, который нужно вложить в сообщение или встречу. Максимальная длина — 2048 символов.</span><span class="sxs-lookup"><span data-stu-id="00378-p132">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="00378-620">String</span><span class="sxs-lookup"><span data-stu-id="00378-620">String</span></span>||<span data-ttu-id="00378-p133">Имя вложения, которое отображается при передаче вложения. Максимальная длина — 255 символов.</span><span class="sxs-lookup"><span data-stu-id="00378-p133">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="00378-623">Object</span><span class="sxs-lookup"><span data-stu-id="00378-623">Object</span></span>| <span data-ttu-id="00378-624">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="00378-624">&lt;optional&gt;</span></span>|<span data-ttu-id="00378-625">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="00378-625">An object literal that contains one or more of the following properties.</span></span>|
| `options.asyncContext` | <span data-ttu-id="00378-626">Object</span><span class="sxs-lookup"><span data-stu-id="00378-626">Object</span></span> | <span data-ttu-id="00378-627">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="00378-627">&lt;optional&gt;</span></span> | <span data-ttu-id="00378-628">В методе обратного вызова разработчики могут указать любой объект, к которому необходимо получить доступ.</span><span class="sxs-lookup"><span data-stu-id="00378-628">Developers can provide any object they wish to access in the callback method.</span></span> |
| `options.isInline` | <span data-ttu-id="00378-629">Boolean</span><span class="sxs-lookup"><span data-stu-id="00378-629">Boolean</span></span> | <span data-ttu-id="00378-630">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="00378-630">&lt;optional&gt;</span></span> | <span data-ttu-id="00378-631">Значение `true` указывает на то, что вложение будет встроено в текст сообщения и не должно отображаться в списке вложений.</span><span class="sxs-lookup"><span data-stu-id="00378-631">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
|`callback`| <span data-ttu-id="00378-632">function</span><span class="sxs-lookup"><span data-stu-id="00378-632">function</span></span>| <span data-ttu-id="00378-633">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="00378-633">&lt;optional&gt;</span></span>|<span data-ttu-id="00378-634">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="00378-634">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="00378-635">После успешного выполнения идентификатор вложения будет представлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="00378-635">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="00378-636">Если добавить вложение не удастся, объект `asyncResult` будет содержать объект `Error` с описанием ошибки.</span><span class="sxs-lookup"><span data-stu-id="00378-636">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="00378-637">Ошибки</span><span class="sxs-lookup"><span data-stu-id="00378-637">Errors</span></span>

| <span data-ttu-id="00378-638">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="00378-638">Error code</span></span> | <span data-ttu-id="00378-639">Описание</span><span class="sxs-lookup"><span data-stu-id="00378-639">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="00378-640">Вложение превышает максимальный размер.</span><span class="sxs-lookup"><span data-stu-id="00378-640">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="00378-641">Расширение вложения не поддерживается.</span><span class="sxs-lookup"><span data-stu-id="00378-641">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="00378-642">Сообщение или встреча содержат слишком много вложений.</span><span class="sxs-lookup"><span data-stu-id="00378-642">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="00378-643">Требования</span><span class="sxs-lookup"><span data-stu-id="00378-643">Requirements</span></span>

|<span data-ttu-id="00378-644">Требование</span><span class="sxs-lookup"><span data-stu-id="00378-644">Requirement</span></span>| <span data-ttu-id="00378-645">Значение</span><span class="sxs-lookup"><span data-stu-id="00378-645">Value</span></span>|
|---|---|
|[<span data-ttu-id="00378-646">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="00378-646">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="00378-647">1.1</span><span class="sxs-lookup"><span data-stu-id="00378-647">1.1</span></span>|
|[<span data-ttu-id="00378-648">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="00378-648">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="00378-649">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="00378-649">ReadWriteItem</span></span>|
|[<span data-ttu-id="00378-650">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="00378-650">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="00378-651">Создание</span><span class="sxs-lookup"><span data-stu-id="00378-651">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="00378-652">Примеры</span><span class="sxs-lookup"><span data-stu-id="00378-652">Examples</span></span>

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

<span data-ttu-id="00378-653">В приведенном ниже примере файл изображения добавляется в качестве встроенного вложения, а ссылка на вложение добавляется в текст сообщения.</span><span class="sxs-lookup"><span data-stu-id="00378-653">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

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

#### <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="00378-654">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="00378-654">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="00378-655">Добавляет к сообщению элемент Exchange, например сообщение, в виде вложения.</span><span class="sxs-lookup"><span data-stu-id="00378-655">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="00378-p134">С помощью метода `addItemAttachmentAsync` можно в элемент формы создания вложить элемент с указанным идентификатором Exchange. Если указать метод обратного вызова, то этот метод вызывается с помощью параметра `asyncResult`, который содержит идентификатор вложения или код, указывающий на ошибки, которые произошли при вложении элемента. При необходимости можно использовать параметр `options` для передачи сведений о состоянии методу обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="00378-p134">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="00378-659">Идентификатор можно использовать с методом [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback), чтобы удалить вложение, добавленное во время текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="00378-659">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="00378-660">Если ваша надстройка Office выполняется в Outlook Web App, метод `addItemAttachmentAsync` обеспечивает вложение элементов в элементы, отличные от редактируемого. Однако это действие не рекомендуем выполнять, так как оно не поддерживается.</span><span class="sxs-lookup"><span data-stu-id="00378-660">If your Office Add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="00378-661">Параметры</span><span class="sxs-lookup"><span data-stu-id="00378-661">Parameters</span></span>

|<span data-ttu-id="00378-662">Имя</span><span class="sxs-lookup"><span data-stu-id="00378-662">Name</span></span>| <span data-ttu-id="00378-663">Тип</span><span class="sxs-lookup"><span data-stu-id="00378-663">Type</span></span>| <span data-ttu-id="00378-664">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="00378-664">Attributes</span></span>| <span data-ttu-id="00378-665">Описание</span><span class="sxs-lookup"><span data-stu-id="00378-665">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="00378-666">String</span><span class="sxs-lookup"><span data-stu-id="00378-666">String</span></span>||<span data-ttu-id="00378-p135">Идентификатор Exchange для вкладываемого элемента. Максимальная длина — 100 символов.</span><span class="sxs-lookup"><span data-stu-id="00378-p135">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="00378-669">String</span><span class="sxs-lookup"><span data-stu-id="00378-669">String</span></span>||<span data-ttu-id="00378-670">Тема вкладываемого элемента.</span><span class="sxs-lookup"><span data-stu-id="00378-670">The subject of the item to be attached.</span></span> <span data-ttu-id="00378-671">Максимальная длина: 255 символов.</span><span class="sxs-lookup"><span data-stu-id="00378-671">The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="00378-672">Object</span><span class="sxs-lookup"><span data-stu-id="00378-672">Object</span></span>| <span data-ttu-id="00378-673">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="00378-673">&lt;optional&gt;</span></span>|<span data-ttu-id="00378-674">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="00378-674">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="00378-675">Object</span><span class="sxs-lookup"><span data-stu-id="00378-675">Object</span></span>| <span data-ttu-id="00378-676">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="00378-676">&lt;optional&gt;</span></span>|<span data-ttu-id="00378-677">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="00378-677">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="00378-678">функция</span><span class="sxs-lookup"><span data-stu-id="00378-678">function</span></span>| <span data-ttu-id="00378-679">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="00378-679">&lt;optional&gt;</span></span>|<span data-ttu-id="00378-680">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="00378-680">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="00378-681">После успешного выполнения идентификатор вложения будет представлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="00378-681">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="00378-682">Если добавить вложение не удастся, объект `asyncResult` будет содержать объект `Error` с описанием ошибки.</span><span class="sxs-lookup"><span data-stu-id="00378-682">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="00378-683">Ошибки</span><span class="sxs-lookup"><span data-stu-id="00378-683">Errors</span></span>

| <span data-ttu-id="00378-684">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="00378-684">Error code</span></span> | <span data-ttu-id="00378-685">Описание</span><span class="sxs-lookup"><span data-stu-id="00378-685">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="00378-686">Сообщение или встреча содержат слишком много вложений.</span><span class="sxs-lookup"><span data-stu-id="00378-686">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="00378-687">Требования</span><span class="sxs-lookup"><span data-stu-id="00378-687">Requirements</span></span>

|<span data-ttu-id="00378-688">Требование</span><span class="sxs-lookup"><span data-stu-id="00378-688">Requirement</span></span>| <span data-ttu-id="00378-689">Значение</span><span class="sxs-lookup"><span data-stu-id="00378-689">Value</span></span>|
|---|---|
|[<span data-ttu-id="00378-690">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="00378-690">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="00378-691">1.1</span><span class="sxs-lookup"><span data-stu-id="00378-691">1.1</span></span>|
|[<span data-ttu-id="00378-692">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="00378-692">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="00378-693">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="00378-693">ReadWriteItem</span></span>|
|[<span data-ttu-id="00378-694">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="00378-694">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="00378-695">Создание</span><span class="sxs-lookup"><span data-stu-id="00378-695">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="00378-696">Пример</span><span class="sxs-lookup"><span data-stu-id="00378-696">Example</span></span>

<span data-ttu-id="00378-697">В следующем примере существующий элемент Outlook добавляется в виде вложения с именем `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="00378-697">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

#### <a name="close"></a><span data-ttu-id="00378-698">close()</span><span class="sxs-lookup"><span data-stu-id="00378-698">close()</span></span>

<span data-ttu-id="00378-699">Закрывает текущий создаваемый элемент.</span><span class="sxs-lookup"><span data-stu-id="00378-699">Closes the current item that is being composed.</span></span>

<span data-ttu-id="00378-p137">Работа метода `close` зависит от текущего состояния создаваемого элемента. Если элемент содержит несохраненные изменения, клиент предложит пользователю сохранить или отклонить их либо отменить действие закрытия.</span><span class="sxs-lookup"><span data-stu-id="00378-p137">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="00378-702">Если в Outlook в Интернете элемент представляет собой встречу, ранее сохраненную с помощью метода `saveAsync`, пользователю предлагается сохранить, отклонить или отменить действие, даже если с момента последнего сохранения элемента не вносились какие-либо изменения.</span><span class="sxs-lookup"><span data-stu-id="00378-702">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="00378-703">Если в клиенте Outlook для настольных ПК сообщение представляет собой ответ в тексте, метод `close` не работает.</span><span class="sxs-lookup"><span data-stu-id="00378-703">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="00378-704">Требования</span><span class="sxs-lookup"><span data-stu-id="00378-704">Requirements</span></span>

|<span data-ttu-id="00378-705">Требование</span><span class="sxs-lookup"><span data-stu-id="00378-705">Requirement</span></span>| <span data-ttu-id="00378-706">Значение</span><span class="sxs-lookup"><span data-stu-id="00378-706">Value</span></span>|
|---|---|
|[<span data-ttu-id="00378-707">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="00378-707">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="00378-708">1.3</span><span class="sxs-lookup"><span data-stu-id="00378-708">1.3</span></span>|
|[<span data-ttu-id="00378-709">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="00378-709">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="00378-710">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="00378-710">Restricted</span></span>|
|[<span data-ttu-id="00378-711">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="00378-711">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="00378-712">Создание</span><span class="sxs-lookup"><span data-stu-id="00378-712">Compose</span></span>|

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="00378-713">displayReplyAllForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="00378-713">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="00378-714">Отображает форму ответа, включающую отправителя и всех получателей выбранного сообщения или организатора и всех участников выбранной встречи.</span><span class="sxs-lookup"><span data-stu-id="00378-714">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="00378-715">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="00378-715">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="00378-716">В Outlook Web App форма ответа отображается в виде всплывающей формы в представлении с 3 либо 1 или 2 колонками.</span><span class="sxs-lookup"><span data-stu-id="00378-716">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="00378-717">Если любой строковый параметр превышает указанные для него ограничения, `displayReplyAllForm` возвращает исключение.</span><span class="sxs-lookup"><span data-stu-id="00378-717">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="00378-p138">Если в параметре `formData.attachments` указаны вложения, Outlook и Outlook Web App пытаются скачать их и вложить в форму ответа. Если какие-либо вложения добавить не удается, в форме отображается сообщение об ошибке. Если сообщения об ошибках не предусмотрены, то они не отображаются.</span><span class="sxs-lookup"><span data-stu-id="00378-p138">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="00378-721">Параметры</span><span class="sxs-lookup"><span data-stu-id="00378-721">Parameters</span></span>

| <span data-ttu-id="00378-722">Имя</span><span class="sxs-lookup"><span data-stu-id="00378-722">Name</span></span> | <span data-ttu-id="00378-723">Тип</span><span class="sxs-lookup"><span data-stu-id="00378-723">Type</span></span> | <span data-ttu-id="00378-724">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="00378-724">Attributes</span></span> | <span data-ttu-id="00378-725">Описание</span><span class="sxs-lookup"><span data-stu-id="00378-725">Description</span></span> |
|---|---|---|---|
|`formData`| <span data-ttu-id="00378-726">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="00378-726">String &#124; Object</span></span>| |<span data-ttu-id="00378-p139">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="00378-p139">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="00378-729">**ИЛИ**</span><span class="sxs-lookup"><span data-stu-id="00378-729">**OR**</span></span><br/><span data-ttu-id="00378-p140">Объект, который содержит текст или данные вложения и функцию обратного вызова. Ниже представлено определение этого объекта.</span><span class="sxs-lookup"><span data-stu-id="00378-p140">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="00378-732">String</span><span class="sxs-lookup"><span data-stu-id="00378-732">String</span></span> | <span data-ttu-id="00378-733">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="00378-733">&lt;optional&gt;</span></span> | <span data-ttu-id="00378-p141">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="00378-p141">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="00378-736">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="00378-736">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="00378-737">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="00378-737">&lt;optional&gt;</span></span> | <span data-ttu-id="00378-738">Массив объектов JSON, представляющих собой вложенные файлы или элементы.</span><span class="sxs-lookup"><span data-stu-id="00378-738">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="00378-739">String</span><span class="sxs-lookup"><span data-stu-id="00378-739">String</span></span> | | <span data-ttu-id="00378-p142">Указывает тип вложения. Допустимые значения: `file` для вложенного файла и `item` для вложенного элемента.</span><span class="sxs-lookup"><span data-stu-id="00378-p142">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="00378-742">String</span><span class="sxs-lookup"><span data-stu-id="00378-742">String</span></span> | | <span data-ttu-id="00378-743">Строка, содержащая имя вложения, длиной до 255 символов.</span><span class="sxs-lookup"><span data-stu-id="00378-743">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="00378-744">String</span><span class="sxs-lookup"><span data-stu-id="00378-744">String</span></span> | | <span data-ttu-id="00378-p143">Используется, только если свойству `type` задано значение `file`. URI расположения файла.</span><span class="sxs-lookup"><span data-stu-id="00378-p143">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.isInline` | <span data-ttu-id="00378-747">Логический</span><span class="sxs-lookup"><span data-stu-id="00378-747">Boolean</span></span> | | <span data-ttu-id="00378-p144">Используется, только если свойству `type` задано значение `file`. Значение `true` указывает на то, что вложение будет встроено в текст сообщения и не должно отображаться в списке вложений.</span><span class="sxs-lookup"><span data-stu-id="00378-p144">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="00378-750">String</span><span class="sxs-lookup"><span data-stu-id="00378-750">String</span></span> | | <span data-ttu-id="00378-p145">Используется, только если свойству `type` задано значение `item`. Идентификатор вложения EWS. Это строка длиной до 100 символов.</span><span class="sxs-lookup"><span data-stu-id="00378-p145">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="00378-754">function</span><span class="sxs-lookup"><span data-stu-id="00378-754">function</span></span> | <span data-ttu-id="00378-755">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="00378-755">&lt;optional&gt;</span></span> | <span data-ttu-id="00378-756">По завершении работы метода функция, переданная параметру `callback`, вызывается с помощью одного параметра `asyncResult`, представляющего собой объект [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="00378-756">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="00378-757">Требования</span><span class="sxs-lookup"><span data-stu-id="00378-757">Requirements</span></span>

|<span data-ttu-id="00378-758">Требование</span><span class="sxs-lookup"><span data-stu-id="00378-758">Requirement</span></span>| <span data-ttu-id="00378-759">Значение</span><span class="sxs-lookup"><span data-stu-id="00378-759">Value</span></span>|
|---|---|
|[<span data-ttu-id="00378-760">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="00378-760">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="00378-761">1.0</span><span class="sxs-lookup"><span data-stu-id="00378-761">1.0</span></span>|
|[<span data-ttu-id="00378-762">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="00378-762">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="00378-763">ReadItem</span><span class="sxs-lookup"><span data-stu-id="00378-763">ReadItem</span></span>|
|[<span data-ttu-id="00378-764">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="00378-764">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="00378-765">Чтение</span><span class="sxs-lookup"><span data-stu-id="00378-765">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="00378-766">Примеры</span><span class="sxs-lookup"><span data-stu-id="00378-766">Examples</span></span>

<span data-ttu-id="00378-767">Приведенный ниже код передает строку в функцию `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="00378-767">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="00378-768">Ответ с пустым текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="00378-768">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="00378-769">Ответ только с текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="00378-769">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="00378-770">Ответ с текстом сообщения и вложенным файлом.</span><span class="sxs-lookup"><span data-stu-id="00378-770">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="00378-771">Ответ с текстом сообщения и вложенным элементом.</span><span class="sxs-lookup"><span data-stu-id="00378-771">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="00378-772">Ответ с текстом сообщения, вложенным файлом, вложенным элементом и обратным вызовом.</span><span class="sxs-lookup"><span data-stu-id="00378-772">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="00378-773">displayReplyForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="00378-773">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="00378-774">Отображает форму ответа, включающую только отправителя выбранного сообщения или организатора выбранной встречи.</span><span class="sxs-lookup"><span data-stu-id="00378-774">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="00378-775">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="00378-775">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="00378-776">В Outlook Web App форма ответа отображается в виде всплывающей формы в представлении с 3 либо 1 или 2 колонками.</span><span class="sxs-lookup"><span data-stu-id="00378-776">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="00378-777">Если любой строковый параметр превышает указанные для него ограничения, `displayReplyForm` возвращает исключение.</span><span class="sxs-lookup"><span data-stu-id="00378-777">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="00378-p146">Если в параметре `formData.attachments` указаны вложения, Outlook и Outlook Web App пытаются скачать их и вложить в форму ответа. Если какие-либо вложения добавить не удается, в форме отображается сообщение об ошибке. Если сообщения об ошибках не предусмотрены, то они не отображаются.</span><span class="sxs-lookup"><span data-stu-id="00378-p146">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="00378-781">Параметры</span><span class="sxs-lookup"><span data-stu-id="00378-781">Parameters</span></span>

| <span data-ttu-id="00378-782">Имя</span><span class="sxs-lookup"><span data-stu-id="00378-782">Name</span></span> | <span data-ttu-id="00378-783">Тип</span><span class="sxs-lookup"><span data-stu-id="00378-783">Type</span></span> | <span data-ttu-id="00378-784">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="00378-784">Attributes</span></span> | <span data-ttu-id="00378-785">Описание</span><span class="sxs-lookup"><span data-stu-id="00378-785">Description</span></span> |
|---|---|---|---|
|`formData`| <span data-ttu-id="00378-786">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="00378-786">String &#124; Object</span></span>| | <span data-ttu-id="00378-p147">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="00378-p147">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="00378-789">**ИЛИ**</span><span class="sxs-lookup"><span data-stu-id="00378-789">**OR**</span></span><br/><span data-ttu-id="00378-p148">Объект, который содержит текст или данные вложения и функцию обратного вызова. Ниже представлено определение этого объекта.</span><span class="sxs-lookup"><span data-stu-id="00378-p148">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="00378-792">String</span><span class="sxs-lookup"><span data-stu-id="00378-792">String</span></span> | <span data-ttu-id="00378-793">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="00378-793">&lt;optional&gt;</span></span> | <span data-ttu-id="00378-p149">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="00378-p149">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="00378-796">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="00378-796">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="00378-797">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="00378-797">&lt;optional&gt;</span></span> | <span data-ttu-id="00378-798">Массив объектов JSON, представляющих собой вложенные файлы или элементы.</span><span class="sxs-lookup"><span data-stu-id="00378-798">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="00378-799">String</span><span class="sxs-lookup"><span data-stu-id="00378-799">String</span></span> | | <span data-ttu-id="00378-p150">Указывает тип вложения. Допустимые значения: `file` для вложенного файла и `item` для вложенного элемента.</span><span class="sxs-lookup"><span data-stu-id="00378-p150">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="00378-802">String</span><span class="sxs-lookup"><span data-stu-id="00378-802">String</span></span> | | <span data-ttu-id="00378-803">Строка, содержащая имя вложения, длиной до 255 символов.</span><span class="sxs-lookup"><span data-stu-id="00378-803">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="00378-804">String</span><span class="sxs-lookup"><span data-stu-id="00378-804">String</span></span> | | <span data-ttu-id="00378-p151">Используется, только если свойству `type` задано значение `file`. URI расположения файла.</span><span class="sxs-lookup"><span data-stu-id="00378-p151">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.isInline` | <span data-ttu-id="00378-807">Логический</span><span class="sxs-lookup"><span data-stu-id="00378-807">Boolean</span></span> | | <span data-ttu-id="00378-p152">Используется, только если свойству `type` задано значение `file`. Значение `true` указывает на то, что вложение будет встроено в текст сообщения и не должно отображаться в списке вложений.</span><span class="sxs-lookup"><span data-stu-id="00378-p152">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="00378-810">String</span><span class="sxs-lookup"><span data-stu-id="00378-810">String</span></span> | | <span data-ttu-id="00378-p153">Используется, только если свойству `type` задано значение `item`. Идентификатор вложения EWS. Это строка длиной до 100 символов.</span><span class="sxs-lookup"><span data-stu-id="00378-p153">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="00378-814">function</span><span class="sxs-lookup"><span data-stu-id="00378-814">function</span></span> | <span data-ttu-id="00378-815">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="00378-815">&lt;optional&gt;</span></span> | <span data-ttu-id="00378-816">По завершении работы метода функция, переданная параметру `callback`, вызывается с помощью одного параметра `asyncResult`, представляющего собой объект [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="00378-816">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="00378-817">Требования</span><span class="sxs-lookup"><span data-stu-id="00378-817">Requirements</span></span>

|<span data-ttu-id="00378-818">Требование</span><span class="sxs-lookup"><span data-stu-id="00378-818">Requirement</span></span>| <span data-ttu-id="00378-819">Значение</span><span class="sxs-lookup"><span data-stu-id="00378-819">Value</span></span>|
|---|---|
|[<span data-ttu-id="00378-820">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="00378-820">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="00378-821">1.0</span><span class="sxs-lookup"><span data-stu-id="00378-821">1.0</span></span>|
|[<span data-ttu-id="00378-822">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="00378-822">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="00378-823">ReadItem</span><span class="sxs-lookup"><span data-stu-id="00378-823">ReadItem</span></span>|
|[<span data-ttu-id="00378-824">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="00378-824">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="00378-825">Чтение</span><span class="sxs-lookup"><span data-stu-id="00378-825">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="00378-826">Примеры</span><span class="sxs-lookup"><span data-stu-id="00378-826">Examples</span></span>

<span data-ttu-id="00378-827">Приведенный ниже код передает строку в функцию `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="00378-827">The following code passes a string to the `displayReplyForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="00378-828">Ответ с пустым текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="00378-828">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="00378-829">Ответ только с текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="00378-829">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="00378-830">Ответ с текстом сообщения и вложенным файлом.</span><span class="sxs-lookup"><span data-stu-id="00378-830">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="00378-831">Ответ с текстом сообщения и вложенным элементом.</span><span class="sxs-lookup"><span data-stu-id="00378-831">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="00378-832">Ответ с текстом сообщения, вложенным файлом, вложенным элементом и обратным вызовом.</span><span class="sxs-lookup"><span data-stu-id="00378-832">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlook15officeentities"></a><span data-ttu-id="00378-833">getEntities() → {[Entities](/javascript/api/outlook_1_5/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="00378-833">getEntities() → {[Entities](/javascript/api/outlook_1_5/office.entities)}</span></span>

<span data-ttu-id="00378-834">Получает сущности, обнаруженные в теле выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="00378-834">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="00378-835">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="00378-835">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="00378-836">Требования</span><span class="sxs-lookup"><span data-stu-id="00378-836">Requirements</span></span>

|<span data-ttu-id="00378-837">Требование</span><span class="sxs-lookup"><span data-stu-id="00378-837">Requirement</span></span>| <span data-ttu-id="00378-838">Значение</span><span class="sxs-lookup"><span data-stu-id="00378-838">Value</span></span>|
|---|---|
|[<span data-ttu-id="00378-839">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="00378-839">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="00378-840">1.0</span><span class="sxs-lookup"><span data-stu-id="00378-840">1.0</span></span>|
|[<span data-ttu-id="00378-841">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="00378-841">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="00378-842">ReadItem</span><span class="sxs-lookup"><span data-stu-id="00378-842">ReadItem</span></span>|
|[<span data-ttu-id="00378-843">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="00378-843">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="00378-844">Чтение</span><span class="sxs-lookup"><span data-stu-id="00378-844">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="00378-845">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="00378-845">Returns:</span></span>

<span data-ttu-id="00378-846">Тип: [Entities](/javascript/api/outlook_1_5/office.entities)</span><span class="sxs-lookup"><span data-stu-id="00378-846">Type: [Entities](/javascript/api/outlook_1_5/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="00378-847">Пример</span><span class="sxs-lookup"><span data-stu-id="00378-847">Example</span></span>

<span data-ttu-id="00378-848">Ниже приведен пример получения доступа к сущностям контактов в теле текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="00378-848">The following example accesses the contacts entities in the current item's body.</span></span>

```javascript
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook15officecontactmeetingsuggestionjavascriptapioutlook15officemeetingsuggestionphonenumberjavascriptapioutlook15officephonenumbertasksuggestionjavascriptapioutlook15officetasksuggestion"></a><span data-ttu-id="00378-849">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="00378-849">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))>}</span></span>

<span data-ttu-id="00378-850">Получает массив всех сущностей указанного типа, обнаруженных в теле выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="00378-850">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="00378-851">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="00378-851">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="00378-852">Параметры</span><span class="sxs-lookup"><span data-stu-id="00378-852">Parameters</span></span>

|<span data-ttu-id="00378-853">Имя</span><span class="sxs-lookup"><span data-stu-id="00378-853">Name</span></span>| <span data-ttu-id="00378-854">Тип</span><span class="sxs-lookup"><span data-stu-id="00378-854">Type</span></span>| <span data-ttu-id="00378-855">Описание</span><span class="sxs-lookup"><span data-stu-id="00378-855">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="00378-856">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="00378-856">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook_1_5/office.mailboxenums.entitytype)|<span data-ttu-id="00378-857">Одно из значений перечисления EntityType.</span><span class="sxs-lookup"><span data-stu-id="00378-857">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="00378-858">Требования</span><span class="sxs-lookup"><span data-stu-id="00378-858">Requirements</span></span>

|<span data-ttu-id="00378-859">Требование</span><span class="sxs-lookup"><span data-stu-id="00378-859">Requirement</span></span>| <span data-ttu-id="00378-860">Значение</span><span class="sxs-lookup"><span data-stu-id="00378-860">Value</span></span>|
|---|---|
|[<span data-ttu-id="00378-861">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="00378-861">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="00378-862">1.0</span><span class="sxs-lookup"><span data-stu-id="00378-862">1.0</span></span>|
|[<span data-ttu-id="00378-863">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="00378-863">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="00378-864">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="00378-864">Restricted</span></span>|
|[<span data-ttu-id="00378-865">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="00378-865">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="00378-866">Чтение</span><span class="sxs-lookup"><span data-stu-id="00378-866">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="00378-867">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="00378-867">Returns:</span></span>

<span data-ttu-id="00378-868">Если значение, переданное в `entityType`, не является допустимым членом перечисления `EntityType`, метод возвращает значение NULL.</span><span class="sxs-lookup"><span data-stu-id="00378-868">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="00378-869">Если в теле элемента отсутствуют сущности указанного типа, метод возвращает пустой массив.</span><span class="sxs-lookup"><span data-stu-id="00378-869">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="00378-870">В противном случае тип объектов в возвращаемом массиве зависит от типа сущности, запрошенной в параметре `entityType`.</span><span class="sxs-lookup"><span data-stu-id="00378-870">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="00378-871">Хотя минимальный уровень разрешений для использования этого метода — **Restricted**, для некоторых типов сущностей требуется доступ на уровне **ReadItem**, как указано в приведенной ниже таблице.</span><span class="sxs-lookup"><span data-stu-id="00378-871">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="00378-872">Значение параметра `entityType`</span><span class="sxs-lookup"><span data-stu-id="00378-872">Value of `entityType`</span></span> | <span data-ttu-id="00378-873">Тип объектов в возвращаемом массиве</span><span class="sxs-lookup"><span data-stu-id="00378-873">Type of objects in returned array</span></span> | <span data-ttu-id="00378-874">Необходимый уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="00378-874">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="00378-875">String</span><span class="sxs-lookup"><span data-stu-id="00378-875">String</span></span> | <span data-ttu-id="00378-876">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="00378-876">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="00378-877">Contact</span><span class="sxs-lookup"><span data-stu-id="00378-877">Contact</span></span> | <span data-ttu-id="00378-878">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="00378-878">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="00378-879">String</span><span class="sxs-lookup"><span data-stu-id="00378-879">String</span></span> | <span data-ttu-id="00378-880">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="00378-880">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="00378-881">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="00378-881">MeetingSuggestion</span></span> | <span data-ttu-id="00378-882">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="00378-882">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="00378-883">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="00378-883">PhoneNumber</span></span> | <span data-ttu-id="00378-884">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="00378-884">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="00378-885">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="00378-885">TaskSuggestion</span></span> | <span data-ttu-id="00378-886">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="00378-886">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="00378-887">String</span><span class="sxs-lookup"><span data-stu-id="00378-887">String</span></span> | <span data-ttu-id="00378-888">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="00378-888">**Restricted**</span></span> |

<span data-ttu-id="00378-889">Тип: Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="00378-889">Type: Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="00378-890">Пример</span><span class="sxs-lookup"><span data-stu-id="00378-890">Example</span></span>

<span data-ttu-id="00378-891">В примере ниже показано, как получить доступ к массиву строк, которые представляют собой почтовые адреса в теле текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="00378-891">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook15officecontactmeetingsuggestionjavascriptapioutlook15officemeetingsuggestionphonenumberjavascriptapioutlook15officephonenumbertasksuggestionjavascriptapioutlook15officetasksuggestion"></a><span data-ttu-id="00378-892">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="00378-892">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))>}</span></span>

<span data-ttu-id="00378-893">Возвращает известные сущности в выбранном элементе, которые проходят через именованный фильтр, определяемый в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="00378-893">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="00378-894">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="00378-894">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="00378-895">Метод `getFilteredEntitiesByName` возвращает сущности, соответствующие регулярному выражению, которое определяется в элементе правила [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) в XML-файле манифеста, с использованием указанного значения элемента `FilterName`.</span><span class="sxs-lookup"><span data-stu-id="00378-895">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="00378-896">Параметры</span><span class="sxs-lookup"><span data-stu-id="00378-896">Parameters</span></span>

|<span data-ttu-id="00378-897">Имя</span><span class="sxs-lookup"><span data-stu-id="00378-897">Name</span></span>| <span data-ttu-id="00378-898">Тип</span><span class="sxs-lookup"><span data-stu-id="00378-898">Type</span></span>| <span data-ttu-id="00378-899">Описание</span><span class="sxs-lookup"><span data-stu-id="00378-899">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="00378-900">String</span><span class="sxs-lookup"><span data-stu-id="00378-900">String</span></span>|<span data-ttu-id="00378-901">Имя элемента правила `ItemHasKnownEntity`, определяющее соответствующий фильтр.</span><span class="sxs-lookup"><span data-stu-id="00378-901">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="00378-902">Требования</span><span class="sxs-lookup"><span data-stu-id="00378-902">Requirements</span></span>

|<span data-ttu-id="00378-903">Требование</span><span class="sxs-lookup"><span data-stu-id="00378-903">Requirement</span></span>| <span data-ttu-id="00378-904">Значение</span><span class="sxs-lookup"><span data-stu-id="00378-904">Value</span></span>|
|---|---|
|[<span data-ttu-id="00378-905">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="00378-905">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="00378-906">1.0</span><span class="sxs-lookup"><span data-stu-id="00378-906">1.0</span></span>|
|[<span data-ttu-id="00378-907">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="00378-907">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="00378-908">ReadItem</span><span class="sxs-lookup"><span data-stu-id="00378-908">ReadItem</span></span>|
|[<span data-ttu-id="00378-909">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="00378-909">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="00378-910">Чтение</span><span class="sxs-lookup"><span data-stu-id="00378-910">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="00378-911">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="00378-911">Returns:</span></span>

<span data-ttu-id="00378-p155">Если в манифесте нет элемента `ItemHasKnownEntity` со значением `FilterName`, соответствующим параметру `name`, метод возвращает `null`. Если параметр `name` соответствует элементу `ItemHasKnownEntity` в манифесте, но при этом в текущем элементе нет соответствующих сущностей, метод возвращает пустой массив.</span><span class="sxs-lookup"><span data-stu-id="00378-p155">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="00378-914">Тип: Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="00378-914">Type: Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))></span></span>

#### <a name="getregexmatches--object"></a><span data-ttu-id="00378-915">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="00378-915">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="00378-916">Возвращает строковые значения в выбранном элементе, которые соответствуют регулярным выражениям, определенным в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="00378-916">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="00378-917">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="00378-917">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="00378-p156">Метод `getRegExMatches` возвращает строки, соответствующие регулярному выражению, которое определяется в каждом элементе правила `ItemHasRegularExpressionMatch` или `ItemHasKnownEntity` в XML-файле манифеста. Для правила `ItemHasRegularExpressionMatch` соответствующую строку должно содержать свойство элемента, указанного этим правилом. Простой тип `PropertyName` определяет поддерживаемые свойства.</span><span class="sxs-lookup"><span data-stu-id="00378-p156">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="00378-921">Например, рассмотрим манифест надстройки, который содержит указанный ниже элемент `Rule`.</span><span class="sxs-lookup"><span data-stu-id="00378-921">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="00378-922">Объект, возвращаемый методом `getRegExMatches`, будет содержать два свойства: `fruits` и `veggies`.</span><span class="sxs-lookup"><span data-stu-id="00378-922">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="00378-p157">Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты. Лучше используйте метод [`Body.getAsync`](/javascript/api/outlook_1_5/office.body#getasync-coerciontype--options--callback-) для этого.</span><span class="sxs-lookup"><span data-stu-id="00378-p157">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook_1_5/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="00378-926">Requirements</span><span class="sxs-lookup"><span data-stu-id="00378-926">Requirements</span></span>

|<span data-ttu-id="00378-927">Требование</span><span class="sxs-lookup"><span data-stu-id="00378-927">Requirement</span></span>| <span data-ttu-id="00378-928">Значение</span><span class="sxs-lookup"><span data-stu-id="00378-928">Value</span></span>|
|---|---|
|[<span data-ttu-id="00378-929">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="00378-929">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="00378-930">1.0</span><span class="sxs-lookup"><span data-stu-id="00378-930">1.0</span></span>|
|[<span data-ttu-id="00378-931">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="00378-931">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="00378-932">ReadItem</span><span class="sxs-lookup"><span data-stu-id="00378-932">ReadItem</span></span>|
|[<span data-ttu-id="00378-933">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="00378-933">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="00378-934">Чтение</span><span class="sxs-lookup"><span data-stu-id="00378-934">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="00378-935">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="00378-935">Returns:</span></span>

<span data-ttu-id="00378-p158">Объект, содержащий массив строк, которые соответствуют регулярным выражениям, определяемым в XML-файле манифеста. Имя каждого массива равно соответствующему значению атрибута `RegExName` подходящего правила `ItemHasRegularExpressionMatch` или атрибута `FilterName` соответствующего правила `ItemHasKnownEntity`.</span><span class="sxs-lookup"><span data-stu-id="00378-p158">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="00378-938">

<dt>Тип</dt>

</span><span class="sxs-lookup"><span data-stu-id="00378-938">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="00378-939">Object</span><span class="sxs-lookup"><span data-stu-id="00378-939">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="00378-940">Пример</span><span class="sxs-lookup"><span data-stu-id="00378-940">Example</span></span>

<span data-ttu-id="00378-941">В примере ниже показано, как получить доступ к массиву совпадений для <rule>элементов регулярного выражения `fruits` и `veggies`, которые указаны в манифесте</rule>.</span><span class="sxs-lookup"><span data-stu-id="00378-941">The following example shows how to access the array of matches for the regular expression <rule>elements `fruits` and `veggies`, which are specified in the manifest.</rule></span></span>

```javascript
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="00378-942">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="00378-942">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="00378-943">Возвращает строковые значения в выбранном элементе, которые соответствуют именованному регулярному выражению, определенному в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="00378-943">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="00378-944">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="00378-944">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="00378-945">Метод `getRegExMatchesByName` возвращает строки, соответствующие регулярному выражению, которое определяется в элементе правила `ItemHasRegularExpressionMatch` в XML-файле манифеста, с использованием указанного значения элемента `RegExName`.</span><span class="sxs-lookup"><span data-stu-id="00378-945">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="00378-p159">Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты.</span><span class="sxs-lookup"><span data-stu-id="00378-p159">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="00378-948">Параметры</span><span class="sxs-lookup"><span data-stu-id="00378-948">Parameters</span></span>

|<span data-ttu-id="00378-949">Имя</span><span class="sxs-lookup"><span data-stu-id="00378-949">Name</span></span>| <span data-ttu-id="00378-950">Тип</span><span class="sxs-lookup"><span data-stu-id="00378-950">Type</span></span>| <span data-ttu-id="00378-951">Описание</span><span class="sxs-lookup"><span data-stu-id="00378-951">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="00378-952">String</span><span class="sxs-lookup"><span data-stu-id="00378-952">String</span></span>|<span data-ttu-id="00378-953">Имя элемента правила `ItemHasRegularExpressionMatch`, определяющее соответствующий фильтр.</span><span class="sxs-lookup"><span data-stu-id="00378-953">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="00378-954">Требования</span><span class="sxs-lookup"><span data-stu-id="00378-954">Requirements</span></span>

|<span data-ttu-id="00378-955">Требование</span><span class="sxs-lookup"><span data-stu-id="00378-955">Requirement</span></span>| <span data-ttu-id="00378-956">Значение</span><span class="sxs-lookup"><span data-stu-id="00378-956">Value</span></span>|
|---|---|
|[<span data-ttu-id="00378-957">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="00378-957">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="00378-958">1.0</span><span class="sxs-lookup"><span data-stu-id="00378-958">1.0</span></span>|
|[<span data-ttu-id="00378-959">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="00378-959">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="00378-960">ReadItem</span><span class="sxs-lookup"><span data-stu-id="00378-960">ReadItem</span></span>|
|[<span data-ttu-id="00378-961">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="00378-961">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="00378-962">Чтение</span><span class="sxs-lookup"><span data-stu-id="00378-962">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="00378-963">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="00378-963">Returns:</span></span>

<span data-ttu-id="00378-964">Массив строк, соответствующих регулярному выражению, определяемому в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="00378-964">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="00378-965">

<dt>Тип</dt>

</span><span class="sxs-lookup"><span data-stu-id="00378-965">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="00378-966">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="00378-966">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="00378-967">Пример</span><span class="sxs-lookup"><span data-stu-id="00378-967">Example</span></span>

```javascript
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

#### <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="00378-968">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="00378-968">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="00378-969">Асинхронно возвращает данные, выбранные в теме или тексте сообщения.</span><span class="sxs-lookup"><span data-stu-id="00378-969">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="00378-p160">Если выделенный фрагмент отсутствует, но курсор находится в тексте или теме, метод возвращает значение NULL для выбранных данных. Если выбраны не текст и не тема, метод возвращает ошибку `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="00378-p160">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="00378-972">Параметры</span><span class="sxs-lookup"><span data-stu-id="00378-972">Parameters</span></span>

|<span data-ttu-id="00378-973">Имя</span><span class="sxs-lookup"><span data-stu-id="00378-973">Name</span></span>| <span data-ttu-id="00378-974">Тип</span><span class="sxs-lookup"><span data-stu-id="00378-974">Type</span></span>| <span data-ttu-id="00378-975">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="00378-975">Attributes</span></span>| <span data-ttu-id="00378-976">Описание</span><span class="sxs-lookup"><span data-stu-id="00378-976">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="00378-977">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="00378-977">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="00378-p161">Запрашивает формат данных. Если задано значение Text, метод возвращает обычный текст как строку, удаляя все имеющиеся HTML-теги. Если задано значение HTML, метод возвращает выделенный текст (обычный текст или HTML).</span><span class="sxs-lookup"><span data-stu-id="00378-p161">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="00378-981">Object</span><span class="sxs-lookup"><span data-stu-id="00378-981">Object</span></span>| <span data-ttu-id="00378-982">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="00378-982">&lt;optional&gt;</span></span>|<span data-ttu-id="00378-983">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="00378-983">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="00378-984">Object</span><span class="sxs-lookup"><span data-stu-id="00378-984">Object</span></span>| <span data-ttu-id="00378-985">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="00378-985">&lt;optional&gt;</span></span>|<span data-ttu-id="00378-986">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="00378-986">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="00378-987">функция</span><span class="sxs-lookup"><span data-stu-id="00378-987">function</span></span>||<span data-ttu-id="00378-988">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="00378-988">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="00378-989">Чтобы получить доступ к выбранным данным из метода обратного вызова, вызовите `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="00378-989">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="00378-990">Для доступа к исходному свойству, представляющему собой источник выбранных данных, вызовите параметр `asyncResult.value.sourceProperty`, который может иметь значение `body` или `subject`.</span><span class="sxs-lookup"><span data-stu-id="00378-990">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="00378-991">Требования</span><span class="sxs-lookup"><span data-stu-id="00378-991">Requirements</span></span>

|<span data-ttu-id="00378-992">Требование</span><span class="sxs-lookup"><span data-stu-id="00378-992">Requirement</span></span>| <span data-ttu-id="00378-993">Значение</span><span class="sxs-lookup"><span data-stu-id="00378-993">Value</span></span>|
|---|---|
|[<span data-ttu-id="00378-994">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="00378-994">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="00378-995">1.2</span><span class="sxs-lookup"><span data-stu-id="00378-995">1.2</span></span>|
|[<span data-ttu-id="00378-996">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="00378-996">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="00378-997">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="00378-997">ReadWriteItem</span></span>|
|[<span data-ttu-id="00378-998">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="00378-998">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="00378-999">Создание</span><span class="sxs-lookup"><span data-stu-id="00378-999">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="00378-1000">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="00378-1000">Returns:</span></span>

<span data-ttu-id="00378-1001">Выбранные данные в виде строки с форматом, определенным в параметре `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="00378-1001">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="00378-1002">

<dt>Тип</dt>

</span><span class="sxs-lookup"><span data-stu-id="00378-1002">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="00378-1003">String</span><span class="sxs-lookup"><span data-stu-id="00378-1003">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="00378-1004">Пример</span><span class="sxs-lookup"><span data-stu-id="00378-1004">Example</span></span>

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

#### <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="00378-1005">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="00378-1005">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="00378-1006">Асинхронно загружает настраиваемые свойства для надстройки для выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="00378-1006">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="00378-p163">Настраиваемые свойства сохраняются в виде пар "ключ-значение" для каждого приложения и каждого элемента. Этот метод возвращает объект `CustomProperties` при обратном вызове, который предоставляет методы для доступа к настраиваемым свойствам, характерным для текущего элемента и текущей надстройки. Настраиваемые свойства не шифруются для элемента, поэтому этот способ хранения не является безопасным.</span><span class="sxs-lookup"><span data-stu-id="00378-p163">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="00378-1010">Параметры</span><span class="sxs-lookup"><span data-stu-id="00378-1010">Parameters</span></span>

|<span data-ttu-id="00378-1011">Имя</span><span class="sxs-lookup"><span data-stu-id="00378-1011">Name</span></span>| <span data-ttu-id="00378-1012">Тип</span><span class="sxs-lookup"><span data-stu-id="00378-1012">Type</span></span>| <span data-ttu-id="00378-1013">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="00378-1013">Attributes</span></span>| <span data-ttu-id="00378-1014">Описание</span><span class="sxs-lookup"><span data-stu-id="00378-1014">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="00378-1015">function</span><span class="sxs-lookup"><span data-stu-id="00378-1015">function</span></span>||<span data-ttu-id="00378-1016">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="00378-1016">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="00378-1017">Настраиваемые свойства предоставляются в виде объекта [`CustomProperties`](/javascript/api/outlook_1_5/office.customproperties) в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="00378-1017">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook_1_5/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="00378-1018">Этот объект позволяет получить, задать и удалить настраиваемые свойства для элемента, а также сохранить изменения, внесенные в набор настраиваемых свойств, на сервере.</span><span class="sxs-lookup"><span data-stu-id="00378-1018">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="00378-1019">Объект</span><span class="sxs-lookup"><span data-stu-id="00378-1019">Object</span></span>| <span data-ttu-id="00378-1020">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="00378-1020">&lt;optional&gt;</span></span>|<span data-ttu-id="00378-1021">Разработчики могут указать любой объект, к которому необходимо получить доступ, в функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="00378-1021">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="00378-1022">Доступ к этому объекту можно получить с помощью свойства `asyncResult.asyncContext` в функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="00378-1022">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="00378-1023">Требования</span><span class="sxs-lookup"><span data-stu-id="00378-1023">Requirements</span></span>

|<span data-ttu-id="00378-1024">Требование</span><span class="sxs-lookup"><span data-stu-id="00378-1024">Requirement</span></span>| <span data-ttu-id="00378-1025">Значение</span><span class="sxs-lookup"><span data-stu-id="00378-1025">Value</span></span>|
|---|---|
|[<span data-ttu-id="00378-1026">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="00378-1026">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="00378-1027">1.0</span><span class="sxs-lookup"><span data-stu-id="00378-1027">1.0</span></span>|
|[<span data-ttu-id="00378-1028">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="00378-1028">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="00378-1029">ReadItem</span><span class="sxs-lookup"><span data-stu-id="00378-1029">ReadItem</span></span>|
|[<span data-ttu-id="00378-1030">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="00378-1030">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="00378-1031">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="00378-1031">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="00378-1032">Пример</span><span class="sxs-lookup"><span data-stu-id="00378-1032">Example</span></span>

<span data-ttu-id="00378-p166">Приведенный ниже пример кода показывает, как асинхронно загружать настраиваемые свойства, характерные для текущего элемента, с помощью метода `loadCustomPropertiesAsync`. Этот пример также показывает, как сохранять эти свойства на сервере с помощью метода `CustomProperties.saveAsync`. После загрузки настраиваемых свойств в этом примере кода метод `CustomProperties.get` используется для считывания настраиваемого свойства `myProp`, метод `CustomProperties.set` — для записи настраиваемого свойства `otherProp`, а метод `saveAsync` — для сохранения настраиваемых свойств.</span><span class="sxs-lookup"><span data-stu-id="00378-p166">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

#### <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="00378-1036">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="00378-1036">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="00378-1037">Удаляет вложение из сообщения или встречи.</span><span class="sxs-lookup"><span data-stu-id="00378-1037">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="00378-p167">Метод `removeAttachmentAsync` удаляет из элемента вложение с указанным идентификатором. Идентификатор вложения рекомендуется использовать для удаления вложения, только если оно добавлено тем же почтовым приложением в ходе текущего сеанса. В Outlook Web App и Outlook Web App для устройств идентификатор вложения действителен только в рамках одного сеанса. Сеанс завершается, когда пользователь закрывает приложение или начинает создавать элемент во встроенной форме, а затем переходит из формы в отдельное окно.</span><span class="sxs-lookup"><span data-stu-id="00378-p167">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="00378-1042">Параметры</span><span class="sxs-lookup"><span data-stu-id="00378-1042">Parameters</span></span>

|<span data-ttu-id="00378-1043">Имя</span><span class="sxs-lookup"><span data-stu-id="00378-1043">Name</span></span>| <span data-ttu-id="00378-1044">Тип</span><span class="sxs-lookup"><span data-stu-id="00378-1044">Type</span></span>| <span data-ttu-id="00378-1045">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="00378-1045">Attributes</span></span>| <span data-ttu-id="00378-1046">Описание</span><span class="sxs-lookup"><span data-stu-id="00378-1046">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="00378-1047">String</span><span class="sxs-lookup"><span data-stu-id="00378-1047">String</span></span>||<span data-ttu-id="00378-1048">Идентификатор удаляемого вложения.</span><span class="sxs-lookup"><span data-stu-id="00378-1048">The identifier of the attachment to remove.</span></span>|
|`options`| <span data-ttu-id="00378-1049">Object</span><span class="sxs-lookup"><span data-stu-id="00378-1049">Object</span></span>| <span data-ttu-id="00378-1050">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="00378-1050">&lt;optional&gt;</span></span>|<span data-ttu-id="00378-1051">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="00378-1051">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="00378-1052">Object</span><span class="sxs-lookup"><span data-stu-id="00378-1052">Object</span></span>| <span data-ttu-id="00378-1053">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="00378-1053">&lt;optional&gt;</span></span>|<span data-ttu-id="00378-1054">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="00378-1054">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="00378-1055">функция</span><span class="sxs-lookup"><span data-stu-id="00378-1055">function</span></span>| <span data-ttu-id="00378-1056">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="00378-1056">&lt;optional&gt;</span></span>|<span data-ttu-id="00378-1057">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="00378-1057">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="00378-1058">Если удалить вложение не удается, свойство `asyncResult.error` содержит код ошибки с указанием ее причины.</span><span class="sxs-lookup"><span data-stu-id="00378-1058">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="00378-1059">Ошибки</span><span class="sxs-lookup"><span data-stu-id="00378-1059">Errors</span></span>

| <span data-ttu-id="00378-1060">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="00378-1060">Error code</span></span> | <span data-ttu-id="00378-1061">Описание</span><span class="sxs-lookup"><span data-stu-id="00378-1061">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="00378-1062">Идентификатор вложения не существует.</span><span class="sxs-lookup"><span data-stu-id="00378-1062">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="00378-1063">Требования</span><span class="sxs-lookup"><span data-stu-id="00378-1063">Requirements</span></span>

|<span data-ttu-id="00378-1064">Требование</span><span class="sxs-lookup"><span data-stu-id="00378-1064">Requirement</span></span>| <span data-ttu-id="00378-1065">Значение</span><span class="sxs-lookup"><span data-stu-id="00378-1065">Value</span></span>|
|---|---|
|[<span data-ttu-id="00378-1066">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="00378-1066">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="00378-1067">1.1</span><span class="sxs-lookup"><span data-stu-id="00378-1067">1.1</span></span>|
|[<span data-ttu-id="00378-1068">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="00378-1068">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="00378-1069">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="00378-1069">ReadWriteItem</span></span>|
|[<span data-ttu-id="00378-1070">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="00378-1070">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="00378-1071">Создание</span><span class="sxs-lookup"><span data-stu-id="00378-1071">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="00378-1072">Пример</span><span class="sxs-lookup"><span data-stu-id="00378-1072">Example</span></span>

<span data-ttu-id="00378-1073">Указанный ниже код удаляет вложение с идентификатором "0".</span><span class="sxs-lookup"><span data-stu-id="00378-1073">The following code removes an attachment with an identifier of '0'.</span></span>

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

#### <a name="saveasyncoptions-callback"></a><span data-ttu-id="00378-1074">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="00378-1074">saveAsync([options], callback)</span></span>

<span data-ttu-id="00378-1075">Асинхронно сохраняет элемент.</span><span class="sxs-lookup"><span data-stu-id="00378-1075">Asynchronously saves an item.</span></span>

<span data-ttu-id="00378-p168">При вызове этот метод сохраняет текущее сообщение в виде черновика и возвращает идентификатор элемента с помощью метода обратного вызова. В Outlook Web App или интерактивном режиме Outlook этот элемент сохраняется на сервере. В Outlook в режиме кэширования этот элемент сохраняется в локальном кэше.</span><span class="sxs-lookup"><span data-stu-id="00378-p168">When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="00378-1079">Если в режиме создания надстройка вызывает для элемента метод `saveAsync`, чтобы получить параметр `itemId` для использования с EWS или REST API, необходимо помнить, что синхронизация элемента с сервером может занять много времени, если Outlook работает в режиме кэширования данных.</span><span class="sxs-lookup"><span data-stu-id="00378-1079">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="00378-1080">До окончания синхронизации элемента применение параметра `itemId` будет приводить к ошибке.</span><span class="sxs-lookup"><span data-stu-id="00378-1080">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="00378-p170">Если метод `saveAsync` вызывается для встречи в режиме создания, она сохраняется как обычная встреча в календаре пользователя, а не как черновик. При сохранении новой встречи приглашения не отправляются. При сохранении существующей встречи уведомления отправляются добавленным или удаленным участникам.</span><span class="sxs-lookup"><span data-stu-id="00378-p170">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="00378-1084">Следующие клиенты отличаются другим поведением `saveAsync` в отношении встреч в режиме создания:</span><span class="sxs-lookup"><span data-stu-id="00378-1084">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="00378-1085">Outlook для Mac не поддерживает сохранение собрания.</span><span class="sxs-lookup"><span data-stu-id="00378-1085">Outlook for Mac does not support saving a meeting.</span></span> <span data-ttu-id="00378-1086">Метод `saveAsync` не работает при вызове из собрания в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="00378-1086">The `saveAsync` method fails when called from a meeting in compose mode.</span></span> <span data-ttu-id="00378-1087">Временное решение представлено в статье [Не удается сохранить встречу как черновик в Outlook для Mac с помощью API JS для Office](https://support.microsoft.com/help/4505745).</span><span class="sxs-lookup"><span data-stu-id="00378-1087">See [Cannot save a meeting as a draft in Outlook for Mac by using Office JS API](https://support.microsoft.com/help/4505745) for a workaround.</span></span>
> - <span data-ttu-id="00378-1088">Outlook в Интернете всегда отправляет приглашение или обновление при вызове `saveAsync` для встречи в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="00378-1088">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="00378-1089">Параметры</span><span class="sxs-lookup"><span data-stu-id="00378-1089">Parameters</span></span>

|<span data-ttu-id="00378-1090">Имя</span><span class="sxs-lookup"><span data-stu-id="00378-1090">Name</span></span>| <span data-ttu-id="00378-1091">Тип</span><span class="sxs-lookup"><span data-stu-id="00378-1091">Type</span></span>| <span data-ttu-id="00378-1092">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="00378-1092">Attributes</span></span>| <span data-ttu-id="00378-1093">Описание</span><span class="sxs-lookup"><span data-stu-id="00378-1093">Description</span></span>|
|---|---|---|---|
|`options`| <span data-ttu-id="00378-1094">Object</span><span class="sxs-lookup"><span data-stu-id="00378-1094">Object</span></span>| <span data-ttu-id="00378-1095">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="00378-1095">&lt;optional&gt;</span></span>|<span data-ttu-id="00378-1096">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="00378-1096">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="00378-1097">Object</span><span class="sxs-lookup"><span data-stu-id="00378-1097">Object</span></span>| <span data-ttu-id="00378-1098">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="00378-1098">&lt;optional&gt;</span></span>|<span data-ttu-id="00378-1099">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="00378-1099">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="00378-1100">функция</span><span class="sxs-lookup"><span data-stu-id="00378-1100">function</span></span>||<span data-ttu-id="00378-1101">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="00378-1101">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="00378-1102">После успешного выполнения идентификатор элемента будет предоставлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="00378-1102">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="00378-1103">Требования</span><span class="sxs-lookup"><span data-stu-id="00378-1103">Requirements</span></span>

|<span data-ttu-id="00378-1104">Требование</span><span class="sxs-lookup"><span data-stu-id="00378-1104">Requirement</span></span>| <span data-ttu-id="00378-1105">Значение</span><span class="sxs-lookup"><span data-stu-id="00378-1105">Value</span></span>|
|---|---|
|[<span data-ttu-id="00378-1106">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="00378-1106">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="00378-1107">1.3</span><span class="sxs-lookup"><span data-stu-id="00378-1107">1.3</span></span>|
|[<span data-ttu-id="00378-1108">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="00378-1108">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="00378-1109">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="00378-1109">ReadWriteItem</span></span>|
|[<span data-ttu-id="00378-1110">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="00378-1110">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="00378-1111">Создание</span><span class="sxs-lookup"><span data-stu-id="00378-1111">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="00378-1112">Примеры</span><span class="sxs-lookup"><span data-stu-id="00378-1112">Examples</span></span>

```javascript
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result.
  });
```

<span data-ttu-id="00378-p172">Ниже приведен пример параметра `result`, переданного функции обратного вызова. Свойство `value` содержит идентификатор элемента.</span><span class="sxs-lookup"><span data-stu-id="00378-p172">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

#### <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="00378-1115">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="00378-1115">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="00378-1116">Асинхронно вставляет данные в текст или тему сообщения.</span><span class="sxs-lookup"><span data-stu-id="00378-1116">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="00378-p173">Метод `setSelectedDataAsync` вставляет указанную строку в местоположение курсора в теме или тексте элемента либо, если текст выделен в редакторе, он заменяет выделенный текст. Если курсор находится вне текста или темы элемента, возвращается ошибка. После вставки курсор помещается в конец вставленного содержимого.</span><span class="sxs-lookup"><span data-stu-id="00378-p173">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="00378-1120">Параметры</span><span class="sxs-lookup"><span data-stu-id="00378-1120">Parameters</span></span>

|<span data-ttu-id="00378-1121">Имя</span><span class="sxs-lookup"><span data-stu-id="00378-1121">Name</span></span>| <span data-ttu-id="00378-1122">Тип</span><span class="sxs-lookup"><span data-stu-id="00378-1122">Type</span></span>| <span data-ttu-id="00378-1123">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="00378-1123">Attributes</span></span>| <span data-ttu-id="00378-1124">Описание</span><span class="sxs-lookup"><span data-stu-id="00378-1124">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="00378-1125">String</span><span class="sxs-lookup"><span data-stu-id="00378-1125">String</span></span>||<span data-ttu-id="00378-p174">Вставляемые данные. Объем данных не должен превышать 1 000 000 символов. Если передано больше 1 000 000 символов, возвращается исключение `ArgumentOutOfRange`.</span><span class="sxs-lookup"><span data-stu-id="00378-p174">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="00378-1129">Object</span><span class="sxs-lookup"><span data-stu-id="00378-1129">Object</span></span>| <span data-ttu-id="00378-1130">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="00378-1130">&lt;optional&gt;</span></span>|<span data-ttu-id="00378-1131">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="00378-1131">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="00378-1132">Object</span><span class="sxs-lookup"><span data-stu-id="00378-1132">Object</span></span>| <span data-ttu-id="00378-1133">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="00378-1133">&lt;optional&gt;</span></span>|<span data-ttu-id="00378-1134">В методе обратного вызова разработчики могут указать любой объект, к которому необходимо получить доступ.</span><span class="sxs-lookup"><span data-stu-id="00378-1134">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`| [<span data-ttu-id="00378-1135">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="00378-1135">Office.CoercionType</span></span>](office.md#coerciontype-string)| <span data-ttu-id="00378-1136">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="00378-1136">&lt;optional&gt;</span></span>|<span data-ttu-id="00378-p175">Если задано значение `text`, текущий стиль применяется в Outlook Web App и Outlook. Если поле представляет собой редактор HTML, вставляются только текстовые данные, даже если они имеют формат HTML.</span><span class="sxs-lookup"><span data-stu-id="00378-p175">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="00378-p176">Если задано значение `html` и поле (не тема) поддерживает HTML, в Outlook Web App применяется текущий стиль, а в Outlook — стиль по умолчанию. Если поле является текстовым, возвращается ошибка `InvalidDataFormat`.</span><span class="sxs-lookup"><span data-stu-id="00378-p176">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="00378-1141">Если свойство `coercionType` не задано, результат зависит от поля: если поле имеет формат HTML, используется текст в формате HTML, а если поле текстовое, применяется обычный текст.</span><span class="sxs-lookup"><span data-stu-id="00378-1141">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="00378-1142">функция</span><span class="sxs-lookup"><span data-stu-id="00378-1142">function</span></span>||<span data-ttu-id="00378-1143">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="00378-1143">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="00378-1144">Требования</span><span class="sxs-lookup"><span data-stu-id="00378-1144">Requirements</span></span>

|<span data-ttu-id="00378-1145">Требование</span><span class="sxs-lookup"><span data-stu-id="00378-1145">Requirement</span></span>| <span data-ttu-id="00378-1146">Значение</span><span class="sxs-lookup"><span data-stu-id="00378-1146">Value</span></span>|
|---|---|
|[<span data-ttu-id="00378-1147">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="00378-1147">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="00378-1148">1.2</span><span class="sxs-lookup"><span data-stu-id="00378-1148">1.2</span></span>|
|[<span data-ttu-id="00378-1149">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="00378-1149">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="00378-1150">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="00378-1150">ReadWriteItem</span></span>|
|[<span data-ttu-id="00378-1151">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="00378-1151">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="00378-1152">Создание</span><span class="sxs-lookup"><span data-stu-id="00378-1152">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="00378-1153">Пример</span><span class="sxs-lookup"><span data-stu-id="00378-1153">Example</span></span>

```javascript
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
