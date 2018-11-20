
# <a name="item"></a><span data-ttu-id="aa147-101">item</span><span class="sxs-lookup"><span data-stu-id="aa147-101">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="aa147-102">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="aa147-102">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="aa147-p101">Пространство имен `item` используется для доступа к выбранному в данный момент сообщению, приглашению на собрание или описанию встречи. Вы можете определить тип пространства имен `item` с помощью свойства [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook17officemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="aa147-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook17officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="aa147-105">Requirements</span><span class="sxs-lookup"><span data-stu-id="aa147-105">Requirements</span></span>

|<span data-ttu-id="aa147-106">Requirement</span><span class="sxs-lookup"><span data-stu-id="aa147-106">Requirement</span></span>|<span data-ttu-id="aa147-107">Значение</span><span class="sxs-lookup"><span data-stu-id="aa147-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="aa147-108">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="aa147-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="aa147-109">1.0</span><span class="sxs-lookup"><span data-stu-id="aa147-109">1.0</span></span>|
|[<span data-ttu-id="aa147-110">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="aa147-110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="aa147-111">Restricted</span><span class="sxs-lookup"><span data-stu-id="aa147-111">Restricted</span></span>|
|[<span data-ttu-id="aa147-112">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="aa147-112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="aa147-113">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="aa147-113">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="aa147-114">Элементы и методы</span><span class="sxs-lookup"><span data-stu-id="aa147-114">Members and methods</span></span>

| <span data-ttu-id="aa147-115">Член</span><span class="sxs-lookup"><span data-stu-id="aa147-115">Member</span></span> | <span data-ttu-id="aa147-116">Тип</span><span class="sxs-lookup"><span data-stu-id="aa147-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="aa147-117">attachments</span><span class="sxs-lookup"><span data-stu-id="aa147-117">attachments</span></span>](#attachments-arrayattachmentdetailsjavascriptapioutlook17officeattachmentdetails) | <span data-ttu-id="aa147-118">Элемент</span><span class="sxs-lookup"><span data-stu-id="aa147-118">Member</span></span> |
| [<span data-ttu-id="aa147-119">bcc</span><span class="sxs-lookup"><span data-stu-id="aa147-119">bcc</span></span>](#bcc-recipientsjavascriptapioutlook17officerecipients) | <span data-ttu-id="aa147-120">Элемент</span><span class="sxs-lookup"><span data-stu-id="aa147-120">Member</span></span> |
| [<span data-ttu-id="aa147-121">body</span><span class="sxs-lookup"><span data-stu-id="aa147-121">body</span></span>](#body-bodyjavascriptapioutlook17officebody) | <span data-ttu-id="aa147-122">Элемент</span><span class="sxs-lookup"><span data-stu-id="aa147-122">Member</span></span> |
| [<span data-ttu-id="aa147-123">cc</span><span class="sxs-lookup"><span data-stu-id="aa147-123">cc</span></span>](#cc-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients) | <span data-ttu-id="aa147-124">Элемент</span><span class="sxs-lookup"><span data-stu-id="aa147-124">Member</span></span> |
| [<span data-ttu-id="aa147-125">conversationId</span><span class="sxs-lookup"><span data-stu-id="aa147-125">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="aa147-126">Элемент</span><span class="sxs-lookup"><span data-stu-id="aa147-126">Member</span></span> |
| [<span data-ttu-id="aa147-127">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="aa147-127">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="aa147-128">Элемент</span><span class="sxs-lookup"><span data-stu-id="aa147-128">Member</span></span> |
| [<span data-ttu-id="aa147-129">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="aa147-129">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="aa147-130">Элемент</span><span class="sxs-lookup"><span data-stu-id="aa147-130">Member</span></span> |
| [<span data-ttu-id="aa147-131">end</span><span class="sxs-lookup"><span data-stu-id="aa147-131">end</span></span>](#end-datetimejavascriptapioutlook17officetime) | <span data-ttu-id="aa147-132">Элемент</span><span class="sxs-lookup"><span data-stu-id="aa147-132">Member</span></span> |
| [<span data-ttu-id="aa147-133">from</span><span class="sxs-lookup"><span data-stu-id="aa147-133">from</span></span>](#from-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsfromjavascriptapioutlook17officefrom) | <span data-ttu-id="aa147-134">Элемент</span><span class="sxs-lookup"><span data-stu-id="aa147-134">Member</span></span> |
| [<span data-ttu-id="aa147-135">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="aa147-135">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="aa147-136">Элемент</span><span class="sxs-lookup"><span data-stu-id="aa147-136">Member</span></span> |
| [<span data-ttu-id="aa147-137">itemClass</span><span class="sxs-lookup"><span data-stu-id="aa147-137">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="aa147-138">Элемент</span><span class="sxs-lookup"><span data-stu-id="aa147-138">Member</span></span> |
| [<span data-ttu-id="aa147-139">itemId</span><span class="sxs-lookup"><span data-stu-id="aa147-139">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="aa147-140">Элемент</span><span class="sxs-lookup"><span data-stu-id="aa147-140">Member</span></span> |
| [<span data-ttu-id="aa147-141">itemType</span><span class="sxs-lookup"><span data-stu-id="aa147-141">itemType</span></span>](#itemtype-officemailboxenumsitemtypejavascriptapioutlook17officemailboxenumsitemtype) | <span data-ttu-id="aa147-142">Элемент</span><span class="sxs-lookup"><span data-stu-id="aa147-142">Member</span></span> |
| [<span data-ttu-id="aa147-143">location</span><span class="sxs-lookup"><span data-stu-id="aa147-143">location</span></span>](#location-stringlocationjavascriptapioutlook17officelocation) | <span data-ttu-id="aa147-144">Элемент</span><span class="sxs-lookup"><span data-stu-id="aa147-144">Member</span></span> |
| [<span data-ttu-id="aa147-145">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="aa147-145">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="aa147-146">Элемент</span><span class="sxs-lookup"><span data-stu-id="aa147-146">Member</span></span> |
| [<span data-ttu-id="aa147-147">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="aa147-147">notificationMessages</span></span>](#notificationmessages-notificationmessagesjavascriptapioutlook17officenotificationmessages) | <span data-ttu-id="aa147-148">Элемент</span><span class="sxs-lookup"><span data-stu-id="aa147-148">Member</span></span> |
| [<span data-ttu-id="aa147-149">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="aa147-149">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients) | <span data-ttu-id="aa147-150">Элемент</span><span class="sxs-lookup"><span data-stu-id="aa147-150">Member</span></span> |
| [<span data-ttu-id="aa147-151">organizer</span><span class="sxs-lookup"><span data-stu-id="aa147-151">organizer</span></span>](#organizer-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsorganizerjavascriptapioutlook17officeorganizer) | <span data-ttu-id="aa147-152">Элемент</span><span class="sxs-lookup"><span data-stu-id="aa147-152">Member</span></span> |
| [<span data-ttu-id="aa147-153">recurrence</span><span class="sxs-lookup"><span data-stu-id="aa147-153">recurrence</span></span>](#nullable-recurrence-recurrencejavascriptapioutlook17officerecurrence) | <span data-ttu-id="aa147-154">Элемент</span><span class="sxs-lookup"><span data-stu-id="aa147-154">Member</span></span> |
| [<span data-ttu-id="aa147-155">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="aa147-155">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients) | <span data-ttu-id="aa147-156">Member</span><span class="sxs-lookup"><span data-stu-id="aa147-156">Member</span></span> |
| [<span data-ttu-id="aa147-157">sender</span><span class="sxs-lookup"><span data-stu-id="aa147-157">sender</span></span>](#sender-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetails) | <span data-ttu-id="aa147-158">Элемент</span><span class="sxs-lookup"><span data-stu-id="aa147-158">Member</span></span> |
| [<span data-ttu-id="aa147-159">seriesId</span><span class="sxs-lookup"><span data-stu-id="aa147-159">seriesId</span></span>](#nullable-seriesid-string) | <span data-ttu-id="aa147-160">Элемент</span><span class="sxs-lookup"><span data-stu-id="aa147-160">Member</span></span> |
| [<span data-ttu-id="aa147-161">start</span><span class="sxs-lookup"><span data-stu-id="aa147-161">start</span></span>](#start-datetimejavascriptapioutlook17officetime) | <span data-ttu-id="aa147-162">Элемент</span><span class="sxs-lookup"><span data-stu-id="aa147-162">Member</span></span> |
| [<span data-ttu-id="aa147-163">subject</span><span class="sxs-lookup"><span data-stu-id="aa147-163">subject</span></span>](#subject-stringsubjectjavascriptapioutlook17officesubject) | <span data-ttu-id="aa147-164">Элемент</span><span class="sxs-lookup"><span data-stu-id="aa147-164">Member</span></span> |
| [<span data-ttu-id="aa147-165">to</span><span class="sxs-lookup"><span data-stu-id="aa147-165">to</span></span>](#to-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients) | <span data-ttu-id="aa147-166">Элемент</span><span class="sxs-lookup"><span data-stu-id="aa147-166">Member</span></span> |
| [<span data-ttu-id="aa147-167">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="aa147-167">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="aa147-168">Метод</span><span class="sxs-lookup"><span data-stu-id="aa147-168">Method</span></span> |
| [<span data-ttu-id="aa147-169">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="aa147-169">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="aa147-170">Метод</span><span class="sxs-lookup"><span data-stu-id="aa147-170">Method</span></span> |
| [<span data-ttu-id="aa147-171">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="aa147-171">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="aa147-172">Метод</span><span class="sxs-lookup"><span data-stu-id="aa147-172">Method</span></span> |
| [<span data-ttu-id="aa147-173">close</span><span class="sxs-lookup"><span data-stu-id="aa147-173">close</span></span>](#close) | <span data-ttu-id="aa147-174">Метод</span><span class="sxs-lookup"><span data-stu-id="aa147-174">Method</span></span> |
| [<span data-ttu-id="aa147-175">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="aa147-175">displayReplyAllForm</span></span>](#displayreplyallformformdata) | <span data-ttu-id="aa147-176">Метод</span><span class="sxs-lookup"><span data-stu-id="aa147-176">Method</span></span> |
| [<span data-ttu-id="aa147-177">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="aa147-177">displayReplyForm</span></span>](#displayreplyformformdata) | <span data-ttu-id="aa147-178">Метод</span><span class="sxs-lookup"><span data-stu-id="aa147-178">Method</span></span> |
| [<span data-ttu-id="aa147-179">getEntities</span><span class="sxs-lookup"><span data-stu-id="aa147-179">getEntities</span></span>](#getentities--entitiesjavascriptapioutlook17officeentities) | <span data-ttu-id="aa147-180">Метод</span><span class="sxs-lookup"><span data-stu-id="aa147-180">Method</span></span> |
| [<span data-ttu-id="aa147-181">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="aa147-181">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook17officecontactmeetingsuggestionjavascriptapioutlook17officemeetingsuggestionphonenumberjavascriptapioutlook17officephonenumbertasksuggestionjavascriptapioutlook17officetasksuggestion) | <span data-ttu-id="aa147-182">Метод</span><span class="sxs-lookup"><span data-stu-id="aa147-182">Method</span></span> |
| [<span data-ttu-id="aa147-183">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="aa147-183">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook17officecontactmeetingsuggestionjavascriptapioutlook17officemeetingsuggestionphonenumberjavascriptapioutlook17officephonenumbertasksuggestionjavascriptapioutlook17officetasksuggestion) | <span data-ttu-id="aa147-184">Метод</span><span class="sxs-lookup"><span data-stu-id="aa147-184">Method</span></span> |
| [<span data-ttu-id="aa147-185">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="aa147-185">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="aa147-186">Метод</span><span class="sxs-lookup"><span data-stu-id="aa147-186">Method</span></span> |
| [<span data-ttu-id="aa147-187">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="aa147-187">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="aa147-188">Метод</span><span class="sxs-lookup"><span data-stu-id="aa147-188">Method</span></span> |
| [<span data-ttu-id="aa147-189">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="aa147-189">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="aa147-190">Метод</span><span class="sxs-lookup"><span data-stu-id="aa147-190">Method</span></span> |
| [<span data-ttu-id="aa147-191">getSelectedEntities</span><span class="sxs-lookup"><span data-stu-id="aa147-191">getSelectedEntities</span></span>](#getselectedentities--entitiesjavascriptapioutlook17officeentities) | <span data-ttu-id="aa147-192">Метод</span><span class="sxs-lookup"><span data-stu-id="aa147-192">Method</span></span> |
| [<span data-ttu-id="aa147-193">getSelectedRegExMatches</span><span class="sxs-lookup"><span data-stu-id="aa147-193">getSelectedRegExMatches</span></span>](#getselectedregexmatches--object) | <span data-ttu-id="aa147-194">Метод</span><span class="sxs-lookup"><span data-stu-id="aa147-194">Method</span></span> |
| [<span data-ttu-id="aa147-195">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="aa147-195">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="aa147-196">Метод</span><span class="sxs-lookup"><span data-stu-id="aa147-196">Method</span></span> |
| [<span data-ttu-id="aa147-197">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="aa147-197">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="aa147-198">Метод</span><span class="sxs-lookup"><span data-stu-id="aa147-198">Method</span></span> |
| [<span data-ttu-id="aa147-199">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="aa147-199">removeHandlerAsync</span></span>](#removehandlerasynceventtype-handler-options-callback) | <span data-ttu-id="aa147-200">Метод</span><span class="sxs-lookup"><span data-stu-id="aa147-200">Method</span></span> |
| [<span data-ttu-id="aa147-201">saveAsync</span><span class="sxs-lookup"><span data-stu-id="aa147-201">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="aa147-202">Метод</span><span class="sxs-lookup"><span data-stu-id="aa147-202">Method</span></span> |
| [<span data-ttu-id="aa147-203">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="aa147-203">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="aa147-204">Метод</span><span class="sxs-lookup"><span data-stu-id="aa147-204">Method</span></span> |

### <a name="example"></a><span data-ttu-id="aa147-205">Пример</span><span class="sxs-lookup"><span data-stu-id="aa147-205">Example</span></span>

<span data-ttu-id="aa147-206">В примере кода JavaScript, приведенном ниже, показано, как получить доступ к свойству `subject` текущего элемента в Outlook.</span><span class="sxs-lookup"><span data-stu-id="aa147-206">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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
}
```

### <a name="members"></a><span data-ttu-id="aa147-207">Элементы</span><span class="sxs-lookup"><span data-stu-id="aa147-207">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook17officeattachmentdetails"></a><span data-ttu-id="aa147-208">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_7/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="aa147-208">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_7/office.attachmentdetails)></span></span>

<span data-ttu-id="aa147-p102">Получает массив вложений для элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="aa147-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="aa147-211">Файлы некоторых типов блокируются в Outlook из-за возможных проблем с безопасностью и поэтому не возвращаются.</span><span class="sxs-lookup"><span data-stu-id="aa147-211">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="aa147-212">Дополнительные сведения см. в статье [Блокировка вложений в Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="aa147-212">For more information see [Payments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="aa147-213">Тип:</span><span class="sxs-lookup"><span data-stu-id="aa147-213">Type:</span></span>

*   <span data-ttu-id="aa147-214">Array.<[AttachmentDetails](/javascript/api/outlook_1_7/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="aa147-214">Array.<[AttachmentDetails](/javascript/api/outlook_1_7/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="aa147-215">Требования</span><span class="sxs-lookup"><span data-stu-id="aa147-215">Requirements</span></span>

|<span data-ttu-id="aa147-216">Requirement</span><span class="sxs-lookup"><span data-stu-id="aa147-216">Requirement</span></span>|<span data-ttu-id="aa147-217">Значение</span><span class="sxs-lookup"><span data-stu-id="aa147-217">Value</span></span>|
|---|---|
|[<span data-ttu-id="aa147-218">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="aa147-218">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="aa147-219">1.0</span><span class="sxs-lookup"><span data-stu-id="aa147-219">1.0</span></span>|
|[<span data-ttu-id="aa147-220">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="aa147-220">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="aa147-221">ReadItem</span><span class="sxs-lookup"><span data-stu-id="aa147-221">ReadItem</span></span>|
|[<span data-ttu-id="aa147-222">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="aa147-222">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="aa147-223">Чтение</span><span class="sxs-lookup"><span data-stu-id="aa147-223">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="aa147-224">Пример</span><span class="sxs-lookup"><span data-stu-id="aa147-224">Example</span></span>

<span data-ttu-id="aa147-225">С помощью приведенного ниже кода можно создать HTML-строку с подробными сведениями обо всех вложениях для текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="aa147-225">The following code builds an HTML string with details of all attachments on the current item.</span></span>

```js
var _Item = Office.context.mailbox.item;
var outputString = "";

if (_Item.attachments.length > 0) {
  for (i = 0 ; i < _Item.attachments.length ; i++) {
    var _att = _Item.attachments[i];
    outputString += "<BR>" + i + ". Name: ";
    outputString += _att.name;
    outputString += "<BR>ID: " + _att.id;
    outputString += "<BR>contentType: " + _att.contentType;
    outputString += "<BR>size: " + _att.size;
    outputString += "<BR>attachmentType: " + _att.attachmentType;
    outputString += "<BR>isInline: " + _att.isInline;
  }
}

// Do something with outputString
```

####  <a name="bcc-recipientsjavascriptapioutlook17officerecipients"></a><span data-ttu-id="aa147-226">bcc :[Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="aa147-226">bcc :[Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

<span data-ttu-id="aa147-227">Извлекает объект, предоставляющий методы для получения или обновления получателей, которые указаны в строке СК (скрытая копия) сообщения.</span><span class="sxs-lookup"><span data-stu-id="aa147-227">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="aa147-228">Только в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="aa147-228">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="aa147-229">Тип:</span><span class="sxs-lookup"><span data-stu-id="aa147-229">Type:</span></span>

*   [<span data-ttu-id="aa147-230">Recipients</span><span class="sxs-lookup"><span data-stu-id="aa147-230">Recipients</span></span>](/javascript/api/outlook_1_7/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="aa147-231">Требования</span><span class="sxs-lookup"><span data-stu-id="aa147-231">Requirements</span></span>

|<span data-ttu-id="aa147-232">Requirement</span><span class="sxs-lookup"><span data-stu-id="aa147-232">Requirement</span></span>|<span data-ttu-id="aa147-233">Значение</span><span class="sxs-lookup"><span data-stu-id="aa147-233">Value</span></span>|
|---|---|
|[<span data-ttu-id="aa147-234">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="aa147-234">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="aa147-235">1.1</span><span class="sxs-lookup"><span data-stu-id="aa147-235">1.1</span></span>|
|[<span data-ttu-id="aa147-236">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="aa147-236">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="aa147-237">ReadItem</span><span class="sxs-lookup"><span data-stu-id="aa147-237">ReadItem</span></span>|
|[<span data-ttu-id="aa147-238">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="aa147-238">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="aa147-239">Создание</span><span class="sxs-lookup"><span data-stu-id="aa147-239">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="aa147-240">Пример</span><span class="sxs-lookup"><span data-stu-id="aa147-240">Example</span></span>

```js
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlook17officebody"></a><span data-ttu-id="aa147-241">body :[Body](/javascript/api/outlook_1_7/office.body)</span><span class="sxs-lookup"><span data-stu-id="aa147-241">body :[Body](/javascript/api/outlook_1_7/office.body)</span></span>

<span data-ttu-id="aa147-242">Получает объект, предоставляющий методы для работы с основным текстом элемента.</span><span class="sxs-lookup"><span data-stu-id="aa147-242">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="aa147-243">Тип:</span><span class="sxs-lookup"><span data-stu-id="aa147-243">Type:</span></span>

*   [<span data-ttu-id="aa147-244">Body</span><span class="sxs-lookup"><span data-stu-id="aa147-244">Body</span></span>](/javascript/api/outlook_1_7/office.body)

##### <a name="requirements"></a><span data-ttu-id="aa147-245">Требования</span><span class="sxs-lookup"><span data-stu-id="aa147-245">Requirements</span></span>

|<span data-ttu-id="aa147-246">Requirement</span><span class="sxs-lookup"><span data-stu-id="aa147-246">Requirement</span></span>|<span data-ttu-id="aa147-247">Значение</span><span class="sxs-lookup"><span data-stu-id="aa147-247">Value</span></span>|
|---|---|
|[<span data-ttu-id="aa147-248">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="aa147-248">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="aa147-249">1.1</span><span class="sxs-lookup"><span data-stu-id="aa147-249">1.1</span></span>|
|[<span data-ttu-id="aa147-250">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="aa147-250">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="aa147-251">ReadItem</span><span class="sxs-lookup"><span data-stu-id="aa147-251">ReadItem</span></span>|
|[<span data-ttu-id="aa147-252">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="aa147-252">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="aa147-253">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="aa147-253">Compose or read</span></span>|

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients"></a><span data-ttu-id="aa147-254">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="aa147-254">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

<span data-ttu-id="aa147-255">Предоставляет доступ к получателям копии сообщения.</span><span class="sxs-lookup"><span data-stu-id="aa147-255">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="aa147-256">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="aa147-256">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="aa147-257">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="aa147-257">Read mode</span></span>

<span data-ttu-id="aa147-p106">Свойство `cc` возвращает массив, который содержит объект `EmailAddressDetails` для каждого получателя, указанного в строке **Копия** сообщения. Коллекция может включать не более 100 элементов.</span><span class="sxs-lookup"><span data-stu-id="aa147-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="aa147-260">Режим создания</span><span class="sxs-lookup"><span data-stu-id="aa147-260">Compose mode</span></span>

<span data-ttu-id="aa147-261">Свойство `cc` возвращает объект `Recipients`, предоставляющий методы для получения или обновления получателей, которые указаны в строке **Копия** сообщения.</span><span class="sxs-lookup"><span data-stu-id="aa147-261">The `cc` property returns a `Recipients` object that provides methods for manipulating the recipients on the **Cc** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="aa147-262">Тип:</span><span class="sxs-lookup"><span data-stu-id="aa147-262">Type:</span></span>

*   <span data-ttu-id="aa147-263">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="aa147-263">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="aa147-264">Требования</span><span class="sxs-lookup"><span data-stu-id="aa147-264">Requirements</span></span>

|<span data-ttu-id="aa147-265">Requirement</span><span class="sxs-lookup"><span data-stu-id="aa147-265">Requirement</span></span>|<span data-ttu-id="aa147-266">Значение</span><span class="sxs-lookup"><span data-stu-id="aa147-266">Value</span></span>|
|---|---|
|[<span data-ttu-id="aa147-267">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="aa147-267">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="aa147-268">1.0</span><span class="sxs-lookup"><span data-stu-id="aa147-268">1.0</span></span>|
|[<span data-ttu-id="aa147-269">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="aa147-269">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="aa147-270">ReadItem</span><span class="sxs-lookup"><span data-stu-id="aa147-270">ReadItem</span></span>|
|[<span data-ttu-id="aa147-271">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="aa147-271">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="aa147-272">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="aa147-272">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="aa147-273">Пример</span><span class="sxs-lookup"><span data-stu-id="aa147-273">Example</span></span>

```js
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="aa147-274">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="aa147-274">(nullable) conversationId :String</span></span>

<span data-ttu-id="aa147-275">Получает идентификатор разговора по электронной почте, содержащего конкретное сообщение.</span><span class="sxs-lookup"><span data-stu-id="aa147-275">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="aa147-p107">Вы можете получить целочисленное значение этого свойства, если ваше почтовое приложение активируется в формах просмотра или формах создания ответов. Если пользователь изменит тему ответа, после его отправки идентификатор беседы будет изменен, и полученное ранее значение будет недействительным.</span><span class="sxs-lookup"><span data-stu-id="aa147-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="aa147-p108">Это свойство имеет значение NULL для нового элемента в форме создания. Свойство `conversationId` вернет значение, если пользователь задаст тему и сохранит элемент.</span><span class="sxs-lookup"><span data-stu-id="aa147-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="aa147-280">Тип:</span><span class="sxs-lookup"><span data-stu-id="aa147-280">Type:</span></span>

*   <span data-ttu-id="aa147-281">String</span><span class="sxs-lookup"><span data-stu-id="aa147-281">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="aa147-282">Требования</span><span class="sxs-lookup"><span data-stu-id="aa147-282">Requirements</span></span>

|<span data-ttu-id="aa147-283">Requirement</span><span class="sxs-lookup"><span data-stu-id="aa147-283">Requirement</span></span>|<span data-ttu-id="aa147-284">Значение</span><span class="sxs-lookup"><span data-stu-id="aa147-284">Value</span></span>|
|---|---|
|[<span data-ttu-id="aa147-285">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="aa147-285">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="aa147-286">1.0</span><span class="sxs-lookup"><span data-stu-id="aa147-286">1.0</span></span>|
|[<span data-ttu-id="aa147-287">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="aa147-287">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="aa147-288">ReadItem</span><span class="sxs-lookup"><span data-stu-id="aa147-288">ReadItem</span></span>|
|[<span data-ttu-id="aa147-289">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="aa147-289">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="aa147-290">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="aa147-290">Compose or read</span></span>|

#### <a name="datetimecreated-date"></a><span data-ttu-id="aa147-291">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="aa147-291">dateTimeCreated :Date</span></span>

<span data-ttu-id="aa147-p109">Получает дату и время создания элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="aa147-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="aa147-294">Тип:</span><span class="sxs-lookup"><span data-stu-id="aa147-294">Type:</span></span>

*   <span data-ttu-id="aa147-295">Date</span><span class="sxs-lookup"><span data-stu-id="aa147-295">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="aa147-296">Требования</span><span class="sxs-lookup"><span data-stu-id="aa147-296">Requirements</span></span>

|<span data-ttu-id="aa147-297">Requirement</span><span class="sxs-lookup"><span data-stu-id="aa147-297">Requirement</span></span>|<span data-ttu-id="aa147-298">Значение</span><span class="sxs-lookup"><span data-stu-id="aa147-298">Value</span></span>|
|---|---|
|[<span data-ttu-id="aa147-299">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="aa147-299">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="aa147-300">1.0</span><span class="sxs-lookup"><span data-stu-id="aa147-300">1.0</span></span>|
|[<span data-ttu-id="aa147-301">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="aa147-301">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="aa147-302">ReadItem</span><span class="sxs-lookup"><span data-stu-id="aa147-302">ReadItem</span></span>|
|[<span data-ttu-id="aa147-303">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="aa147-303">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="aa147-304">Чтение</span><span class="sxs-lookup"><span data-stu-id="aa147-304">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="aa147-305">Пример</span><span class="sxs-lookup"><span data-stu-id="aa147-305">Example</span></span>

```js
var created = Office.context.mailbox.item.dateTimeCreated;
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="aa147-306">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="aa147-306">dateTimeModified :Date</span></span>

<span data-ttu-id="aa147-p110">Получает дату и время последнего изменения элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="aa147-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="aa147-309">Этот элемент не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="aa147-309">Note: This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="aa147-310">Тип:</span><span class="sxs-lookup"><span data-stu-id="aa147-310">Type:</span></span>

*   <span data-ttu-id="aa147-311">Date</span><span class="sxs-lookup"><span data-stu-id="aa147-311">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="aa147-312">Требования</span><span class="sxs-lookup"><span data-stu-id="aa147-312">Requirements</span></span>

|<span data-ttu-id="aa147-313">Requirement</span><span class="sxs-lookup"><span data-stu-id="aa147-313">Requirement</span></span>|<span data-ttu-id="aa147-314">Значение</span><span class="sxs-lookup"><span data-stu-id="aa147-314">Value</span></span>|
|---|---|
|[<span data-ttu-id="aa147-315">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="aa147-315">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="aa147-316">1.0</span><span class="sxs-lookup"><span data-stu-id="aa147-316">1.0</span></span>|
|[<span data-ttu-id="aa147-317">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="aa147-317">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="aa147-318">ReadItem</span><span class="sxs-lookup"><span data-stu-id="aa147-318">ReadItem</span></span>|
|[<span data-ttu-id="aa147-319">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="aa147-319">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="aa147-320">Чтение</span><span class="sxs-lookup"><span data-stu-id="aa147-320">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="aa147-321">Пример</span><span class="sxs-lookup"><span data-stu-id="aa147-321">Example</span></span>

```js
var modified = Office.context.mailbox.item.dateTimeModified;
```

####  <a name="end-datetimejavascriptapioutlook17officetime"></a><span data-ttu-id="aa147-322">end :Date|[Time](/javascript/api/outlook_1_7/office.time)</span><span class="sxs-lookup"><span data-stu-id="aa147-322">end :Date|[Time](/javascript/api/outlook_1_7/office.time)</span></span>

<span data-ttu-id="aa147-323">Получает или задает дату и время окончания встречи.</span><span class="sxs-lookup"><span data-stu-id="aa147-323">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="aa147-p111">Свойство `end` представлено в виде значения даты и времени в формате UTC. Преобразовать значение свойства end в местные значения даты и времени клиента можно с помощью метода [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook17officelocalclienttime).</span><span class="sxs-lookup"><span data-stu-id="aa147-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook17officelocalclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="aa147-326">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="aa147-326">Read mode</span></span>

<span data-ttu-id="aa147-327">Свойство `end` возвращает объект `Date`.</span><span class="sxs-lookup"><span data-stu-id="aa147-327">The `end` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="aa147-328">Режим создания</span><span class="sxs-lookup"><span data-stu-id="aa147-328">Compose mode</span></span>

<span data-ttu-id="aa147-329">Свойство `end` возвращает объект `Time`.</span><span class="sxs-lookup"><span data-stu-id="aa147-329">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="aa147-330">Если вы задаете время окончания с помощью метода [`Time.setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-), необходимо использовать метод [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) для преобразования местного времени на клиенте в формат UTC для сервера.</span><span class="sxs-lookup"><span data-stu-id="aa147-330">When you use the [`Time.setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="aa147-331">Тип:</span><span class="sxs-lookup"><span data-stu-id="aa147-331">Type:</span></span>

*   <span data-ttu-id="aa147-332">Date | [Time](/javascript/api/outlook_1_7/office.time)</span><span class="sxs-lookup"><span data-stu-id="aa147-332">Date | [Time](/javascript/api/outlook_1_7/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="aa147-333">Требования</span><span class="sxs-lookup"><span data-stu-id="aa147-333">Requirements</span></span>

|<span data-ttu-id="aa147-334">Requirement</span><span class="sxs-lookup"><span data-stu-id="aa147-334">Requirement</span></span>|<span data-ttu-id="aa147-335">Значение</span><span class="sxs-lookup"><span data-stu-id="aa147-335">Value</span></span>|
|---|---|
|[<span data-ttu-id="aa147-336">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="aa147-336">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="aa147-337">1.0</span><span class="sxs-lookup"><span data-stu-id="aa147-337">1.0</span></span>|
|[<span data-ttu-id="aa147-338">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="aa147-338">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="aa147-339">ReadItem</span><span class="sxs-lookup"><span data-stu-id="aa147-339">ReadItem</span></span>|
|[<span data-ttu-id="aa147-340">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="aa147-340">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="aa147-341">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="aa147-341">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="aa147-342">Пример</span><span class="sxs-lookup"><span data-stu-id="aa147-342">Example</span></span>

<span data-ttu-id="aa147-343">В примере ниже показано, как с помощью метода [`setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) объекта `Time` задать время окончания встречи в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="aa147-343">The following example sets the end time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

```js
var endTime = new Date("3/14/2015");
var options = {
  // Pass information that can be used
  // in the callback
  asyncContext: {verb:"Set"}
}
Office.context.mailbox.item.end.setAsync(endTime, options, function(result) {
  if (result.error) {
    console.debug(result.error);
  } else {
    // Access the asyncContext that was passed to the setAsync function
    console.debug("End Time " + result.asyncContext.verb);
  }
});
```

#### <a name="from-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsfromjavascriptapioutlook17officefrom"></a><span data-ttu-id="aa147-344">from :[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)|[From](/javascript/api/outlook_1_7/office.from)</span><span class="sxs-lookup"><span data-stu-id="aa147-344">from :[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)|[From](/javascript/api/outlook_1_7/office.from)</span></span>

<span data-ttu-id="aa147-345">Получает адрес электронной почты отправителя сообщения.</span><span class="sxs-lookup"><span data-stu-id="aa147-345">Gets the email address of the sender of a message.</span></span>

<span data-ttu-id="aa147-p112">Свойства `from` и [`sender`](#sender-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetails) представляют одно лицо, если сообщение не отправлено представителем. В противном случае свойство `from` представляет лицо, делегировавшее полномочия, а свойство sender — представителя.</span><span class="sxs-lookup"><span data-stu-id="aa147-p112">The `from` and [`sender`](#sender-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="aa147-348">Свойству `recipientType`, принадлежащему объекту `EmailAddressDetails` в свойстве `from`, задано значение `undefined`.</span><span class="sxs-lookup"><span data-stu-id="aa147-348">Note: The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="aa147-349">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="aa147-349">Read mode</span></span>

<span data-ttu-id="aa147-350">Свойство `from` возвращает объект `EmailAddressDetails`.</span><span class="sxs-lookup"><span data-stu-id="aa147-350">The `from` property returns a `EmailAddressDetails` object.</span></span>

```js
var subject = Office.context.mailbox.item.from;
```

##### <a name="compose-mode"></a><span data-ttu-id="aa147-351">Режим создания</span><span class="sxs-lookup"><span data-stu-id="aa147-351">Compose mode</span></span>

<span data-ttu-id="aa147-352">Свойство `from` возвращает объект `From`, который предоставляет метод для получения значения отправителя.</span><span class="sxs-lookup"><span data-stu-id="aa147-352">The `from` property returns a `From` object that provides a method to get the from value.</span></span>

```js
Office.context.mailbox.item.from.getAsync(callback);

function callback(asyncResult) {
  var from = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="aa147-353">Тип:</span><span class="sxs-lookup"><span data-stu-id="aa147-353">Type:</span></span>

*   <span data-ttu-id="aa147-354">[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails) | [From](/javascript/api/outlook_1_7/office.from)</span><span class="sxs-lookup"><span data-stu-id="aa147-354">[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails) | [From](/javascript/api/outlook_1_7/office.from)</span></span>

##### <a name="requirements"></a><span data-ttu-id="aa147-355">Требования</span><span class="sxs-lookup"><span data-stu-id="aa147-355">Requirements</span></span>

|<span data-ttu-id="aa147-356">Требование</span><span class="sxs-lookup"><span data-stu-id="aa147-356">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="aa147-357">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="aa147-357">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="aa147-358">1.0</span><span class="sxs-lookup"><span data-stu-id="aa147-358">1.0</span></span>|<span data-ttu-id="aa147-359">1.7</span><span class="sxs-lookup"><span data-stu-id="aa147-359">ExcelApi 1.7 Beta</span></span>|
|[<span data-ttu-id="aa147-360">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="aa147-360">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="aa147-361">ReadItem</span><span class="sxs-lookup"><span data-stu-id="aa147-361">ReadItem</span></span>|<span data-ttu-id="aa147-362">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="aa147-362">ReadWriteItem</span></span>|
|[<span data-ttu-id="aa147-363">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="aa147-363">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="aa147-364">Чтение</span><span class="sxs-lookup"><span data-stu-id="aa147-364">Read</span></span>|<span data-ttu-id="aa147-365">Создание</span><span class="sxs-lookup"><span data-stu-id="aa147-365">Compose</span></span>|

#### <a name="internetmessageid-string"></a><span data-ttu-id="aa147-366">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="aa147-366">internetMessageId :String</span></span>

<span data-ttu-id="aa147-p113">Получает идентификатор интернет-сообщения для электронного сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="aa147-p113">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="aa147-369">Тип:</span><span class="sxs-lookup"><span data-stu-id="aa147-369">Type:</span></span>

*   <span data-ttu-id="aa147-370">String</span><span class="sxs-lookup"><span data-stu-id="aa147-370">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="aa147-371">Требования</span><span class="sxs-lookup"><span data-stu-id="aa147-371">Requirements</span></span>

|<span data-ttu-id="aa147-372">Requirement</span><span class="sxs-lookup"><span data-stu-id="aa147-372">Requirement</span></span>|<span data-ttu-id="aa147-373">Значение</span><span class="sxs-lookup"><span data-stu-id="aa147-373">Value</span></span>|
|---|---|
|[<span data-ttu-id="aa147-374">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="aa147-374">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="aa147-375">1.0</span><span class="sxs-lookup"><span data-stu-id="aa147-375">1.0</span></span>|
|[<span data-ttu-id="aa147-376">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="aa147-376">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="aa147-377">ReadItem</span><span class="sxs-lookup"><span data-stu-id="aa147-377">ReadItem</span></span>|
|[<span data-ttu-id="aa147-378">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="aa147-378">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="aa147-379">Чтение</span><span class="sxs-lookup"><span data-stu-id="aa147-379">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="aa147-380">Пример</span><span class="sxs-lookup"><span data-stu-id="aa147-380">Example</span></span>

```js
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="aa147-381">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="aa147-381">itemClass :String</span></span>

<span data-ttu-id="aa147-p114">Получает класс элемента веб-служб Exchange для выбранного элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="aa147-p114">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="aa147-p115">Свойство `itemClass` указывает класс сообщения выбранного элемента. Ниже приводятся классы сообщения по умолчанию для элемента сообщения или встречи.</span><span class="sxs-lookup"><span data-stu-id="aa147-p115">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

|<span data-ttu-id="aa147-386">Тип</span><span class="sxs-lookup"><span data-stu-id="aa147-386">Type</span></span>|<span data-ttu-id="aa147-387">Описание</span><span class="sxs-lookup"><span data-stu-id="aa147-387">Description</span></span>|<span data-ttu-id="aa147-388">Класс элемента</span><span class="sxs-lookup"><span data-stu-id="aa147-388">item class</span></span>|
|---|---|---|
|<span data-ttu-id="aa147-389">Элементы встречи</span><span class="sxs-lookup"><span data-stu-id="aa147-389">Appointment items</span></span>|<span data-ttu-id="aa147-390">Это элементы календаря для класса элемента `IPM.Appointment` или `IPM.Appointment.Occurence`.</span><span class="sxs-lookup"><span data-stu-id="aa147-390">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurence`.</span></span>|`IPM.Appointment`<br />`IPM.Appointment.Occurence`|
|<span data-ttu-id="aa147-391">Элементы сообщения</span><span class="sxs-lookup"><span data-stu-id="aa147-391">Message items</span></span>|<span data-ttu-id="aa147-392">Сюда входят электронные сообщения, для которых по умолчанию задан класс сообщения `IPM.Note`, а также приглашения на собрания, ответы на них и уведомления об их отмене, использующие `IPM.Schedule.Meeting` в качестве базового класса сообщения.</span><span class="sxs-lookup"><span data-stu-id="aa147-392">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span>|`IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled`|

<span data-ttu-id="aa147-393">Можно создавать настраиваемые классы сообщения, расширяющие классы сообщения по умолчанию, например настраиваемый класс сообщения о встрече `IPM.Appointment.Contoso`.</span><span class="sxs-lookup"><span data-stu-id="aa147-393">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="aa147-394">Тип:</span><span class="sxs-lookup"><span data-stu-id="aa147-394">Type:</span></span>

*   <span data-ttu-id="aa147-395">String</span><span class="sxs-lookup"><span data-stu-id="aa147-395">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="aa147-396">Требования</span><span class="sxs-lookup"><span data-stu-id="aa147-396">Requirements</span></span>

|<span data-ttu-id="aa147-397">Requirement</span><span class="sxs-lookup"><span data-stu-id="aa147-397">Requirement</span></span>|<span data-ttu-id="aa147-398">Значение</span><span class="sxs-lookup"><span data-stu-id="aa147-398">Value</span></span>|
|---|---|
|[<span data-ttu-id="aa147-399">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="aa147-399">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="aa147-400">1.0</span><span class="sxs-lookup"><span data-stu-id="aa147-400">1.0</span></span>|
|[<span data-ttu-id="aa147-401">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="aa147-401">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="aa147-402">ReadItem</span><span class="sxs-lookup"><span data-stu-id="aa147-402">ReadItem</span></span>|
|[<span data-ttu-id="aa147-403">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="aa147-403">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="aa147-404">Чтение</span><span class="sxs-lookup"><span data-stu-id="aa147-404">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="aa147-405">Пример</span><span class="sxs-lookup"><span data-stu-id="aa147-405">Example</span></span>

```js
var itemClass = Office.context.mailbox.item.itemClass;
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="aa147-406">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="aa147-406">(nullable) itemId :String</span></span>

<span data-ttu-id="aa147-p116">Получает идентификатор элемента веб-служб Exchange для текущего элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="aa147-p116">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="aa147-409">Идентификатор, возвращаемый свойством `itemId`, совпадает с идентификатором элемента веб-служб Exchange.</span><span class="sxs-lookup"><span data-stu-id="aa147-409">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="aa147-410">Свойство `itemId` не совпадает с идентификатором записи Outlook, а также идентификатором, который используется REST API Outlook.</span><span class="sxs-lookup"><span data-stu-id="aa147-410">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="aa147-411">Прежде чем совершать вызовы REST API, используя это значение, его необходимо преобразовать с помощью [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="aa147-411">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="aa147-412">Дополнительные сведения см. в статье [Использование REST API Outlook из надстройки Outlook](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="aa147-412">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="aa147-p118">Свойство `itemId` недоступно в режиме создания. Если требуется идентификатор элемента, с помощью метода [`saveAsync`](#saveasyncoptions-callback) можно сохранить элемент в хранилище. При этом в параметре [`AsyncResult.value`](/javascript/api/office/office.asyncresult) функции обратного вызова возвращается идентификатор элемента.</span><span class="sxs-lookup"><span data-stu-id="aa147-p118">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="aa147-415">Тип:</span><span class="sxs-lookup"><span data-stu-id="aa147-415">Type:</span></span>

*   <span data-ttu-id="aa147-416">String</span><span class="sxs-lookup"><span data-stu-id="aa147-416">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="aa147-417">Требования</span><span class="sxs-lookup"><span data-stu-id="aa147-417">Requirements</span></span>

|<span data-ttu-id="aa147-418">Requirement</span><span class="sxs-lookup"><span data-stu-id="aa147-418">Requirement</span></span>|<span data-ttu-id="aa147-419">Значение</span><span class="sxs-lookup"><span data-stu-id="aa147-419">Value</span></span>|
|---|---|
|[<span data-ttu-id="aa147-420">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="aa147-420">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="aa147-421">1.0</span><span class="sxs-lookup"><span data-stu-id="aa147-421">1.0</span></span>|
|[<span data-ttu-id="aa147-422">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="aa147-422">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="aa147-423">ReadItem</span><span class="sxs-lookup"><span data-stu-id="aa147-423">ReadItem</span></span>|
|[<span data-ttu-id="aa147-424">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="aa147-424">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="aa147-425">Чтение</span><span class="sxs-lookup"><span data-stu-id="aa147-425">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="aa147-426">Пример</span><span class="sxs-lookup"><span data-stu-id="aa147-426">Example</span></span>

<span data-ttu-id="aa147-p119">Указанный ниже код проверяет наличие идентификатора элемента. Если свойство `itemId` возвращает значение `null` или `undefined`, элемент будет сохранен в хранилище, а из асинхронного результата будет получен идентификатор элемента.</span><span class="sxs-lookup"><span data-stu-id="aa147-p119">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```js
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result){
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook17officemailboxenumsitemtype"></a><span data-ttu-id="aa147-429">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_7/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="aa147-429">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_7/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="aa147-430">Получает тип элемента, который представляет экземпляр.</span><span class="sxs-lookup"><span data-stu-id="aa147-430">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="aa147-431">Свойство `itemType` возвращает одно из значений перечисления `ItemType`, которое указывает, является ли экземпляр объекта `item` сообщением или встречей.</span><span class="sxs-lookup"><span data-stu-id="aa147-431">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="aa147-432">Тип:</span><span class="sxs-lookup"><span data-stu-id="aa147-432">Type:</span></span>

*   [<span data-ttu-id="aa147-433">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="aa147-433">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook_1_7/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="aa147-434">Требования</span><span class="sxs-lookup"><span data-stu-id="aa147-434">Requirements</span></span>

|<span data-ttu-id="aa147-435">Requirement</span><span class="sxs-lookup"><span data-stu-id="aa147-435">Requirement</span></span>|<span data-ttu-id="aa147-436">Значение</span><span class="sxs-lookup"><span data-stu-id="aa147-436">Value</span></span>|
|---|---|
|[<span data-ttu-id="aa147-437">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="aa147-437">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="aa147-438">1.0</span><span class="sxs-lookup"><span data-stu-id="aa147-438">1.0</span></span>|
|[<span data-ttu-id="aa147-439">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="aa147-439">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="aa147-440">ReadItem</span><span class="sxs-lookup"><span data-stu-id="aa147-440">ReadItem</span></span>|
|[<span data-ttu-id="aa147-441">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="aa147-441">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="aa147-442">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="aa147-442">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="aa147-443">Пример</span><span class="sxs-lookup"><span data-stu-id="aa147-443">Example</span></span>

```js
if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Message)
  // do something
else
  // do something else
```

####  <a name="location-stringlocationjavascriptapioutlook17officelocation"></a><span data-ttu-id="aa147-444">location :String|[Location](/javascript/api/outlook_1_7/office.location)</span><span class="sxs-lookup"><span data-stu-id="aa147-444">location :String|[Location](/javascript/api/outlook_1_7/office.location)</span></span>

<span data-ttu-id="aa147-445">Получает или задает место встречи.</span><span class="sxs-lookup"><span data-stu-id="aa147-445">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="aa147-446">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="aa147-446">Read mode</span></span>

<span data-ttu-id="aa147-447">Свойство `location` возвращает строку, содержащую сведения о месте встречи.</span><span class="sxs-lookup"><span data-stu-id="aa147-447">The `location` property returns a string that contains the location of the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="aa147-448">Режим создания</span><span class="sxs-lookup"><span data-stu-id="aa147-448">Compose mode</span></span>

<span data-ttu-id="aa147-449">Свойство `location` возвращает объект `Location`, предоставляющий методы, которые используются для получения и задания места встречи.</span><span class="sxs-lookup"><span data-stu-id="aa147-449">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="aa147-450">Тип:</span><span class="sxs-lookup"><span data-stu-id="aa147-450">Type:</span></span>

*   <span data-ttu-id="aa147-451">String | [Location](/javascript/api/outlook_1_7/office.location)</span><span class="sxs-lookup"><span data-stu-id="aa147-451">String | [Location](/javascript/api/outlook_1_7/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="aa147-452">Требования</span><span class="sxs-lookup"><span data-stu-id="aa147-452">Requirements</span></span>

|<span data-ttu-id="aa147-453">Requirement</span><span class="sxs-lookup"><span data-stu-id="aa147-453">Requirement</span></span>|<span data-ttu-id="aa147-454">Значение</span><span class="sxs-lookup"><span data-stu-id="aa147-454">Value</span></span>|
|---|---|
|[<span data-ttu-id="aa147-455">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="aa147-455">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="aa147-456">1.0</span><span class="sxs-lookup"><span data-stu-id="aa147-456">1.0</span></span>|
|[<span data-ttu-id="aa147-457">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="aa147-457">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="aa147-458">ReadItem</span><span class="sxs-lookup"><span data-stu-id="aa147-458">ReadItem</span></span>|
|[<span data-ttu-id="aa147-459">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="aa147-459">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="aa147-460">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="aa147-460">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="aa147-461">Пример</span><span class="sxs-lookup"><span data-stu-id="aa147-461">Example</span></span>

```js
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

#### <a name="normalizedsubject-string"></a><span data-ttu-id="aa147-462">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="aa147-462">normalizedSubject :String</span></span>

<span data-ttu-id="aa147-p120">Получает тему элемента со всеми удаленными префиксами (включая `RE:` и `FWD:`). Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="aa147-p120">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="aa147-p121">Свойство normalizedSubject получает тему элемента со стандартными префиксами (такими как `RE:` и `FW:`), добавляемыми почтовыми программами. Для получения темы элемента с неизмененными префиксами используйте свойство [`subject`](#subject-stringsubjectjavascriptapioutlook17officesubject).</span><span class="sxs-lookup"><span data-stu-id="aa147-p121">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubjectjavascriptapioutlook17officesubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="aa147-467">Тип:</span><span class="sxs-lookup"><span data-stu-id="aa147-467">Type:</span></span>

*   <span data-ttu-id="aa147-468">String</span><span class="sxs-lookup"><span data-stu-id="aa147-468">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="aa147-469">Требования</span><span class="sxs-lookup"><span data-stu-id="aa147-469">Requirements</span></span>

|<span data-ttu-id="aa147-470">Requirement</span><span class="sxs-lookup"><span data-stu-id="aa147-470">Requirement</span></span>|<span data-ttu-id="aa147-471">Значение</span><span class="sxs-lookup"><span data-stu-id="aa147-471">Value</span></span>|
|---|---|
|[<span data-ttu-id="aa147-472">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="aa147-472">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="aa147-473">1.0</span><span class="sxs-lookup"><span data-stu-id="aa147-473">1.0</span></span>|
|[<span data-ttu-id="aa147-474">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="aa147-474">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="aa147-475">ReadItem</span><span class="sxs-lookup"><span data-stu-id="aa147-475">ReadItem</span></span>|
|[<span data-ttu-id="aa147-476">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="aa147-476">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="aa147-477">Чтение</span><span class="sxs-lookup"><span data-stu-id="aa147-477">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="aa147-478">Пример</span><span class="sxs-lookup"><span data-stu-id="aa147-478">Example</span></span>

```js
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
```

####  <a name="notificationmessages-notificationmessagesjavascriptapioutlook17officenotificationmessages"></a><span data-ttu-id="aa147-479">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_7/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="aa147-479">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_7/office.notificationmessages)</span></span>

<span data-ttu-id="aa147-480">Получает сообщения уведомления для элемента.</span><span class="sxs-lookup"><span data-stu-id="aa147-480">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="aa147-481">Тип:</span><span class="sxs-lookup"><span data-stu-id="aa147-481">Type:</span></span>

*   [<span data-ttu-id="aa147-482">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="aa147-482">NotificationMessages</span></span>](/javascript/api/outlook_1_7/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="aa147-483">Требования</span><span class="sxs-lookup"><span data-stu-id="aa147-483">Requirements</span></span>

|<span data-ttu-id="aa147-484">Requirement</span><span class="sxs-lookup"><span data-stu-id="aa147-484">Requirement</span></span>|<span data-ttu-id="aa147-485">Значение</span><span class="sxs-lookup"><span data-stu-id="aa147-485">Value</span></span>|
|---|---|
|[<span data-ttu-id="aa147-486">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="aa147-486">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="aa147-487">1.3</span><span class="sxs-lookup"><span data-stu-id="aa147-487">1.3</span></span>|
|[<span data-ttu-id="aa147-488">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="aa147-488">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="aa147-489">ReadItem</span><span class="sxs-lookup"><span data-stu-id="aa147-489">ReadItem</span></span>|
|[<span data-ttu-id="aa147-490">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="aa147-490">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="aa147-491">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="aa147-491">Compose or read</span></span>|

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients"></a><span data-ttu-id="aa147-492">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="aa147-492">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

<span data-ttu-id="aa147-493">Предоставляет доступ к необязательным участникам события.</span><span class="sxs-lookup"><span data-stu-id="aa147-493">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="aa147-494">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="aa147-494">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="aa147-495">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="aa147-495">Read mode</span></span>

<span data-ttu-id="aa147-496">Свойство `optionalAttendees` возвращает массив, содержащий объект `EmailAddressDetails` для каждого необязательного участника собрания.</span><span class="sxs-lookup"><span data-stu-id="aa147-496">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="aa147-497">Режим создания</span><span class="sxs-lookup"><span data-stu-id="aa147-497">Compose mode</span></span>

<span data-ttu-id="aa147-498">Свойство `optionalAttendees` возвращает объект `Recipients`, который предоставляет методы для получения или обновления необязательных участников собрания.</span><span class="sxs-lookup"><span data-stu-id="aa147-498">The `optionalAttendees` property returns a `Recipients` object that provides methods to get and set the optional attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="aa147-499">Тип:</span><span class="sxs-lookup"><span data-stu-id="aa147-499">Type:</span></span>

*   <span data-ttu-id="aa147-500">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="aa147-500">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="aa147-501">Требования</span><span class="sxs-lookup"><span data-stu-id="aa147-501">Requirements</span></span>

|<span data-ttu-id="aa147-502">Requirement</span><span class="sxs-lookup"><span data-stu-id="aa147-502">Requirement</span></span>|<span data-ttu-id="aa147-503">Значение</span><span class="sxs-lookup"><span data-stu-id="aa147-503">Value</span></span>|
|---|---|
|[<span data-ttu-id="aa147-504">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="aa147-504">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="aa147-505">1.0</span><span class="sxs-lookup"><span data-stu-id="aa147-505">1.0</span></span>|
|[<span data-ttu-id="aa147-506">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="aa147-506">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="aa147-507">ReadItem</span><span class="sxs-lookup"><span data-stu-id="aa147-507">ReadItem</span></span>|
|[<span data-ttu-id="aa147-508">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="aa147-508">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="aa147-509">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="aa147-509">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="aa147-510">Пример</span><span class="sxs-lookup"><span data-stu-id="aa147-510">Example</span></span>

```js
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsorganizerjavascriptapioutlook17officeorganizer"></a><span data-ttu-id="aa147-511">organizer :[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)|[Organizer](/javascript/api/outlook_1_7/office.organizer)</span><span class="sxs-lookup"><span data-stu-id="aa147-511">organizer :[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)|[Organizer](/javascript/api/outlook_1_7/office.organizer)</span></span>

<span data-ttu-id="aa147-512">Получает адрес электронной почты организатора указанного собрания.</span><span class="sxs-lookup"><span data-stu-id="aa147-512">Gets the email address of the meeting organizer for a specified meeting.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="aa147-513">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="aa147-513">Read mode</span></span>

<span data-ttu-id="aa147-514">Свойство `organizer` возвращает объект [EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails), представляющий организатора собрания.</span><span class="sxs-lookup"><span data-stu-id="aa147-514">The `organizer` property returns an [EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails) object that represents the meeting organizer.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="aa147-515">Режим создания</span><span class="sxs-lookup"><span data-stu-id="aa147-515">Compose mode</span></span>

<span data-ttu-id="aa147-516">Свойство `organizer` возвращает объект [Organizer](/javascript/api/outlook_1_7/office.organizer), который предоставляет метод для получения значения организатора.</span><span class="sxs-lookup"><span data-stu-id="aa147-516">The `organizer` property returns an [Organizer](/javascript/api/outlook_1_7/office.organizer) object that provides a method to get the organizer value.</span></span>

##### <a name="type"></a><span data-ttu-id="aa147-517">Тип:</span><span class="sxs-lookup"><span data-stu-id="aa147-517">Type:</span></span>

*   <span data-ttu-id="aa147-518">[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails) | [Organizer](/javascript/api/outlook_1_7/office.organizer)</span><span class="sxs-lookup"><span data-stu-id="aa147-518">[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails) | [Organizer](/javascript/api/outlook_1_7/office.organizer)</span></span>

##### <a name="requirements"></a><span data-ttu-id="aa147-519">Требования</span><span class="sxs-lookup"><span data-stu-id="aa147-519">Requirements</span></span>

|<span data-ttu-id="aa147-520">Требование</span><span class="sxs-lookup"><span data-stu-id="aa147-520">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="aa147-521">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="aa147-521">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="aa147-522">1.0</span><span class="sxs-lookup"><span data-stu-id="aa147-522">1.0</span></span>|<span data-ttu-id="aa147-523">1.7</span><span class="sxs-lookup"><span data-stu-id="aa147-523">ExcelApi 1.7 Beta</span></span>|
|[<span data-ttu-id="aa147-524">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="aa147-524">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="aa147-525">ReadItem</span><span class="sxs-lookup"><span data-stu-id="aa147-525">ReadItem</span></span>|<span data-ttu-id="aa147-526">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="aa147-526">ReadWriteItem</span></span>|
|[<span data-ttu-id="aa147-527">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="aa147-527">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="aa147-528">Чтение</span><span class="sxs-lookup"><span data-stu-id="aa147-528">Read</span></span>|<span data-ttu-id="aa147-529">Создание</span><span class="sxs-lookup"><span data-stu-id="aa147-529">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="aa147-530">Пример</span><span class="sxs-lookup"><span data-stu-id="aa147-530">Example</span></span>

```js
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
```

#### <a name="nullable-recurrence-recurrencejavascriptapioutlook17officerecurrence"></a><span data-ttu-id="aa147-531">(nullable) recurrence :[Recurrence](/javascript/api/outlook_1_7/office.recurrence)</span><span class="sxs-lookup"><span data-stu-id="aa147-531">(nullable) recurrence :[Recurrence](/javascript/api/outlook_1_7/office.recurrence)</span></span>

<span data-ttu-id="aa147-532">Получает или задает расписание повторения для встречи.</span><span class="sxs-lookup"><span data-stu-id="aa147-532">Gets or sets the location of an appointment.</span></span> <span data-ttu-id="aa147-533">Получает расписание повторения для приглашения на собрание.</span><span class="sxs-lookup"><span data-stu-id="aa147-533">Gets the recurrence pattern of a meeting request.</span></span> <span data-ttu-id="aa147-534">Доступно в режимах чтения и создания для элементов встречи.</span><span class="sxs-lookup"><span data-stu-id="aa147-534">Read and compose modes for appointment items.</span></span> <span data-ttu-id="aa147-535">Доступно в режиме чтения для элементов приглашения на собрание.</span><span class="sxs-lookup"><span data-stu-id="aa147-535">Read mode for meeting request items.</span></span>

<span data-ttu-id="aa147-536">Свойство `recurrence` возвращает объект [recurrence](/javascript/api/outlook_1_7/office.recurrence) для повторяющихся встреч или приглашений на собрание, если элемент представляет собой серию или экземпляр в пределах серии.</span><span class="sxs-lookup"><span data-stu-id="aa147-536">The `recurrence` property returns a [recurrence](/javascript/api/outlook_1_7/office.recurrence) object for recurring appointments or meetings requests if an item is a series or an instance in a series.</span></span> <span data-ttu-id="aa147-537">Значение `null` возвращается для отдельных встреч и приглашений на собрания, связанных с одной встречей.</span><span class="sxs-lookup"><span data-stu-id="aa147-537">`null` is returned for single appointments and meeting requests of single appointments.</span></span> <span data-ttu-id="aa147-538">Значение `undefined` возвращается для сообщений, которые не являются приглашениями на собрания.</span><span class="sxs-lookup"><span data-stu-id="aa147-538">`undefined` is returned for messages that are not meeting requests.</span></span>

> <span data-ttu-id="aa147-539">Примечание. Приглашения на собрания имеют значение `itemClass` для класса IPM.Schedule.Meeting.Request.</span><span class="sxs-lookup"><span data-stu-id="aa147-539">Note: Meeting requests have an `itemClass` value of IPM.Schedule.Meeting.Request.</span></span>

> <span data-ttu-id="aa147-540">Примечание. Если объект recurrence имеет значение `null`, он представляет собой отдельную встречу или приглашение на собрание, связанное с одной встречей, и НЕ входит в серию.</span><span class="sxs-lookup"><span data-stu-id="aa147-540">Note: If the recurrence object is `null`, this indicates that the object is a single appointment or a meeting request of a single appointment and NOT a part of a series.</span></span>

##### <a name="type"></a><span data-ttu-id="aa147-541">Тип:</span><span class="sxs-lookup"><span data-stu-id="aa147-541">Type:</span></span>

* [<span data-ttu-id="aa147-542">Recurrence</span><span class="sxs-lookup"><span data-stu-id="aa147-542">recurrence</span></span>](/javascript/api/outlook_1_7/office.recurrence)

|<span data-ttu-id="aa147-543">Требование</span><span class="sxs-lookup"><span data-stu-id="aa147-543">Requirement</span></span>|<span data-ttu-id="aa147-544">Значение</span><span class="sxs-lookup"><span data-stu-id="aa147-544">Value</span></span>|
|---|---|
|[<span data-ttu-id="aa147-545">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="aa147-545">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="aa147-546">1.7</span><span class="sxs-lookup"><span data-stu-id="aa147-546">ExcelApi 1.7 Beta</span></span>|
|[<span data-ttu-id="aa147-547">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="aa147-547">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="aa147-548">ReadItem</span><span class="sxs-lookup"><span data-stu-id="aa147-548">ReadItem</span></span>|
|[<span data-ttu-id="aa147-549">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="aa147-549">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="aa147-550">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="aa147-550">Compose or read</span></span>|

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients"></a><span data-ttu-id="aa147-551">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="aa147-551">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

<span data-ttu-id="aa147-552">Предоставляет доступ к обязательным участникам события.</span><span class="sxs-lookup"><span data-stu-id="aa147-552">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="aa147-553">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="aa147-553">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="aa147-554">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="aa147-554">Read mode</span></span>

<span data-ttu-id="aa147-555">Свойство `requiredAttendees` возвращает массив, содержащий объект `EmailAddressDetails` для каждого обязательного участника собрания.</span><span class="sxs-lookup"><span data-stu-id="aa147-555">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="aa147-556">Режим создания</span><span class="sxs-lookup"><span data-stu-id="aa147-556">Compose mode</span></span>

<span data-ttu-id="aa147-557">Свойство `requiredAttendees` возвращает объект `Recipients`, предоставляющий методы, с помощью которых можно получить или обновить сведения об обязательных участниках собрания.</span><span class="sxs-lookup"><span data-stu-id="aa147-557">The `requiredAttendees` property returns a `Recipients` object that provides methods to get and set the required attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="aa147-558">Тип:</span><span class="sxs-lookup"><span data-stu-id="aa147-558">Type:</span></span>

*   <span data-ttu-id="aa147-559">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="aa147-559">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="aa147-560">Требования</span><span class="sxs-lookup"><span data-stu-id="aa147-560">Requirements</span></span>

|<span data-ttu-id="aa147-561">Requirement</span><span class="sxs-lookup"><span data-stu-id="aa147-561">Requirement</span></span>|<span data-ttu-id="aa147-562">Значение</span><span class="sxs-lookup"><span data-stu-id="aa147-562">Value</span></span>|
|---|---|
|[<span data-ttu-id="aa147-563">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="aa147-563">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="aa147-564">1.0</span><span class="sxs-lookup"><span data-stu-id="aa147-564">1.0</span></span>|
|[<span data-ttu-id="aa147-565">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="aa147-565">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="aa147-566">ReadItem</span><span class="sxs-lookup"><span data-stu-id="aa147-566">ReadItem</span></span>|
|[<span data-ttu-id="aa147-567">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="aa147-567">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="aa147-568">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="aa147-568">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="aa147-569">Пример</span><span class="sxs-lookup"><span data-stu-id="aa147-569">Example</span></span>

```js
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
}
```

#### <a name="sender-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetails"></a><span data-ttu-id="aa147-570">sender :[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="aa147-570">sender :[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)</span></span>

<span data-ttu-id="aa147-p126">Получает электронный адрес отправителя электронного сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="aa147-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="aa147-p127">Свойства [`from`](#from-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsfromjavascriptapioutlook17officefrom) и `sender` представляют одно лицо, если сообщение не отправлено представителем. В противном случае свойство `from` представляет лицо, делегировавшее полномочия, а свойство sender — представителя.</span><span class="sxs-lookup"><span data-stu-id="aa147-p127">The [`from`](#from-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsfromjavascriptapioutlook17officefrom) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="aa147-575">Свойству `recipientType`, принадлежащему объекту `EmailAddressDetails` в свойстве `sender`, задано значение `undefined`.</span><span class="sxs-lookup"><span data-stu-id="aa147-575">Note: The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="aa147-576">Тип:</span><span class="sxs-lookup"><span data-stu-id="aa147-576">Type:</span></span>

*   [<span data-ttu-id="aa147-577">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="aa147-577">EmailAddressDetails</span></span>](/javascript/api/outlook_1_7/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="aa147-578">Требования</span><span class="sxs-lookup"><span data-stu-id="aa147-578">Requirements</span></span>

|<span data-ttu-id="aa147-579">Requirement</span><span class="sxs-lookup"><span data-stu-id="aa147-579">Requirement</span></span>|<span data-ttu-id="aa147-580">Значение</span><span class="sxs-lookup"><span data-stu-id="aa147-580">Value</span></span>|
|---|---|
|[<span data-ttu-id="aa147-581">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="aa147-581">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="aa147-582">1.0</span><span class="sxs-lookup"><span data-stu-id="aa147-582">1.0</span></span>|
|[<span data-ttu-id="aa147-583">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="aa147-583">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="aa147-584">ReadItem</span><span class="sxs-lookup"><span data-stu-id="aa147-584">ReadItem</span></span>|
|[<span data-ttu-id="aa147-585">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="aa147-585">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="aa147-586">Чтение</span><span class="sxs-lookup"><span data-stu-id="aa147-586">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="aa147-587">Пример</span><span class="sxs-lookup"><span data-stu-id="aa147-587">Example</span></span>

```js
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
```

#### <a name="nullable-seriesid-string"></a><span data-ttu-id="aa147-588">(nullable) seriesId :String</span><span class="sxs-lookup"><span data-stu-id="aa147-588">(nullable) seriesId :String</span></span>

<span data-ttu-id="aa147-589">Получает идентификатор серии, к которой относится экземпляр.</span><span class="sxs-lookup"><span data-stu-id="aa147-589">Gets the id of the series that an instance belongs to.</span></span>

<span data-ttu-id="aa147-590">В Outlook Web App и Outlook свойство `seriesId` возвращает идентификатор веб-служб Exchange (EWS) родительского элемента (серии), к которому относится этот элемент.</span><span class="sxs-lookup"><span data-stu-id="aa147-590">In OWA and Outlook, the `seriesId` returns the Exchange Web Services (EWS) ID of the parent (series) item that this item belongs to.</span></span> <span data-ttu-id="aa147-591">Однако в iOS и Android свойство `seriesId` возвращает идентификатор REST родительского элемента.</span><span class="sxs-lookup"><span data-stu-id="aa147-591">However, in iOS and Android, the `seriesId` returns the REST ID of the parent item.</span></span>

> [!NOTE]
> <span data-ttu-id="aa147-592">Идентификатор, возвращаемый свойством `seriesId`, совпадает с идентификатором элемента веб-служб Exchange.</span><span class="sxs-lookup"><span data-stu-id="aa147-592">The identifier returned by the `seriesId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="aa147-593">Свойство `seriesId` не совпадает с идентификаторами Outlook, которые используются в REST API Outlook.</span><span class="sxs-lookup"><span data-stu-id="aa147-593">The `seriesId` property is not identical to the Outlook IDs used by the Outlook REST API.</span></span> <span data-ttu-id="aa147-594">Прежде чем совершать вызовы REST API, используя это значение, его необходимо преобразовать с помощью [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="aa147-594">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="aa147-595">Дополнительные сведения см. в статье [Использование REST API Outlook из надстройки Outlook](https://docs.microsoft.com/outlook/add-ins/use-rest-api).</span><span class="sxs-lookup"><span data-stu-id="aa147-595">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api).</span></span>

<span data-ttu-id="aa147-596">Свойство `seriesId` возвращает значение `null` для элементов, у которых нет родительских элементов, например отдельных встреч, элементов серий или приглашений на собрания, и возвращает значение `undefined` для всех других элементов, которые не представляют собой приглашения на собрания.</span><span class="sxs-lookup"><span data-stu-id="aa147-596">The `seriesId` property returns `null` for items that do not have parent items such as single appointments, series items, or meeting requests and returns `undefined` for any other items that are not meeting requests.</span></span>

##### <a name="type"></a><span data-ttu-id="aa147-597">Тип:</span><span class="sxs-lookup"><span data-stu-id="aa147-597">Type:</span></span>

* <span data-ttu-id="aa147-598">String</span><span class="sxs-lookup"><span data-stu-id="aa147-598">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="aa147-599">Требования</span><span class="sxs-lookup"><span data-stu-id="aa147-599">Requirements</span></span>

|<span data-ttu-id="aa147-600">Requirement</span><span class="sxs-lookup"><span data-stu-id="aa147-600">Requirement</span></span>|<span data-ttu-id="aa147-601">Значение</span><span class="sxs-lookup"><span data-stu-id="aa147-601">Value</span></span>|
|---|---|
|[<span data-ttu-id="aa147-602">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="aa147-602">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="aa147-603">1.7</span><span class="sxs-lookup"><span data-stu-id="aa147-603">ExcelApi 1.7 Beta</span></span>|
|[<span data-ttu-id="aa147-604">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="aa147-604">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="aa147-605">ReadItem</span><span class="sxs-lookup"><span data-stu-id="aa147-605">ReadItem</span></span>|
|[<span data-ttu-id="aa147-606">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="aa147-606">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="aa147-607">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="aa147-607">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="aa147-608">Пример</span><span class="sxs-lookup"><span data-stu-id="aa147-608">Example</span></span>

```js
var seriesId = Office.context.mailbox.item.seriesId;
var isSeries = (seriesId == null);
```

####  <a name="start-datetimejavascriptapioutlook17officetime"></a><span data-ttu-id="aa147-609">start :Date|[Time](/javascript/api/outlook_1_7/office.time)</span><span class="sxs-lookup"><span data-stu-id="aa147-609">start :Date|[Time](/javascript/api/outlook_1_7/office.time)</span></span>

<span data-ttu-id="aa147-610">Получает или задает дату и время начала встречи.</span><span class="sxs-lookup"><span data-stu-id="aa147-610">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="aa147-p130">Свойство `start` представлено в виде значения даты и времени в формате UTC. Это значение можно преобразовать в местные значения даты и времени клиента с помощью метода [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook17officelocalclienttime).</span><span class="sxs-lookup"><span data-stu-id="aa147-p130">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook17officelocalclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="aa147-613">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="aa147-613">Read mode</span></span>

<span data-ttu-id="aa147-614">Свойство `start` возвращает объект `Date`.</span><span class="sxs-lookup"><span data-stu-id="aa147-614">The `start` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="aa147-615">Режим создания</span><span class="sxs-lookup"><span data-stu-id="aa147-615">Compose mode</span></span>

<span data-ttu-id="aa147-616">Свойство `start` возвращает объект `Time`.</span><span class="sxs-lookup"><span data-stu-id="aa147-616">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="aa147-617">Если вы задаете время начала с помощью метода [`Time.setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-), необходимо использовать метод [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) для преобразования местного времени на клиенте в формат UTC для сервера.</span><span class="sxs-lookup"><span data-stu-id="aa147-617">When you use the [`Time.setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="aa147-618">Тип:</span><span class="sxs-lookup"><span data-stu-id="aa147-618">Type:</span></span>

*   <span data-ttu-id="aa147-619">Date | [Time](/javascript/api/outlook_1_7/office.time)</span><span class="sxs-lookup"><span data-stu-id="aa147-619">Date | [Time](/javascript/api/outlook_1_7/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="aa147-620">Требования</span><span class="sxs-lookup"><span data-stu-id="aa147-620">Requirements</span></span>

|<span data-ttu-id="aa147-621">Requirement</span><span class="sxs-lookup"><span data-stu-id="aa147-621">Requirement</span></span>|<span data-ttu-id="aa147-622">Значение</span><span class="sxs-lookup"><span data-stu-id="aa147-622">Value</span></span>|
|---|---|
|[<span data-ttu-id="aa147-623">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="aa147-623">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="aa147-624">1.0</span><span class="sxs-lookup"><span data-stu-id="aa147-624">1.0</span></span>|
|[<span data-ttu-id="aa147-625">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="aa147-625">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="aa147-626">ReadItem</span><span class="sxs-lookup"><span data-stu-id="aa147-626">ReadItem</span></span>|
|[<span data-ttu-id="aa147-627">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="aa147-627">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="aa147-628">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="aa147-628">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="aa147-629">Пример</span><span class="sxs-lookup"><span data-stu-id="aa147-629">Example</span></span>

<span data-ttu-id="aa147-630">В примере ниже с помощью метода [`setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) объекта `Time` задается время начала встречи в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="aa147-630">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

```js
var startTime = new Date("3/14/2015");
var options = {
  // Pass information that can be used
  // in the callback
  asyncContext: {verb:"Set"}
}
Office.context.mailbox.item.start.setAsync(startTime, options, function(result) {
  if (result.error) {
    console.debug(result.error);
  } else {
    // Access the asyncContext that was passed to the setAsync function
    console.debug("Start Time " + result.asyncContext.verb);
  }
});
```

####  <a name="subject-stringsubjectjavascriptapioutlook17officesubject"></a><span data-ttu-id="aa147-631">subject :String|[Subject](/javascript/api/outlook_1_7/office.subject)</span><span class="sxs-lookup"><span data-stu-id="aa147-631">subject :String|[Subject](/javascript/api/outlook_1_7/office.subject)</span></span>

<span data-ttu-id="aa147-632">Получает или задает описание, которое отображается в поле темы элемента.</span><span class="sxs-lookup"><span data-stu-id="aa147-632">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="aa147-633">Свойство `subject` получает или задает всю тему элемента для отправки с почтового сервера.</span><span class="sxs-lookup"><span data-stu-id="aa147-633">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="aa147-634">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="aa147-634">Read mode</span></span>

<span data-ttu-id="aa147-p131">Свойство `subject` возвращает строку. С помощью свойства [`normalizedSubject`](#normalizedsubject-string) можно получить тему без начальных префиксов, таких как `RE:` и `FW:`.</span><span class="sxs-lookup"><span data-stu-id="aa147-p131">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```js
var subject = Office.context.mailbox.item.subject;
```

##### <a name="compose-mode"></a><span data-ttu-id="aa147-637">Режим создания</span><span class="sxs-lookup"><span data-stu-id="aa147-637">Compose mode</span></span>

<span data-ttu-id="aa147-638">Свойство `subject` возвращает объект `Subject`, который предоставляет методы для получения и задания темы.</span><span class="sxs-lookup"><span data-stu-id="aa147-638">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```js
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="aa147-639">Тип:</span><span class="sxs-lookup"><span data-stu-id="aa147-639">Type:</span></span>

*   <span data-ttu-id="aa147-640">String | [Subject](/javascript/api/outlook_1_7/office.subject)</span><span class="sxs-lookup"><span data-stu-id="aa147-640">String | [Subject](/javascript/api/outlook_1_7/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="aa147-641">Требования</span><span class="sxs-lookup"><span data-stu-id="aa147-641">Requirements</span></span>

|<span data-ttu-id="aa147-642">Requirement</span><span class="sxs-lookup"><span data-stu-id="aa147-642">Requirement</span></span>|<span data-ttu-id="aa147-643">Значение</span><span class="sxs-lookup"><span data-stu-id="aa147-643">Value</span></span>|
|---|---|
|[<span data-ttu-id="aa147-644">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="aa147-644">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="aa147-645">1.0</span><span class="sxs-lookup"><span data-stu-id="aa147-645">1.0</span></span>|
|[<span data-ttu-id="aa147-646">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="aa147-646">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="aa147-647">ReadItem</span><span class="sxs-lookup"><span data-stu-id="aa147-647">ReadItem</span></span>|
|[<span data-ttu-id="aa147-648">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="aa147-648">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="aa147-649">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="aa147-649">Compose or read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients"></a><span data-ttu-id="aa147-650">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="aa147-650">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

<span data-ttu-id="aa147-651">Предоставляет доступ к получателям, указанным в строке **Кому** сообщения.</span><span class="sxs-lookup"><span data-stu-id="aa147-651">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="aa147-652">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="aa147-652">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="aa147-653">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="aa147-653">Read mode</span></span>

<span data-ttu-id="aa147-p133">Свойство `to` возвращает массив, содержащий объект `EmailAddressDetails` для каждого получателя в строке **Кому** сообщения. Коллекция может включать не более 100 элементов.</span><span class="sxs-lookup"><span data-stu-id="aa147-p133">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="aa147-656">Режим создания</span><span class="sxs-lookup"><span data-stu-id="aa147-656">Compose mode</span></span>

<span data-ttu-id="aa147-657">Свойство `to` возвращает объект `Recipients`, предоставляющий методы для получения или обновления получателей, которые указаны в строке **Кому** сообщения.</span><span class="sxs-lookup"><span data-stu-id="aa147-657">The to`to` property returns a Recipients`Recipients` object that provides methods for manipulating the recipients on the **To** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="aa147-658">Тип:</span><span class="sxs-lookup"><span data-stu-id="aa147-658">Type:</span></span>

*   <span data-ttu-id="aa147-659">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="aa147-659">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="aa147-660">Требования</span><span class="sxs-lookup"><span data-stu-id="aa147-660">Requirements</span></span>

|<span data-ttu-id="aa147-661">Requirement</span><span class="sxs-lookup"><span data-stu-id="aa147-661">Requirement</span></span>|<span data-ttu-id="aa147-662">Значение</span><span class="sxs-lookup"><span data-stu-id="aa147-662">Value</span></span>|
|---|---|
|[<span data-ttu-id="aa147-663">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="aa147-663">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="aa147-664">1.0</span><span class="sxs-lookup"><span data-stu-id="aa147-664">1.0</span></span>|
|[<span data-ttu-id="aa147-665">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="aa147-665">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="aa147-666">ReadItem</span><span class="sxs-lookup"><span data-stu-id="aa147-666">ReadItem</span></span>|
|[<span data-ttu-id="aa147-667">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="aa147-667">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="aa147-668">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="aa147-668">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="aa147-669">Пример</span><span class="sxs-lookup"><span data-stu-id="aa147-669">Example</span></span>

```js
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

### <a name="methods"></a><span data-ttu-id="aa147-670">Методы</span><span class="sxs-lookup"><span data-stu-id="aa147-670">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="aa147-671">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="aa147-671">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="aa147-672">Добавляет файл в сообщение или встречу в качестве вложения.</span><span class="sxs-lookup"><span data-stu-id="aa147-672">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="aa147-673">Метод `addFileAttachmentAsync` передает файл по указанному универсальному коду ресурса (URI) и вкладывает его в элемент в форме создания.</span><span class="sxs-lookup"><span data-stu-id="aa147-673">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="aa147-674">Идентификатор можно последовательно использовать с методом [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback), чтобы удалить вложение, добавленное во время текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="aa147-674">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="aa147-675">Параметры</span><span class="sxs-lookup"><span data-stu-id="aa147-675">Parameters:</span></span>
|<span data-ttu-id="aa147-676">Имя</span><span class="sxs-lookup"><span data-stu-id="aa147-676">Name</span></span>|<span data-ttu-id="aa147-677">Тип</span><span class="sxs-lookup"><span data-stu-id="aa147-677">Type</span></span>|<span data-ttu-id="aa147-678">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="aa147-678">Attributes</span></span>|<span data-ttu-id="aa147-679">Описание</span><span class="sxs-lookup"><span data-stu-id="aa147-679">Description</span></span>|
|---|---|---|---|
|`uri`|<span data-ttu-id="aa147-680">String</span><span class="sxs-lookup"><span data-stu-id="aa147-680">String</span></span>||<span data-ttu-id="aa147-p134">Универсальный код ресурса (URI), представляющий расположение файла, который нужно вложить в сообщение или встречу. Максимальная длина — 2048 символов.</span><span class="sxs-lookup"><span data-stu-id="aa147-p134">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="aa147-683">String</span><span class="sxs-lookup"><span data-stu-id="aa147-683">String</span></span>||<span data-ttu-id="aa147-p135">Имя вложения, которое отображается при передаче вложения. Максимальная длина — 255 символов.</span><span class="sxs-lookup"><span data-stu-id="aa147-p135">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="aa147-686">Object</span><span class="sxs-lookup"><span data-stu-id="aa147-686">Object</span></span>|<span data-ttu-id="aa147-687">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="aa147-687">&lt;optional&gt;</span></span>|<span data-ttu-id="aa147-688">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="aa147-688">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="aa147-689">Object</span><span class="sxs-lookup"><span data-stu-id="aa147-689">Object</span></span>|<span data-ttu-id="aa147-690">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="aa147-690">&lt;optional&gt;</span></span>|<span data-ttu-id="aa147-691">В методе обратного вызова разработчики могут указать любой объект, к которому необходимо получить доступ.</span><span class="sxs-lookup"><span data-stu-id="aa147-691">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="aa147-692">Boolean</span><span class="sxs-lookup"><span data-stu-id="aa147-692">Boolean</span></span>|<span data-ttu-id="aa147-693">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="aa147-693">&lt;optional&gt;</span></span>|<span data-ttu-id="aa147-694">Значение `true` указывает на то, что вложение будет встроено в текст сообщения и не должно отображаться в списке вложений.</span><span class="sxs-lookup"><span data-stu-id="aa147-694">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="aa147-695">function</span><span class="sxs-lookup"><span data-stu-id="aa147-695">function</span></span>|<span data-ttu-id="aa147-696">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="aa147-696">&lt;optional&gt;</span></span>|<span data-ttu-id="aa147-697">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="aa147-697">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="aa147-698">После успешного выполнения идентификатор вложения будет представлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="aa147-698">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="aa147-699">Если добавить вложение не удастся, объект `asyncResult` будет содержать объект `Error` с описанием ошибки.</span><span class="sxs-lookup"><span data-stu-id="aa147-699">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="aa147-700">Ошибки</span><span class="sxs-lookup"><span data-stu-id="aa147-700">Errors</span></span>

|<span data-ttu-id="aa147-701">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="aa147-701">Error code</span></span>|<span data-ttu-id="aa147-702">Описание</span><span class="sxs-lookup"><span data-stu-id="aa147-702">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="aa147-703">Вложение превышает максимальный размер.</span><span class="sxs-lookup"><span data-stu-id="aa147-703">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="aa147-704">Расширение вложения не поддерживается.</span><span class="sxs-lookup"><span data-stu-id="aa147-704">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="aa147-705">Сообщение или встреча содержат слишком много вложений.</span><span class="sxs-lookup"><span data-stu-id="aa147-705">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="aa147-706">Требования</span><span class="sxs-lookup"><span data-stu-id="aa147-706">Requirements</span></span>

|<span data-ttu-id="aa147-707">Requirement</span><span class="sxs-lookup"><span data-stu-id="aa147-707">Requirement</span></span>|<span data-ttu-id="aa147-708">Значение</span><span class="sxs-lookup"><span data-stu-id="aa147-708">Value</span></span>|
|---|---|
|[<span data-ttu-id="aa147-709">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="aa147-709">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="aa147-710">1.1</span><span class="sxs-lookup"><span data-stu-id="aa147-710">1.1</span></span>|
|[<span data-ttu-id="aa147-711">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="aa147-711">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="aa147-712">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="aa147-712">ReadWriteItem</span></span>|
|[<span data-ttu-id="aa147-713">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="aa147-713">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="aa147-714">Создание</span><span class="sxs-lookup"><span data-stu-id="aa147-714">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="aa147-715">Примеры</span><span class="sxs-lookup"><span data-stu-id="aa147-715">Examples</span></span>

```js
function callback(result) {
  if (result.error) {
    showMessage(result.error);
  } else {
    showMessage("Attachment added");
  }
}

function addAttachment() {
  // The values in asyncContext can be accessed in the callback
  var options = { 'asyncContext': { var1: 1, var2: 2 } };

  var attachmentURL = "https://contoso.com/rtm/icon.png";
  Office.context.mailbox.item.addFileAttachmentAsync(attachmentURL, attachmentURL, options, callback);
}
```

<span data-ttu-id="aa147-716">В примере ниже файл изображения добавляется в качестве встроенного вложения, а ссылка на вложение добавляется в текст сообщения.</span><span class="sxs-lookup"><span data-stu-id="aa147-716">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

```js
Office.context.mailbox.item.addFileAttachmentAsync
(
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
        
      }
    );
  }
);
```

####  <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="aa147-717">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="aa147-717">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="aa147-718">Добавляет обработчик для поддерживаемого события.</span><span class="sxs-lookup"><span data-stu-id="aa147-718">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="aa147-719">Сейчас поддерживаются следующие типы событий: `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged` и `Office.EventType.RecurrenceChanged`.</span><span class="sxs-lookup"><span data-stu-id="aa147-719">Currently the supported event types are `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`</span></span>

##### <a name="parameters"></a><span data-ttu-id="aa147-720">Параметры</span><span class="sxs-lookup"><span data-stu-id="aa147-720">Parameters:</span></span>

| <span data-ttu-id="aa147-721">Имя</span><span class="sxs-lookup"><span data-stu-id="aa147-721">Name</span></span> | <span data-ttu-id="aa147-722">Тип</span><span class="sxs-lookup"><span data-stu-id="aa147-722">Type</span></span> | <span data-ttu-id="aa147-723">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="aa147-723">Attributes</span></span> | <span data-ttu-id="aa147-724">Описание</span><span class="sxs-lookup"><span data-stu-id="aa147-724">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="aa147-725">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="aa147-725">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="aa147-726">Событие, которое должно вызвать обработчик.</span><span class="sxs-lookup"><span data-stu-id="aa147-726">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="aa147-727">Function</span><span class="sxs-lookup"><span data-stu-id="aa147-727">Function</span></span> || <span data-ttu-id="aa147-p136">Функция для обработки события. Функция должна принимать один параметр, представляющий собой объектный литерал. Значение свойства `type` параметра совпадет со значением параметра `eventType`, переданного методу `addHandlerAsync`.</span><span class="sxs-lookup"><span data-stu-id="aa147-p136">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="aa147-731">Объект</span><span class="sxs-lookup"><span data-stu-id="aa147-731">Object</span></span> | <span data-ttu-id="aa147-732">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="aa147-732">&lt;optional&gt;</span></span> | <span data-ttu-id="aa147-733">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="aa147-733">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="aa147-734">Object</span><span class="sxs-lookup"><span data-stu-id="aa147-734">Object</span></span> | <span data-ttu-id="aa147-735">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="aa147-735">&lt;optional&gt;</span></span> | <span data-ttu-id="aa147-736">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="aa147-736">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="aa147-737">функция</span><span class="sxs-lookup"><span data-stu-id="aa147-737">function</span></span>| <span data-ttu-id="aa147-738">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="aa147-738">&lt;optional&gt;</span></span>|<span data-ttu-id="aa147-739">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="aa147-739">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="aa147-740">Требования</span><span class="sxs-lookup"><span data-stu-id="aa147-740">Requirements</span></span>

|<span data-ttu-id="aa147-741">Requirement</span><span class="sxs-lookup"><span data-stu-id="aa147-741">Requirement</span></span>| <span data-ttu-id="aa147-742">Значение</span><span class="sxs-lookup"><span data-stu-id="aa147-742">Value</span></span>|
|---|---|
|[<span data-ttu-id="aa147-743">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="aa147-743">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="aa147-744">1.7</span><span class="sxs-lookup"><span data-stu-id="aa147-744">ExcelApi 1.7 Beta</span></span> |
|[<span data-ttu-id="aa147-745">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="aa147-745">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="aa147-746">ReadItem</span><span class="sxs-lookup"><span data-stu-id="aa147-746">ReadItem</span></span> |
|[<span data-ttu-id="aa147-747">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="aa147-747">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="aa147-748">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="aa147-748">Compose or read</span></span> |

##### <a name="example"></a><span data-ttu-id="aa147-749">Пример</span><span class="sxs-lookup"><span data-stu-id="aa147-749">Example</span></span>

```js
Office.initialize = function (reason) {
  $(document).ready(function () {
    Office.context.mailbox.item.addHandlerAsync(Office.EventType.RecurrenceChanged, loadNewItem, function (result) {
      if (result.status === Office.AsyncResultStatus.Failed) {
        // Handle error
      }
    });
  });
};

function loadNewItem(eventArgs) {
  // Load the properties of the newly selected item
  loadProps(Office.context.mailbox.item);
};
```

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="aa147-750">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="aa147-750">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="aa147-751">Добавляет к сообщению элемент Exchange, например сообщение, в виде вложения.</span><span class="sxs-lookup"><span data-stu-id="aa147-751">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="aa147-p137">С помощью метода `addItemAttachmentAsync` можно в элемент формы создания вложить элемент с указанным идентификатором Exchange. Если указать метод обратного вызова, то этот метод вызывается с помощью параметра `asyncResult`, который содержит идентификатор вложения или код, указывающий на ошибки, которые произошли при вложении элемента. При необходимости можно использовать параметр `options` для передачи сведений о состоянии методу обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="aa147-p137">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="aa147-755">Идентификатор можно последовательно использовать с методом [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback), чтобы удалить вложение, добавленное во время текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="aa147-755">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="aa147-756">Если ваша надстройка Office выполняется в Outlook Web App, метод `addItemAttachmentAsync` обеспечивает вложение элементов в элементы, отличные от редактируемого. Однако мы не рекомендуем выполнять это действие, так как оно не поддерживается.</span><span class="sxs-lookup"><span data-stu-id="aa147-756">If your Office add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="aa147-757">Параметры</span><span class="sxs-lookup"><span data-stu-id="aa147-757">Parameters:</span></span>

|<span data-ttu-id="aa147-758">Имя</span><span class="sxs-lookup"><span data-stu-id="aa147-758">Name</span></span>|<span data-ttu-id="aa147-759">Тип</span><span class="sxs-lookup"><span data-stu-id="aa147-759">Type</span></span>|<span data-ttu-id="aa147-760">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="aa147-760">Attributes</span></span>|<span data-ttu-id="aa147-761">Описание</span><span class="sxs-lookup"><span data-stu-id="aa147-761">Description</span></span>|
|---|---|---|---|
|`itemId`|<span data-ttu-id="aa147-762">String</span><span class="sxs-lookup"><span data-stu-id="aa147-762">String</span></span>||<span data-ttu-id="aa147-p138">Идентификатор Exchange для вкладываемого элемента. Максимальная длина — 100 символов.</span><span class="sxs-lookup"><span data-stu-id="aa147-p138">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="aa147-765">String</span><span class="sxs-lookup"><span data-stu-id="aa147-765">String</span></span>||<span data-ttu-id="aa147-p139">Тема вкладываемого элемента. Максимальная длина — 255 символов.</span><span class="sxs-lookup"><span data-stu-id="aa147-p139">The sujbect of the item to be attached. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="aa147-768">Object</span><span class="sxs-lookup"><span data-stu-id="aa147-768">Object</span></span>|<span data-ttu-id="aa147-769">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="aa147-769">&lt;optional&gt;</span></span>|<span data-ttu-id="aa147-770">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="aa147-770">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="aa147-771">Object</span><span class="sxs-lookup"><span data-stu-id="aa147-771">Object</span></span>|<span data-ttu-id="aa147-772">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="aa147-772">&lt;optional&gt;</span></span>|<span data-ttu-id="aa147-773">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="aa147-773">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="aa147-774">функция</span><span class="sxs-lookup"><span data-stu-id="aa147-774">function</span></span>|<span data-ttu-id="aa147-775">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="aa147-775">&lt;optional&gt;</span></span>|<span data-ttu-id="aa147-776">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="aa147-776">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="aa147-777">После успешного выполнения идентификатор вложения будет представлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="aa147-777">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="aa147-778">Если добавить вложение не удастся, объект `asyncResult` будет содержать объект `Error` с описанием ошибки.</span><span class="sxs-lookup"><span data-stu-id="aa147-778">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="aa147-779">Ошибки</span><span class="sxs-lookup"><span data-stu-id="aa147-779">Errors</span></span>

|<span data-ttu-id="aa147-780">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="aa147-780">Error code</span></span>|<span data-ttu-id="aa147-781">Описание</span><span class="sxs-lookup"><span data-stu-id="aa147-781">Description</span></span>|
|------------|-------------|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="aa147-782">Сообщение или встреча содержат слишком много вложений.</span><span class="sxs-lookup"><span data-stu-id="aa147-782">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="aa147-783">Требования</span><span class="sxs-lookup"><span data-stu-id="aa147-783">Requirements</span></span>

|<span data-ttu-id="aa147-784">Requirement</span><span class="sxs-lookup"><span data-stu-id="aa147-784">Requirement</span></span>|<span data-ttu-id="aa147-785">Значение</span><span class="sxs-lookup"><span data-stu-id="aa147-785">Value</span></span>|
|---|---|
|[<span data-ttu-id="aa147-786">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="aa147-786">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="aa147-787">1.1</span><span class="sxs-lookup"><span data-stu-id="aa147-787">1.1</span></span>|
|[<span data-ttu-id="aa147-788">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="aa147-788">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="aa147-789">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="aa147-789">ReadWriteItem</span></span>|
|[<span data-ttu-id="aa147-790">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="aa147-790">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="aa147-791">Создание</span><span class="sxs-lookup"><span data-stu-id="aa147-791">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="aa147-792">Пример</span><span class="sxs-lookup"><span data-stu-id="aa147-792">Example</span></span>

<span data-ttu-id="aa147-793">В следующем примере существующий элемент Outlook добавляется в виде вложения с именем `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="aa147-793">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

```js
function callback(result) {
  if (result.error) {
    showMessage(result.error);
  } else {
    showMessage("Attachment added");
  }
}

function addAttachment() {
  // EWS ID of item to attach
  // (Shortened for readability)
  var itemId = "AAMkADI1...AAA=";

  // The values in asyncContext can be accessed in the callback
  var options = { 'asyncContext': { var1: 1, var2: 2 } };

  Office.context.mailbox.item.addItemAttachmentAsync(itemId, "My Attachment", options, callback);
}
```

####  <a name="close"></a><span data-ttu-id="aa147-794">close()</span><span class="sxs-lookup"><span data-stu-id="aa147-794">close()</span></span>

<span data-ttu-id="aa147-795">Закрывает текущий создаваемый элемент.</span><span class="sxs-lookup"><span data-stu-id="aa147-795">Closes the current item that is being composed.</span></span>

<span data-ttu-id="aa147-p140">Работа метода `close` зависит от текущего состояния создаваемого элемента. Если элемент содержит несохраненные изменения, клиент предложит пользователю сохранить или отклонить их либо отменить действие закрытия.</span><span class="sxs-lookup"><span data-stu-id="aa147-p140">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="aa147-798">Если в Outlook в Интернете элемент представляет собой встречу, ранее сохраненную с помощью метода `saveAsync`, пользователю предлагается сохранить, отклонить или отменить действие, даже если с момента последнего сохранения элемента не вносилось никаких изменений.</span><span class="sxs-lookup"><span data-stu-id="aa147-798">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="aa147-799">Если в клиенте Outlook для настольных ПК сообщение представляет собой ответ в тексте, метод `close` не приносит результатов.</span><span class="sxs-lookup"><span data-stu-id="aa147-799">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="aa147-800">Требования</span><span class="sxs-lookup"><span data-stu-id="aa147-800">Requirements</span></span>

|<span data-ttu-id="aa147-801">Requirement</span><span class="sxs-lookup"><span data-stu-id="aa147-801">Requirement</span></span>|<span data-ttu-id="aa147-802">Значение</span><span class="sxs-lookup"><span data-stu-id="aa147-802">Value</span></span>|
|---|---|
|[<span data-ttu-id="aa147-803">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="aa147-803">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="aa147-804">1.3</span><span class="sxs-lookup"><span data-stu-id="aa147-804">1.3</span></span>|
|[<span data-ttu-id="aa147-805">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="aa147-805">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="aa147-806">Restricted</span><span class="sxs-lookup"><span data-stu-id="aa147-806">Restricted</span></span>|
|[<span data-ttu-id="aa147-807">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="aa147-807">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="aa147-808">Создание</span><span class="sxs-lookup"><span data-stu-id="aa147-808">Compose</span></span>|

#### <a name="displayreplyallformformdata"></a><span data-ttu-id="aa147-809">displayReplyAllForm(formData)</span><span class="sxs-lookup"><span data-stu-id="aa147-809">displayReplyAllForm(formData)</span></span>

<span data-ttu-id="aa147-810">Отображает форму ответа, включающую отправителя и всех получателей выбранного сообщения или организатора и всех участников выбранной встречи.</span><span class="sxs-lookup"><span data-stu-id="aa147-810">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="aa147-811">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="aa147-811">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="aa147-812">В Outlook Web App форма ответа отображается в виде всплывающей формы в представлении с 3 либо 1 или 2 колонками.</span><span class="sxs-lookup"><span data-stu-id="aa147-812">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="aa147-813">Если любой строковый параметр превышает указанные для него ограничения, `displayReplyAllForm` возвращает исключение.</span><span class="sxs-lookup"><span data-stu-id="aa147-813">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="aa147-p141">Если в параметре `formData.attachments` указаны вложения, Outlook и Outlook Web App пытаются скачать их и вложить в форму ответа. Если какие-либо вложения добавить не удается, в форме отображается сообщение об ошибке. Если сообщения об ошибках не предусмотрены, то они не отображаются.</span><span class="sxs-lookup"><span data-stu-id="aa147-p141">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="aa147-817">Параметры</span><span class="sxs-lookup"><span data-stu-id="aa147-817">Parameters:</span></span>

|<span data-ttu-id="aa147-818">Имя</span><span class="sxs-lookup"><span data-stu-id="aa147-818">Name</span></span>|<span data-ttu-id="aa147-819">Тип</span><span class="sxs-lookup"><span data-stu-id="aa147-819">Type</span></span>|<span data-ttu-id="aa147-820">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="aa147-820">Attributes</span></span>|<span data-ttu-id="aa147-821">Описание</span><span class="sxs-lookup"><span data-stu-id="aa147-821">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="aa147-822">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="aa147-822">String &#124; Object</span></span>||<span data-ttu-id="aa147-p142">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="aa147-p142">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="aa147-825">**ИЛИ**</span><span class="sxs-lookup"><span data-stu-id="aa147-825">**OR**</span></span><br/><span data-ttu-id="aa147-p143">Объект, который содержит текст или данные вложения и функцию обратного вызова. Ниже представлено определение этого объекта.</span><span class="sxs-lookup"><span data-stu-id="aa147-p143">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="aa147-828">String</span><span class="sxs-lookup"><span data-stu-id="aa147-828">String</span></span>|<span data-ttu-id="aa147-829">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="aa147-829">&lt;optional&gt;</span></span>|<span data-ttu-id="aa147-p144">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="aa147-p144">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="aa147-832">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="aa147-832">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="aa147-833">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="aa147-833">&lt;optional&gt;</span></span>|<span data-ttu-id="aa147-834">Массив объектов JSON, представляющих собой вложенные файлы или элементы.</span><span class="sxs-lookup"><span data-stu-id="aa147-834">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="aa147-835">String</span><span class="sxs-lookup"><span data-stu-id="aa147-835">String</span></span>||<span data-ttu-id="aa147-p145">Указывает тип вложения. Допустимые значения: `file` для вложенного файла и `item` для вложенного элемента.</span><span class="sxs-lookup"><span data-stu-id="aa147-p145">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="aa147-838">String</span><span class="sxs-lookup"><span data-stu-id="aa147-838">String</span></span>||<span data-ttu-id="aa147-839">Строка, содержащая имя вложения, длиной до 255 символов.</span><span class="sxs-lookup"><span data-stu-id="aa147-839">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="aa147-840">String</span><span class="sxs-lookup"><span data-stu-id="aa147-840">String</span></span>||<span data-ttu-id="aa147-p146">Используется, только если свойству `type` задано значение `file`. URI расположения файла.</span><span class="sxs-lookup"><span data-stu-id="aa147-p146">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="aa147-843">Логический</span><span class="sxs-lookup"><span data-stu-id="aa147-843">Boolean</span></span>||<span data-ttu-id="aa147-p147">Используется, только если свойству `type` задано значение `file`. Значение `true` указывает на то, что вложение будет встроено в текст сообщения и не должно отображаться в списке вложений.</span><span class="sxs-lookup"><span data-stu-id="aa147-p147">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="aa147-846">String</span><span class="sxs-lookup"><span data-stu-id="aa147-846">String</span></span>||<span data-ttu-id="aa147-p148">Используется, только если свойству `type` задано значение `item`. Идентификатор вложения EWS. Это строка длиной до 100 символов.</span><span class="sxs-lookup"><span data-stu-id="aa147-p148">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="aa147-850">function</span><span class="sxs-lookup"><span data-stu-id="aa147-850">function</span></span>|<span data-ttu-id="aa147-851">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="aa147-851">&lt;optional&gt;</span></span>|<span data-ttu-id="aa147-852">По завершении работы метода функция, переданная параметру `callback`, вызывается с помощью одного параметра `asyncResult`, представляющего собой объект [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="aa147-852">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="aa147-853">Требования</span><span class="sxs-lookup"><span data-stu-id="aa147-853">Requirements</span></span>

|<span data-ttu-id="aa147-854">Requirement</span><span class="sxs-lookup"><span data-stu-id="aa147-854">Requirement</span></span>|<span data-ttu-id="aa147-855">Значение</span><span class="sxs-lookup"><span data-stu-id="aa147-855">Value</span></span>|
|---|---|
|[<span data-ttu-id="aa147-856">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="aa147-856">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="aa147-857">1.0</span><span class="sxs-lookup"><span data-stu-id="aa147-857">1.0</span></span>|
|[<span data-ttu-id="aa147-858">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="aa147-858">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="aa147-859">ReadItem</span><span class="sxs-lookup"><span data-stu-id="aa147-859">ReadItem</span></span>|
|[<span data-ttu-id="aa147-860">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="aa147-860">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="aa147-861">Чтение</span><span class="sxs-lookup"><span data-stu-id="aa147-861">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="aa147-862">Примеры</span><span class="sxs-lookup"><span data-stu-id="aa147-862">Examples</span></span>

<span data-ttu-id="aa147-863">Приведенный ниже код передает строку в функцию `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="aa147-863">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="aa147-864">Ответ с пустым текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="aa147-864">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="aa147-865">Ответ только с текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="aa147-865">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="aa147-866">Ответ с текстом сообщения и вложенным файлом.</span><span class="sxs-lookup"><span data-stu-id="aa147-866">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="aa147-867">Ответ с текстом сообщения и вложенным элементом.</span><span class="sxs-lookup"><span data-stu-id="aa147-867">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="aa147-868">Ответ с текстом сообщения, вложенным файлом, вложенным элементом и обратным вызовом.</span><span class="sxs-lookup"><span data-stu-id="aa147-868">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata"></a><span data-ttu-id="aa147-869">displayReplyForm(formData)</span><span class="sxs-lookup"><span data-stu-id="aa147-869">displayReplyForm(formData)</span></span>

<span data-ttu-id="aa147-870">Отображает форму ответа, включающую только отправителя выбранного сообщения или организатора выбранной встречи.</span><span class="sxs-lookup"><span data-stu-id="aa147-870">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="aa147-871">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="aa147-871">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="aa147-872">В Outlook Web App форма ответа отображается в виде всплывающей формы в представлении с 3 либо 1 или 2 колонками.</span><span class="sxs-lookup"><span data-stu-id="aa147-872">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="aa147-873">Если любой строковый параметр превышает указанные для него ограничения, `displayReplyForm` возвращает исключение.</span><span class="sxs-lookup"><span data-stu-id="aa147-873">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="aa147-p149">Если в параметре `formData.attachments` указаны вложения, Outlook и Outlook Web App пытаются скачать их и вложить в форму ответа. Если какие-либо вложения добавить не удается, в форме отображается сообщение об ошибке. Если сообщения об ошибках не предусмотрены, то они не отображаются.</span><span class="sxs-lookup"><span data-stu-id="aa147-p149">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="aa147-877">Параметры</span><span class="sxs-lookup"><span data-stu-id="aa147-877">Parameters:</span></span>

|<span data-ttu-id="aa147-878">Имя</span><span class="sxs-lookup"><span data-stu-id="aa147-878">Name</span></span>|<span data-ttu-id="aa147-879">Тип</span><span class="sxs-lookup"><span data-stu-id="aa147-879">Type</span></span>|<span data-ttu-id="aa147-880">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="aa147-880">Attributes</span></span>|<span data-ttu-id="aa147-881">Описание</span><span class="sxs-lookup"><span data-stu-id="aa147-881">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="aa147-882">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="aa147-882">String &#124; Object</span></span>||<span data-ttu-id="aa147-p150">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="aa147-p150">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="aa147-885">**ИЛИ**</span><span class="sxs-lookup"><span data-stu-id="aa147-885">**OR**</span></span><br/><span data-ttu-id="aa147-p151">Объект, который содержит текст или данные вложения и функцию обратного вызова. Ниже представлено определение этого объекта.</span><span class="sxs-lookup"><span data-stu-id="aa147-p151">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="aa147-888">String</span><span class="sxs-lookup"><span data-stu-id="aa147-888">String</span></span>|<span data-ttu-id="aa147-889">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="aa147-889">&lt;optional&gt;</span></span>|<span data-ttu-id="aa147-p152">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="aa147-p152">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="aa147-892">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="aa147-892">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="aa147-893">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="aa147-893">&lt;optional&gt;</span></span>|<span data-ttu-id="aa147-894">Массив объектов JSON, представляющих собой вложенные файлы или элементы.</span><span class="sxs-lookup"><span data-stu-id="aa147-894">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="aa147-895">String</span><span class="sxs-lookup"><span data-stu-id="aa147-895">String</span></span>||<span data-ttu-id="aa147-p153">Указывает тип вложения. Допустимые значения: `file` для вложенного файла и `item` для вложенного элемента.</span><span class="sxs-lookup"><span data-stu-id="aa147-p153">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="aa147-898">String</span><span class="sxs-lookup"><span data-stu-id="aa147-898">String</span></span>||<span data-ttu-id="aa147-899">Строка, содержащая имя вложения, длиной до 255 символов.</span><span class="sxs-lookup"><span data-stu-id="aa147-899">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="aa147-900">String</span><span class="sxs-lookup"><span data-stu-id="aa147-900">String</span></span>||<span data-ttu-id="aa147-p154">Используется, только если свойству `type` задано значение `file`. URI расположения файла.</span><span class="sxs-lookup"><span data-stu-id="aa147-p154">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="aa147-903">Логический</span><span class="sxs-lookup"><span data-stu-id="aa147-903">Boolean</span></span>||<span data-ttu-id="aa147-p155">Используется, только если свойству `type` задано значение `file`. Значение `true` указывает на то, что вложение будет встроено в текст сообщения и не должно отображаться в списке вложений.</span><span class="sxs-lookup"><span data-stu-id="aa147-p155">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="aa147-906">String</span><span class="sxs-lookup"><span data-stu-id="aa147-906">String</span></span>||<span data-ttu-id="aa147-p156">Используется, только если свойству `type` задано значение `item`. Идентификатор вложения EWS. Это строка длиной до 100 символов.</span><span class="sxs-lookup"><span data-stu-id="aa147-p156">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="aa147-910">function</span><span class="sxs-lookup"><span data-stu-id="aa147-910">function</span></span>|<span data-ttu-id="aa147-911">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="aa147-911">&lt;optional&gt;</span></span>|<span data-ttu-id="aa147-912">По завершении работы метода функция, переданная параметру `callback`, вызывается с помощью одного параметра `asyncResult`, представляющего собой объект [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="aa147-912">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="aa147-913">Требования</span><span class="sxs-lookup"><span data-stu-id="aa147-913">Requirements</span></span>

|<span data-ttu-id="aa147-914">Requirement</span><span class="sxs-lookup"><span data-stu-id="aa147-914">Requirement</span></span>|<span data-ttu-id="aa147-915">Значение</span><span class="sxs-lookup"><span data-stu-id="aa147-915">Value</span></span>|
|---|---|
|[<span data-ttu-id="aa147-916">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="aa147-916">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="aa147-917">1.0</span><span class="sxs-lookup"><span data-stu-id="aa147-917">1.0</span></span>|
|[<span data-ttu-id="aa147-918">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="aa147-918">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="aa147-919">ReadItem</span><span class="sxs-lookup"><span data-stu-id="aa147-919">ReadItem</span></span>|
|[<span data-ttu-id="aa147-920">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="aa147-920">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="aa147-921">Чтение</span><span class="sxs-lookup"><span data-stu-id="aa147-921">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="aa147-922">Примеры</span><span class="sxs-lookup"><span data-stu-id="aa147-922">Examples</span></span>

<span data-ttu-id="aa147-923">Приведенный ниже код передает строку в функцию `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="aa147-923">The following code passes a string to the `displayReplyForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="aa147-924">Ответ с пустым текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="aa147-924">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="aa147-925">Ответ только с текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="aa147-925">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="aa147-926">Ответ с текстом сообщения и вложенным файлом.</span><span class="sxs-lookup"><span data-stu-id="aa147-926">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="aa147-927">Ответ с текстом сообщения и вложенным элементом.</span><span class="sxs-lookup"><span data-stu-id="aa147-927">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="aa147-928">Ответ с текстом сообщения, вложенным файлом, вложенным элементом и обратным вызовом.</span><span class="sxs-lookup"><span data-stu-id="aa147-928">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlook17officeentities"></a><span data-ttu-id="aa147-929">getEntities() → {[Entities](/javascript/api/outlook_1_7/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="aa147-929">getEntities() → {[Entities](/javascript/api/outlook_1_7/office.entities)}</span></span>

<span data-ttu-id="aa147-930">Получает сущности, обнаруженные в тексте выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="aa147-930">Gets the entities found in the selected item.</span></span>

> [!NOTE]
> <span data-ttu-id="aa147-931">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="aa147-931">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="aa147-932">Требования</span><span class="sxs-lookup"><span data-stu-id="aa147-932">Requirements</span></span>

|<span data-ttu-id="aa147-933">Requirement</span><span class="sxs-lookup"><span data-stu-id="aa147-933">Requirement</span></span>|<span data-ttu-id="aa147-934">Значение</span><span class="sxs-lookup"><span data-stu-id="aa147-934">Value</span></span>|
|---|---|
|[<span data-ttu-id="aa147-935">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="aa147-935">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="aa147-936">1.0</span><span class="sxs-lookup"><span data-stu-id="aa147-936">1.0</span></span>|
|[<span data-ttu-id="aa147-937">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="aa147-937">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="aa147-938">ReadItem</span><span class="sxs-lookup"><span data-stu-id="aa147-938">ReadItem</span></span>|
|[<span data-ttu-id="aa147-939">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="aa147-939">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="aa147-940">Чтение</span><span class="sxs-lookup"><span data-stu-id="aa147-940">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="aa147-941">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="aa147-941">Returns:</span></span>

<span data-ttu-id="aa147-942">Тип: [Entities](/javascript/api/outlook_1_7/office.entities)</span><span class="sxs-lookup"><span data-stu-id="aa147-942">Type: [Entities](/javascript/api/outlook_1_7/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="aa147-943">Пример</span><span class="sxs-lookup"><span data-stu-id="aa147-943">Example</span></span>

<span data-ttu-id="aa147-944">Ниже приведен пример получения доступа к сущностям контактов в тексте текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="aa147-944">The following example accesses the contacts entities on the current item.</span></span>

```js
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook17officecontactmeetingsuggestionjavascriptapioutlook17officemeetingsuggestionphonenumberjavascriptapioutlook17officephonenumbertasksuggestionjavascriptapioutlook17officetasksuggestion"></a><span data-ttu-id="aa147-945">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="aa147-945">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))>}</span></span>

<span data-ttu-id="aa147-946">Получает массив всех сущностей указанного типа, обнаруженных в тексте выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="aa147-946">Gets an array of all the entities of the specified entity type found in the selected item.</span></span>

> [!NOTE]
> <span data-ttu-id="aa147-947">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="aa147-947">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="aa147-948">Параметры</span><span class="sxs-lookup"><span data-stu-id="aa147-948">Parameters:</span></span>

|<span data-ttu-id="aa147-949">Имя</span><span class="sxs-lookup"><span data-stu-id="aa147-949">Name</span></span>|<span data-ttu-id="aa147-950">Тип</span><span class="sxs-lookup"><span data-stu-id="aa147-950">Type</span></span>|<span data-ttu-id="aa147-951">Описание</span><span class="sxs-lookup"><span data-stu-id="aa147-951">Description</span></span>|
|---|---|---|
|`entityType`|[<span data-ttu-id="aa147-952">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="aa147-952">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook_1_7/office.mailboxenums.entitytype)|<span data-ttu-id="aa147-953">Одно из значений перечисления EntityType.</span><span class="sxs-lookup"><span data-stu-id="aa147-953">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="aa147-954">Требования</span><span class="sxs-lookup"><span data-stu-id="aa147-954">Requirements</span></span>

|<span data-ttu-id="aa147-955">Requirement</span><span class="sxs-lookup"><span data-stu-id="aa147-955">Requirement</span></span>|<span data-ttu-id="aa147-956">Значение</span><span class="sxs-lookup"><span data-stu-id="aa147-956">Value</span></span>|
|---|---|
|[<span data-ttu-id="aa147-957">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="aa147-957">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="aa147-958">1.0</span><span class="sxs-lookup"><span data-stu-id="aa147-958">1.0</span></span>|
|[<span data-ttu-id="aa147-959">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="aa147-959">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="aa147-960">Restricted</span><span class="sxs-lookup"><span data-stu-id="aa147-960">Restricted</span></span>|
|[<span data-ttu-id="aa147-961">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="aa147-961">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="aa147-962">Чтение</span><span class="sxs-lookup"><span data-stu-id="aa147-962">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="aa147-963">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="aa147-963">Returns:</span></span>

<span data-ttu-id="aa147-964">Если значение, переданное в `entityType`, не является допустимым членом перечисления `EntityType`, метод возвращает значение NULL.</span><span class="sxs-lookup"><span data-stu-id="aa147-964">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="aa147-965">Если в тексте элемента отсутствуют сущности указанного типа, метод возвращает пустой массив.</span><span class="sxs-lookup"><span data-stu-id="aa147-965">If no entities of the specified type are present on the item, the method returns an empty array.</span></span> <span data-ttu-id="aa147-966">В противном случае тип объектов в возвращаемом массиве зависит от типа сущности, запрошенной в параметре `entityType`.</span><span class="sxs-lookup"><span data-stu-id="aa147-966">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="aa147-967">Хотя минимальный уровень разрешений для использования этого метода — **Restricted**, для некоторых типов сущностей требуется доступ на уровне **ReadItem**, как указано в приведенной ниже таблице.</span><span class="sxs-lookup"><span data-stu-id="aa147-967">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

|<span data-ttu-id="aa147-968">Значение параметра `entityType`</span><span class="sxs-lookup"><span data-stu-id="aa147-968">Value of `entityType`</span></span>|<span data-ttu-id="aa147-969">Тип объектов в возвращаемом массиве</span><span class="sxs-lookup"><span data-stu-id="aa147-969">Type of objects in returned array</span></span>|<span data-ttu-id="aa147-970">Необходимый уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="aa147-970">Required Permission Level</span></span>|
|---|---|---|
|`Address`|<span data-ttu-id="aa147-971">String</span><span class="sxs-lookup"><span data-stu-id="aa147-971">String</span></span>|<span data-ttu-id="aa147-972">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="aa147-972">**Restricted**</span></span>|
|`Contact`|<span data-ttu-id="aa147-973">Contact</span><span class="sxs-lookup"><span data-stu-id="aa147-973">Contact</span></span>|<span data-ttu-id="aa147-974">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="aa147-974">**ReadItem**</span></span>|
|`EmailAddress`|<span data-ttu-id="aa147-975">String</span><span class="sxs-lookup"><span data-stu-id="aa147-975">String</span></span>|<span data-ttu-id="aa147-976">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="aa147-976">**ReadItem**</span></span>|
|`MeetingSuggestion`|<span data-ttu-id="aa147-977">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="aa147-977">MeetingSuggestion</span></span>|<span data-ttu-id="aa147-978">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="aa147-978">**ReadItem**</span></span>|
|`PhoneNumber`|<span data-ttu-id="aa147-979">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="aa147-979">PhoneNumber</span></span>|<span data-ttu-id="aa147-980">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="aa147-980">**Restricted**</span></span>|
|`TaskSuggestion`|<span data-ttu-id="aa147-981">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="aa147-981">TaskSuggestion</span></span>|<span data-ttu-id="aa147-982">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="aa147-982">**ReadItem**</span></span>|
|`URL`|<span data-ttu-id="aa147-983">String</span><span class="sxs-lookup"><span data-stu-id="aa147-983">String</span></span>|<span data-ttu-id="aa147-984">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="aa147-984">**Restricted**</span></span>|

<span data-ttu-id="aa147-985">Тип: Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="aa147-985">Type: Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="aa147-986">Пример</span><span class="sxs-lookup"><span data-stu-id="aa147-986">Example</span></span>

<span data-ttu-id="aa147-987">В примере ниже показано, как получить доступ к массиву строк, которые представляют собой почтовые адреса в тексте текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="aa147-987">The following example shows how to access an array of strings that represent postal addresses in the subject or body of the current item.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook17officecontactmeetingsuggestionjavascriptapioutlook17officemeetingsuggestionphonenumberjavascriptapioutlook17officephonenumbertasksuggestionjavascriptapioutlook17officetasksuggestion"></a><span data-ttu-id="aa147-988">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="aa147-988">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))>}</span></span>

<span data-ttu-id="aa147-989">Возвращает известные сущности в выбранном элементе, которые проходят через именованный фильтр, определяемый в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="aa147-989">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="aa147-990">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="aa147-990">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="aa147-991">Метод `getFilteredEntitiesByName` возвращает сущности, соответствующие регулярному выражению, которое определяется в элементе правила [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) в XML-файле манифеста, с использованием указанного значения элемента `FilterName`.</span><span class="sxs-lookup"><span data-stu-id="aa147-991">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="aa147-992">Параметры</span><span class="sxs-lookup"><span data-stu-id="aa147-992">Parameters:</span></span>

|<span data-ttu-id="aa147-993">Имя</span><span class="sxs-lookup"><span data-stu-id="aa147-993">Name</span></span>|<span data-ttu-id="aa147-994">Тип</span><span class="sxs-lookup"><span data-stu-id="aa147-994">Type</span></span>|<span data-ttu-id="aa147-995">Описание</span><span class="sxs-lookup"><span data-stu-id="aa147-995">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="aa147-996">String</span><span class="sxs-lookup"><span data-stu-id="aa147-996">String</span></span>|<span data-ttu-id="aa147-997">Имя элемента правила `ItemHasKnownEntity`, определяющее соответствующий фильтр.</span><span class="sxs-lookup"><span data-stu-id="aa147-997">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="aa147-998">Требования</span><span class="sxs-lookup"><span data-stu-id="aa147-998">Requirements</span></span>

|<span data-ttu-id="aa147-999">Requirement</span><span class="sxs-lookup"><span data-stu-id="aa147-999">Requirement</span></span>|<span data-ttu-id="aa147-1000">Значение</span><span class="sxs-lookup"><span data-stu-id="aa147-1000">Value</span></span>|
|---|---|
|[<span data-ttu-id="aa147-1001">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="aa147-1001">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="aa147-1002">1.0</span><span class="sxs-lookup"><span data-stu-id="aa147-1002">1.0</span></span>|
|[<span data-ttu-id="aa147-1003">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="aa147-1003">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="aa147-1004">ReadItem</span><span class="sxs-lookup"><span data-stu-id="aa147-1004">ReadItem</span></span>|
|[<span data-ttu-id="aa147-1005">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="aa147-1005">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="aa147-1006">Чтение</span><span class="sxs-lookup"><span data-stu-id="aa147-1006">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="aa147-1007">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="aa147-1007">Returns:</span></span>

<span data-ttu-id="aa147-p158">Если в манифесте нет элемента `ItemHasKnownEntity` со значением `FilterName`, соответствующим параметру `name`, метод возвращает `null`. Если параметр `name` соответствует элементу `ItemHasKnownEntity` в манифесте, но при этом в текущем элементе нет соответствующих сущностей, метод возвращает пустой массив.</span><span class="sxs-lookup"><span data-stu-id="aa147-p158">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="aa147-1010">Тип: Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="aa147-1010">Type: Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))></span></span>

#### <a name="getregexmatches--object"></a><span data-ttu-id="aa147-1011">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="aa147-1011">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="aa147-1012">Возвращает строковые значения в выбранном элементе, которые соответствуют регулярным выражениям, определенным в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="aa147-1012">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="aa147-1013">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="aa147-1013">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="aa147-p159">Метод `getRegExMatches` возвращает строки, соответствующие регулярному выражению, которое определяется в каждом элементе правила `ItemHasRegularExpressionMatch` или `ItemHasKnownEntity` в XML-файле манифеста. Для правила `ItemHasRegularExpressionMatch` соответствующую строку должно содержать свойство элемента, указанного этим правилом. Простой тип `PropertyName` определяет поддерживаемые свойства.</span><span class="sxs-lookup"><span data-stu-id="aa147-p159">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="aa147-1017">Например, рассмотрим манифест надстройки, который содержит указанный ниже элемент `Rule`.</span><span class="sxs-lookup"><span data-stu-id="aa147-1017">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="aa147-1018">Объект, возвращаемый методом `getRegExMatches`, будет содержать два свойства: `fruits` и `veggies`.</span><span class="sxs-lookup"><span data-stu-id="aa147-1018">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="aa147-p160">Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты. Лучше используйте метод [`Body.getAsync`](/javascript/api/outlook_1_7/office.body#getasync-coerciontype--options--callback-) для этого.</span><span class="sxs-lookup"><span data-stu-id="aa147-p160">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook_1_7/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="aa147-1022">Requirements</span><span class="sxs-lookup"><span data-stu-id="aa147-1022">Requirements</span></span>

|<span data-ttu-id="aa147-1023">Requirement</span><span class="sxs-lookup"><span data-stu-id="aa147-1023">Requirement</span></span>|<span data-ttu-id="aa147-1024">Значение</span><span class="sxs-lookup"><span data-stu-id="aa147-1024">Value</span></span>|
|---|---|
|[<span data-ttu-id="aa147-1025">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="aa147-1025">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="aa147-1026">1.0</span><span class="sxs-lookup"><span data-stu-id="aa147-1026">1.0</span></span>|
|[<span data-ttu-id="aa147-1027">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="aa147-1027">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="aa147-1028">ReadItem</span><span class="sxs-lookup"><span data-stu-id="aa147-1028">ReadItem</span></span>|
|[<span data-ttu-id="aa147-1029">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="aa147-1029">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="aa147-1030">Чтение</span><span class="sxs-lookup"><span data-stu-id="aa147-1030">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="aa147-1031">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="aa147-1031">Returns:</span></span>

<span data-ttu-id="aa147-p161">Объект, содержащий массив строк, которые соответствуют регулярным выражениям, определяемым в XML-файле манифеста. Имя каждого массива равно соответствующему значению атрибута `RegExName` подходящего правила `ItemHasRegularExpressionMatch` или атрибута `FilterName` соответствующего правила `ItemHasKnownEntity`.</span><span class="sxs-lookup"><span data-stu-id="aa147-p161">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="aa147-1034">

<dt>Тип</dt>

</span><span class="sxs-lookup"><span data-stu-id="aa147-1034">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="aa147-1035">Object</span><span class="sxs-lookup"><span data-stu-id="aa147-1035">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="aa147-1036">Пример</span><span class="sxs-lookup"><span data-stu-id="aa147-1036">Example</span></span>

<span data-ttu-id="aa147-1037">В приведенном ниже примере показано, как получить доступ к массиву совпадений с элементами `fruits` и `veggies` правил активации регулярных выражений, указанными в манифесте.</span><span class="sxs-lookup"><span data-stu-id="aa147-1037">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```js
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veges = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="aa147-1038">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="aa147-1038">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="aa147-1039">Возвращает строковые значения в выбранном элементе, которые соответствуют именованному регулярному выражению, определенному в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="aa147-1039">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="aa147-1040">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="aa147-1040">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="aa147-1041">Метод `getRegExMatchesByName` возвращает строки, соответствующие регулярному выражению, которое определяется в элементе правила `ItemHasRegularExpressionMatch` в XML-файле манифеста, с использованием указанного значения элемента `RegExName`.</span><span class="sxs-lookup"><span data-stu-id="aa147-1041">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="aa147-p162">Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты.</span><span class="sxs-lookup"><span data-stu-id="aa147-p162">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="aa147-1044">Параметры</span><span class="sxs-lookup"><span data-stu-id="aa147-1044">Parameters:</span></span>

|<span data-ttu-id="aa147-1045">Имя</span><span class="sxs-lookup"><span data-stu-id="aa147-1045">Name</span></span>|<span data-ttu-id="aa147-1046">Тип</span><span class="sxs-lookup"><span data-stu-id="aa147-1046">Type</span></span>|<span data-ttu-id="aa147-1047">Описание</span><span class="sxs-lookup"><span data-stu-id="aa147-1047">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="aa147-1048">String</span><span class="sxs-lookup"><span data-stu-id="aa147-1048">String</span></span>|<span data-ttu-id="aa147-1049">Имя элемента правила `ItemHasRegularExpressionMatch`, определяющее соответствующий фильтр.</span><span class="sxs-lookup"><span data-stu-id="aa147-1049">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="aa147-1050">Требования</span><span class="sxs-lookup"><span data-stu-id="aa147-1050">Requirements</span></span>

|<span data-ttu-id="aa147-1051">Requirement</span><span class="sxs-lookup"><span data-stu-id="aa147-1051">Requirement</span></span>|<span data-ttu-id="aa147-1052">Значение</span><span class="sxs-lookup"><span data-stu-id="aa147-1052">Value</span></span>|
|---|---|
|[<span data-ttu-id="aa147-1053">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="aa147-1053">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="aa147-1054">1.0</span><span class="sxs-lookup"><span data-stu-id="aa147-1054">1.0</span></span>|
|[<span data-ttu-id="aa147-1055">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="aa147-1055">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="aa147-1056">ReadItem</span><span class="sxs-lookup"><span data-stu-id="aa147-1056">ReadItem</span></span>|
|[<span data-ttu-id="aa147-1057">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="aa147-1057">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="aa147-1058">Чтение</span><span class="sxs-lookup"><span data-stu-id="aa147-1058">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="aa147-1059">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="aa147-1059">Returns:</span></span>

<span data-ttu-id="aa147-1060">Массив строк, соответствующих регулярному выражению, определяемому в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="aa147-1060">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="aa147-1061">

<dt>Тип</dt>

</span><span class="sxs-lookup"><span data-stu-id="aa147-1061">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="aa147-1062">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="aa147-1062">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="aa147-1063">Пример</span><span class="sxs-lookup"><span data-stu-id="aa147-1063">Example</span></span>

```js
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="aa147-1064">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="aa147-1064">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="aa147-1065">Асинхронно возвращает данные, выбранные в теме или тексте сообщения.</span><span class="sxs-lookup"><span data-stu-id="aa147-1065">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="aa147-p163">Если выделенный фрагмент отсутствует, но курсор находится в тексте или теме, метод возвращает значение NULL для выбранных данных. Если выбраны не текст и не тема, метод возвращает ошибку `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="aa147-p163">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="aa147-1068">Параметры</span><span class="sxs-lookup"><span data-stu-id="aa147-1068">Parameters:</span></span>

|<span data-ttu-id="aa147-1069">Имя</span><span class="sxs-lookup"><span data-stu-id="aa147-1069">Name</span></span>|<span data-ttu-id="aa147-1070">Тип</span><span class="sxs-lookup"><span data-stu-id="aa147-1070">Type</span></span>|<span data-ttu-id="aa147-1071">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="aa147-1071">Attributes</span></span>|<span data-ttu-id="aa147-1072">Описание</span><span class="sxs-lookup"><span data-stu-id="aa147-1072">Description</span></span>|
|---|---|---|---|
|`coercionType`|[<span data-ttu-id="aa147-1073">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="aa147-1073">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="aa147-p164">Запрашивает формат данных. Если задано значение Text, метод возвращает обычный текст как строку, удаляя все имеющиеся HTML-теги. Если задано значение HTML, метод возвращает выделенный текст (обычный текст или HTML).</span><span class="sxs-lookup"><span data-stu-id="aa147-p164">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`|<span data-ttu-id="aa147-1077">Object</span><span class="sxs-lookup"><span data-stu-id="aa147-1077">Object</span></span>|<span data-ttu-id="aa147-1078">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="aa147-1078">&lt;optional&gt;</span></span>|<span data-ttu-id="aa147-1079">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="aa147-1079">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="aa147-1080">Object</span><span class="sxs-lookup"><span data-stu-id="aa147-1080">Object</span></span>|<span data-ttu-id="aa147-1081">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="aa147-1081">&lt;optional&gt;</span></span>|<span data-ttu-id="aa147-1082">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="aa147-1082">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="aa147-1083">функция</span><span class="sxs-lookup"><span data-stu-id="aa147-1083">function</span></span>||<span data-ttu-id="aa147-1084">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="aa147-1084">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="aa147-1085">Чтобы получить доступ к выбранным данным из метода обратного вызова, вызовите `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="aa147-1085">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="aa147-1086">Для доступа к исходному свойству, представляющему собой источник выбранных данных, вызовите параметр `asyncResult.value.sourceProperty`, который может иметь значение `body` или `subject`.</span><span class="sxs-lookup"><span data-stu-id="aa147-1086">To access the source property that the selection comes from, call , which will be either  or .|</span></span>|

##### <a name="requirements"></a><span data-ttu-id="aa147-1087">Требования</span><span class="sxs-lookup"><span data-stu-id="aa147-1087">Requirements</span></span>

|<span data-ttu-id="aa147-1088">Requirement</span><span class="sxs-lookup"><span data-stu-id="aa147-1088">Requirement</span></span>|<span data-ttu-id="aa147-1089">Значение</span><span class="sxs-lookup"><span data-stu-id="aa147-1089">Value</span></span>|
|---|---|
|[<span data-ttu-id="aa147-1090">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="aa147-1090">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="aa147-1091">1.2</span><span class="sxs-lookup"><span data-stu-id="aa147-1091">1.2</span></span>|
|[<span data-ttu-id="aa147-1092">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="aa147-1092">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="aa147-1093">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="aa147-1093">ReadWriteItem</span></span>|
|[<span data-ttu-id="aa147-1094">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="aa147-1094">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="aa147-1095">Создание</span><span class="sxs-lookup"><span data-stu-id="aa147-1095">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="aa147-1096">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="aa147-1096">Returns:</span></span>

<span data-ttu-id="aa147-1097">Выбранные данные в виде строки с форматом, определенным в параметре `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="aa147-1097">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="aa147-1098">

<dt>Тип</dt>

</span><span class="sxs-lookup"><span data-stu-id="aa147-1098">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="aa147-1099">String</span><span class="sxs-lookup"><span data-stu-id="aa147-1099">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="aa147-1100">Пример</span><span class="sxs-lookup"><span data-stu-id="aa147-1100">Example</span></span>

```js
// getting selected data
Office.initialize = function () {
    Office.context.mailbox.item.getSelectedDataAsync(Office.CoercionType.Text, {}, getCallback);
}

function getCallback(asyncResult) {
    var text = asyncResult.value.data;
    var prop = asyncResult.value.sourceProperty;

    Office.context.mailbox.item.setSelectedDataAsync('Setting ' + prop + ': ' + text, {}, setCallback);
}

function setCallback(asyncResult) {
    // check for errors
}
```

#### <a name="getselectedentities--entitiesjavascriptapioutlook17officeentities"></a><span data-ttu-id="aa147-1101">getSelectedEntities() → {[Entities](/javascript/api/outlook_1_7/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="aa147-1101">getSelectedEntities() → {[Entities](/javascript/api/outlook_1_7/office.entities)}</span></span>

<span data-ttu-id="aa147-p166">Возвращает сущности, найденные в выделенном совпадении, выбранном пользователем. Выделенные совпадения применяются к [контекстным надстройкам](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="aa147-p166">Gets the entities found in a highlighted match a user has selected. Highlighted matches apply to [contextual add-ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="aa147-1104">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="aa147-1104">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="aa147-1105">Требования</span><span class="sxs-lookup"><span data-stu-id="aa147-1105">Requirements</span></span>

|<span data-ttu-id="aa147-1106">Requirement</span><span class="sxs-lookup"><span data-stu-id="aa147-1106">Requirement</span></span>|<span data-ttu-id="aa147-1107">Значение</span><span class="sxs-lookup"><span data-stu-id="aa147-1107">Value</span></span>|
|---|---|
|[<span data-ttu-id="aa147-1108">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="aa147-1108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="aa147-1109">1.6</span><span class="sxs-lookup"><span data-stu-id="aa147-1109">1.6</span></span>|
|[<span data-ttu-id="aa147-1110">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="aa147-1110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="aa147-1111">ReadItem</span><span class="sxs-lookup"><span data-stu-id="aa147-1111">ReadItem</span></span>|
|[<span data-ttu-id="aa147-1112">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="aa147-1112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="aa147-1113">Чтение</span><span class="sxs-lookup"><span data-stu-id="aa147-1113">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="aa147-1114">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="aa147-1114">Returns:</span></span>

<span data-ttu-id="aa147-1115">Тип: [Entities](/javascript/api/outlook_1_7/office.entities)</span><span class="sxs-lookup"><span data-stu-id="aa147-1115">Type: [Entities](/javascript/api/outlook_1_7/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="aa147-1116">Пример</span><span class="sxs-lookup"><span data-stu-id="aa147-1116">Example</span></span>

<span data-ttu-id="aa147-1117">В приведенном ниже примере показано, как получить доступ к сущностям адресов в выделенном совпадении, выбранном пользователем.</span><span class="sxs-lookup"><span data-stu-id="aa147-1117">The following example accesses the addresses entities in the highlighted match selected by the user.</span></span>

```js
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

#### <a name="getselectedregexmatches--object"></a><span data-ttu-id="aa147-1118">getSelectedRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="aa147-1118">getSelectedRegExMatches() → {Object}</span></span>

<span data-ttu-id="aa147-p167">Возвращает строковые значения в выделенном совпадении, которые соответствуют регулярным выражениям, определенным в XML-файле манифеста. Выделенные совпадения применяются к [контекстным надстройкам](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="aa147-p167">Returns string values in a highlighted match that match the regular expressions defined in the manifest XML file. Highlighted matches apply to [contextual add-ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="aa147-1121">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="aa147-1121">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="aa147-p168">Метод `getSelectedRegExMatches` возвращает строки, соответствующие регулярному выражению, которое определяется в каждом элементе правила `ItemHasRegularExpressionMatch` или `ItemHasKnownEntity` в XML-файле манифеста. Для правила `ItemHasRegularExpressionMatch` соответствующую строку должно содержать свойство элемента, указанного этим правилом. Простой тип `PropertyName` определяет поддерживаемые свойства.</span><span class="sxs-lookup"><span data-stu-id="aa147-p168">The `getSelectedRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="aa147-1125">Например, рассмотрим манифест надстройки, который содержит указанный ниже элемент `Rule`.</span><span class="sxs-lookup"><span data-stu-id="aa147-1125">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="aa147-1126">Объект, возвращаемый методом `getRegExMatches`, будет содержать два свойства: `fruits` и `veggies`.</span><span class="sxs-lookup"><span data-stu-id="aa147-1126">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="aa147-p169">Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты. Лучше используйте метод [`Body.getAsync`](/javascript/api/outlook_1_7/office.body#getasync-coerciontype--options--callback-) для этого.</span><span class="sxs-lookup"><span data-stu-id="aa147-p169">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook_1_7/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="aa147-1130">Requirements</span><span class="sxs-lookup"><span data-stu-id="aa147-1130">Requirements</span></span>

|<span data-ttu-id="aa147-1131">Requirement</span><span class="sxs-lookup"><span data-stu-id="aa147-1131">Requirement</span></span>|<span data-ttu-id="aa147-1132">Значение</span><span class="sxs-lookup"><span data-stu-id="aa147-1132">Value</span></span>|
|---|---|
|[<span data-ttu-id="aa147-1133">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="aa147-1133">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="aa147-1134">1.6</span><span class="sxs-lookup"><span data-stu-id="aa147-1134">1.6</span></span>|
|[<span data-ttu-id="aa147-1135">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="aa147-1135">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="aa147-1136">ReadItem</span><span class="sxs-lookup"><span data-stu-id="aa147-1136">ReadItem</span></span>|
|[<span data-ttu-id="aa147-1137">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="aa147-1137">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="aa147-1138">Чтение</span><span class="sxs-lookup"><span data-stu-id="aa147-1138">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="aa147-1139">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="aa147-1139">Returns:</span></span>

<span data-ttu-id="aa147-p170">Объект, содержащий массив строк, которые соответствуют регулярным выражениям, определяемым в XML-файле манифеста. Имя каждого массива равно соответствующему значению атрибута `RegExName` подходящего правила `ItemHasRegularExpressionMatch` или атрибута `FilterName` соответствующего правила `ItemHasKnownEntity`.</span><span class="sxs-lookup"><span data-stu-id="aa147-p170">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

##### <a name="example"></a><span data-ttu-id="aa147-1142">Пример</span><span class="sxs-lookup"><span data-stu-id="aa147-1142">Example</span></span>

<span data-ttu-id="aa147-1143">В приведенном ниже примере показано, как получить доступ к массиву совпадений с элементами `fruits` и `veggies` правил активации регулярных выражений, указанными в манифесте.</span><span class="sxs-lookup"><span data-stu-id="aa147-1143">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```js
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="aa147-1144">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="aa147-1144">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="aa147-1145">Асинхронно загружает настраиваемые свойства для надстройки для выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="aa147-1145">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="aa147-p171">Настраиваемые свойства сохраняются в виде пар "ключ-значение" для каждого приложения и каждого элемента. Этот метод возвращает объект `CustomProperties` при обратном вызове, который предоставляет методы для доступа к настраиваемым свойствам, характерным для текущего элемента и текущей надстройки. Настраиваемые свойства не шифруются для элемента, поэтому этот способ хранения не является безопасным.</span><span class="sxs-lookup"><span data-stu-id="aa147-p171">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="aa147-1149">Параметры</span><span class="sxs-lookup"><span data-stu-id="aa147-1149">Parameters:</span></span>

|<span data-ttu-id="aa147-1150">Имя</span><span class="sxs-lookup"><span data-stu-id="aa147-1150">Name</span></span>|<span data-ttu-id="aa147-1151">Тип</span><span class="sxs-lookup"><span data-stu-id="aa147-1151">Type</span></span>|<span data-ttu-id="aa147-1152">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="aa147-1152">Attributes</span></span>|<span data-ttu-id="aa147-1153">Описание</span><span class="sxs-lookup"><span data-stu-id="aa147-1153">Description</span></span>|
|---|---|---|---|
|`callback`|<span data-ttu-id="aa147-1154">function</span><span class="sxs-lookup"><span data-stu-id="aa147-1154">function</span></span>||<span data-ttu-id="aa147-1155">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="aa147-1155">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="aa147-1156">Настраиваемые свойства предоставляются в виде объекта [`CustomProperties`](/javascript/api/outlook_1_7/office.customproperties) в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="aa147-1156">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook_1_7/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="aa147-1157">Этот объект позволяет получить, задать и удалить настраиваемые свойства из элемента, а также сохранить изменения, внесенные в набор настраиваемых свойств, на сервере.</span><span class="sxs-lookup"><span data-stu-id="aa147-1157">The custom properties are provided as a CustomProperties object in the asyncResult.value property. This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`|<span data-ttu-id="aa147-1158">Object</span><span class="sxs-lookup"><span data-stu-id="aa147-1158">Object</span></span>|<span data-ttu-id="aa147-1159">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="aa147-1159">&lt;optional&gt;</span></span>|<span data-ttu-id="aa147-1160">Разработчики могут указать любой объект, к которому необходимо получить доступ, в функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="aa147-1160">Developers can provide any object they wish to access in the callback method.</span></span> <span data-ttu-id="aa147-1161">Доступ к этому объекту можно получить с помощью свойства `asyncResult.asyncContext` в функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="aa147-1161">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.|</span></span>|

##### <a name="requirements"></a><span data-ttu-id="aa147-1162">Требования</span><span class="sxs-lookup"><span data-stu-id="aa147-1162">Requirements</span></span>

|<span data-ttu-id="aa147-1163">Requirement</span><span class="sxs-lookup"><span data-stu-id="aa147-1163">Requirement</span></span>|<span data-ttu-id="aa147-1164">Значение</span><span class="sxs-lookup"><span data-stu-id="aa147-1164">Value</span></span>|
|---|---|
|[<span data-ttu-id="aa147-1165">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="aa147-1165">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="aa147-1166">1.0</span><span class="sxs-lookup"><span data-stu-id="aa147-1166">1.0</span></span>|
|[<span data-ttu-id="aa147-1167">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="aa147-1167">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="aa147-1168">ReadItem</span><span class="sxs-lookup"><span data-stu-id="aa147-1168">ReadItem</span></span>|
|[<span data-ttu-id="aa147-1169">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="aa147-1169">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="aa147-1170">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="aa147-1170">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="aa147-1171">Пример</span><span class="sxs-lookup"><span data-stu-id="aa147-1171">Example</span></span>

<span data-ttu-id="aa147-p174">Приведенный ниже пример кода показывает, как асинхронно загружать настраиваемые свойства, характерные для текущего элемента, с помощью метода `loadCustomPropertiesAsync`. Этот пример также показывает, как сохранять эти свойства на сервере с помощью метода `CustomProperties.saveAsync`. После загрузки настраиваемых свойств в этом примере кода метод `CustomProperties.get` используется для считывания настраиваемого свойства `myProp`, метод `CustomProperties.set` — для записи настраиваемого свойства `otherProp`, а метод `saveAsync` — для сохранения настраиваемых свойств.</span><span class="sxs-lookup"><span data-stu-id="aa147-p174">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="aa147-1175">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="aa147-1175">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="aa147-1176">Удаляет вложение из сообщения или встречи.</span><span class="sxs-lookup"><span data-stu-id="aa147-1176">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="aa147-p175">Метод `removeAttachmentAsync` удаляет из элемента вложение с указанным идентификатором. Идентификатор вложения рекомендуется использовать для удаления вложения, только если оно добавлено тем же почтовым приложением в ходе текущего сеанса. В Outlook Web App и Outlook Web App для устройств идентификатор вложения действителен только в рамках одного сеанса. Сеанс завершается, когда пользователь закрывает приложение или начинает создавать элемент во встроенной форме, а затем переходит из формы в отдельное окно.</span><span class="sxs-lookup"><span data-stu-id="aa147-p175">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="aa147-1181">Параметры</span><span class="sxs-lookup"><span data-stu-id="aa147-1181">Parameters:</span></span>

|<span data-ttu-id="aa147-1182">Имя</span><span class="sxs-lookup"><span data-stu-id="aa147-1182">Name</span></span>|<span data-ttu-id="aa147-1183">Тип</span><span class="sxs-lookup"><span data-stu-id="aa147-1183">Type</span></span>|<span data-ttu-id="aa147-1184">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="aa147-1184">Attributes</span></span>|<span data-ttu-id="aa147-1185">Описание</span><span class="sxs-lookup"><span data-stu-id="aa147-1185">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="aa147-1186">String</span><span class="sxs-lookup"><span data-stu-id="aa147-1186">String</span></span>||<span data-ttu-id="aa147-p176">Идентификатор удаляемого вложения. Максимальная длина строки — 100 символов.</span><span class="sxs-lookup"><span data-stu-id="aa147-p176">The identifier of the attachment to remove. The maximum length of the string is 100 characters.</span></span>|
|`options`|<span data-ttu-id="aa147-1189">Object</span><span class="sxs-lookup"><span data-stu-id="aa147-1189">Object</span></span>|<span data-ttu-id="aa147-1190">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="aa147-1190">&lt;optional&gt;</span></span>|<span data-ttu-id="aa147-1191">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="aa147-1191">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="aa147-1192">Object</span><span class="sxs-lookup"><span data-stu-id="aa147-1192">Object</span></span>|<span data-ttu-id="aa147-1193">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="aa147-1193">&lt;optional&gt;</span></span>|<span data-ttu-id="aa147-1194">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="aa147-1194">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="aa147-1195">функция</span><span class="sxs-lookup"><span data-stu-id="aa147-1195">function</span></span>|<span data-ttu-id="aa147-1196">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="aa147-1196">&lt;optional&gt;</span></span>|<span data-ttu-id="aa147-1197">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="aa147-1197">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="aa147-1198">Если удалить вложение не удается, свойство `asyncResult.error` содержит код ошибки с указанием ее причины.</span><span class="sxs-lookup"><span data-stu-id="aa147-1198">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="aa147-1199">Ошибки</span><span class="sxs-lookup"><span data-stu-id="aa147-1199">Errors</span></span>

|<span data-ttu-id="aa147-1200">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="aa147-1200">Error code</span></span>|<span data-ttu-id="aa147-1201">Описание</span><span class="sxs-lookup"><span data-stu-id="aa147-1201">Description</span></span>|
|------------|-------------|
|`InvalidAttachmentId`|<span data-ttu-id="aa147-1202">Идентификатор вложения не существует.</span><span class="sxs-lookup"><span data-stu-id="aa147-1202">The attachment identifier does not exist.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="aa147-1203">Требования</span><span class="sxs-lookup"><span data-stu-id="aa147-1203">Requirements</span></span>

|<span data-ttu-id="aa147-1204">Requirement</span><span class="sxs-lookup"><span data-stu-id="aa147-1204">Requirement</span></span>|<span data-ttu-id="aa147-1205">Значение</span><span class="sxs-lookup"><span data-stu-id="aa147-1205">Value</span></span>|
|---|---|
|[<span data-ttu-id="aa147-1206">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="aa147-1206">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="aa147-1207">1.1</span><span class="sxs-lookup"><span data-stu-id="aa147-1207">1.1</span></span>|
|[<span data-ttu-id="aa147-1208">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="aa147-1208">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="aa147-1209">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="aa147-1209">ReadWriteItem</span></span>|
|[<span data-ttu-id="aa147-1210">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="aa147-1210">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="aa147-1211">Создание</span><span class="sxs-lookup"><span data-stu-id="aa147-1211">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="aa147-1212">Пример</span><span class="sxs-lookup"><span data-stu-id="aa147-1212">Example</span></span>

<span data-ttu-id="aa147-1213">Указанный ниже код удаляет вложение с идентификатором "0".</span><span class="sxs-lookup"><span data-stu-id="aa147-1213">The following code removes an attachment with an identifier of '0'.</span></span>

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

####  <a name="removehandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="aa147-1214">removeHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="aa147-1214">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="aa147-1215">Удаляет обработчик для поддерживаемого события.</span><span class="sxs-lookup"><span data-stu-id="aa147-1215">Removes an event handler for a</span></span>

<span data-ttu-id="aa147-1216">Сейчас поддерживаются следующие типы событий: `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged` и `Office.EventType.RecurrenceChanged`.</span><span class="sxs-lookup"><span data-stu-id="aa147-1216">Currently the supported event types are `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`</span></span>

##### <a name="parameters"></a><span data-ttu-id="aa147-1217">Параметры</span><span class="sxs-lookup"><span data-stu-id="aa147-1217">Parameters:</span></span>

| <span data-ttu-id="aa147-1218">Имя</span><span class="sxs-lookup"><span data-stu-id="aa147-1218">Name</span></span> | <span data-ttu-id="aa147-1219">Тип</span><span class="sxs-lookup"><span data-stu-id="aa147-1219">Type</span></span> | <span data-ttu-id="aa147-1220">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="aa147-1220">Attributes</span></span> | <span data-ttu-id="aa147-1221">Описание</span><span class="sxs-lookup"><span data-stu-id="aa147-1221">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="aa147-1222">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="aa147-1222">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="aa147-1223">Событие, которое должно вызвать обработчик.</span><span class="sxs-lookup"><span data-stu-id="aa147-1223">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="aa147-1224">Function</span><span class="sxs-lookup"><span data-stu-id="aa147-1224">Function</span></span> || <span data-ttu-id="aa147-p177">Функция для обработки события. Функция должна принимать один параметр, представляющий собой объектный литерал. Значение свойства `type` параметра совпадет со значением параметра `eventType`, переданного методу `removeHandlerAsync`.</span><span class="sxs-lookup"><span data-stu-id="aa147-p177">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `removeHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="aa147-1228">Объект</span><span class="sxs-lookup"><span data-stu-id="aa147-1228">Object</span></span> | <span data-ttu-id="aa147-1229">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="aa147-1229">&lt;optional&gt;</span></span> | <span data-ttu-id="aa147-1230">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="aa147-1230">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="aa147-1231">Object</span><span class="sxs-lookup"><span data-stu-id="aa147-1231">Object</span></span> | <span data-ttu-id="aa147-1232">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="aa147-1232">&lt;optional&gt;</span></span> | <span data-ttu-id="aa147-1233">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="aa147-1233">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="aa147-1234">функция</span><span class="sxs-lookup"><span data-stu-id="aa147-1234">function</span></span>| <span data-ttu-id="aa147-1235">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="aa147-1235">&lt;optional&gt;</span></span>|<span data-ttu-id="aa147-1236">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="aa147-1236">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="aa147-1237">Требования</span><span class="sxs-lookup"><span data-stu-id="aa147-1237">Requirements</span></span>

|<span data-ttu-id="aa147-1238">Requirement</span><span class="sxs-lookup"><span data-stu-id="aa147-1238">Requirement</span></span>| <span data-ttu-id="aa147-1239">Значение</span><span class="sxs-lookup"><span data-stu-id="aa147-1239">Value</span></span>|
|---|---|
|[<span data-ttu-id="aa147-1240">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="aa147-1240">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="aa147-1241">1.7</span><span class="sxs-lookup"><span data-stu-id="aa147-1241">ExcelApi 1.7 Beta</span></span> |
|[<span data-ttu-id="aa147-1242">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="aa147-1242">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="aa147-1243">ReadItem</span><span class="sxs-lookup"><span data-stu-id="aa147-1243">ReadItem</span></span> |
|[<span data-ttu-id="aa147-1244">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="aa147-1244">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="aa147-1245">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="aa147-1245">Compose or read</span></span> |

##### <a name="example"></a><span data-ttu-id="aa147-1246">Пример</span><span class="sxs-lookup"><span data-stu-id="aa147-1246">Example</span></span>

```js
Office.initialize = function (reason) {
  $(document).ready(function () {
    Office.context.mailbox.item.removeHandlerAsync(Office.EventType.RecurrenceChanged, loadNewItem, function (result) {
      if (result.status === Office.AsyncResultStatus.Failed) {
        // Handle error
      }
    });
  });
};

function loadNewItem(eventArgs) {
  // Load the properties of the newly selected item
  loadProps(Office.context.mailbox.item);
};
```

####  <a name="saveasyncoptions-callback"></a><span data-ttu-id="aa147-1247">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="aa147-1247">saveAsync([options], callback)</span></span>

<span data-ttu-id="aa147-1248">Асинхронно сохраняет элемент.</span><span class="sxs-lookup"><span data-stu-id="aa147-1248">Asynchronously saves an item.</span></span>

<span data-ttu-id="aa147-p178">При вызове этот метод сохраняет текущее сообщение в виде черновика и возвращает идентификатор элемента с помощью метода обратного вызова. В Outlook Web App или интерактивном режиме Outlook этот элемент сохраняется на сервере. В Outlook в режиме кэширования этот элемент сохраняется в локальном кэше.</span><span class="sxs-lookup"><span data-stu-id="aa147-p178">When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="aa147-1252">Если в режиме создания надстройка вызывает для элемента метод `saveAsync`, чтобы получить параметр `itemId` для использования с EWS или REST API, помните, что синхронизация элемента с сервером может занять много времени, если Outlook работает в режиме кэширования данных.</span><span class="sxs-lookup"><span data-stu-id="aa147-1252">Note: If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server. Until the item is synced, using the  will return an error.</span></span> <span data-ttu-id="aa147-1253">До окончания синхронизации применение параметра `itemId` будет приводить к ошибке.</span><span class="sxs-lookup"><span data-stu-id="aa147-1253">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="aa147-p180">Если метод `saveAsync` вызывается для встречи в режиме создания, она сохраняется как обычная встреча в календаре пользователя, а не как черновик. При сохранении новой встречи приглашения не отправляются. При сохранении существующей встречи уведомления отправляются добавленным или удаленным участникам.</span><span class="sxs-lookup"><span data-stu-id="aa147-p180">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="aa147-1257">Следующие клиенты отличаются другим поведением метода `saveAsync` в отношении встреч в режиме создания:</span><span class="sxs-lookup"><span data-stu-id="aa147-1257">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="aa147-1258">Outlook для Mac не поддерживает метод `saveAsync` для собраний в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="aa147-1258">Note: Mac Outlook does not support `saveAsync` on a meeting in compose mode. Calling  on a meeting in Mac Outlook will return an error.</span></span> <span data-ttu-id="aa147-1259">Метод `saveAsync`, вызванный для собрания в Outlook для Mac, возвращает ошибку.</span><span class="sxs-lookup"><span data-stu-id="aa147-1259">Note: Mac Outlook does not support  on a meeting in compose mode. Calling `saveAsync` on a meeting in Mac Outlook will return an error.</span></span>
> - <span data-ttu-id="aa147-1260">Outlook в Интернете всегда отправляет приглашение или обновление при вызове метода `saveAsync` для встречи в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="aa147-1260">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="aa147-1261">Параметры</span><span class="sxs-lookup"><span data-stu-id="aa147-1261">Parameters:</span></span>

|<span data-ttu-id="aa147-1262">Имя</span><span class="sxs-lookup"><span data-stu-id="aa147-1262">Name</span></span>|<span data-ttu-id="aa147-1263">Тип</span><span class="sxs-lookup"><span data-stu-id="aa147-1263">Type</span></span>|<span data-ttu-id="aa147-1264">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="aa147-1264">Attributes</span></span>|<span data-ttu-id="aa147-1265">Описание</span><span class="sxs-lookup"><span data-stu-id="aa147-1265">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="aa147-1266">Object</span><span class="sxs-lookup"><span data-stu-id="aa147-1266">Object</span></span>|<span data-ttu-id="aa147-1267">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="aa147-1267">&lt;optional&gt;</span></span>|<span data-ttu-id="aa147-1268">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="aa147-1268">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="aa147-1269">Object</span><span class="sxs-lookup"><span data-stu-id="aa147-1269">Object</span></span>|<span data-ttu-id="aa147-1270">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="aa147-1270">&lt;optional&gt;</span></span>|<span data-ttu-id="aa147-1271">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="aa147-1271">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="aa147-1272">функция</span><span class="sxs-lookup"><span data-stu-id="aa147-1272">function</span></span>||<span data-ttu-id="aa147-1273">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="aa147-1273">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="aa147-1274">После успешного выполнения идентификатор элемента будет предоставлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="aa147-1274">On success, the item identifier is provided in the `asyncResult.value` property.|</span></span>|

##### <a name="requirements"></a><span data-ttu-id="aa147-1275">Требования</span><span class="sxs-lookup"><span data-stu-id="aa147-1275">Requirements</span></span>

|<span data-ttu-id="aa147-1276">Requirement</span><span class="sxs-lookup"><span data-stu-id="aa147-1276">Requirement</span></span>|<span data-ttu-id="aa147-1277">Значение</span><span class="sxs-lookup"><span data-stu-id="aa147-1277">Value</span></span>|
|---|---|
|[<span data-ttu-id="aa147-1278">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="aa147-1278">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="aa147-1279">1.3</span><span class="sxs-lookup"><span data-stu-id="aa147-1279">1.3</span></span>|
|[<span data-ttu-id="aa147-1280">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="aa147-1280">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="aa147-1281">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="aa147-1281">ReadWriteItem</span></span>|
|[<span data-ttu-id="aa147-1282">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="aa147-1282">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="aa147-1283">Создание</span><span class="sxs-lookup"><span data-stu-id="aa147-1283">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="aa147-1284">Примеры</span><span class="sxs-lookup"><span data-stu-id="aa147-1284">Examples</span></span>

```js
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result
  });
```

<span data-ttu-id="aa147-p182">Ниже приведен пример параметра `result`, переданного функции обратного вызова. Свойство `value` содержит идентификатор элемента.</span><span class="sxs-lookup"><span data-stu-id="aa147-p182">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

####  <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="aa147-1287">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="aa147-1287">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="aa147-1288">Асинхронно вставляет данные в текст или тему сообщения.</span><span class="sxs-lookup"><span data-stu-id="aa147-1288">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="aa147-p183">Метод `setSelectedDataAsync` вставляет указанную строку в местоположение курсора в теме или тексте элемента либо, если текст выделен в редакторе, он заменяет выделенный текст. Если курсор находится вне текста или темы элемента, возвращается ошибка. После вставки курсор помещается в конец вставленного содержимого.</span><span class="sxs-lookup"><span data-stu-id="aa147-p183">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="aa147-1292">Параметры:</span><span class="sxs-lookup"><span data-stu-id="aa147-1292">Parameters:</span></span>

|<span data-ttu-id="aa147-1293">Имя</span><span class="sxs-lookup"><span data-stu-id="aa147-1293">Name</span></span>|<span data-ttu-id="aa147-1294">Тип</span><span class="sxs-lookup"><span data-stu-id="aa147-1294">Type</span></span>|<span data-ttu-id="aa147-1295">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="aa147-1295">Attributes</span></span>|<span data-ttu-id="aa147-1296">Описание</span><span class="sxs-lookup"><span data-stu-id="aa147-1296">Description</span></span>|
|---|---|---|---|
|`data`|<span data-ttu-id="aa147-1297">String</span><span class="sxs-lookup"><span data-stu-id="aa147-1297">String</span></span>||<span data-ttu-id="aa147-p184">Вставляемые данные. Объем данных не должен превышать 1 000 000 символов. Если передано больше 1 000 000 символов, возвращается исключение `ArgumentOutOfRange`.</span><span class="sxs-lookup"><span data-stu-id="aa147-p184">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`|<span data-ttu-id="aa147-1301">Object</span><span class="sxs-lookup"><span data-stu-id="aa147-1301">Object</span></span>|<span data-ttu-id="aa147-1302">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="aa147-1302">&lt;optional&gt;</span></span>|<span data-ttu-id="aa147-1303">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="aa147-1303">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="aa147-1304">Object</span><span class="sxs-lookup"><span data-stu-id="aa147-1304">Object</span></span>|<span data-ttu-id="aa147-1305">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="aa147-1305">&lt;optional&gt;</span></span>|<span data-ttu-id="aa147-1306">В методе обратного вызова разработчики могут указать любой объект, к которому необходимо получить доступ.</span><span class="sxs-lookup"><span data-stu-id="aa147-1306">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="aa147-1307">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="aa147-1307">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="aa147-1308">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="aa147-1308">&lt;optional&gt;</span></span>|<span data-ttu-id="aa147-p185">Если задано значение `text`, текущий стиль применяется в Outlook Web App и Outlook. Если поле представляет собой редактор HTML, вставляются только текстовые данные, даже если они имеют формат HTML.</span><span class="sxs-lookup"><span data-stu-id="aa147-p185">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="aa147-p186">Если задано значение `html` и поле (не тема) поддерживает HTML, в Outlook Web App применяется текущий стиль, а в Outlook — стиль по умолчанию. Если поле является текстовым, возвращается ошибка `InvalidDataFormat`.</span><span class="sxs-lookup"><span data-stu-id="aa147-p186">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="aa147-1313">Если свойство `coercionType` не задано, результат зависит от поля: если поле имеет формат HTML, используется текст в формате HTML, а если поле текстовое, применяется обычный текст.</span><span class="sxs-lookup"><span data-stu-id="aa147-1313">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`|<span data-ttu-id="aa147-1314">функция</span><span class="sxs-lookup"><span data-stu-id="aa147-1314">function</span></span>||<span data-ttu-id="aa147-1315">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="aa147-1315">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="aa147-1316">Требования</span><span class="sxs-lookup"><span data-stu-id="aa147-1316">Requirements</span></span>

|<span data-ttu-id="aa147-1317">Requirement</span><span class="sxs-lookup"><span data-stu-id="aa147-1317">Requirement</span></span>|<span data-ttu-id="aa147-1318">Значение</span><span class="sxs-lookup"><span data-stu-id="aa147-1318">Value</span></span>|
|---|---|
|[<span data-ttu-id="aa147-1319">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="aa147-1319">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="aa147-1320">1.2</span><span class="sxs-lookup"><span data-stu-id="aa147-1320">1.2</span></span>|
|[<span data-ttu-id="aa147-1321">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="aa147-1321">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="aa147-1322">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="aa147-1322">ReadWriteItem</span></span>|
|[<span data-ttu-id="aa147-1323">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="aa147-1323">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="aa147-1324">Создание</span><span class="sxs-lookup"><span data-stu-id="aa147-1324">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="aa147-1325">Пример</span><span class="sxs-lookup"><span data-stu-id="aa147-1325">Example</span></span>

```js
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```