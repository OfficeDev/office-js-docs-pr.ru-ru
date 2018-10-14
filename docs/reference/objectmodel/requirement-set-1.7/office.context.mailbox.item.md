
# <a name="item"></a><span data-ttu-id="91d32-101">item</span><span class="sxs-lookup"><span data-stu-id="91d32-101">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="91d32-102">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="91d32-102">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="91d32-p101">Пространство имен `item` используется для доступа к выбранному в данный момент сообщению, приглашению на собрание или описанию встречи. Вы можете определить тип пространства имен `item` с помощью свойства [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook17officemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="91d32-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook17officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="91d32-105">Требования</span><span class="sxs-lookup"><span data-stu-id="91d32-105">Requirements</span></span>

|<span data-ttu-id="91d32-106">Требование</span><span class="sxs-lookup"><span data-stu-id="91d32-106">Requirement</span></span>|<span data-ttu-id="91d32-107">Значение</span><span class="sxs-lookup"><span data-stu-id="91d32-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="91d32-108">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="91d32-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="91d32-109">1.0</span><span class="sxs-lookup"><span data-stu-id="91d32-109">1.0</span></span>|
|[<span data-ttu-id="91d32-110">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="91d32-110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="91d32-111">Ограниченный доступ</span><span class="sxs-lookup"><span data-stu-id="91d32-111">Restricted</span></span>|
|[<span data-ttu-id="91d32-112">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="91d32-112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="91d32-113">Compose или read</span><span class="sxs-lookup"><span data-stu-id="91d32-113">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="91d32-114">Члены и методы</span><span class="sxs-lookup"><span data-stu-id="91d32-114">Members and methods</span></span>

| <span data-ttu-id="91d32-115">Член</span><span class="sxs-lookup"><span data-stu-id="91d32-115">Member</span></span> | <span data-ttu-id="91d32-116">Тип</span><span class="sxs-lookup"><span data-stu-id="91d32-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="91d32-117">attachments</span><span class="sxs-lookup"><span data-stu-id="91d32-117">attachments</span></span>](#attachments-arrayattachmentdetailsjavascriptapioutlook17officeattachmentdetails) | <span data-ttu-id="91d32-118">Член</span><span class="sxs-lookup"><span data-stu-id="91d32-118">Member</span></span> |
| [<span data-ttu-id="91d32-119">bcc</span><span class="sxs-lookup"><span data-stu-id="91d32-119">bcc</span></span>](#bcc-recipientsjavascriptapioutlook17officerecipients) | <span data-ttu-id="91d32-120">Член</span><span class="sxs-lookup"><span data-stu-id="91d32-120">Member</span></span> |
| [<span data-ttu-id="91d32-121">body</span><span class="sxs-lookup"><span data-stu-id="91d32-121">body</span></span>](#body-bodyjavascriptapioutlook17officebody) | <span data-ttu-id="91d32-122">Член</span><span class="sxs-lookup"><span data-stu-id="91d32-122">Member</span></span> |
| [<span data-ttu-id="91d32-123">cc</span><span class="sxs-lookup"><span data-stu-id="91d32-123">cc</span></span>](#cc-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients) | <span data-ttu-id="91d32-124">Член</span><span class="sxs-lookup"><span data-stu-id="91d32-124">Member</span></span> |
| [<span data-ttu-id="91d32-125">conversationId</span><span class="sxs-lookup"><span data-stu-id="91d32-125">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="91d32-126">Член</span><span class="sxs-lookup"><span data-stu-id="91d32-126">Member</span></span> |
| [<span data-ttu-id="91d32-127">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="91d32-127">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="91d32-128">Член</span><span class="sxs-lookup"><span data-stu-id="91d32-128">Member</span></span> |
| [<span data-ttu-id="91d32-129">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="91d32-129">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="91d32-130">Член</span><span class="sxs-lookup"><span data-stu-id="91d32-130">Member</span></span> |
| [<span data-ttu-id="91d32-131">end</span><span class="sxs-lookup"><span data-stu-id="91d32-131">end</span></span>](#end-datetimejavascriptapioutlook17officetime) | <span data-ttu-id="91d32-132">Член</span><span class="sxs-lookup"><span data-stu-id="91d32-132">Member</span></span> |
| [<span data-ttu-id="91d32-133">from</span><span class="sxs-lookup"><span data-stu-id="91d32-133">from</span></span>](#from-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsfromjavascriptapioutlook17officefrom) | <span data-ttu-id="91d32-134">Член</span><span class="sxs-lookup"><span data-stu-id="91d32-134">Member</span></span> |
| [<span data-ttu-id="91d32-135">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="91d32-135">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="91d32-136">Член</span><span class="sxs-lookup"><span data-stu-id="91d32-136">Member</span></span> |
| [<span data-ttu-id="91d32-137">itemClass</span><span class="sxs-lookup"><span data-stu-id="91d32-137">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="91d32-138">Член</span><span class="sxs-lookup"><span data-stu-id="91d32-138">Member</span></span> |
| [<span data-ttu-id="91d32-139">itemId</span><span class="sxs-lookup"><span data-stu-id="91d32-139">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="91d32-140">Член</span><span class="sxs-lookup"><span data-stu-id="91d32-140">Member</span></span> |
| [<span data-ttu-id="91d32-141">itemType</span><span class="sxs-lookup"><span data-stu-id="91d32-141">itemType</span></span>](#itemtype-officemailboxenumsitemtypejavascriptapioutlook17officemailboxenumsitemtype) | <span data-ttu-id="91d32-142">Член</span><span class="sxs-lookup"><span data-stu-id="91d32-142">Member</span></span> |
| [<span data-ttu-id="91d32-143">location</span><span class="sxs-lookup"><span data-stu-id="91d32-143">location</span></span>](#location-stringlocationjavascriptapioutlook17officelocation) | <span data-ttu-id="91d32-144">Член</span><span class="sxs-lookup"><span data-stu-id="91d32-144">Member</span></span> |
| [<span data-ttu-id="91d32-145">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="91d32-145">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="91d32-146">Член</span><span class="sxs-lookup"><span data-stu-id="91d32-146">Member</span></span> |
| [<span data-ttu-id="91d32-147">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="91d32-147">notificationMessages</span></span>](#notificationmessages-notificationmessagesjavascriptapioutlook17officenotificationmessages) | <span data-ttu-id="91d32-148">Член</span><span class="sxs-lookup"><span data-stu-id="91d32-148">Member</span></span> |
| [<span data-ttu-id="91d32-149">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="91d32-149">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients) | <span data-ttu-id="91d32-150">Член</span><span class="sxs-lookup"><span data-stu-id="91d32-150">Member</span></span> |
| [<span data-ttu-id="91d32-151">organizer</span><span class="sxs-lookup"><span data-stu-id="91d32-151">organizer</span></span>](#organizer-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsorganizerjavascriptapioutlook17officeorganizer) | <span data-ttu-id="91d32-152">Член</span><span class="sxs-lookup"><span data-stu-id="91d32-152">Member</span></span> |
| [<span data-ttu-id="91d32-153">recurrence</span><span class="sxs-lookup"><span data-stu-id="91d32-153">recurrence</span></span>](#nullable-recurrence-recurrencejavascriptapioutlook17officerecurrence) | <span data-ttu-id="91d32-154">Член</span><span class="sxs-lookup"><span data-stu-id="91d32-154">Member</span></span> |
| [<span data-ttu-id="91d32-155">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="91d32-155">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients) | <span data-ttu-id="91d32-156">Член</span><span class="sxs-lookup"><span data-stu-id="91d32-156">Member</span></span> |
| [<span data-ttu-id="91d32-157">sender</span><span class="sxs-lookup"><span data-stu-id="91d32-157">sender</span></span>](#sender-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetails) | <span data-ttu-id="91d32-158">Член</span><span class="sxs-lookup"><span data-stu-id="91d32-158">Member</span></span> |
| [<span data-ttu-id="91d32-159">seriesId</span><span class="sxs-lookup"><span data-stu-id="91d32-159">seriesId</span></span>](#nullable-seriesid-string) | <span data-ttu-id="91d32-160">Член</span><span class="sxs-lookup"><span data-stu-id="91d32-160">Member</span></span> |
| [<span data-ttu-id="91d32-161">start</span><span class="sxs-lookup"><span data-stu-id="91d32-161">start</span></span>](#start-datetimejavascriptapioutlook17officetime) | <span data-ttu-id="91d32-162">Член</span><span class="sxs-lookup"><span data-stu-id="91d32-162">Member</span></span> |
| [<span data-ttu-id="91d32-163">subject</span><span class="sxs-lookup"><span data-stu-id="91d32-163">subject</span></span>](#subject-stringsubjectjavascriptapioutlook17officesubject) | <span data-ttu-id="91d32-164">Член</span><span class="sxs-lookup"><span data-stu-id="91d32-164">Member</span></span> |
| [<span data-ttu-id="91d32-165">to</span><span class="sxs-lookup"><span data-stu-id="91d32-165">to</span></span>](#to-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients) | <span data-ttu-id="91d32-166">Член</span><span class="sxs-lookup"><span data-stu-id="91d32-166">Member</span></span> |
| [<span data-ttu-id="91d32-167">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="91d32-167">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="91d32-168">Метод</span><span class="sxs-lookup"><span data-stu-id="91d32-168">Method</span></span> |
| [<span data-ttu-id="91d32-169">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="91d32-169">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="91d32-170">Метод</span><span class="sxs-lookup"><span data-stu-id="91d32-170">Method</span></span> |
| [<span data-ttu-id="91d32-171">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="91d32-171">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="91d32-172">Метод</span><span class="sxs-lookup"><span data-stu-id="91d32-172">Method</span></span> |
| [<span data-ttu-id="91d32-173">close</span><span class="sxs-lookup"><span data-stu-id="91d32-173">close</span></span>](#close) | <span data-ttu-id="91d32-174">Метод</span><span class="sxs-lookup"><span data-stu-id="91d32-174">Method</span></span> |
| [<span data-ttu-id="91d32-175">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="91d32-175">displayReplyAllForm</span></span>](#displayreplyallformformdata) | <span data-ttu-id="91d32-176">Метод</span><span class="sxs-lookup"><span data-stu-id="91d32-176">Method</span></span> |
| [<span data-ttu-id="91d32-177">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="91d32-177">displayReplyForm</span></span>](#displayreplyformformdata) | <span data-ttu-id="91d32-178">Метод</span><span class="sxs-lookup"><span data-stu-id="91d32-178">Method</span></span> |
| [<span data-ttu-id="91d32-179">getEntities</span><span class="sxs-lookup"><span data-stu-id="91d32-179">getEntities</span></span>](#getentities--entitiesjavascriptapioutlook17officeentities) | <span data-ttu-id="91d32-180">Метод</span><span class="sxs-lookup"><span data-stu-id="91d32-180">Method</span></span> |
| [<span data-ttu-id="91d32-181">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="91d32-181">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook17officecontactmeetingsuggestionjavascriptapioutlook17officemeetingsuggestionphonenumberjavascriptapioutlook17officephonenumbertasksuggestionjavascriptapioutlook17officetasksuggestion) | <span data-ttu-id="91d32-182">Метод</span><span class="sxs-lookup"><span data-stu-id="91d32-182">Method</span></span> |
| [<span data-ttu-id="91d32-183">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="91d32-183">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook17officecontactmeetingsuggestionjavascriptapioutlook17officemeetingsuggestionphonenumberjavascriptapioutlook17officephonenumbertasksuggestionjavascriptapioutlook17officetasksuggestion) | <span data-ttu-id="91d32-184">Метод</span><span class="sxs-lookup"><span data-stu-id="91d32-184">Method</span></span> |
| [<span data-ttu-id="91d32-185">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="91d32-185">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="91d32-186">Метод</span><span class="sxs-lookup"><span data-stu-id="91d32-186">Method</span></span> |
| [<span data-ttu-id="91d32-187">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="91d32-187">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="91d32-188">Метод</span><span class="sxs-lookup"><span data-stu-id="91d32-188">Method</span></span> |
| [<span data-ttu-id="91d32-189">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="91d32-189">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="91d32-190">Метод</span><span class="sxs-lookup"><span data-stu-id="91d32-190">Method</span></span> |
| [<span data-ttu-id="91d32-191">getSelectedEntities</span><span class="sxs-lookup"><span data-stu-id="91d32-191">getSelectedEntities</span></span>](#getselectedentities--entitiesjavascriptapioutlook17officeentities) | <span data-ttu-id="91d32-192">Метод</span><span class="sxs-lookup"><span data-stu-id="91d32-192">Method</span></span> |
| [<span data-ttu-id="91d32-193">getSelectedRegExMatches</span><span class="sxs-lookup"><span data-stu-id="91d32-193">getSelectedRegExMatches</span></span>](#getselectedregexmatches--object) | <span data-ttu-id="91d32-194">Метод</span><span class="sxs-lookup"><span data-stu-id="91d32-194">Method</span></span> |
| [<span data-ttu-id="91d32-195">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="91d32-195">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="91d32-196">Метод</span><span class="sxs-lookup"><span data-stu-id="91d32-196">Method</span></span> |
| [<span data-ttu-id="91d32-197">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="91d32-197">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="91d32-198">Метод</span><span class="sxs-lookup"><span data-stu-id="91d32-198">Method</span></span> |
| [<span data-ttu-id="91d32-199">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="91d32-199">removeHandlerAsync</span></span>](#removehandlerasynceventtype-handler-options-callback) | <span data-ttu-id="91d32-200">Метод</span><span class="sxs-lookup"><span data-stu-id="91d32-200">Method</span></span> |
| [<span data-ttu-id="91d32-201">saveAsync</span><span class="sxs-lookup"><span data-stu-id="91d32-201">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="91d32-202">Метод</span><span class="sxs-lookup"><span data-stu-id="91d32-202">Method</span></span> |
| [<span data-ttu-id="91d32-203">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="91d32-203">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="91d32-204">Метод</span><span class="sxs-lookup"><span data-stu-id="91d32-204">Method</span></span> |

### <a name="example"></a><span data-ttu-id="91d32-205">Пример</span><span class="sxs-lookup"><span data-stu-id="91d32-205">Example</span></span>

<span data-ttu-id="91d32-206">В приведенном ниже примере кода JavaScript показано, как получить доступ к свойству `subject` текущего элемента в Outlook.</span><span class="sxs-lookup"><span data-stu-id="91d32-206">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

```
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

### <a name="members"></a><span data-ttu-id="91d32-207">Члены</span><span class="sxs-lookup"><span data-stu-id="91d32-207">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook17officeattachmentdetails"></a><span data-ttu-id="91d32-208">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_7/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="91d32-208">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_7/office.attachmentdetails)></span></span>

<span data-ttu-id="91d32-p102">Получает массив вложений для элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="91d32-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="91d32-211">Некоторые типы файлов блокируются Outlook из-за потенциальных проблем безопасности и поэтому не возвращаются.</span><span class="sxs-lookup"><span data-stu-id="91d32-211">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="91d32-212">Дополнительные сведения см. в статье [Блокированные вложения в Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="91d32-212">For more information see [Payments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="91d32-213">Тип:</span><span class="sxs-lookup"><span data-stu-id="91d32-213">Type:</span></span>

*   <span data-ttu-id="91d32-214">Array.<[AttachmentDetails](/javascript/api/outlook_1_7/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="91d32-214">Array.<[AttachmentDetails](/javascript/api/outlook_1_7/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="91d32-215">Требования</span><span class="sxs-lookup"><span data-stu-id="91d32-215">Requirements</span></span>

|<span data-ttu-id="91d32-216">Требование</span><span class="sxs-lookup"><span data-stu-id="91d32-216">Requirement</span></span>|<span data-ttu-id="91d32-217">Значение</span><span class="sxs-lookup"><span data-stu-id="91d32-217">Value</span></span>|
|---|---|
|[<span data-ttu-id="91d32-218">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="91d32-218">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="91d32-219">1.0</span><span class="sxs-lookup"><span data-stu-id="91d32-219">1.0</span></span>|
|[<span data-ttu-id="91d32-220">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="91d32-220">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="91d32-221">ReadItem</span><span class="sxs-lookup"><span data-stu-id="91d32-221">ReadItem</span></span>|
|[<span data-ttu-id="91d32-222">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="91d32-222">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="91d32-223">Чтение</span><span class="sxs-lookup"><span data-stu-id="91d32-223">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="91d32-224">Пример</span><span class="sxs-lookup"><span data-stu-id="91d32-224">Example</span></span>

<span data-ttu-id="91d32-225">С помощью приведенного ниже кода можно создать HTML-строку с подробными сведениями обо всех вложениях для текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="91d32-225">The following code builds an HTML string with details of all attachments on the current item.</span></span>

```
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

####  <a name="bcc-recipientsjavascriptapioutlook17officerecipients"></a><span data-ttu-id="91d32-226">bcc :[Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="91d32-226">bcc :[Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

<span data-ttu-id="91d32-227">Получает объект, который предоставляет методы для получения или обновления получателей в строке Bcc (скрытой копии) сообщения.</span><span class="sxs-lookup"><span data-stu-id="91d32-227">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="91d32-228">Только в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="91d32-228">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="91d32-229">Тип:</span><span class="sxs-lookup"><span data-stu-id="91d32-229">Type:</span></span>

*   [<span data-ttu-id="91d32-230">Recipients</span><span class="sxs-lookup"><span data-stu-id="91d32-230">Recipients</span></span>](/javascript/api/outlook_1_7/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="91d32-231">Требования</span><span class="sxs-lookup"><span data-stu-id="91d32-231">Requirements</span></span>

|<span data-ttu-id="91d32-232">Требование</span><span class="sxs-lookup"><span data-stu-id="91d32-232">Requirement</span></span>|<span data-ttu-id="91d32-233">Значение</span><span class="sxs-lookup"><span data-stu-id="91d32-233">Value</span></span>|
|---|---|
|[<span data-ttu-id="91d32-234">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="91d32-234">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="91d32-235">1.1</span><span class="sxs-lookup"><span data-stu-id="91d32-235">1.1</span></span>|
|[<span data-ttu-id="91d32-236">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="91d32-236">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="91d32-237">ReadItem</span><span class="sxs-lookup"><span data-stu-id="91d32-237">ReadItem</span></span>|
|[<span data-ttu-id="91d32-238">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="91d32-238">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="91d32-239">Создание</span><span class="sxs-lookup"><span data-stu-id="91d32-239">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="91d32-240">Пример</span><span class="sxs-lookup"><span data-stu-id="91d32-240">Example</span></span>

```
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlook17officebody"></a><span data-ttu-id="91d32-241">body :[Body](/javascript/api/outlook_1_7/office.body)</span><span class="sxs-lookup"><span data-stu-id="91d32-241">body :[Body](/javascript/api/outlook_1_7/office.body)</span></span>

<span data-ttu-id="91d32-242">Получает объект, предоставляющий методы для работы с основным текстом элемента.</span><span class="sxs-lookup"><span data-stu-id="91d32-242">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="91d32-243">Тип:</span><span class="sxs-lookup"><span data-stu-id="91d32-243">Type:</span></span>

*   [<span data-ttu-id="91d32-244">Body</span><span class="sxs-lookup"><span data-stu-id="91d32-244">Body</span></span>](/javascript/api/outlook_1_7/office.body)

##### <a name="requirements"></a><span data-ttu-id="91d32-245">Требования</span><span class="sxs-lookup"><span data-stu-id="91d32-245">Requirements</span></span>

|<span data-ttu-id="91d32-246">Требование</span><span class="sxs-lookup"><span data-stu-id="91d32-246">Requirement</span></span>|<span data-ttu-id="91d32-247">Значение</span><span class="sxs-lookup"><span data-stu-id="91d32-247">Value</span></span>|
|---|---|
|[<span data-ttu-id="91d32-248">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="91d32-248">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="91d32-249">1.1</span><span class="sxs-lookup"><span data-stu-id="91d32-249">1.1</span></span>|
|[<span data-ttu-id="91d32-250">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="91d32-250">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="91d32-251">ReadItem</span><span class="sxs-lookup"><span data-stu-id="91d32-251">ReadItem</span></span>|
|[<span data-ttu-id="91d32-252">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="91d32-252">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="91d32-253">Compose (создание) или read (чтение)</span><span class="sxs-lookup"><span data-stu-id="91d32-253">Compose or read</span></span>|

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients"></a><span data-ttu-id="91d32-254">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="91d32-254">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

<span data-ttu-id="91d32-255">Предоставляет доступ к получателям Cc (копии) сообщения.</span><span class="sxs-lookup"><span data-stu-id="91d32-255">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="91d32-256">Уровень доступа и тип объекта, зависит от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="91d32-256">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="91d32-257">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="91d32-257">Read mode</span></span>

<span data-ttu-id="91d32-p106">Свойство `cc` возвращает массив, который содержит объект `EmailAddressDetails`, каждому получателю, указанному в строке **Cc (копия)** сообщения. Коллекция может включать не более 100 членов.</span><span class="sxs-lookup"><span data-stu-id="91d32-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="91d32-260">Режим создания</span><span class="sxs-lookup"><span data-stu-id="91d32-260">Compose mode</span></span>

<span data-ttu-id="91d32-261">Свойство `cc` возвращает объект `Recipients`, предоставляющий методы для получения или обновления получателей, которые указаны в строке **Cc (копия)** сообщения.</span><span class="sxs-lookup"><span data-stu-id="91d32-261">The `cc` property returns a `Recipients` object that provides methods for manipulating the recipients on the **Cc** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="91d32-262">Тип:</span><span class="sxs-lookup"><span data-stu-id="91d32-262">Type:</span></span>

*   <span data-ttu-id="91d32-263">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="91d32-263">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="91d32-264">Требования</span><span class="sxs-lookup"><span data-stu-id="91d32-264">Requirements</span></span>

|<span data-ttu-id="91d32-265">Требование</span><span class="sxs-lookup"><span data-stu-id="91d32-265">Requirement</span></span>|<span data-ttu-id="91d32-266">Значение</span><span class="sxs-lookup"><span data-stu-id="91d32-266">Value</span></span>|
|---|---|
|[<span data-ttu-id="91d32-267">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="91d32-267">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="91d32-268">1.0</span><span class="sxs-lookup"><span data-stu-id="91d32-268">1.0</span></span>|
|[<span data-ttu-id="91d32-269">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="91d32-269">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="91d32-270">ReadItem</span><span class="sxs-lookup"><span data-stu-id="91d32-270">ReadItem</span></span>|
|[<span data-ttu-id="91d32-271">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="91d32-271">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="91d32-272">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="91d32-272">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="91d32-273">Пример</span><span class="sxs-lookup"><span data-stu-id="91d32-273">Example</span></span>

```
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="91d32-274">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="91d32-274">(nullable) conversationId :String</span></span>

<span data-ttu-id="91d32-275">Получает идентификатор разговора по электронной почте, содержащего конкретное сообщение.</span><span class="sxs-lookup"><span data-stu-id="91d32-275">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="91d32-p107">Вы можете получить целочисленное значение этого свойства, если ваше почтовое приложение активируется в формах просмотра или формах создания ответов. Если пользователь в свою очередь изменит тему ответа, после его отправки идентификатор беседы будет изменен, и полученное ранее значение будет недействительным.</span><span class="sxs-lookup"><span data-stu-id="91d32-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="91d32-p108">Это свойство имеет значение NULL для нового элемента в форме создания. Свойство `conversationId` вернет значение, если пользователь задаст тему и сохранит элемент.</span><span class="sxs-lookup"><span data-stu-id="91d32-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="91d32-280">Тип:</span><span class="sxs-lookup"><span data-stu-id="91d32-280">Type:</span></span>

*   <span data-ttu-id="91d32-281">String</span><span class="sxs-lookup"><span data-stu-id="91d32-281">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="91d32-282">Требования</span><span class="sxs-lookup"><span data-stu-id="91d32-282">Requirements</span></span>

|<span data-ttu-id="91d32-283">Требование</span><span class="sxs-lookup"><span data-stu-id="91d32-283">Requirement</span></span>|<span data-ttu-id="91d32-284">Значение</span><span class="sxs-lookup"><span data-stu-id="91d32-284">Value</span></span>|
|---|---|
|[<span data-ttu-id="91d32-285">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="91d32-285">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="91d32-286">1.0</span><span class="sxs-lookup"><span data-stu-id="91d32-286">1.0</span></span>|
|[<span data-ttu-id="91d32-287">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="91d32-287">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="91d32-288">ReadItem</span><span class="sxs-lookup"><span data-stu-id="91d32-288">ReadItem</span></span>|
|[<span data-ttu-id="91d32-289">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="91d32-289">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="91d32-290">Compose (создание) или read (чтение)</span><span class="sxs-lookup"><span data-stu-id="91d32-290">Compose or read</span></span>|

#### <a name="datetimecreated-date"></a><span data-ttu-id="91d32-291">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="91d32-291">dateTimeCreated :Date</span></span>

<span data-ttu-id="91d32-p109">Получает дату и время создания элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="91d32-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="91d32-294">Тип:</span><span class="sxs-lookup"><span data-stu-id="91d32-294">Type:</span></span>

*   <span data-ttu-id="91d32-295">Date</span><span class="sxs-lookup"><span data-stu-id="91d32-295">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="91d32-296">Требования</span><span class="sxs-lookup"><span data-stu-id="91d32-296">Requirements</span></span>

|<span data-ttu-id="91d32-297">Требование</span><span class="sxs-lookup"><span data-stu-id="91d32-297">Requirement</span></span>|<span data-ttu-id="91d32-298">Значение</span><span class="sxs-lookup"><span data-stu-id="91d32-298">Value</span></span>|
|---|---|
|[<span data-ttu-id="91d32-299">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="91d32-299">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="91d32-300">1.0</span><span class="sxs-lookup"><span data-stu-id="91d32-300">1.0</span></span>|
|[<span data-ttu-id="91d32-301">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="91d32-301">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="91d32-302">ReadItem</span><span class="sxs-lookup"><span data-stu-id="91d32-302">ReadItem</span></span>|
|[<span data-ttu-id="91d32-303">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="91d32-303">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="91d32-304">Чтение</span><span class="sxs-lookup"><span data-stu-id="91d32-304">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="91d32-305">Пример</span><span class="sxs-lookup"><span data-stu-id="91d32-305">Example</span></span>

```
var created = Office.context.mailbox.item.dateTimeCreated;
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="91d32-306">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="91d32-306">dateTimeModified :Date</span></span>

<span data-ttu-id="91d32-p110">Получает дату и время последнего изменения элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="91d32-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="91d32-309">Этот элемент не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="91d32-309">Note: This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="91d32-310">Тип:</span><span class="sxs-lookup"><span data-stu-id="91d32-310">Type:</span></span>

*   <span data-ttu-id="91d32-311">Date</span><span class="sxs-lookup"><span data-stu-id="91d32-311">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="91d32-312">Требования</span><span class="sxs-lookup"><span data-stu-id="91d32-312">Requirements</span></span>

|<span data-ttu-id="91d32-313">Требование</span><span class="sxs-lookup"><span data-stu-id="91d32-313">Requirement</span></span>|<span data-ttu-id="91d32-314">Значение</span><span class="sxs-lookup"><span data-stu-id="91d32-314">Value</span></span>|
|---|---|
|[<span data-ttu-id="91d32-315">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="91d32-315">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="91d32-316">1.0</span><span class="sxs-lookup"><span data-stu-id="91d32-316">1.0</span></span>|
|[<span data-ttu-id="91d32-317">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="91d32-317">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="91d32-318">ReadItem</span><span class="sxs-lookup"><span data-stu-id="91d32-318">ReadItem</span></span>|
|[<span data-ttu-id="91d32-319">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="91d32-319">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="91d32-320">Чтение</span><span class="sxs-lookup"><span data-stu-id="91d32-320">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="91d32-321">Пример</span><span class="sxs-lookup"><span data-stu-id="91d32-321">Example</span></span>

```
var modified = Office.context.mailbox.item.dateTimeModified;
```

####  <a name="end-datetimejavascriptapioutlook17officetime"></a><span data-ttu-id="91d32-322">end :Date|[Time](/javascript/api/outlook_1_7/office.time)</span><span class="sxs-lookup"><span data-stu-id="91d32-322">end :Date|[Time](/javascript/api/outlook_1_7/office.time)</span></span>

<span data-ttu-id="91d32-323">Получает или задает дату и время окончания встречи.</span><span class="sxs-lookup"><span data-stu-id="91d32-323">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="91d32-p111">Свойство `end` представлено в виде значения даты и времени в формате UTC. Преобразовать значение свойства end в местные значения даты и времени клиента можно с помощью метода [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook17officelocalclienttime).</span><span class="sxs-lookup"><span data-stu-id="91d32-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook17officelocalclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="91d32-326">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="91d32-326">Read mode</span></span>

<span data-ttu-id="91d32-327">Свойство `end` возвращает объект `Date`.</span><span class="sxs-lookup"><span data-stu-id="91d32-327">The `end` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="91d32-328">Режим создания</span><span class="sxs-lookup"><span data-stu-id="91d32-328">Compose mode</span></span>

<span data-ttu-id="91d32-329">Свойство `end` возвращает объект `Time`.</span><span class="sxs-lookup"><span data-stu-id="91d32-329">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="91d32-330">Когда вы используете метод [`Time.setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) для того, чтобы задать время окончания, вы должны использовать метод [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) , чтобы преобразовать местное время на клиенте в формат UTC.</span><span class="sxs-lookup"><span data-stu-id="91d32-330">When you use the [`Time.setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="91d32-331">Тип:</span><span class="sxs-lookup"><span data-stu-id="91d32-331">Type:</span></span>

*   <span data-ttu-id="91d32-332">Date | [Time](/javascript/api/outlook_1_7/office.time)</span><span class="sxs-lookup"><span data-stu-id="91d32-332">Date | [Time](/javascript/api/outlook_1_7/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="91d32-333">Требования</span><span class="sxs-lookup"><span data-stu-id="91d32-333">Requirements</span></span>

|<span data-ttu-id="91d32-334">Требование</span><span class="sxs-lookup"><span data-stu-id="91d32-334">Requirement</span></span>|<span data-ttu-id="91d32-335">Значение</span><span class="sxs-lookup"><span data-stu-id="91d32-335">Value</span></span>|
|---|---|
|[<span data-ttu-id="91d32-336">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="91d32-336">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="91d32-337">1.0</span><span class="sxs-lookup"><span data-stu-id="91d32-337">1.0</span></span>|
|[<span data-ttu-id="91d32-338">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="91d32-338">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="91d32-339">ReadItem</span><span class="sxs-lookup"><span data-stu-id="91d32-339">ReadItem</span></span>|
|[<span data-ttu-id="91d32-340">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="91d32-340">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="91d32-341">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="91d32-341">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="91d32-342">Пример</span><span class="sxs-lookup"><span data-stu-id="91d32-342">Example</span></span>

<span data-ttu-id="91d32-343">В примере ниже показано, как с помощью метода [`setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) объекта `Time` задать время окончания встречи в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="91d32-343">The following example sets the end time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

```
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

#### <a name="from-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsfromjavascriptapioutlook17officefrom"></a><span data-ttu-id="91d32-344">from :[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)|[From](/javascript/api/outlook_1_7/office.from)</span><span class="sxs-lookup"><span data-stu-id="91d32-344">from :[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)|[From](/javascript/api/outlook_1_7/office.from)</span></span>

<span data-ttu-id="91d32-345">Получает адрес электронной почты отправителя сообщения.</span><span class="sxs-lookup"><span data-stu-id="91d32-345">Gets the email address of the sender of a message.</span></span>

<span data-ttu-id="91d32-p112">Свойства `from` и [`sender`](#sender-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetails) представляют одно лицо, пока сообщение не будет отправлено представителем. В противном случае свойство `from` представляет лицо, делегировавшее полномочия, а свойство sender — представителя.</span><span class="sxs-lookup"><span data-stu-id="91d32-p112">The `from` and [`sender`](#sender-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="91d32-348">Свойство `recipientType` объекта `EmailAddressDetails` в свойстве `from` — `undefined`.</span><span class="sxs-lookup"><span data-stu-id="91d32-348">Note: The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="91d32-349">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="91d32-349">Read mode</span></span>

<span data-ttu-id="91d32-350">Свойство `from` возвращает объект `EmailAddressDetails`.</span><span class="sxs-lookup"><span data-stu-id="91d32-350">The `from` property returns a `EmailAddressDetails` object.</span></span>

```
var subject = Office.context.mailbox.item.from;
```

##### <a name="compose-mode"></a><span data-ttu-id="91d32-351">Режим создания</span><span class="sxs-lookup"><span data-stu-id="91d32-351">Compose mode</span></span>

<span data-ttu-id="91d32-352">Свойство `from` возвращает объект `From`, который обеспечивает метод получения объекта из значения.</span><span class="sxs-lookup"><span data-stu-id="91d32-352">The `from` property returns a `From` object that provides a method to get the from value.</span></span>

```
Office.context.mailbox.item.from.getAsync(callback);

function callback(asyncResult) {
  var from = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="91d32-353">Тип:</span><span class="sxs-lookup"><span data-stu-id="91d32-353">Type:</span></span>

*   <span data-ttu-id="91d32-354">[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails) | [From](/javascript/api/outlook_1_7/office.from)</span><span class="sxs-lookup"><span data-stu-id="91d32-354">[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails) | [From](/javascript/api/outlook_1_7/office.from)</span></span>

##### <a name="requirements"></a><span data-ttu-id="91d32-355">Требования</span><span class="sxs-lookup"><span data-stu-id="91d32-355">Requirements</span></span>

|<span data-ttu-id="91d32-356">Требование</span><span class="sxs-lookup"><span data-stu-id="91d32-356">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="91d32-357">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="91d32-357">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="91d32-358">1.0</span><span class="sxs-lookup"><span data-stu-id="91d32-358">1.0</span></span>|<span data-ttu-id="91d32-359">1.7</span><span class="sxs-lookup"><span data-stu-id="91d32-359">17 </span></span>|
|[<span data-ttu-id="91d32-360">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="91d32-360">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="91d32-361">ReadItem</span><span class="sxs-lookup"><span data-stu-id="91d32-361">ReadItem</span></span>|<span data-ttu-id="91d32-362">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="91d32-362">ReadWriteItem</span></span>|
|[<span data-ttu-id="91d32-363">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="91d32-363">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="91d32-364">Чтение</span><span class="sxs-lookup"><span data-stu-id="91d32-364">Read</span></span>|<span data-ttu-id="91d32-365">Compose</span><span class="sxs-lookup"><span data-stu-id="91d32-365">Compose</span></span>|

#### <a name="internetmessageid-string"></a><span data-ttu-id="91d32-366">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="91d32-366">internetMessageId :String</span></span>

<span data-ttu-id="91d32-p113">Получает идентификатор интернет-сообщения для электронного сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="91d32-p113">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="91d32-369">Тип:</span><span class="sxs-lookup"><span data-stu-id="91d32-369">Type:</span></span>

*   <span data-ttu-id="91d32-370">String</span><span class="sxs-lookup"><span data-stu-id="91d32-370">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="91d32-371">Требования</span><span class="sxs-lookup"><span data-stu-id="91d32-371">Requirements</span></span>

|<span data-ttu-id="91d32-372">Требование</span><span class="sxs-lookup"><span data-stu-id="91d32-372">Requirement</span></span>|<span data-ttu-id="91d32-373">Значение</span><span class="sxs-lookup"><span data-stu-id="91d32-373">Value</span></span>|
|---|---|
|[<span data-ttu-id="91d32-374">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="91d32-374">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="91d32-375">1.0</span><span class="sxs-lookup"><span data-stu-id="91d32-375">1.0</span></span>|
|[<span data-ttu-id="91d32-376">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="91d32-376">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="91d32-377">ReadItem</span><span class="sxs-lookup"><span data-stu-id="91d32-377">ReadItem</span></span>|
|[<span data-ttu-id="91d32-378">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="91d32-378">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="91d32-379">Чтение</span><span class="sxs-lookup"><span data-stu-id="91d32-379">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="91d32-380">Пример</span><span class="sxs-lookup"><span data-stu-id="91d32-380">Example</span></span>

```
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="91d32-381">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="91d32-381">itemClass :String</span></span>

<span data-ttu-id="91d32-p114">Получает класс элемента веб-служб Exchange для выбранного элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="91d32-p114">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="91d32-p115">Свойство `itemClass` указывает класс сообщения выбранного элемента. Ниже приводятся классы сообщения по умолчанию для элемента сообщения или встречи.</span><span class="sxs-lookup"><span data-stu-id="91d32-p115">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

|<span data-ttu-id="91d32-386">Тип</span><span class="sxs-lookup"><span data-stu-id="91d32-386">Type</span></span>|<span data-ttu-id="91d32-387">Описание</span><span class="sxs-lookup"><span data-stu-id="91d32-387">Description</span></span>|<span data-ttu-id="91d32-388">item class</span><span class="sxs-lookup"><span data-stu-id="91d32-388">item class</span></span>|
|---|---|---|
|<span data-ttu-id="91d32-389">Элементы встречи</span><span class="sxs-lookup"><span data-stu-id="91d32-389">Appointment items</span></span>|<span data-ttu-id="91d32-390">Это элементы календаря для класса элемента `IPM.Appointment` или `IPM.Appointment.Occurence`.</span><span class="sxs-lookup"><span data-stu-id="91d32-390">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurence`.</span></span>|`IPM.Appointment`<br />`IPM.Appointment.Occurence`|
|<span data-ttu-id="91d32-391">Элементы сообщения</span><span class="sxs-lookup"><span data-stu-id="91d32-391">Message items</span></span>|<span data-ttu-id="91d32-392">Сюда входят электронные сообщения, для которых по умолчанию задан класс сообщения `IPM.Note`, а также приглашения на собрания, ответы на них и уведомления об их отмене, использующие `IPM.Schedule.Meeting` в качестве базового класса сообщений.</span><span class="sxs-lookup"><span data-stu-id="91d32-392">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span>|`IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled`|

<span data-ttu-id="91d32-393">Вы можете создавать настраиваемые классы сообщения, расширяющие классы сообщения по умолчанию, например, настраиваемый класс сообщения о встрече `IPM.Appointment.Contoso`.</span><span class="sxs-lookup"><span data-stu-id="91d32-393">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="91d32-394">Тип:</span><span class="sxs-lookup"><span data-stu-id="91d32-394">Type:</span></span>

*   <span data-ttu-id="91d32-395">String</span><span class="sxs-lookup"><span data-stu-id="91d32-395">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="91d32-396">Требования</span><span class="sxs-lookup"><span data-stu-id="91d32-396">Requirements</span></span>

|<span data-ttu-id="91d32-397">Требование</span><span class="sxs-lookup"><span data-stu-id="91d32-397">Requirement</span></span>|<span data-ttu-id="91d32-398">Значение</span><span class="sxs-lookup"><span data-stu-id="91d32-398">Value</span></span>|
|---|---|
|[<span data-ttu-id="91d32-399">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="91d32-399">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="91d32-400">1.0</span><span class="sxs-lookup"><span data-stu-id="91d32-400">1.0</span></span>|
|[<span data-ttu-id="91d32-401">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="91d32-401">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="91d32-402">ReadItem</span><span class="sxs-lookup"><span data-stu-id="91d32-402">ReadItem</span></span>|
|[<span data-ttu-id="91d32-403">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="91d32-403">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="91d32-404">Чтение</span><span class="sxs-lookup"><span data-stu-id="91d32-404">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="91d32-405">Пример</span><span class="sxs-lookup"><span data-stu-id="91d32-405">Example</span></span>

```
var itemClass = Office.context.mailbox.item.itemClass;
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="91d32-406">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="91d32-406">(nullable) itemId :String</span></span>

<span data-ttu-id="91d32-p116">Получает идентификатор элемента веб-служб Exchange для текущего элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="91d32-p116">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="91d32-409">Идентификатор, возвращаемый свойством `itemId`, совпадает с идентификатором элемента веб-служб Exchange.</span><span class="sxs-lookup"><span data-stu-id="91d32-409">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="91d32-410">Свойство  `itemId` не совпадает с идентификатором записи Outlook или идентификатором, используемым API-Интерфейсом REST Outlook.</span><span class="sxs-lookup"><span data-stu-id="91d32-410">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="91d32-411">Прежде чем осуществлять вызовы API-Интерфейса REST с помощью этого значения, его следует преобразовать с помощью [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="91d32-411">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="91d32-412">Дополнительные сведения см. в статье [Использование API REST для Outlook из надстройки Outlook](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="91d32-412">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="91d32-p118">Свойство `itemId` недоступно в режиме создания. Если требуется идентификатор элемента, с помощью метода [`saveAsync`](#saveasyncoptions-callback) можно сохранить элемент в хранилище. При этом в параметре [`AsyncResult.value`](/javascript/api/office/office.asyncresult) функции обратного вызова возвращается идентификатор элемента.</span><span class="sxs-lookup"><span data-stu-id="91d32-p118">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="91d32-415">Тип:</span><span class="sxs-lookup"><span data-stu-id="91d32-415">Type:</span></span>

*   <span data-ttu-id="91d32-416">String</span><span class="sxs-lookup"><span data-stu-id="91d32-416">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="91d32-417">Требования</span><span class="sxs-lookup"><span data-stu-id="91d32-417">Requirements</span></span>

|<span data-ttu-id="91d32-418">Требование</span><span class="sxs-lookup"><span data-stu-id="91d32-418">Requirement</span></span>|<span data-ttu-id="91d32-419">Значение</span><span class="sxs-lookup"><span data-stu-id="91d32-419">Value</span></span>|
|---|---|
|[<span data-ttu-id="91d32-420">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="91d32-420">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="91d32-421">1.0</span><span class="sxs-lookup"><span data-stu-id="91d32-421">1.0</span></span>|
|[<span data-ttu-id="91d32-422">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="91d32-422">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="91d32-423">ReadItem</span><span class="sxs-lookup"><span data-stu-id="91d32-423">ReadItem</span></span>|
|[<span data-ttu-id="91d32-424">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="91d32-424">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="91d32-425">Чтение</span><span class="sxs-lookup"><span data-stu-id="91d32-425">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="91d32-426">Пример</span><span class="sxs-lookup"><span data-stu-id="91d32-426">Example</span></span>

<span data-ttu-id="91d32-p119">Указанный ниже код проверяет наличие идентификатора элемента. Если свойство `itemId` возвращает значение `null` или `undefined`, элемент будет сохранен в хранилище, а из асинхронного результата будет получен идентификатор элемента.</span><span class="sxs-lookup"><span data-stu-id="91d32-p119">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result){
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook17officemailboxenumsitemtype"></a><span data-ttu-id="91d32-429">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_7/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="91d32-429">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_7/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="91d32-430">Получает тип элемента, который представляет экземпляр.</span><span class="sxs-lookup"><span data-stu-id="91d32-430">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="91d32-431">Свойство `itemType` возвращает одно из значений перечисления `ItemType`, которое указывает, является ли экземпляр объекта `item` сообщением или встречей.</span><span class="sxs-lookup"><span data-stu-id="91d32-431">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="91d32-432">Тип:</span><span class="sxs-lookup"><span data-stu-id="91d32-432">Type:</span></span>

*   [<span data-ttu-id="91d32-433">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="91d32-433">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook_1_7/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="91d32-434">Требования</span><span class="sxs-lookup"><span data-stu-id="91d32-434">Requirements</span></span>

|<span data-ttu-id="91d32-435">Требование</span><span class="sxs-lookup"><span data-stu-id="91d32-435">Requirement</span></span>|<span data-ttu-id="91d32-436">Значение</span><span class="sxs-lookup"><span data-stu-id="91d32-436">Value</span></span>|
|---|---|
|[<span data-ttu-id="91d32-437">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="91d32-437">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="91d32-438">1.0</span><span class="sxs-lookup"><span data-stu-id="91d32-438">1.0</span></span>|
|[<span data-ttu-id="91d32-439">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="91d32-439">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="91d32-440">ReadItem</span><span class="sxs-lookup"><span data-stu-id="91d32-440">ReadItem</span></span>|
|[<span data-ttu-id="91d32-441">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="91d32-441">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="91d32-442">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="91d32-442">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="91d32-443">Пример</span><span class="sxs-lookup"><span data-stu-id="91d32-443">Example</span></span>

```
if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Message)
  // do something
else
  // do something else
```

####  <a name="location-stringlocationjavascriptapioutlook17officelocation"></a><span data-ttu-id="91d32-444">location :String|[Location](/javascript/api/outlook_1_7/office.location)</span><span class="sxs-lookup"><span data-stu-id="91d32-444">location :String|[Location](/javascript/api/outlook_1_7/office.location)</span></span>

<span data-ttu-id="91d32-445">Получает или задает место встречи.</span><span class="sxs-lookup"><span data-stu-id="91d32-445">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="91d32-446">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="91d32-446">Read mode</span></span>

<span data-ttu-id="91d32-447">Свойство `location` возвращает строку, содержащую сведения о месте встречи.</span><span class="sxs-lookup"><span data-stu-id="91d32-447">The `location` property returns a string that contains the location of the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="91d32-448">Режим создания</span><span class="sxs-lookup"><span data-stu-id="91d32-448">Compose mode</span></span>

<span data-ttu-id="91d32-449">Свойство `location` возвращает объект `Location`, предоставляющий методы, которые используются для получения и задания места встречи.</span><span class="sxs-lookup"><span data-stu-id="91d32-449">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="91d32-450">Тип:</span><span class="sxs-lookup"><span data-stu-id="91d32-450">Type:</span></span>

*   <span data-ttu-id="91d32-451">String | [Location](/javascript/api/outlook_1_7/office.location)</span><span class="sxs-lookup"><span data-stu-id="91d32-451">String | [Location](/javascript/api/outlook_1_7/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="91d32-452">Требования</span><span class="sxs-lookup"><span data-stu-id="91d32-452">Requirements</span></span>

|<span data-ttu-id="91d32-453">Требование</span><span class="sxs-lookup"><span data-stu-id="91d32-453">Requirement</span></span>|<span data-ttu-id="91d32-454">Значение</span><span class="sxs-lookup"><span data-stu-id="91d32-454">Value</span></span>|
|---|---|
|[<span data-ttu-id="91d32-455">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="91d32-455">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="91d32-456">1.0</span><span class="sxs-lookup"><span data-stu-id="91d32-456">1.0</span></span>|
|[<span data-ttu-id="91d32-457">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="91d32-457">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="91d32-458">ReadItem</span><span class="sxs-lookup"><span data-stu-id="91d32-458">ReadItem</span></span>|
|[<span data-ttu-id="91d32-459">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="91d32-459">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="91d32-460">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="91d32-460">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="91d32-461">Пример</span><span class="sxs-lookup"><span data-stu-id="91d32-461">Example</span></span>

```
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

#### <a name="normalizedsubject-string"></a><span data-ttu-id="91d32-462">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="91d32-462">normalizedSubject :String</span></span>

<span data-ttu-id="91d32-p120">Получает тему элемента со всеми удаленными префиксами (включая `RE:` и `FWD:`). Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="91d32-p120">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="91d32-p121">Свойство normalizedSubject получает тему элемента со стандартными префиксами (такими как `RE:` и `FW:`), добавляемыми почтовыми программами. Для получения темы элемента с неизмененными префиксами используйте свойство [`subject`](#subject-stringsubjectjavascriptapioutlook17officesubject).</span><span class="sxs-lookup"><span data-stu-id="91d32-p121">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubjectjavascriptapioutlook17officesubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="91d32-467">Тип:</span><span class="sxs-lookup"><span data-stu-id="91d32-467">Type:</span></span>

*   <span data-ttu-id="91d32-468">String</span><span class="sxs-lookup"><span data-stu-id="91d32-468">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="91d32-469">Требования</span><span class="sxs-lookup"><span data-stu-id="91d32-469">Requirements</span></span>

|<span data-ttu-id="91d32-470">Требование</span><span class="sxs-lookup"><span data-stu-id="91d32-470">Requirement</span></span>|<span data-ttu-id="91d32-471">Значение</span><span class="sxs-lookup"><span data-stu-id="91d32-471">Value</span></span>|
|---|---|
|[<span data-ttu-id="91d32-472">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="91d32-472">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="91d32-473">1.0</span><span class="sxs-lookup"><span data-stu-id="91d32-473">1.0</span></span>|
|[<span data-ttu-id="91d32-474">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="91d32-474">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="91d32-475">ReadItem</span><span class="sxs-lookup"><span data-stu-id="91d32-475">ReadItem</span></span>|
|[<span data-ttu-id="91d32-476">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="91d32-476">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="91d32-477">Чтение</span><span class="sxs-lookup"><span data-stu-id="91d32-477">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="91d32-478">Пример</span><span class="sxs-lookup"><span data-stu-id="91d32-478">Example</span></span>

```
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
```

####  <a name="notificationmessages-notificationmessagesjavascriptapioutlook17officenotificationmessages"></a><span data-ttu-id="91d32-479">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_7/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="91d32-479">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_7/office.notificationmessages)</span></span>

<span data-ttu-id="91d32-480">Получает сообщения уведомления для элемента.</span><span class="sxs-lookup"><span data-stu-id="91d32-480">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="91d32-481">Тип:</span><span class="sxs-lookup"><span data-stu-id="91d32-481">Type:</span></span>

*   [<span data-ttu-id="91d32-482">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="91d32-482">NotificationMessages</span></span>](/javascript/api/outlook_1_7/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="91d32-483">Требования</span><span class="sxs-lookup"><span data-stu-id="91d32-483">Requirements</span></span>

|<span data-ttu-id="91d32-484">Требование</span><span class="sxs-lookup"><span data-stu-id="91d32-484">Requirement</span></span>|<span data-ttu-id="91d32-485">Значение</span><span class="sxs-lookup"><span data-stu-id="91d32-485">Value</span></span>|
|---|---|
|[<span data-ttu-id="91d32-486">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="91d32-486">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="91d32-487">1.3</span><span class="sxs-lookup"><span data-stu-id="91d32-487">1.3</span></span>|
|[<span data-ttu-id="91d32-488">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="91d32-488">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="91d32-489">ReadItem</span><span class="sxs-lookup"><span data-stu-id="91d32-489">ReadItem</span></span>|
|[<span data-ttu-id="91d32-490">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="91d32-490">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="91d32-491">Compose (создание) или read (чтение)</span><span class="sxs-lookup"><span data-stu-id="91d32-491">Compose or read</span></span>|

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients"></a><span data-ttu-id="91d32-492">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="91d32-492">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

<span data-ttu-id="91d32-493">Предоставляет доступ к необязательным участникам события.</span><span class="sxs-lookup"><span data-stu-id="91d32-493">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="91d32-494">Уровень доступа и тип объекта, зависит от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="91d32-494">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="91d32-495">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="91d32-495">Read mode</span></span>

<span data-ttu-id="91d32-496">Свойство `optionalAttendees` возвращает массив, содержащий объект `EmailAddressDetails` для каждого необязательного участника собрания.</span><span class="sxs-lookup"><span data-stu-id="91d32-496">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="91d32-497">Режим создания</span><span class="sxs-lookup"><span data-stu-id="91d32-497">Compose mode</span></span>

<span data-ttu-id="91d32-498">Свойство `optionalAttendees` возвращает объект `Recipients`, который предоставляет методы для получения и обновления необязательных участников собрания.</span><span class="sxs-lookup"><span data-stu-id="91d32-498">The `optionalAttendees` property returns a `Recipients` object that provides methods to get and set the optional attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="91d32-499">Тип:</span><span class="sxs-lookup"><span data-stu-id="91d32-499">Type:</span></span>

*   <span data-ttu-id="91d32-500">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="91d32-500">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="91d32-501">Требования</span><span class="sxs-lookup"><span data-stu-id="91d32-501">Requirements</span></span>

|<span data-ttu-id="91d32-502">Требование</span><span class="sxs-lookup"><span data-stu-id="91d32-502">Requirement</span></span>|<span data-ttu-id="91d32-503">Значение</span><span class="sxs-lookup"><span data-stu-id="91d32-503">Value</span></span>|
|---|---|
|[<span data-ttu-id="91d32-504">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="91d32-504">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="91d32-505">1.0</span><span class="sxs-lookup"><span data-stu-id="91d32-505">1.0</span></span>|
|[<span data-ttu-id="91d32-506">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="91d32-506">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="91d32-507">ReadItem</span><span class="sxs-lookup"><span data-stu-id="91d32-507">ReadItem</span></span>|
|[<span data-ttu-id="91d32-508">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="91d32-508">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="91d32-509">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="91d32-509">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="91d32-510">Пример</span><span class="sxs-lookup"><span data-stu-id="91d32-510">Example</span></span>

```
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsorganizerjavascriptapioutlook17officeorganizer"></a><span data-ttu-id="91d32-511">organizer :[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)|[Organizer](/javascript/api/outlook_1_7/office.organizer)</span><span class="sxs-lookup"><span data-stu-id="91d32-511">organizer :[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)|[Organizer](/javascript/api/outlook_1_7/office.organizer)</span></span>

<span data-ttu-id="91d32-512">Получает адрес электронной почты организатора указанного собрания.</span><span class="sxs-lookup"><span data-stu-id="91d32-512">Gets the email address of the meeting organizer for a specified meeting.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="91d32-513">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="91d32-513">Read mode</span></span>

<span data-ttu-id="91d32-514">Свойство `organizer` возвращает объект [EmailAddressDetails,](/javascript/api/outlook_1_7/office.emailaddressdetails) который представляет организатора собрания.</span><span class="sxs-lookup"><span data-stu-id="91d32-514">The `organizer` property returns an [EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails) object that represents the meeting organizer.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="91d32-515">Режим создания</span><span class="sxs-lookup"><span data-stu-id="91d32-515">Compose mode</span></span>

<span data-ttu-id="91d32-516">Свойство `organizer` возвращает объект [Organizer](/javascript/api/outlook_1_7/office.organizer), который предоставляет метод для получения значения организатора.</span><span class="sxs-lookup"><span data-stu-id="91d32-516">The `organizer` property returns an [Organizer](/javascript/api/outlook_1_7/office.organizer) object that provides a method to get the organizer value.</span></span>

##### <a name="type"></a><span data-ttu-id="91d32-517">Тип:</span><span class="sxs-lookup"><span data-stu-id="91d32-517">Type:</span></span>

*   <span data-ttu-id="91d32-518">[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails) | [Organizer](/javascript/api/outlook_1_7/office.organizer)</span><span class="sxs-lookup"><span data-stu-id="91d32-518">[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails) | [Organizer](/javascript/api/outlook_1_7/office.organizer)</span></span>

##### <a name="requirements"></a><span data-ttu-id="91d32-519">Требования</span><span class="sxs-lookup"><span data-stu-id="91d32-519">Requirements</span></span>

|<span data-ttu-id="91d32-520">Требование</span><span class="sxs-lookup"><span data-stu-id="91d32-520">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="91d32-521">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="91d32-521">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="91d32-522">1.0</span><span class="sxs-lookup"><span data-stu-id="91d32-522">1.0</span></span>|<span data-ttu-id="91d32-523">1.7</span><span class="sxs-lookup"><span data-stu-id="91d32-523">17 </span></span>|
|[<span data-ttu-id="91d32-524">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="91d32-524">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="91d32-525">ReadItem</span><span class="sxs-lookup"><span data-stu-id="91d32-525">ReadItem</span></span>|<span data-ttu-id="91d32-526">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="91d32-526">ReadWriteItem</span></span>|
|[<span data-ttu-id="91d32-527">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="91d32-527">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="91d32-528">Чтение</span><span class="sxs-lookup"><span data-stu-id="91d32-528">Read</span></span>|<span data-ttu-id="91d32-529">Compose</span><span class="sxs-lookup"><span data-stu-id="91d32-529">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="91d32-530">Пример</span><span class="sxs-lookup"><span data-stu-id="91d32-530">Example</span></span>

```
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
```

#### <a name="nullable-recurrence-recurrencejavascriptapioutlook17officerecurrence"></a><span data-ttu-id="91d32-531">(nullable) recurrence :[Recurrence](/javascript/api/outlook_1_7/office.recurrence)</span><span class="sxs-lookup"><span data-stu-id="91d32-531">(nullable) recurrence :[Recurrence](/javascript/api/outlook_1_7/office.recurrence)</span></span>

<span data-ttu-id="91d32-532">Получает или задает расписание повторения встречи.</span><span class="sxs-lookup"><span data-stu-id="91d32-532">Gets or sets the location of an appointment.</span></span> <span data-ttu-id="91d32-533">Получает расписание повторения приглашения на собрание.</span><span class="sxs-lookup"><span data-stu-id="91d32-533">Gets the recurrence pattern of a meeting request.</span></span> <span data-ttu-id="91d32-534">Чтение и создание режимов для элементов встречи.</span><span class="sxs-lookup"><span data-stu-id="91d32-534">Read and compose modes for appointment items.</span></span> <span data-ttu-id="91d32-535">Режим чтения для элементов запроса на собрание.</span><span class="sxs-lookup"><span data-stu-id="91d32-535">Read mode for meeting request items.</span></span>

<span data-ttu-id="91d32-536">Свойство `recurrence` возвращает объект [recurrence](/javascript/api/outlook_1_7/office.recurrence) для повторения запросов на встречи или собрания, если элемент или экземпляр являются серийными.</span><span class="sxs-lookup"><span data-stu-id="91d32-536">The `recurrence` property returns a [recurrence](/javascript/api/outlook_1_7/office.recurrence) object for recurring appointments or meetings requests if an item is a series or an instance in a series.</span></span> <span data-ttu-id="91d32-537">`null` возвращается для одиночных встреч и запросов на собрания одиночных встреч.</span><span class="sxs-lookup"><span data-stu-id="91d32-537">`null` is returned for single appointments and meeting requests of single appointments.</span></span> <span data-ttu-id="91d32-538">`undefined` возвращается для сообщений, которые не являются запросами на собрания.</span><span class="sxs-lookup"><span data-stu-id="91d32-538">`undefined` is returned for messages that are not meeting requests.</span></span>

> <span data-ttu-id="91d32-539">Примечание: запросы на собрание имеют значение IPM.Schedule.Meeting.Request `itemClass`.</span><span class="sxs-lookup"><span data-stu-id="91d32-539">Note: Meeting requests have an `itemClass` value of IPM.Schedule.Meeting.Request.</span></span>

> <span data-ttu-id="91d32-540">Примечание: если объектом повторения является `null`, это указывает на то, что объект является одиночной встречей или запросом на собрание одиночной встречи и НЕ является частью серии.</span><span class="sxs-lookup"><span data-stu-id="91d32-540">Note: If the recurrence object is `null`, this indicates that the object is a single appointment or a meeting request of a single appointment and NOT a part of a series.</span></span>

##### <a name="type"></a><span data-ttu-id="91d32-541">Тип:</span><span class="sxs-lookup"><span data-stu-id="91d32-541">Type:</span></span>

* [<span data-ttu-id="91d32-542">Recurrence</span><span class="sxs-lookup"><span data-stu-id="91d32-542">recurrence</span></span>](/javascript/api/outlook_1_7/office.recurrence)

|<span data-ttu-id="91d32-543">Требование</span><span class="sxs-lookup"><span data-stu-id="91d32-543">Requirement</span></span>|<span data-ttu-id="91d32-544">Значение</span><span class="sxs-lookup"><span data-stu-id="91d32-544">Value</span></span>|
|---|---|
|[<span data-ttu-id="91d32-545">Версия минимального набора требований для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="91d32-545">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="91d32-546">1.7</span><span class="sxs-lookup"><span data-stu-id="91d32-546">17 </span></span>|
|[<span data-ttu-id="91d32-547">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="91d32-547">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="91d32-548">ReadItem</span><span class="sxs-lookup"><span data-stu-id="91d32-548">ReadItem</span></span>|
|[<span data-ttu-id="91d32-549">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="91d32-549">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="91d32-550">Compose (создание) или read (чтение)</span><span class="sxs-lookup"><span data-stu-id="91d32-550">Compose or read</span></span>|

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients"></a><span data-ttu-id="91d32-551">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="91d32-551">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

<span data-ttu-id="91d32-552">Предоставляет доступ к обязательным участникам события.</span><span class="sxs-lookup"><span data-stu-id="91d32-552">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="91d32-553">Уровень доступа и тип объекта, зависит от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="91d32-553">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="91d32-554">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="91d32-554">Read mode</span></span>

<span data-ttu-id="91d32-555">Свойство `requiredAttendees` возвращает массив, содержащий объект `EmailAddressDetails`, каждому обязательному участнику собрания.</span><span class="sxs-lookup"><span data-stu-id="91d32-555">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="91d32-556">Режим создания</span><span class="sxs-lookup"><span data-stu-id="91d32-556">Compose mode</span></span>

<span data-ttu-id="91d32-557">Свойство `requiredAttendees` возвращает объект `Recipients`, который предоставляет методы для получения и обновления обязательных участников собрания.</span><span class="sxs-lookup"><span data-stu-id="91d32-557">The `requiredAttendees` property returns a `Recipients` object that provides methods to get and set the required attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="91d32-558">Тип:</span><span class="sxs-lookup"><span data-stu-id="91d32-558">Type:</span></span>

*   <span data-ttu-id="91d32-559">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="91d32-559">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="91d32-560">Требования</span><span class="sxs-lookup"><span data-stu-id="91d32-560">Requirements</span></span>

|<span data-ttu-id="91d32-561">Требование</span><span class="sxs-lookup"><span data-stu-id="91d32-561">Requirement</span></span>|<span data-ttu-id="91d32-562">Значение</span><span class="sxs-lookup"><span data-stu-id="91d32-562">Value</span></span>|
|---|---|
|[<span data-ttu-id="91d32-563">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="91d32-563">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="91d32-564">1.0</span><span class="sxs-lookup"><span data-stu-id="91d32-564">1.0</span></span>|
|[<span data-ttu-id="91d32-565">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="91d32-565">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="91d32-566">ReadItem</span><span class="sxs-lookup"><span data-stu-id="91d32-566">ReadItem</span></span>|
|[<span data-ttu-id="91d32-567">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="91d32-567">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="91d32-568">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="91d32-568">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="91d32-569">Пример</span><span class="sxs-lookup"><span data-stu-id="91d32-569">Example</span></span>

```
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
}
```

#### <a name="sender-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetails"></a><span data-ttu-id="91d32-570">sender :[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="91d32-570">sender :[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)</span></span>

<span data-ttu-id="91d32-p126">Получает электронный адрес отправителя электронного сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="91d32-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="91d32-p127">Свойства [`from`](#from-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsfromjavascriptapioutlook17officefrom) и `sender` представляют одно и то же лицо, если сообщение не отправлено делегатом. В противном случае свойство `from` представляет лицо, делегировавшее полномочия, а свойство sender — делегата.</span><span class="sxs-lookup"><span data-stu-id="91d32-p127">The [`from`](#from-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsfromjavascriptapioutlook17officefrom) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="91d32-575">Свойство `recipientType` объекта `EmailAddressDetails` в свойстве `sender` — `undefined`.</span><span class="sxs-lookup"><span data-stu-id="91d32-575">Note: The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="91d32-576">Тип:</span><span class="sxs-lookup"><span data-stu-id="91d32-576">Type:</span></span>

*   [<span data-ttu-id="91d32-577">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="91d32-577">EmailAddressDetails</span></span>](/javascript/api/outlook_1_7/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="91d32-578">Требования</span><span class="sxs-lookup"><span data-stu-id="91d32-578">Requirements</span></span>

|<span data-ttu-id="91d32-579">Требование</span><span class="sxs-lookup"><span data-stu-id="91d32-579">Requirement</span></span>|<span data-ttu-id="91d32-580">Значение</span><span class="sxs-lookup"><span data-stu-id="91d32-580">Value</span></span>|
|---|---|
|[<span data-ttu-id="91d32-581">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="91d32-581">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="91d32-582">1.0</span><span class="sxs-lookup"><span data-stu-id="91d32-582">1.0</span></span>|
|[<span data-ttu-id="91d32-583">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="91d32-583">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="91d32-584">ReadItem</span><span class="sxs-lookup"><span data-stu-id="91d32-584">ReadItem</span></span>|
|[<span data-ttu-id="91d32-585">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="91d32-585">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="91d32-586">Чтение</span><span class="sxs-lookup"><span data-stu-id="91d32-586">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="91d32-587">Пример</span><span class="sxs-lookup"><span data-stu-id="91d32-587">Example</span></span>

```
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
```

#### <a name="nullable-seriesid-string"></a><span data-ttu-id="91d32-588">(nullable) seriesId :String</span><span class="sxs-lookup"><span data-stu-id="91d32-588">(nullable) seriesId :String</span></span>

<span data-ttu-id="91d32-589">Получает идентификатор серии, к которой принадлежит экземпляр.</span><span class="sxs-lookup"><span data-stu-id="91d32-589">Gets the id of the series that an instance belongs to.</span></span>

<span data-ttu-id="91d32-590">В OWA и Outlook `seriesId` возвращает идентификатор веб-служб Exchange (EWS) родительского (серийного) элемента, к которому принадлежит этот элемент.</span><span class="sxs-lookup"><span data-stu-id="91d32-590">In OWA and Outlook, the `seriesId` returns the Exchange Web Services (EWS) ID of the parent (series) item that this item belongs to.</span></span> <span data-ttu-id="91d32-591">Однако в iOS и Android `seriesId` возвращает REST идентификатор родительского элемента.</span><span class="sxs-lookup"><span data-stu-id="91d32-591">However, in iOS and Android, the `seriesId` returns the REST ID of the parent item.</span></span>

> [!NOTE]
> <span data-ttu-id="91d32-592">Идентификатор, возвращаемый свойством `seriesId`, совпадает с идентификатором элемента веб-служб Exchange.</span><span class="sxs-lookup"><span data-stu-id="91d32-592">The identifier returned by the `seriesId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="91d32-593">Свойство `seriesId` не идентично идентификаторам Outlook, используемым API-Интерфейсом REST Outlook.</span><span class="sxs-lookup"><span data-stu-id="91d32-593">The `seriesId` property is not identical to the Outlook IDs used by the Outlook REST API.</span></span> <span data-ttu-id="91d32-594">Прежде чем осуществлять вызовы API-Интерфейса REST с помощью этого значения, его следует преобразовать с помощью [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="91d32-594">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="91d32-595">Для получения дополнительных сведений см. [Использование API REST Outlook из надстройки Outlook](https://docs.microsoft.com/outlook/add-ins/use-rest-api).</span><span class="sxs-lookup"><span data-stu-id="91d32-595">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api).</span></span>

<span data-ttu-id="91d32-596">Свойство `seriesId` возвращает `null` для элементов, у которых нет родительских элементов, таких как одиночные встречи, элементы серии или запросы на собрания и возвращает `undefined` для любых других элементов, которые не являются запросами на собрание.</span><span class="sxs-lookup"><span data-stu-id="91d32-596">The `seriesId` property returns `null` for items that do not have parent items such as single appointments, series items, or meeting requests and returns `undefined` for any other items that are not meeting requests.</span></span>

##### <a name="type"></a><span data-ttu-id="91d32-597">Тип:</span><span class="sxs-lookup"><span data-stu-id="91d32-597">Type:</span></span>

* <span data-ttu-id="91d32-598">String</span><span class="sxs-lookup"><span data-stu-id="91d32-598">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="91d32-599">Требования</span><span class="sxs-lookup"><span data-stu-id="91d32-599">Requirements</span></span>

|<span data-ttu-id="91d32-600">Требование</span><span class="sxs-lookup"><span data-stu-id="91d32-600">Requirement</span></span>|<span data-ttu-id="91d32-601">Значение</span><span class="sxs-lookup"><span data-stu-id="91d32-601">Value</span></span>|
|---|---|
|[<span data-ttu-id="91d32-602">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="91d32-602">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="91d32-603">1.7</span><span class="sxs-lookup"><span data-stu-id="91d32-603">17 </span></span>|
|[<span data-ttu-id="91d32-604">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="91d32-604">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="91d32-605">ReadItem</span><span class="sxs-lookup"><span data-stu-id="91d32-605">ReadItem</span></span>|
|[<span data-ttu-id="91d32-606">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="91d32-606">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="91d32-607">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="91d32-607">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="91d32-608">Пример</span><span class="sxs-lookup"><span data-stu-id="91d32-608">Example</span></span>

```
var seriesId = Office.context.mailbox.item.seriesId; 
var isSeries = (seriesId == null);
```

####  <a name="start-datetimejavascriptapioutlook17officetime"></a><span data-ttu-id="91d32-609">start :Date|[Time](/javascript/api/outlook_1_7/office.time)</span><span class="sxs-lookup"><span data-stu-id="91d32-609">start :Date|[Time](/javascript/api/outlook_1_7/office.time)</span></span>

<span data-ttu-id="91d32-610">Получает или задает дату и время начала встречи.</span><span class="sxs-lookup"><span data-stu-id="91d32-610">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="91d32-p130">Свойство `start` представлено в виде значения даты и времени в формате UTC. Это значение можно преобразовать в местные значения даты и времени клиента с помощью метода [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook17officelocalclienttime).</span><span class="sxs-lookup"><span data-stu-id="91d32-p130">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook17officelocalclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="91d32-613">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="91d32-613">Read mode</span></span>

<span data-ttu-id="91d32-614">Свойство `start` возвращает объект `Date`.</span><span class="sxs-lookup"><span data-stu-id="91d32-614">The `start` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="91d32-615">Режим создания</span><span class="sxs-lookup"><span data-stu-id="91d32-615">Compose mode</span></span>

<span data-ttu-id="91d32-616">Свойство `start` возвращает объект `Time`.</span><span class="sxs-lookup"><span data-stu-id="91d32-616">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="91d32-617">Если вы задаете время начала с помощью метода [`Time.setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-), необходимо использовать метод [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) для преобразования местного времени на клиенте в формат UTC для сервера.</span><span class="sxs-lookup"><span data-stu-id="91d32-617">When you use the [`Time.setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="91d32-618">Тип:</span><span class="sxs-lookup"><span data-stu-id="91d32-618">Type:</span></span>

*   <span data-ttu-id="91d32-619">Date | [Time](/javascript/api/outlook_1_7/office.time)</span><span class="sxs-lookup"><span data-stu-id="91d32-619">Date | [Time](/javascript/api/outlook_1_7/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="91d32-620">Требования</span><span class="sxs-lookup"><span data-stu-id="91d32-620">Requirements</span></span>

|<span data-ttu-id="91d32-621">Требование</span><span class="sxs-lookup"><span data-stu-id="91d32-621">Requirement</span></span>|<span data-ttu-id="91d32-622">Значение</span><span class="sxs-lookup"><span data-stu-id="91d32-622">Value</span></span>|
|---|---|
|[<span data-ttu-id="91d32-623">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="91d32-623">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="91d32-624">1.0</span><span class="sxs-lookup"><span data-stu-id="91d32-624">1.0</span></span>|
|[<span data-ttu-id="91d32-625">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="91d32-625">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="91d32-626">ReadItem</span><span class="sxs-lookup"><span data-stu-id="91d32-626">ReadItem</span></span>|
|[<span data-ttu-id="91d32-627">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="91d32-627">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="91d32-628">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="91d32-628">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="91d32-629">Пример</span><span class="sxs-lookup"><span data-stu-id="91d32-629">Example</span></span>

<span data-ttu-id="91d32-630">В примере ниже с помощью метода [`setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) объекта `Time` задается время начала встречи в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="91d32-630">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

```
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

####  <a name="subject-stringsubjectjavascriptapioutlook17officesubject"></a><span data-ttu-id="91d32-631">subject :String|[Subject](/javascript/api/outlook_1_7/office.subject)</span><span class="sxs-lookup"><span data-stu-id="91d32-631">subject :String|[Subject](/javascript/api/outlook_1_7/office.subject)</span></span>

<span data-ttu-id="91d32-632">Получает или задает описание, которое отображается в поле темы элемента.</span><span class="sxs-lookup"><span data-stu-id="91d32-632">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="91d32-633">Свойство `subject` получает или задает всю тему элемента для отправки с почтового сервера.</span><span class="sxs-lookup"><span data-stu-id="91d32-633">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="91d32-634">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="91d32-634">Read mode</span></span>

<span data-ttu-id="91d32-p131">Свойство `subject` возвращает строку. С помощью свойства [`normalizedSubject`](#normalizedsubject-string) можно получить тему без начальных префиксов, например, `RE:` и `FW:`.</span><span class="sxs-lookup"><span data-stu-id="91d32-p131">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```
var subject = Office.context.mailbox.item.subject;
```

##### <a name="compose-mode"></a><span data-ttu-id="91d32-637">Режим создания</span><span class="sxs-lookup"><span data-stu-id="91d32-637">Compose mode</span></span>

<span data-ttu-id="91d32-638">Свойство `subject` возвращает объект `Subject`, который предоставляет методы для получения и задания темы.</span><span class="sxs-lookup"><span data-stu-id="91d32-638">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="91d32-639">Тип:</span><span class="sxs-lookup"><span data-stu-id="91d32-639">Type:</span></span>

*   <span data-ttu-id="91d32-640">String | [Subject](/javascript/api/outlook_1_7/office.subject)</span><span class="sxs-lookup"><span data-stu-id="91d32-640">String | [Subject](/javascript/api/outlook_1_7/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="91d32-641">Требования</span><span class="sxs-lookup"><span data-stu-id="91d32-641">Requirements</span></span>

|<span data-ttu-id="91d32-642">Требование</span><span class="sxs-lookup"><span data-stu-id="91d32-642">Requirement</span></span>|<span data-ttu-id="91d32-643">Значение</span><span class="sxs-lookup"><span data-stu-id="91d32-643">Value</span></span>|
|---|---|
|[<span data-ttu-id="91d32-644">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="91d32-644">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="91d32-645">1.0</span><span class="sxs-lookup"><span data-stu-id="91d32-645">1.0</span></span>|
|[<span data-ttu-id="91d32-646">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="91d32-646">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="91d32-647">ReadItem</span><span class="sxs-lookup"><span data-stu-id="91d32-647">ReadItem</span></span>|
|[<span data-ttu-id="91d32-648">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="91d32-648">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="91d32-649">Compose (создание) или read (чтение)</span><span class="sxs-lookup"><span data-stu-id="91d32-649">Compose or read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients"></a><span data-ttu-id="91d32-650">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="91d32-650">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

<span data-ttu-id="91d32-651">Предоставляет доступ получателей к строке **To (Кому)** в сообщении.</span><span class="sxs-lookup"><span data-stu-id="91d32-651">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="91d32-652">Уровень доступа и тип объекта, зависит от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="91d32-652">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="91d32-653">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="91d32-653">Read mode</span></span>

<span data-ttu-id="91d32-p133">Свойство `to` возвращает массив, содержащий объект `EmailAddressDetails` для каждого получателя в строке **To (Кому)** сообщения. Коллекция может включать не более 100 элементов.</span><span class="sxs-lookup"><span data-stu-id="91d32-p133">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="91d32-656">Режим создания</span><span class="sxs-lookup"><span data-stu-id="91d32-656">Compose mode</span></span>

<span data-ttu-id="91d32-657">Свойство `to` возвращает объект `Recipients`, предоставляющий методы для получения или обновления получателей, которые указаны в строке **To (кому)** сообщения.</span><span class="sxs-lookup"><span data-stu-id="91d32-657">The to`to` property returns a Recipients`Recipients` object that provides methods for manipulating the recipients on the **To** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="91d32-658">Тип:</span><span class="sxs-lookup"><span data-stu-id="91d32-658">Type:</span></span>

*   <span data-ttu-id="91d32-659">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="91d32-659">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="91d32-660">Требования</span><span class="sxs-lookup"><span data-stu-id="91d32-660">Requirements</span></span>

|<span data-ttu-id="91d32-661">Требование</span><span class="sxs-lookup"><span data-stu-id="91d32-661">Requirement</span></span>|<span data-ttu-id="91d32-662">Значение</span><span class="sxs-lookup"><span data-stu-id="91d32-662">Value</span></span>|
|---|---|
|[<span data-ttu-id="91d32-663">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="91d32-663">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="91d32-664">1.0</span><span class="sxs-lookup"><span data-stu-id="91d32-664">1.0</span></span>|
|[<span data-ttu-id="91d32-665">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="91d32-665">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="91d32-666">ReadItem</span><span class="sxs-lookup"><span data-stu-id="91d32-666">ReadItem</span></span>|
|[<span data-ttu-id="91d32-667">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="91d32-667">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="91d32-668">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="91d32-668">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="91d32-669">Пример</span><span class="sxs-lookup"><span data-stu-id="91d32-669">Example</span></span>

```
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

### <a name="methods"></a><span data-ttu-id="91d32-670">Методы</span><span class="sxs-lookup"><span data-stu-id="91d32-670">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="91d32-671">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="91d32-671">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="91d32-672">Добавляет файл в сообщение или встречу в качестве вложения.</span><span class="sxs-lookup"><span data-stu-id="91d32-672">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="91d32-673">Метод `addFileAttachmentAsync` передает файл по указанному универсальному коду ресурса (URI) и вкладывает его в элемент в форме создания.</span><span class="sxs-lookup"><span data-stu-id="91d32-673">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="91d32-674">Идентификатор можно последовательно использовать с методом [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback), чтобы удалить вложение, добавленное во время текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="91d32-674">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="91d32-675">Параметры:</span><span class="sxs-lookup"><span data-stu-id="91d32-675">Parameters:</span></span>
|<span data-ttu-id="91d32-676">Имя</span><span class="sxs-lookup"><span data-stu-id="91d32-676">Name</span></span>|<span data-ttu-id="91d32-677">Тип</span><span class="sxs-lookup"><span data-stu-id="91d32-677">Type</span></span>|<span data-ttu-id="91d32-678">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="91d32-678">Attributes</span></span>|<span data-ttu-id="91d32-679">Описание</span><span class="sxs-lookup"><span data-stu-id="91d32-679">Description</span></span>|
|---|---|---|---|
|`uri`|<span data-ttu-id="91d32-680">String</span><span class="sxs-lookup"><span data-stu-id="91d32-680">String</span></span>||<span data-ttu-id="91d32-p134">URI-адрес, представляющий расположение файла, который нужно вложить в сообщение или встречу. Максимальная длина — 2048 символов.</span><span class="sxs-lookup"><span data-stu-id="91d32-p134">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="91d32-683">String</span><span class="sxs-lookup"><span data-stu-id="91d32-683">String</span></span>||<span data-ttu-id="91d32-p135">Имя вложения, которое отображается при передаче вложения. Максимальная длина — 255 символов.</span><span class="sxs-lookup"><span data-stu-id="91d32-p135">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="91d32-686">Объект</span><span class="sxs-lookup"><span data-stu-id="91d32-686">Object</span></span>|<span data-ttu-id="91d32-687">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="91d32-687">&lt;optional&gt;</span></span>|<span data-ttu-id="91d32-688">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="91d32-688">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="91d32-689">Объект</span><span class="sxs-lookup"><span data-stu-id="91d32-689">Object</span></span>|<span data-ttu-id="91d32-690">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="91d32-690">&lt;optional&gt;</span></span>|<span data-ttu-id="91d32-691">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="91d32-691">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="91d32-692">Логическое значение</span><span class="sxs-lookup"><span data-stu-id="91d32-692">Boolean</span></span>|<span data-ttu-id="91d32-693">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="91d32-693">&lt;optional&gt;</span></span>|<span data-ttu-id="91d32-694">Значение `true` указывает на то, что вложение будет встроено в текст сообщения и не должно отображаться в списке вложений.</span><span class="sxs-lookup"><span data-stu-id="91d32-694">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="91d32-695">function</span><span class="sxs-lookup"><span data-stu-id="91d32-695">function</span></span>|<span data-ttu-id="91d32-696">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="91d32-696">&lt;optional&gt;</span></span>|<span data-ttu-id="91d32-697">Когда метод завершает выполнение, переданная в параметре `callback` функция вызывается с единственным параметром `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="91d32-697">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="91d32-698">После успешного выполнения идентификатор вложения будет представлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="91d32-698">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="91d32-699">Если добавить вложение не удастся, объект `asyncResult` будет содержать объект `Error` с описанием ошибки.</span><span class="sxs-lookup"><span data-stu-id="91d32-699">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="91d32-700">Ошибки</span><span class="sxs-lookup"><span data-stu-id="91d32-700">Errors</span></span>

|<span data-ttu-id="91d32-701">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="91d32-701">Error code</span></span>|<span data-ttu-id="91d32-702">Описание</span><span class="sxs-lookup"><span data-stu-id="91d32-702">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="91d32-703">Вложение превышает максимальный размер.</span><span class="sxs-lookup"><span data-stu-id="91d32-703">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="91d32-704">Расширение вложения не поддерживается.</span><span class="sxs-lookup"><span data-stu-id="91d32-704">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="91d32-705">Сообщение или встреча содержат слишком много вложений.</span><span class="sxs-lookup"><span data-stu-id="91d32-705">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="91d32-706">Требования</span><span class="sxs-lookup"><span data-stu-id="91d32-706">Requirements</span></span>

|<span data-ttu-id="91d32-707">Требование</span><span class="sxs-lookup"><span data-stu-id="91d32-707">Requirement</span></span>|<span data-ttu-id="91d32-708">Значение</span><span class="sxs-lookup"><span data-stu-id="91d32-708">Value</span></span>|
|---|---|
|[<span data-ttu-id="91d32-709">Версия минимального набора требований для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="91d32-709">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="91d32-710">1.1</span><span class="sxs-lookup"><span data-stu-id="91d32-710">1.1</span></span>|
|[<span data-ttu-id="91d32-711">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="91d32-711">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="91d32-712">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="91d32-712">ReadWriteItem</span></span>|
|[<span data-ttu-id="91d32-713">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="91d32-713">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="91d32-714">Создание</span><span class="sxs-lookup"><span data-stu-id="91d32-714">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="91d32-715">Примеры</span><span class="sxs-lookup"><span data-stu-id="91d32-715">Examples</span></span>

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

<span data-ttu-id="91d32-716">В приведенном ниже примере файл изображения добавляется в качестве встроенного вложения, а ссылка на вложение добавляется в текст сообщения.</span><span class="sxs-lookup"><span data-stu-id="91d32-716">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

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

####  <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="91d32-717">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="91d32-717">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="91d32-718">Добавляет обработчик для поддерживаемого события.</span><span class="sxs-lookup"><span data-stu-id="91d32-718">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="91d32-719">В настоящее время поддерживаемые типы событий — `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, и `Office.EventType.RecurrenceChanged`</span><span class="sxs-lookup"><span data-stu-id="91d32-719">Currently the supported event types are `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`</span></span>

##### <a name="parameters"></a><span data-ttu-id="91d32-720">Параметры:</span><span class="sxs-lookup"><span data-stu-id="91d32-720">Parameters:</span></span>

| <span data-ttu-id="91d32-721">Имя</span><span class="sxs-lookup"><span data-stu-id="91d32-721">Name</span></span> | <span data-ttu-id="91d32-722">Тип</span><span class="sxs-lookup"><span data-stu-id="91d32-722">Type</span></span> | <span data-ttu-id="91d32-723">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="91d32-723">Attributes</span></span> | <span data-ttu-id="91d32-724">Описание</span><span class="sxs-lookup"><span data-stu-id="91d32-724">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="91d32-725">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="91d32-725">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="91d32-726">Событие, которое должно вызвать обработчик.</span><span class="sxs-lookup"><span data-stu-id="91d32-726">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="91d32-727">Функция</span><span class="sxs-lookup"><span data-stu-id="91d32-727">Function</span></span> || <span data-ttu-id="91d32-p136">Функция для обработки события. Функция должна принимать один параметр, представляющий собой объектный литерал. Значение свойства `type` параметра совпадет со значением параметра `eventType`, переданного методу `addHandlerAsync`.</span><span class="sxs-lookup"><span data-stu-id="91d32-p136">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="91d32-731">Object</span><span class="sxs-lookup"><span data-stu-id="91d32-731">Object</span></span> | <span data-ttu-id="91d32-732">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="91d32-732">&lt;optional&gt;</span></span> | <span data-ttu-id="91d32-733">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="91d32-733">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="91d32-734">Объект</span><span class="sxs-lookup"><span data-stu-id="91d32-734">Object</span></span> | <span data-ttu-id="91d32-735">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="91d32-735">&lt;optional&gt;</span></span> | <span data-ttu-id="91d32-736">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="91d32-736">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="91d32-737">function</span><span class="sxs-lookup"><span data-stu-id="91d32-737">function</span></span>| <span data-ttu-id="91d32-738">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="91d32-738">&lt;optional&gt;</span></span>|<span data-ttu-id="91d32-739">Когда метод завершает выполнение, переданная в параметре `callback` функция вызывается с единственным параметром `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="91d32-739">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="91d32-740">Требования</span><span class="sxs-lookup"><span data-stu-id="91d32-740">Requirements</span></span>

|<span data-ttu-id="91d32-741">Требование</span><span class="sxs-lookup"><span data-stu-id="91d32-741">Requirement</span></span>| <span data-ttu-id="91d32-742">Значение</span><span class="sxs-lookup"><span data-stu-id="91d32-742">Value</span></span>|
|---|---|
|[<span data-ttu-id="91d32-743">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="91d32-743">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="91d32-744">1.7</span><span class="sxs-lookup"><span data-stu-id="91d32-744">17 </span></span> |
|[<span data-ttu-id="91d32-745">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="91d32-745">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="91d32-746">ReadItem</span><span class="sxs-lookup"><span data-stu-id="91d32-746">ReadItem</span></span> |
|[<span data-ttu-id="91d32-747">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="91d32-747">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="91d32-748">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="91d32-748">Compose or read</span></span> |

##### <a name="example"></a><span data-ttu-id="91d32-749">Пример</span><span class="sxs-lookup"><span data-stu-id="91d32-749">Example</span></span>

```
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

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="91d32-750">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="91d32-750">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="91d32-751">Добавляет к сообщению или встрече элемент Exchange (например, сообщение) в виде вложения.</span><span class="sxs-lookup"><span data-stu-id="91d32-751">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="91d32-p137">С помощью метода `addItemAttachmentAsync` в элемент формы создания можно вложить элемент с указанным идентификатором Exchange. Если указать метод обратного вызова, то этот метод вызывается с помощью параметра `asyncResult`, который содержит идентификатор вложения или код, указывающий на ошибки, которые произошли при вложении элемента. При необходимости можно использовать параметр `options` для передачи сведений о состоянии в метод обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="91d32-p137">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="91d32-755">Идентификатор можно последовательно использовать с методом [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback), чтобы удалить вложение, добавленное во время текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="91d32-755">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="91d32-756">Если ваша надстройка Office выполняется в веб-приложении Outlook, метод `addItemAttachmentAsync` обеспечивает вложение элементов в элементы, отличные от редактируемого. Однако это действие не рекомендуется выполнять, так как оно не поддерживается.</span><span class="sxs-lookup"><span data-stu-id="91d32-756">If your Office add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="91d32-757">Параметры:</span><span class="sxs-lookup"><span data-stu-id="91d32-757">Parameters:</span></span>

|<span data-ttu-id="91d32-758">Имя</span><span class="sxs-lookup"><span data-stu-id="91d32-758">Name</span></span>|<span data-ttu-id="91d32-759">Тип</span><span class="sxs-lookup"><span data-stu-id="91d32-759">Type</span></span>|<span data-ttu-id="91d32-760">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="91d32-760">Attributes</span></span>|<span data-ttu-id="91d32-761">Описание</span><span class="sxs-lookup"><span data-stu-id="91d32-761">Description</span></span>|
|---|---|---|---|
|`itemId`|<span data-ttu-id="91d32-762">String</span><span class="sxs-lookup"><span data-stu-id="91d32-762">String</span></span>||<span data-ttu-id="91d32-p138">Идентификатор Exchange для вкладываемого элемента. Максимальная длина — 100 символов.</span><span class="sxs-lookup"><span data-stu-id="91d32-p138">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="91d32-765">String</span><span class="sxs-lookup"><span data-stu-id="91d32-765">String</span></span>||<span data-ttu-id="91d32-p139">Тема вкладываемого элемента. Максимальная длина — 255 символов.</span><span class="sxs-lookup"><span data-stu-id="91d32-p139">The sujbect of the item to be attached. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="91d32-768">Объект</span><span class="sxs-lookup"><span data-stu-id="91d32-768">Object</span></span>|<span data-ttu-id="91d32-769">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="91d32-769">&lt;optional&gt;</span></span>|<span data-ttu-id="91d32-770">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="91d32-770">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="91d32-771">Объект</span><span class="sxs-lookup"><span data-stu-id="91d32-771">Object</span></span>|<span data-ttu-id="91d32-772">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="91d32-772">&lt;optional&gt;</span></span>|<span data-ttu-id="91d32-773">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="91d32-773">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="91d32-774">function</span><span class="sxs-lookup"><span data-stu-id="91d32-774">function</span></span>|<span data-ttu-id="91d32-775">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="91d32-775">&lt;optional&gt;</span></span>|<span data-ttu-id="91d32-776">Когда метод завершает выполнение, переданная в параметре `callback` функция вызывается с единственным параметром `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="91d32-776">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="91d32-777">После успешного выполнения идентификатор вложения будет представлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="91d32-777">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="91d32-778">Если добавить вложение не удастся, объект `asyncResult` будет содержать объект `Error` с описанием ошибки.</span><span class="sxs-lookup"><span data-stu-id="91d32-778">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="91d32-779">Ошибки</span><span class="sxs-lookup"><span data-stu-id="91d32-779">Errors</span></span>

|<span data-ttu-id="91d32-780">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="91d32-780">Error code</span></span>|<span data-ttu-id="91d32-781">Описание</span><span class="sxs-lookup"><span data-stu-id="91d32-781">Description</span></span>|
|------------|-------------|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="91d32-782">Сообщение или встреча содержат слишком много вложений.</span><span class="sxs-lookup"><span data-stu-id="91d32-782">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="91d32-783">Требования</span><span class="sxs-lookup"><span data-stu-id="91d32-783">Requirements</span></span>

|<span data-ttu-id="91d32-784">Требование</span><span class="sxs-lookup"><span data-stu-id="91d32-784">Requirement</span></span>|<span data-ttu-id="91d32-785">Значение</span><span class="sxs-lookup"><span data-stu-id="91d32-785">Value</span></span>|
|---|---|
|[<span data-ttu-id="91d32-786">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="91d32-786">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="91d32-787">1.1</span><span class="sxs-lookup"><span data-stu-id="91d32-787">1.1</span></span>|
|[<span data-ttu-id="91d32-788">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="91d32-788">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="91d32-789">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="91d32-789">ReadWriteItem</span></span>|
|[<span data-ttu-id="91d32-790">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="91d32-790">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="91d32-791">Создание</span><span class="sxs-lookup"><span data-stu-id="91d32-791">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="91d32-792">Пример</span><span class="sxs-lookup"><span data-stu-id="91d32-792">Example</span></span>

<span data-ttu-id="91d32-793">В следующем примере существующий элемент Outlook добавляется в виде вложения с именем `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="91d32-793">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

```
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

####  <a name="close"></a><span data-ttu-id="91d32-794">close()</span><span class="sxs-lookup"><span data-stu-id="91d32-794">close()</span></span>

<span data-ttu-id="91d32-795">Закрывает текущий создаваемый элемент.</span><span class="sxs-lookup"><span data-stu-id="91d32-795">Closes the current item that is being composed.</span></span>

<span data-ttu-id="91d32-p140">Работа метода `close` зависит от текущего состояния создаваемого элемента. Если элемент содержит несохраненные изменения, клиент предложит пользователю сохранить или отклонить их либо отменить действие закрытия.</span><span class="sxs-lookup"><span data-stu-id="91d32-p140">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="91d32-798">Если элемент является встречей в Outlook в Интернете, и он был ранее сохранен с помощью `saveAsync`, пользователю предлагается сохранить, отменить или удалить его, даже если не произошло каких-либо изменений, поскольку этот элемент был последним сохраненным.</span><span class="sxs-lookup"><span data-stu-id="91d32-798">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="91d32-799">Если в клиенте Outlook для настольных ПК сообщение представляет собой ответ в тексте, метод `close` не работает.</span><span class="sxs-lookup"><span data-stu-id="91d32-799">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="91d32-800">Требования</span><span class="sxs-lookup"><span data-stu-id="91d32-800">Requirements</span></span>

|<span data-ttu-id="91d32-801">Требование</span><span class="sxs-lookup"><span data-stu-id="91d32-801">Requirement</span></span>|<span data-ttu-id="91d32-802">Значение</span><span class="sxs-lookup"><span data-stu-id="91d32-802">Value</span></span>|
|---|---|
|[<span data-ttu-id="91d32-803">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="91d32-803">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="91d32-804">1.3</span><span class="sxs-lookup"><span data-stu-id="91d32-804">1.3</span></span>|
|[<span data-ttu-id="91d32-805">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="91d32-805">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="91d32-806">Ограниченный доступ</span><span class="sxs-lookup"><span data-stu-id="91d32-806">Restricted</span></span>|
|[<span data-ttu-id="91d32-807">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="91d32-807">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="91d32-808">Создание</span><span class="sxs-lookup"><span data-stu-id="91d32-808">Compose</span></span>|

#### <a name="displayreplyallformformdata"></a><span data-ttu-id="91d32-809">displayReplyAllForm(formData)</span><span class="sxs-lookup"><span data-stu-id="91d32-809">displayReplyAllForm(formData)</span></span>

<span data-ttu-id="91d32-810">Отображает форму ответа, включающую отправителя и всех получателей выбранного сообщения или организатора и всех участников выбранной встречи.</span><span class="sxs-lookup"><span data-stu-id="91d32-810">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="91d32-811">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="91d32-811">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="91d32-812">В веб-приложении Outlook форма ответа отображается в виде всплывающей формы в представлении с 3 колонками либо всплывающей формы в представлении с 1 или 2 колонками.</span><span class="sxs-lookup"><span data-stu-id="91d32-812">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="91d32-813">Если любой строчный параметр превышает указанные для него ограничения, `displayReplyAllForm` возвращает исключение.</span><span class="sxs-lookup"><span data-stu-id="91d32-813">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="91d32-p141">Если в параметре `formData.attachments` указаны вложения, Outlook и веб-приложение Outlook пытаются скачать их и вложить в форму ответа. Если какие-либо вложения добавить не удается, в форме отображается сообщение об ошибке. Если сообщения об ошибках не предусмотрены, то они не отображаются.</span><span class="sxs-lookup"><span data-stu-id="91d32-p141">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="91d32-817">Параметры:</span><span class="sxs-lookup"><span data-stu-id="91d32-817">Parameters:</span></span>

|<span data-ttu-id="91d32-818">Имя</span><span class="sxs-lookup"><span data-stu-id="91d32-818">Name</span></span>|<span data-ttu-id="91d32-819">Тип</span><span class="sxs-lookup"><span data-stu-id="91d32-819">Type</span></span>|<span data-ttu-id="91d32-820">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="91d32-820">Attributes</span></span>|<span data-ttu-id="91d32-821">Описание</span><span class="sxs-lookup"><span data-stu-id="91d32-821">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="91d32-822">String | Object</span><span class="sxs-lookup"><span data-stu-id="91d32-822">String &#124; Object</span></span>||<span data-ttu-id="91d32-p142">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="91d32-p142">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="91d32-825">**ИЛИ**</span><span class="sxs-lookup"><span data-stu-id="91d32-825">**OR**</span></span><br/><span data-ttu-id="91d32-p143">Объект, который содержит текст или данные вложения и функцию обратного вызова. Ниже представлено определение этого объекта.</span><span class="sxs-lookup"><span data-stu-id="91d32-p143">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="91d32-828">String</span><span class="sxs-lookup"><span data-stu-id="91d32-828">String</span></span>|<span data-ttu-id="91d32-829">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="91d32-829">&lt;optional&gt;</span></span>|<span data-ttu-id="91d32-p144">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="91d32-p144">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="91d32-832">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="91d32-832">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="91d32-833">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="91d32-833">&lt;optional&gt;</span></span>|<span data-ttu-id="91d32-834">Массив объектов JSON, представляющих собой вложенные файлы или элементы.</span><span class="sxs-lookup"><span data-stu-id="91d32-834">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="91d32-835">String</span><span class="sxs-lookup"><span data-stu-id="91d32-835">String</span></span>||<span data-ttu-id="91d32-p145">Указывает тип вложения. Допустимые значения: `file` для вложенного файла и `item` для вложенного элемента.</span><span class="sxs-lookup"><span data-stu-id="91d32-p145">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="91d32-838">String</span><span class="sxs-lookup"><span data-stu-id="91d32-838">String</span></span>||<span data-ttu-id="91d32-839">Строка, содержащая имя вложения, длиной до 255 символов.</span><span class="sxs-lookup"><span data-stu-id="91d32-839">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="91d32-840">String</span><span class="sxs-lookup"><span data-stu-id="91d32-840">String</span></span>||<span data-ttu-id="91d32-p146">Используется, только если свойству `type` задано значение `file`. URI расположения файла.</span><span class="sxs-lookup"><span data-stu-id="91d32-p146">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="91d32-843">Логическое значение</span><span class="sxs-lookup"><span data-stu-id="91d32-843">Boolean</span></span>||<span data-ttu-id="91d32-p147">Используется только в том случае, если свойству `type` задано значение `file`. Значение `true` указывает на то, что вложение будет встроено в текст сообщения и не должно отображаться в списке вложений.</span><span class="sxs-lookup"><span data-stu-id="91d32-p147">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="91d32-846">String</span><span class="sxs-lookup"><span data-stu-id="91d32-846">String</span></span>||<span data-ttu-id="91d32-p148">Используется, только если свойству `type` задано значение `item`. Идентификатор вложения EWS. Это строка длиной до 100 символов.</span><span class="sxs-lookup"><span data-stu-id="91d32-p148">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="91d32-850">function</span><span class="sxs-lookup"><span data-stu-id="91d32-850">function</span></span>|<span data-ttu-id="91d32-851">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="91d32-851">&lt;optional&gt;</span></span>|<span data-ttu-id="91d32-852">По завершении работы метода функция, переданная параметру `callback`, вызывается с помощью одного параметра `asyncResult`, представляющего собой объект [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="91d32-852">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="91d32-853">Требования</span><span class="sxs-lookup"><span data-stu-id="91d32-853">Requirements</span></span>

|<span data-ttu-id="91d32-854">Требование</span><span class="sxs-lookup"><span data-stu-id="91d32-854">Requirement</span></span>|<span data-ttu-id="91d32-855">Значение</span><span class="sxs-lookup"><span data-stu-id="91d32-855">Value</span></span>|
|---|---|
|[<span data-ttu-id="91d32-856">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="91d32-856">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="91d32-857">1.0</span><span class="sxs-lookup"><span data-stu-id="91d32-857">1.0</span></span>|
|[<span data-ttu-id="91d32-858">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="91d32-858">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="91d32-859">ReadItem</span><span class="sxs-lookup"><span data-stu-id="91d32-859">ReadItem</span></span>|
|[<span data-ttu-id="91d32-860">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="91d32-860">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="91d32-861">Чтение</span><span class="sxs-lookup"><span data-stu-id="91d32-861">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="91d32-862">Примеры</span><span class="sxs-lookup"><span data-stu-id="91d32-862">Examples</span></span>

<span data-ttu-id="91d32-863">Приведенный ниже код передает строку в функцию `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="91d32-863">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="91d32-864">Ответ с пустым текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="91d32-864">Reply with an empty body.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="91d32-865">Ответ только с текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="91d32-865">Reply with just a body.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="91d32-866">Ответ с текстом сообщения и вложенным файлом.</span><span class="sxs-lookup"><span data-stu-id="91d32-866">Reply with a body and a file attachment.</span></span>

```
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

<span data-ttu-id="91d32-867">Ответ с текстом сообщения и вложенным элементом.</span><span class="sxs-lookup"><span data-stu-id="91d32-867">Reply with a body and an item attachment.</span></span>

```
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

<span data-ttu-id="91d32-868">Ответ с текстом сообщения, вложенным файлом, вложенным элементом и обратным вызовом.</span><span class="sxs-lookup"><span data-stu-id="91d32-868">Reply with a body, file attachment, item attachment, and a callback.</span></span>

```
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

#### <a name="displayreplyformformdata"></a><span data-ttu-id="91d32-869">displayReplyForm(formData)</span><span class="sxs-lookup"><span data-stu-id="91d32-869">displayReplyForm(formData)</span></span>

<span data-ttu-id="91d32-870">Отображает форму ответа, включающую только отправителя выбранного сообщения или организатора выбранной встречи.</span><span class="sxs-lookup"><span data-stu-id="91d32-870">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="91d32-871">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="91d32-871">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="91d32-872">В веб-приложении Outlook форма ответа отображается в виде всплывающей формы в представлении с 3 колонками либо всплывающей формы в представлении с 1 или 2 колонками.</span><span class="sxs-lookup"><span data-stu-id="91d32-872">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="91d32-873">Если любой строчный параметр превышает указанные для него ограничения, `displayReplyForm` возвращает исключение.</span><span class="sxs-lookup"><span data-stu-id="91d32-873">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="91d32-p149">Если в параметре `formData.attachments` указаны вложения, Outlook и веб-приложение Outlook пытаются скачать их и вложить в форму ответа. Если какие-либо вложения добавить не удается, в форме отображается сообщение об ошибке. Если сообщения об ошибках не предусмотрены, то они не отображаются.</span><span class="sxs-lookup"><span data-stu-id="91d32-p149">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="91d32-877">Параметры:</span><span class="sxs-lookup"><span data-stu-id="91d32-877">Parameters:</span></span>

|<span data-ttu-id="91d32-878">Имя</span><span class="sxs-lookup"><span data-stu-id="91d32-878">Name</span></span>|<span data-ttu-id="91d32-879">Тип</span><span class="sxs-lookup"><span data-stu-id="91d32-879">Type</span></span>|<span data-ttu-id="91d32-880">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="91d32-880">Attributes</span></span>|<span data-ttu-id="91d32-881">Описание</span><span class="sxs-lookup"><span data-stu-id="91d32-881">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="91d32-882">String | Object</span><span class="sxs-lookup"><span data-stu-id="91d32-882">String &#124; Object</span></span>||<span data-ttu-id="91d32-p150">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="91d32-p150">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="91d32-885">**ИЛИ**</span><span class="sxs-lookup"><span data-stu-id="91d32-885">**OR**</span></span><br/><span data-ttu-id="91d32-p151">Объект, который содержит текст или данные вложения и функцию обратного вызова. Ниже представлено определение этого объекта.</span><span class="sxs-lookup"><span data-stu-id="91d32-p151">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="91d32-888">String</span><span class="sxs-lookup"><span data-stu-id="91d32-888">String</span></span>|<span data-ttu-id="91d32-889">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="91d32-889">&lt;optional&gt;</span></span>|<span data-ttu-id="91d32-p152">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="91d32-p152">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="91d32-892">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="91d32-892">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="91d32-893">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="91d32-893">&lt;optional&gt;</span></span>|<span data-ttu-id="91d32-894">Массив объектов JSON, представляющих собой вложенные файлы или элементы.</span><span class="sxs-lookup"><span data-stu-id="91d32-894">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="91d32-895">String</span><span class="sxs-lookup"><span data-stu-id="91d32-895">String</span></span>||<span data-ttu-id="91d32-p153">Указывает тип вложения. Допустимые значения: `file` для вложенного файла и `item` для вложенного элемента.</span><span class="sxs-lookup"><span data-stu-id="91d32-p153">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="91d32-898">String</span><span class="sxs-lookup"><span data-stu-id="91d32-898">String</span></span>||<span data-ttu-id="91d32-899">Строка, содержащая имя вложения, длиной до 255 символов.</span><span class="sxs-lookup"><span data-stu-id="91d32-899">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="91d32-900">String</span><span class="sxs-lookup"><span data-stu-id="91d32-900">String</span></span>||<span data-ttu-id="91d32-p154">Используется, только если свойству `type` задано значение `file`. URI расположения файла.</span><span class="sxs-lookup"><span data-stu-id="91d32-p154">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="91d32-903">Логическое значение</span><span class="sxs-lookup"><span data-stu-id="91d32-903">Boolean</span></span>||<span data-ttu-id="91d32-p155">Используется только в том случае, если свойству `type` задано значение `file`. Значение `true` указывает на то, что вложение будет встроено в текст сообщения и не должно отображаться в списке вложений.</span><span class="sxs-lookup"><span data-stu-id="91d32-p155">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="91d32-906">String</span><span class="sxs-lookup"><span data-stu-id="91d32-906">String</span></span>||<span data-ttu-id="91d32-p156">Используется, только если свойству `type` задано значение `item`. Идентификатор вложения EWS. Это строка длиной до 100 символов.</span><span class="sxs-lookup"><span data-stu-id="91d32-p156">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="91d32-910">function</span><span class="sxs-lookup"><span data-stu-id="91d32-910">function</span></span>|<span data-ttu-id="91d32-911">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="91d32-911">&lt;optional&gt;</span></span>|<span data-ttu-id="91d32-912">По завершении работы метода функция, переданная параметру `callback`, вызывается с помощью одного параметра `asyncResult`, представляющего собой объект [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="91d32-912">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="91d32-913">Требования</span><span class="sxs-lookup"><span data-stu-id="91d32-913">Requirements</span></span>

|<span data-ttu-id="91d32-914">Требование</span><span class="sxs-lookup"><span data-stu-id="91d32-914">Requirement</span></span>|<span data-ttu-id="91d32-915">Значение</span><span class="sxs-lookup"><span data-stu-id="91d32-915">Value</span></span>|
|---|---|
|[<span data-ttu-id="91d32-916">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="91d32-916">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="91d32-917">1.0</span><span class="sxs-lookup"><span data-stu-id="91d32-917">1.0</span></span>|
|[<span data-ttu-id="91d32-918">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="91d32-918">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="91d32-919">ReadItem</span><span class="sxs-lookup"><span data-stu-id="91d32-919">ReadItem</span></span>|
|[<span data-ttu-id="91d32-920">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="91d32-920">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="91d32-921">Чтение</span><span class="sxs-lookup"><span data-stu-id="91d32-921">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="91d32-922">Примеры</span><span class="sxs-lookup"><span data-stu-id="91d32-922">Examples</span></span>

<span data-ttu-id="91d32-923">Приведенный ниже код передает строку в функцию `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="91d32-923">The following code passes a string to the `displayReplyForm` function.</span></span>

```
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="91d32-924">Ответ с пустым текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="91d32-924">Reply with an empty body.</span></span>

```
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="91d32-925">Ответ только с текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="91d32-925">Reply with just a body.</span></span>

```
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="91d32-926">Ответ с текстом сообщения и вложенным файлом.</span><span class="sxs-lookup"><span data-stu-id="91d32-926">Reply with a body and a file attachment.</span></span>

```
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

<span data-ttu-id="91d32-927">Ответ с текстом сообщения и вложенным элементом.</span><span class="sxs-lookup"><span data-stu-id="91d32-927">Reply with a body and an item attachment.</span></span>

```
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

<span data-ttu-id="91d32-928">Ответ с текстом сообщения, вложенным файлом, вложенным элементом и обратным вызовом.</span><span class="sxs-lookup"><span data-stu-id="91d32-928">Reply with a body, file attachment, item attachment, and a callback.</span></span>

```
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

#### <a name="getentities--entitiesjavascriptapioutlook17officeentities"></a><span data-ttu-id="91d32-929">getEntities() → {[Entities](/javascript/api/outlook_1_7/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="91d32-929">getEntities() → {[Entities](/javascript/api/outlook_1_7/office.entities)}</span></span>

<span data-ttu-id="91d32-930">Получает сущности, обнаруженные в выбранном тексте элемента.</span><span class="sxs-lookup"><span data-stu-id="91d32-930">Gets the entities found in the selected item.</span></span>

> [!NOTE]
> <span data-ttu-id="91d32-931">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="91d32-931">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="91d32-932">Требования</span><span class="sxs-lookup"><span data-stu-id="91d32-932">Requirements</span></span>

|<span data-ttu-id="91d32-933">Требование</span><span class="sxs-lookup"><span data-stu-id="91d32-933">Requirement</span></span>|<span data-ttu-id="91d32-934">Значение</span><span class="sxs-lookup"><span data-stu-id="91d32-934">Value</span></span>|
|---|---|
|[<span data-ttu-id="91d32-935">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="91d32-935">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="91d32-936">1.0</span><span class="sxs-lookup"><span data-stu-id="91d32-936">1.0</span></span>|
|[<span data-ttu-id="91d32-937">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="91d32-937">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="91d32-938">ReadItem</span><span class="sxs-lookup"><span data-stu-id="91d32-938">ReadItem</span></span>|
|[<span data-ttu-id="91d32-939">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="91d32-939">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="91d32-940">Чтение</span><span class="sxs-lookup"><span data-stu-id="91d32-940">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="91d32-941">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="91d32-941">Returns:</span></span>

<span data-ttu-id="91d32-942">Тип: [Entities](/javascript/api/outlook_1_7/office.entities)</span><span class="sxs-lookup"><span data-stu-id="91d32-942">Type: [Entities](/javascript/api/outlook_1_7/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="91d32-943">Пример</span><span class="sxs-lookup"><span data-stu-id="91d32-943">Example</span></span>

<span data-ttu-id="91d32-944">Ниже приведен пример получения доступа к сущностям контактов в тексте текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="91d32-944">The following example accesses the contacts entities on the current item.</span></span>

```
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook17officecontactmeetingsuggestionjavascriptapioutlook17officemeetingsuggestionphonenumberjavascriptapioutlook17officephonenumbertasksuggestionjavascriptapioutlook17officetasksuggestion"></a><span data-ttu-id="91d32-945">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="91d32-945">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))>}</span></span>

<span data-ttu-id="91d32-946">Получает массив всех сущностей указанного типа, обнаруженных в тексте выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="91d32-946">Gets an array of all the entities of the specified entity type found in the selected item.</span></span>

> [!NOTE]
> <span data-ttu-id="91d32-947">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="91d32-947">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="91d32-948">Параметры:</span><span class="sxs-lookup"><span data-stu-id="91d32-948">Parameters:</span></span>

|<span data-ttu-id="91d32-949">Имя</span><span class="sxs-lookup"><span data-stu-id="91d32-949">Name</span></span>|<span data-ttu-id="91d32-950">Тип</span><span class="sxs-lookup"><span data-stu-id="91d32-950">Type</span></span>|<span data-ttu-id="91d32-951">Описание</span><span class="sxs-lookup"><span data-stu-id="91d32-951">Description</span></span>|
|---|---|---|
|`entityType`|[<span data-ttu-id="91d32-952">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="91d32-952">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook_1_7/office.mailboxenums.entitytype)|<span data-ttu-id="91d32-953">Одно из значений перечисления EntityType.</span><span class="sxs-lookup"><span data-stu-id="91d32-953">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="91d32-954">Требования</span><span class="sxs-lookup"><span data-stu-id="91d32-954">Requirements</span></span>

|<span data-ttu-id="91d32-955">Требование</span><span class="sxs-lookup"><span data-stu-id="91d32-955">Requirement</span></span>|<span data-ttu-id="91d32-956">Значение</span><span class="sxs-lookup"><span data-stu-id="91d32-956">Value</span></span>|
|---|---|
|[<span data-ttu-id="91d32-957">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="91d32-957">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="91d32-958">1.0</span><span class="sxs-lookup"><span data-stu-id="91d32-958">1.0</span></span>|
|[<span data-ttu-id="91d32-959">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="91d32-959">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="91d32-960">Ограниченный доступ</span><span class="sxs-lookup"><span data-stu-id="91d32-960">Restricted</span></span>|
|[<span data-ttu-id="91d32-961">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="91d32-961">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="91d32-962">Чтение</span><span class="sxs-lookup"><span data-stu-id="91d32-962">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="91d32-963">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="91d32-963">Returns:</span></span>

<span data-ttu-id="91d32-964">Если значение, переданное в `entityType`, не является допустимым членом перечисления `EntityType`, метод возвращает значение NULL.</span><span class="sxs-lookup"><span data-stu-id="91d32-964">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="91d32-965">Если в тексте элемента отсутствуют сущности указанного типа, метод возвращает пустой массив.</span><span class="sxs-lookup"><span data-stu-id="91d32-965">If no entities of the specified type are present on the item, the method returns an empty array.</span></span> <span data-ttu-id="91d32-966">В противном случае тип объектов в возвращаемом массиве зависит от типа сущности, запрошенной в параметре `entityType`.</span><span class="sxs-lookup"><span data-stu-id="91d32-966">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="91d32-967">Хотя минимальный уровень разрешений для использования этого метода — **Restricted**, для некоторых типов сущностей требуется доступ на уровне **ReadItem**, как указано в приведенной ниже таблице.</span><span class="sxs-lookup"><span data-stu-id="91d32-967">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

|<span data-ttu-id="91d32-968">Значение параметра `entityType`</span><span class="sxs-lookup"><span data-stu-id="91d32-968">Value of `entityType`</span></span>|<span data-ttu-id="91d32-969">Тип объектов в возвращаемом массиве</span><span class="sxs-lookup"><span data-stu-id="91d32-969">Type of objects in returned array</span></span>|<span data-ttu-id="91d32-970">Необходимый уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="91d32-970">Required Permission Level</span></span>|
|---|---|---|
|`Address`|<span data-ttu-id="91d32-971">String</span><span class="sxs-lookup"><span data-stu-id="91d32-971">String</span></span>|<span data-ttu-id="91d32-972">**С ограничениями**</span><span class="sxs-lookup"><span data-stu-id="91d32-972">**Restricted**</span></span>|
|`Contact`|<span data-ttu-id="91d32-973">Contact</span><span class="sxs-lookup"><span data-stu-id="91d32-973">Contact</span></span>|<span data-ttu-id="91d32-974">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="91d32-974">**ReadItem**</span></span>|
|`EmailAddress`|<span data-ttu-id="91d32-975">String</span><span class="sxs-lookup"><span data-stu-id="91d32-975">String</span></span>|<span data-ttu-id="91d32-976">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="91d32-976">**ReadItem**</span></span>|
|`MeetingSuggestion`|<span data-ttu-id="91d32-977">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="91d32-977">MeetingSuggestion</span></span>|<span data-ttu-id="91d32-978">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="91d32-978">**ReadItem**</span></span>|
|`PhoneNumber`|<span data-ttu-id="91d32-979">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="91d32-979">PhoneNumber</span></span>|<span data-ttu-id="91d32-980">**С ограничениями**</span><span class="sxs-lookup"><span data-stu-id="91d32-980">**Restricted**</span></span>|
|`TaskSuggestion`|<span data-ttu-id="91d32-981">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="91d32-981">TaskSuggestion</span></span>|<span data-ttu-id="91d32-982">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="91d32-982">**ReadItem**</span></span>|
|`URL`|<span data-ttu-id="91d32-983">String</span><span class="sxs-lookup"><span data-stu-id="91d32-983">String</span></span>|<span data-ttu-id="91d32-984">**С ограничениями**</span><span class="sxs-lookup"><span data-stu-id="91d32-984">**Restricted**</span></span>|

<span data-ttu-id="91d32-985">Type: Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="91d32-985">Type: Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="91d32-986">Пример</span><span class="sxs-lookup"><span data-stu-id="91d32-986">Example</span></span>

<span data-ttu-id="91d32-987">В примере ниже показано, как получить доступ к массиву строк, которые представляют собой почтовые адреса в тексте текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="91d32-987">The following example shows how to access an array of strings that represent postal addresses in the subject or body of the current item.</span></span>

```
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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook17officecontactmeetingsuggestionjavascriptapioutlook17officemeetingsuggestionphonenumberjavascriptapioutlook17officephonenumbertasksuggestionjavascriptapioutlook17officetasksuggestion"></a><span data-ttu-id="91d32-988">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="91d32-988">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))>}</span></span>

<span data-ttu-id="91d32-989">Возвращает известные сущности в выбранном элементе, которые проходят через именованный фильтр, определяемый в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="91d32-989">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="91d32-990">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="91d32-990">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="91d32-991">Метод `getFilteredEntitiesByName` возвращает сущности, соответствующие регулярному выражению, которое определяется в элементе правила [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) в XML-файле манифеста, с использованием указанного значения элемента `FilterName`.</span><span class="sxs-lookup"><span data-stu-id="91d32-991">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="91d32-992">Параметры:</span><span class="sxs-lookup"><span data-stu-id="91d32-992">Parameters:</span></span>

|<span data-ttu-id="91d32-993">Имя</span><span class="sxs-lookup"><span data-stu-id="91d32-993">Name</span></span>|<span data-ttu-id="91d32-994">Тип</span><span class="sxs-lookup"><span data-stu-id="91d32-994">Type</span></span>|<span data-ttu-id="91d32-995">Описание</span><span class="sxs-lookup"><span data-stu-id="91d32-995">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="91d32-996">String</span><span class="sxs-lookup"><span data-stu-id="91d32-996">String</span></span>|<span data-ttu-id="91d32-997">Имя элемента правила `ItemHasKnownEntity`, определяющее соответствующий фильтр.</span><span class="sxs-lookup"><span data-stu-id="91d32-997">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="91d32-998">Требования</span><span class="sxs-lookup"><span data-stu-id="91d32-998">Requirements</span></span>

|<span data-ttu-id="91d32-999">Требование</span><span class="sxs-lookup"><span data-stu-id="91d32-999">Requirement</span></span>|<span data-ttu-id="91d32-1000">Значение</span><span class="sxs-lookup"><span data-stu-id="91d32-1000">Value</span></span>|
|---|---|
|[<span data-ttu-id="91d32-1001">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="91d32-1001">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="91d32-1002">1.0</span><span class="sxs-lookup"><span data-stu-id="91d32-1002">1.0</span></span>|
|[<span data-ttu-id="91d32-1003">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="91d32-1003">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="91d32-1004">ReadItem</span><span class="sxs-lookup"><span data-stu-id="91d32-1004">ReadItem</span></span>|
|[<span data-ttu-id="91d32-1005">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="91d32-1005">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="91d32-1006">Чтение</span><span class="sxs-lookup"><span data-stu-id="91d32-1006">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="91d32-1007">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="91d32-1007">Returns:</span></span>

<span data-ttu-id="91d32-p158">Если в манифесте нет элемента `ItemHasKnownEntity` со значением `FilterName`, соответствующим параметру `name`, метод возвращает `null`. Если параметр `name` соответствует элементу `ItemHasKnownEntity` в манифесте, но при этом в текущем элементе нет соответствующих сущностей, метод возвращает пустой массив.</span><span class="sxs-lookup"><span data-stu-id="91d32-p158">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="91d32-1010">Тип: Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="91d32-1010">Type: Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))></span></span>

#### <a name="getregexmatches--object"></a><span data-ttu-id="91d32-1011">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="91d32-1011">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="91d32-1012">Возвращает строчные значения в выбранном элементе, которые соответствуют регулярным выражениям, определенным в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="91d32-1012">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="91d32-1013">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="91d32-1013">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="91d32-p159">Метод `getRegExMatches` возвращает строки, соответствующие регулярному выражению, которое определяется в каждом элементе правила `ItemHasRegularExpressionMatch` или `ItemHasKnownEntity` в XML-файле манифеста. Для правила `ItemHasRegularExpressionMatch` свойство элемента, указанного этим правилом, должно содержать соответствующую строку. Простой тип `PropertyName` определяет поддерживаемые свойства.</span><span class="sxs-lookup"><span data-stu-id="91d32-p159">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="91d32-1017">Например, рассмотрим манифест надстройки, который содержит указанный ниже элемент `Rule`.</span><span class="sxs-lookup"><span data-stu-id="91d32-1017">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="91d32-1018">Объект, возвращаемый методом `getRegExMatches`, будет содержать два свойства: `fruits` и `veggies`.</span><span class="sxs-lookup"><span data-stu-id="91d32-1018">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="91d32-p160">Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты. Лучше используйте для этого метод [`Body.getAsync`](/javascript/api/outlook_1_7/office.body#getasync-coerciontype--options--callback-).</span><span class="sxs-lookup"><span data-stu-id="91d32-p160">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook_1_7/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="91d32-1022">Требования</span><span class="sxs-lookup"><span data-stu-id="91d32-1022">Requirements</span></span>

|<span data-ttu-id="91d32-1023">Требование</span><span class="sxs-lookup"><span data-stu-id="91d32-1023">Requirement</span></span>|<span data-ttu-id="91d32-1024">Значение</span><span class="sxs-lookup"><span data-stu-id="91d32-1024">Value</span></span>|
|---|---|
|[<span data-ttu-id="91d32-1025">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="91d32-1025">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="91d32-1026">1.0</span><span class="sxs-lookup"><span data-stu-id="91d32-1026">1.0</span></span>|
|[<span data-ttu-id="91d32-1027">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="91d32-1027">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="91d32-1028">ReadItem</span><span class="sxs-lookup"><span data-stu-id="91d32-1028">ReadItem</span></span>|
|[<span data-ttu-id="91d32-1029">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="91d32-1029">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="91d32-1030">Чтение</span><span class="sxs-lookup"><span data-stu-id="91d32-1030">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="91d32-1031">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="91d32-1031">Returns:</span></span>

<span data-ttu-id="91d32-p161">Объект, содержащий массив строк, которые соответствуют регулярным выражениям, определяемым в XML-файле манифеста. Имя каждого массива равно соответствующему значению атрибута `RegExName` правила сопоставления `ItemHasRegularExpressionMatch` или атрибута `FilterName` правила сопоставления `ItemHasKnownEntity`.</span><span class="sxs-lookup"><span data-stu-id="91d32-p161">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="91d32-1034">

<dt>Тип</dt>

</span><span class="sxs-lookup"><span data-stu-id="91d32-1034">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="91d32-1035">Объект</span><span class="sxs-lookup"><span data-stu-id="91d32-1035">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="91d32-1036">Пример</span><span class="sxs-lookup"><span data-stu-id="91d32-1036">Example</span></span>

<span data-ttu-id="91d32-1037">В приведенном ниже примере показано, как получить доступ к массиву совпадений с элементами `fruits` и `veggies` правил активации регулярных выражений, указанными в манифесте.</span><span class="sxs-lookup"><span data-stu-id="91d32-1037">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veges = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="91d32-1038">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="91d32-1038">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="91d32-1039">Возвращает строчные значения в выбранном элементе, которые соответствуют именованному регулярному выражению, определенному в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="91d32-1039">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="91d32-1040">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="91d32-1040">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="91d32-1041">Метод `getRegExMatchesByName` возвращает строки, соответствующие регулярному выражению, которое определяется в элементе правила `ItemHasRegularExpressionMatch` в XML-файле манифеста, с использованием указанного значения элемента `RegExName`.</span><span class="sxs-lookup"><span data-stu-id="91d32-1041">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="91d32-p162">Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты.</span><span class="sxs-lookup"><span data-stu-id="91d32-p162">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="91d32-1044">Параметры:</span><span class="sxs-lookup"><span data-stu-id="91d32-1044">Parameters:</span></span>

|<span data-ttu-id="91d32-1045">Имя</span><span class="sxs-lookup"><span data-stu-id="91d32-1045">Name</span></span>|<span data-ttu-id="91d32-1046">Тип</span><span class="sxs-lookup"><span data-stu-id="91d32-1046">Type</span></span>|<span data-ttu-id="91d32-1047">Описание</span><span class="sxs-lookup"><span data-stu-id="91d32-1047">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="91d32-1048">String</span><span class="sxs-lookup"><span data-stu-id="91d32-1048">String</span></span>|<span data-ttu-id="91d32-1049">Имя элемента правила `ItemHasRegularExpressionMatch`, определяющее соответствующий фильтр.</span><span class="sxs-lookup"><span data-stu-id="91d32-1049">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="91d32-1050">Требования</span><span class="sxs-lookup"><span data-stu-id="91d32-1050">Requirements</span></span>

|<span data-ttu-id="91d32-1051">Требование</span><span class="sxs-lookup"><span data-stu-id="91d32-1051">Requirement</span></span>|<span data-ttu-id="91d32-1052">Значение</span><span class="sxs-lookup"><span data-stu-id="91d32-1052">Value</span></span>|
|---|---|
|[<span data-ttu-id="91d32-1053">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="91d32-1053">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="91d32-1054">1.0</span><span class="sxs-lookup"><span data-stu-id="91d32-1054">1.0</span></span>|
|[<span data-ttu-id="91d32-1055">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="91d32-1055">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="91d32-1056">ReadItem</span><span class="sxs-lookup"><span data-stu-id="91d32-1056">ReadItem</span></span>|
|[<span data-ttu-id="91d32-1057">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="91d32-1057">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="91d32-1058">Чтение</span><span class="sxs-lookup"><span data-stu-id="91d32-1058">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="91d32-1059">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="91d32-1059">Returns:</span></span>

<span data-ttu-id="91d32-1060">Массив строк, соответствующих регулярному выражению, определяемому в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="91d32-1060">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="91d32-1061">

<dt>Тип</dt>

</span><span class="sxs-lookup"><span data-stu-id="91d32-1061">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="91d32-1062">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="91d32-1062">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="91d32-1063">Пример</span><span class="sxs-lookup"><span data-stu-id="91d32-1063">Example</span></span>

```
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="91d32-1064">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="91d32-1064">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="91d32-1065">Асинхронно возвращает данные, выбранные в теме или тексте сообщения.</span><span class="sxs-lookup"><span data-stu-id="91d32-1065">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="91d32-p163">Если выделенный фрагмент отсутствует, но курсор находится в тексте или теме, метод возвращает значение NULL для выбранных данных. Если выбраны не текст и не тема, метод возвращает ошибку `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="91d32-p163">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="91d32-1068">Параметры:</span><span class="sxs-lookup"><span data-stu-id="91d32-1068">Parameters:</span></span>

|<span data-ttu-id="91d32-1069">Имя</span><span class="sxs-lookup"><span data-stu-id="91d32-1069">Name</span></span>|<span data-ttu-id="91d32-1070">Тип</span><span class="sxs-lookup"><span data-stu-id="91d32-1070">Type</span></span>|<span data-ttu-id="91d32-1071">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="91d32-1071">Attributes</span></span>|<span data-ttu-id="91d32-1072">Описание</span><span class="sxs-lookup"><span data-stu-id="91d32-1072">Description</span></span>|
|---|---|---|---|
|`coercionType`|[<span data-ttu-id="91d32-1073">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="91d32-1073">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="91d32-p164">Запрашивает формат данных. Если задано значение Text, метод возвращает обычный текст как строку, удаляя все имеющиеся HTML-теги. Если задано значение HTML, метод возвращает выделенный текст (обычный текст или HTML).</span><span class="sxs-lookup"><span data-stu-id="91d32-p164">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`|<span data-ttu-id="91d32-1077">Объект</span><span class="sxs-lookup"><span data-stu-id="91d32-1077">Object</span></span>|<span data-ttu-id="91d32-1078">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="91d32-1078">&lt;optional&gt;</span></span>|<span data-ttu-id="91d32-1079">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="91d32-1079">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="91d32-1080">Объект</span><span class="sxs-lookup"><span data-stu-id="91d32-1080">Object</span></span>|<span data-ttu-id="91d32-1081">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="91d32-1081">&lt;optional&gt;</span></span>|<span data-ttu-id="91d32-1082">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="91d32-1082">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="91d32-1083">функция</span><span class="sxs-lookup"><span data-stu-id="91d32-1083">function</span></span>||<span data-ttu-id="91d32-1084">Когда метод завершает выполнение, переданная в параметре `callback` функция вызывается с единственным параметром `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="91d32-1084">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="91d32-1085">Чтобы получить доступ к выбранным данным из метода обратного вызова, вызовите `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="91d32-1085">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="91d32-1086">Для доступа к исходному свойству, на основе которого созданы выбранные данные, вызовите  `asyncResult.value.sourceProperty`, который может иметь значение `body` или `subject`.</span><span class="sxs-lookup"><span data-stu-id="91d32-1086">To access the source property that the selection comes from, call , which will be either  or .|</span></span>|

##### <a name="requirements"></a><span data-ttu-id="91d32-1087">Требования</span><span class="sxs-lookup"><span data-stu-id="91d32-1087">Requirements</span></span>

|<span data-ttu-id="91d32-1088">Требование</span><span class="sxs-lookup"><span data-stu-id="91d32-1088">Requirement</span></span>|<span data-ttu-id="91d32-1089">Значение</span><span class="sxs-lookup"><span data-stu-id="91d32-1089">Value</span></span>|
|---|---|
|[<span data-ttu-id="91d32-1090">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="91d32-1090">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="91d32-1091">1.2</span><span class="sxs-lookup"><span data-stu-id="91d32-1091">1.2</span></span>|
|[<span data-ttu-id="91d32-1092">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="91d32-1092">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="91d32-1093">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="91d32-1093">ReadWriteItem</span></span>|
|[<span data-ttu-id="91d32-1094">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="91d32-1094">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="91d32-1095">Создание</span><span class="sxs-lookup"><span data-stu-id="91d32-1095">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="91d32-1096">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="91d32-1096">Returns:</span></span>

<span data-ttu-id="91d32-1097">Выбранные данные в виде строки с форматом, определенным в параметре `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="91d32-1097">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="91d32-1098">

<dt>Тип</dt>

</span><span class="sxs-lookup"><span data-stu-id="91d32-1098">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="91d32-1099">String</span><span class="sxs-lookup"><span data-stu-id="91d32-1099">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="91d32-1100">Пример</span><span class="sxs-lookup"><span data-stu-id="91d32-1100">Example</span></span>

```
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

#### <a name="getselectedentities--entitiesjavascriptapioutlook17officeentities"></a><span data-ttu-id="91d32-1101">getSelectedEntities() → {[Entities](/javascript/api/outlook_1_7/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="91d32-1101">getSelectedEntities() → {[Entities](/javascript/api/outlook_1_7/office.entities)}</span></span>

<span data-ttu-id="91d32-p166">Возвращает сущности, найденные в выделенном совпадении, выбранном пользователем. Выделенные совпадения применяются к [контекстным надстройкам](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="91d32-p166">Gets the entities found in a highlighted match a user has selected. Highlighted matches apply to [contextual add-ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="91d32-1104">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="91d32-1104">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="91d32-1105">Требования</span><span class="sxs-lookup"><span data-stu-id="91d32-1105">Requirements</span></span>

|<span data-ttu-id="91d32-1106">Требование</span><span class="sxs-lookup"><span data-stu-id="91d32-1106">Requirement</span></span>|<span data-ttu-id="91d32-1107">Значение</span><span class="sxs-lookup"><span data-stu-id="91d32-1107">Value</span></span>|
|---|---|
|[<span data-ttu-id="91d32-1108">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="91d32-1108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="91d32-1109">1.6</span><span class="sxs-lookup"><span data-stu-id="91d32-1109">1.6</span></span>|
|[<span data-ttu-id="91d32-1110">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="91d32-1110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="91d32-1111">ReadItem</span><span class="sxs-lookup"><span data-stu-id="91d32-1111">ReadItem</span></span>|
|[<span data-ttu-id="91d32-1112">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="91d32-1112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="91d32-1113">Чтение</span><span class="sxs-lookup"><span data-stu-id="91d32-1113">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="91d32-1114">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="91d32-1114">Returns:</span></span>

<span data-ttu-id="91d32-1115">Тип: [Entities](/javascript/api/outlook_1_7/office.entities)</span><span class="sxs-lookup"><span data-stu-id="91d32-1115">Type: [Entities](/javascript/api/outlook_1_7/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="91d32-1116">Пример</span><span class="sxs-lookup"><span data-stu-id="91d32-1116">Example</span></span>

<span data-ttu-id="91d32-1117">В приведенном ниже примере показано, как получить доступ к сущностям адресов в выделенном совпадении, выбранном пользователем.</span><span class="sxs-lookup"><span data-stu-id="91d32-1117">The following example accesses the addresses entities in the highlighted match selected by the user.</span></span>

```
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

#### <a name="getselectedregexmatches--object"></a><span data-ttu-id="91d32-1118">getSelectedRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="91d32-1118">getSelectedRegExMatches() → {Object}</span></span>

<span data-ttu-id="91d32-p167">Возвращает строковые значения в выделенном совпадении, которые соответствуют регулярным выражениям, определенным в XML-файле манифеста. Выделенные совпадения применяются к [контекстным надстройкам](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="91d32-p167">Returns string values in a highlighted match that match the regular expressions defined in the manifest XML file. Highlighted matches apply to [contextual add-ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="91d32-1121">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="91d32-1121">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="91d32-p168">Метод `getSelectedRegExMatches` возвращает строки, соответствующие регулярному выражению, которое определяется в каждом элементе правила `ItemHasRegularExpressionMatch` или `ItemHasKnownEntity` в XML-файле манифеста. Для правила `ItemHasRegularExpressionMatch` свойство элемента, указанного этим правилом, должно содержать соответствующую строку. Простой тип `PropertyName` определяет поддерживаемые свойства.</span><span class="sxs-lookup"><span data-stu-id="91d32-p168">The `getSelectedRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="91d32-1125">Например, рассмотрим манифест надстройки, который содержит указанный ниже элемент `Rule`.</span><span class="sxs-lookup"><span data-stu-id="91d32-1125">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="91d32-1126">Объект, возвращаемый методом `getRegExMatches`, будет содержать два свойства: `fruits` и `veggies`.</span><span class="sxs-lookup"><span data-stu-id="91d32-1126">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="91d32-p169">Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты. Лучше используйте для этого метод [`Body.getAsync`](/javascript/api/outlook_1_7/office.body#getasync-coerciontype--options--callback-).</span><span class="sxs-lookup"><span data-stu-id="91d32-p169">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook_1_7/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="91d32-1130">Требования</span><span class="sxs-lookup"><span data-stu-id="91d32-1130">Requirements</span></span>

|<span data-ttu-id="91d32-1131">Требование</span><span class="sxs-lookup"><span data-stu-id="91d32-1131">Requirement</span></span>|<span data-ttu-id="91d32-1132">Значение</span><span class="sxs-lookup"><span data-stu-id="91d32-1132">Value</span></span>|
|---|---|
|[<span data-ttu-id="91d32-1133">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="91d32-1133">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="91d32-1134">1.6</span><span class="sxs-lookup"><span data-stu-id="91d32-1134">1.6</span></span>|
|[<span data-ttu-id="91d32-1135">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="91d32-1135">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="91d32-1136">ReadItem</span><span class="sxs-lookup"><span data-stu-id="91d32-1136">ReadItem</span></span>|
|[<span data-ttu-id="91d32-1137">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="91d32-1137">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="91d32-1138">Чтение</span><span class="sxs-lookup"><span data-stu-id="91d32-1138">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="91d32-1139">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="91d32-1139">Returns:</span></span>

<span data-ttu-id="91d32-p170">Объект, содержащий массив строк, которые соответствуют регулярным выражениям, определяемым в XML-файле манифеста. Имя каждого массива равно соответствующему значению атрибута `RegExName` правила сопоставления `ItemHasRegularExpressionMatch` или атрибута `FilterName` правила сопоставления `ItemHasKnownEntity`.</span><span class="sxs-lookup"><span data-stu-id="91d32-p170">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

##### <a name="example"></a><span data-ttu-id="91d32-1142">Пример</span><span class="sxs-lookup"><span data-stu-id="91d32-1142">Example</span></span>

<span data-ttu-id="91d32-1143">В приведенном ниже примере показано, как получить доступ к массиву совпадений с элементами `fruits` и `veggies` правил активации регулярных выражений, указанными в манифесте.</span><span class="sxs-lookup"><span data-stu-id="91d32-1143">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="91d32-1144">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="91d32-1144">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="91d32-1145">Асинхронно загружает настраиваемые свойства для надстройки выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="91d32-1145">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="91d32-p171">Настраиваемые свойства сохраняются в виде пар "ключ-значение" для каждого приложения и каждого элемента. Этот метод возвращает объект `CustomProperties` при обратном вызове, который предоставляет методы для доступа к настраиваемым свойствам, характерным для текущего элемента и текущей надстройки. Настраиваемые свойства не шифруются для элемента, поэтому этот способ хранения не является безопасным.</span><span class="sxs-lookup"><span data-stu-id="91d32-p171">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="91d32-1149">Параметры:</span><span class="sxs-lookup"><span data-stu-id="91d32-1149">Parameters:</span></span>

|<span data-ttu-id="91d32-1150">Имя</span><span class="sxs-lookup"><span data-stu-id="91d32-1150">Name</span></span>|<span data-ttu-id="91d32-1151">Тип</span><span class="sxs-lookup"><span data-stu-id="91d32-1151">Type</span></span>|<span data-ttu-id="91d32-1152">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="91d32-1152">Attributes</span></span>|<span data-ttu-id="91d32-1153">Описание</span><span class="sxs-lookup"><span data-stu-id="91d32-1153">Description</span></span>|
|---|---|---|---|
|`callback`|<span data-ttu-id="91d32-1154">function</span><span class="sxs-lookup"><span data-stu-id="91d32-1154">function</span></span>||<span data-ttu-id="91d32-1155">Когда метод завершает выполнение, переданная в параметре `callback` функция вызывается с единственным параметром `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="91d32-1155">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="91d32-1156">Настраиваемые свойства предоставляются в виде объекта [`CustomProperties`](/javascript/api/outlook_1_7/office.customproperties) в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="91d32-1156">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook_1_7/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="91d32-1157">Этот объект можно использовать для получения, задания и удаления настраиваемых свойств из элемента и сохранения изменений настраиваемого свойства на сервере.</span><span class="sxs-lookup"><span data-stu-id="91d32-1157">The custom properties are provided as a CustomProperties object in the asyncResult.value property. This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`|<span data-ttu-id="91d32-1158">Объект</span><span class="sxs-lookup"><span data-stu-id="91d32-1158">Object</span></span>|<span data-ttu-id="91d32-1159">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="91d32-1159">&lt;optional&gt;</span></span>|<span data-ttu-id="91d32-1160">Разработчики могут указать любой объект, к которому необходимо получить доступ, в функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="91d32-1160">Developers can provide any object they wish to access in the callback method.</span></span> <span data-ttu-id="91d32-1161">Доступ к этому объекту можно получить с помощью свойства `asyncResult.asyncContext` в функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="91d32-1161">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.|</span></span>|

##### <a name="requirements"></a><span data-ttu-id="91d32-1162">Требования</span><span class="sxs-lookup"><span data-stu-id="91d32-1162">Requirements</span></span>

|<span data-ttu-id="91d32-1163">Требование</span><span class="sxs-lookup"><span data-stu-id="91d32-1163">Requirement</span></span>|<span data-ttu-id="91d32-1164">Значение</span><span class="sxs-lookup"><span data-stu-id="91d32-1164">Value</span></span>|
|---|---|
|[<span data-ttu-id="91d32-1165">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="91d32-1165">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="91d32-1166">1.0</span><span class="sxs-lookup"><span data-stu-id="91d32-1166">1.0</span></span>|
|[<span data-ttu-id="91d32-1167">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="91d32-1167">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="91d32-1168">ReadItem</span><span class="sxs-lookup"><span data-stu-id="91d32-1168">ReadItem</span></span>|
|[<span data-ttu-id="91d32-1169">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="91d32-1169">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="91d32-1170">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="91d32-1170">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="91d32-1171">Пример</span><span class="sxs-lookup"><span data-stu-id="91d32-1171">Example</span></span>

<span data-ttu-id="91d32-p174">В приведенном ниже примере кода показано, как асинхронно загружать настраиваемые свойства, характерные для текущего элемента, с помощью метода `loadCustomPropertiesAsync`. Этот пример также показывает, как сохранять эти свойства на сервере с помощью метода `CustomProperties.saveAsync`. В этом примере кода, после того как выполнена загрузка настраиваемых свойств, метод `CustomProperties.get` используется для считывания настраиваемого свойства `myProp`, метод `CustomProperties.set` — для записи настраиваемого свойства `otherProp`, а метод `saveAsync` — для сохранения настраиваемых свойств.</span><span class="sxs-lookup"><span data-stu-id="91d32-p174">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

```
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

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="91d32-1175">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="91d32-1175">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="91d32-1176">Удаляет вложение из сообщения или встречи.</span><span class="sxs-lookup"><span data-stu-id="91d32-1176">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="91d32-p175">Метод `removeAttachmentAsync` удаляет из элемента вложение с указанным идентификатором. Идентификатор вложения рекомендуется использовать для удаления вложения, только если оно добавлено тем же почтовым приложением в ходе текущего сеанса. В веб-приложении Outlook и веб-приложении Outlook для устройств идентификатор вложения действителен только в рамках одного сеанса. Сеанс завершается, когда пользователь закрывает приложение или начинает создавать элемент во встроенной форме, а затем переходит из формы в отдельное окно.</span><span class="sxs-lookup"><span data-stu-id="91d32-p175">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="91d32-1181">Параметры:</span><span class="sxs-lookup"><span data-stu-id="91d32-1181">Parameters:</span></span>

|<span data-ttu-id="91d32-1182">Имя</span><span class="sxs-lookup"><span data-stu-id="91d32-1182">Name</span></span>|<span data-ttu-id="91d32-1183">Тип</span><span class="sxs-lookup"><span data-stu-id="91d32-1183">Type</span></span>|<span data-ttu-id="91d32-1184">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="91d32-1184">Attributes</span></span>|<span data-ttu-id="91d32-1185">Описание</span><span class="sxs-lookup"><span data-stu-id="91d32-1185">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="91d32-1186">String</span><span class="sxs-lookup"><span data-stu-id="91d32-1186">String</span></span>||<span data-ttu-id="91d32-p176">Идентификатор удаляемого вложения. Максимальная длина строки — 100 символов.</span><span class="sxs-lookup"><span data-stu-id="91d32-p176">The identifier of the attachment to remove. The maximum length of the string is 100 characters.</span></span>|
|`options`|<span data-ttu-id="91d32-1189">Объект</span><span class="sxs-lookup"><span data-stu-id="91d32-1189">Object</span></span>|<span data-ttu-id="91d32-1190">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="91d32-1190">&lt;optional&gt;</span></span>|<span data-ttu-id="91d32-1191">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="91d32-1191">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="91d32-1192">Объект</span><span class="sxs-lookup"><span data-stu-id="91d32-1192">Object</span></span>|<span data-ttu-id="91d32-1193">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="91d32-1193">&lt;optional&gt;</span></span>|<span data-ttu-id="91d32-1194">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="91d32-1194">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="91d32-1195">function</span><span class="sxs-lookup"><span data-stu-id="91d32-1195">function</span></span>|<span data-ttu-id="91d32-1196">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="91d32-1196">&lt;optional&gt;</span></span>|<span data-ttu-id="91d32-1197">Когда метод завершает выполнение, переданная в параметре `callback` функция вызывается с единственным параметром `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="91d32-1197">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="91d32-1198">Если удалить вложение не удается, свойство `asyncResult.error` содержит код ошибки с указанием ее причины.</span><span class="sxs-lookup"><span data-stu-id="91d32-1198">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="91d32-1199">Ошибки</span><span class="sxs-lookup"><span data-stu-id="91d32-1199">Errors</span></span>

|<span data-ttu-id="91d32-1200">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="91d32-1200">Error code</span></span>|<span data-ttu-id="91d32-1201">Описание</span><span class="sxs-lookup"><span data-stu-id="91d32-1201">Description</span></span>|
|------------|-------------|
|`InvalidAttachmentId`|<span data-ttu-id="91d32-1202">Идентификатор вложения не существует.</span><span class="sxs-lookup"><span data-stu-id="91d32-1202">The attachment identifier does not exist.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="91d32-1203">Требования</span><span class="sxs-lookup"><span data-stu-id="91d32-1203">Requirements</span></span>

|<span data-ttu-id="91d32-1204">Требование</span><span class="sxs-lookup"><span data-stu-id="91d32-1204">Requirement</span></span>|<span data-ttu-id="91d32-1205">Значение</span><span class="sxs-lookup"><span data-stu-id="91d32-1205">Value</span></span>|
|---|---|
|[<span data-ttu-id="91d32-1206">Версия минимального набора требований для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="91d32-1206">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="91d32-1207">1.1</span><span class="sxs-lookup"><span data-stu-id="91d32-1207">1.1</span></span>|
|[<span data-ttu-id="91d32-1208">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="91d32-1208">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="91d32-1209">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="91d32-1209">ReadWriteItem</span></span>|
|[<span data-ttu-id="91d32-1210">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="91d32-1210">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="91d32-1211">Создание</span><span class="sxs-lookup"><span data-stu-id="91d32-1211">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="91d32-1212">Пример</span><span class="sxs-lookup"><span data-stu-id="91d32-1212">Example</span></span>

<span data-ttu-id="91d32-1213">Указанный ниже код удаляет вложение с идентификатором "0".</span><span class="sxs-lookup"><span data-stu-id="91d32-1213">The following code removes an attachment with an identifier of '0'.</span></span>

```
Office.context.mailbox.item.removeAttachmentAsync(
  '0',
  { asyncContext : null },
  function (asyncResult)
  {
    console.log(asyncResult.status);
  }
);
```

####  <a name="removehandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="91d32-1214">removeHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="91d32-1214">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="91d32-1215">Удаляет обработчик для поддерживаемого события.</span><span class="sxs-lookup"><span data-stu-id="91d32-1215">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="91d32-1216">В настоящее время поддерживаемые типы событий, `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, и `Office.EventType.RecurrenceChanged`</span><span class="sxs-lookup"><span data-stu-id="91d32-1216">Currently the supported event types are `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`</span></span>

##### <a name="parameters"></a><span data-ttu-id="91d32-1217">Параметры:</span><span class="sxs-lookup"><span data-stu-id="91d32-1217">Parameters:</span></span>

| <span data-ttu-id="91d32-1218">Имя</span><span class="sxs-lookup"><span data-stu-id="91d32-1218">Name</span></span> | <span data-ttu-id="91d32-1219">Тип</span><span class="sxs-lookup"><span data-stu-id="91d32-1219">Type</span></span> | <span data-ttu-id="91d32-1220">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="91d32-1220">Attributes</span></span> | <span data-ttu-id="91d32-1221">Описание</span><span class="sxs-lookup"><span data-stu-id="91d32-1221">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="91d32-1222">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="91d32-1222">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="91d32-1223">Событие, которое должно вызвать обработчик.</span><span class="sxs-lookup"><span data-stu-id="91d32-1223">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="91d32-1224">Функция</span><span class="sxs-lookup"><span data-stu-id="91d32-1224">Function</span></span> || <span data-ttu-id="91d32-p177">Функция для обработки события. Функция должна принимать один параметр, представляющий собой объектный литерал. Значение свойства `type` параметра совпадет со значением параметра `eventType`, переданного методу `removeHandlerAsync`.</span><span class="sxs-lookup"><span data-stu-id="91d32-p177">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `removeHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="91d32-1228">Object</span><span class="sxs-lookup"><span data-stu-id="91d32-1228">Object</span></span> | <span data-ttu-id="91d32-1229">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="91d32-1229">&lt;optional&gt;</span></span> | <span data-ttu-id="91d32-1230">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="91d32-1230">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="91d32-1231">Объект</span><span class="sxs-lookup"><span data-stu-id="91d32-1231">Object</span></span> | <span data-ttu-id="91d32-1232">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="91d32-1232">&lt;optional&gt;</span></span> | <span data-ttu-id="91d32-1233">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="91d32-1233">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="91d32-1234">function</span><span class="sxs-lookup"><span data-stu-id="91d32-1234">function</span></span>| <span data-ttu-id="91d32-1235">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="91d32-1235">&lt;optional&gt;</span></span>|<span data-ttu-id="91d32-1236">Когда метод завершает выполнение, переданная в параметре `callback` функция вызывается с единственным параметром `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="91d32-1236">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="91d32-1237">Требования</span><span class="sxs-lookup"><span data-stu-id="91d32-1237">Requirements</span></span>

|<span data-ttu-id="91d32-1238">Требование</span><span class="sxs-lookup"><span data-stu-id="91d32-1238">Requirement</span></span>| <span data-ttu-id="91d32-1239">Значение</span><span class="sxs-lookup"><span data-stu-id="91d32-1239">Value</span></span>|
|---|---|
|[<span data-ttu-id="91d32-1240">Версия минимального набора требований для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="91d32-1240">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="91d32-1241">1.7</span><span class="sxs-lookup"><span data-stu-id="91d32-1241">17 </span></span> |
|[<span data-ttu-id="91d32-1242">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="91d32-1242">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="91d32-1243">ReadItem</span><span class="sxs-lookup"><span data-stu-id="91d32-1243">ReadItem</span></span> |
|[<span data-ttu-id="91d32-1244">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="91d32-1244">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="91d32-1245">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="91d32-1245">Compose or read</span></span> |

##### <a name="example"></a><span data-ttu-id="91d32-1246">Пример</span><span class="sxs-lookup"><span data-stu-id="91d32-1246">Example</span></span>

```
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

####  <a name="saveasyncoptions-callback"></a><span data-ttu-id="91d32-1247">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="91d32-1247">saveAsync([options], callback)</span></span>

<span data-ttu-id="91d32-1248">Асинхронно сохраняет элемент.</span><span class="sxs-lookup"><span data-stu-id="91d32-1248">Asynchronously saves an item.</span></span>

<span data-ttu-id="91d32-p178">При вызове этот метод сохраняет текущее сообщение в виде черновика и возвращает идентификатор элемента с помощью метода обратного вызова. В веб-приложернии Outlook или интерактивном режиме Outlook этот элемент сохраняется на сервере. В Outlook в режиме кэширования этот элемент сохраняется в локальном кэше.</span><span class="sxs-lookup"><span data-stu-id="91d32-p178">When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="91d32-1252">Если в режиме создания надстройка вызывает для элемента метод `saveAsync`, чтобы получить параметр `itemId` для использования с EWS или REST API, необходимо помнить, что синхронизация элемента с сервером может занять много времени, если Outlook работает в режиме кэширования данных.</span><span class="sxs-lookup"><span data-stu-id="91d32-1252">Note: If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server. Until the item is synced, using the  will return an error.</span></span> <span data-ttu-id="91d32-1253">До окончания синхронизации применение параметра `itemId` будет приводить к ошибке.</span><span class="sxs-lookup"><span data-stu-id="91d32-1253">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="91d32-p180">Так как для встреч не предусмотрено состояние черновика, если `saveAsync` вызывается для встречи в режиме создания, она сохраняется как обычная встреча в календаре пользователя. При сохранении новой встречи приглашения не отправляются. При сохранении существующей встречи уведомления отправляются добавленным или удаленным участникам.</span><span class="sxs-lookup"><span data-stu-id="91d32-p180">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="91d32-1257">Следующие клиенты имеют разную реакцию на событие для `saveAsync` для встреч в режиме создания:</span><span class="sxs-lookup"><span data-stu-id="91d32-1257">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="91d32-1258">Mac Outlook не поддерживает `saveAsync` на собрании в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="91d32-1258">Note: Mac Outlook does not support `saveAsync` on a meeting in compose mode. Calling  on a meeting in Mac Outlook will return an error.</span></span> <span data-ttu-id="91d32-1259">Вызов `saveAsync` на собрании в Mac Outlook возвращает ошибку.</span><span class="sxs-lookup"><span data-stu-id="91d32-1259">Note: Mac Outlook does not support  on a meeting in compose mode. Calling `saveAsync` on a meeting in Mac Outlook will return an error.</span></span>
> - <span data-ttu-id="91d32-1260">Outlook в Интернете всегда отправляет приглашение или обновления при вызове `saveAsync` на встрече в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="91d32-1260">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="91d32-1261">Параметры:</span><span class="sxs-lookup"><span data-stu-id="91d32-1261">Parameters:</span></span>

|<span data-ttu-id="91d32-1262">Имя</span><span class="sxs-lookup"><span data-stu-id="91d32-1262">Name</span></span>|<span data-ttu-id="91d32-1263">Тип</span><span class="sxs-lookup"><span data-stu-id="91d32-1263">Type</span></span>|<span data-ttu-id="91d32-1264">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="91d32-1264">Attributes</span></span>|<span data-ttu-id="91d32-1265">Описание</span><span class="sxs-lookup"><span data-stu-id="91d32-1265">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="91d32-1266">Oбъект</span><span class="sxs-lookup"><span data-stu-id="91d32-1266">Object</span></span>|<span data-ttu-id="91d32-1267">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="91d32-1267">&lt;optional&gt;</span></span>|<span data-ttu-id="91d32-1268">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="91d32-1268">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="91d32-1269">Объект</span><span class="sxs-lookup"><span data-stu-id="91d32-1269">Object</span></span>|<span data-ttu-id="91d32-1270">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="91d32-1270">&lt;optional&gt;</span></span>|<span data-ttu-id="91d32-1271">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="91d32-1271">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="91d32-1272">функция</span><span class="sxs-lookup"><span data-stu-id="91d32-1272">function</span></span>||<span data-ttu-id="91d32-1273">Когда метод завершает выполнение, переданная в параметре `callback` функция вызывается с единственным параметром `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="91d32-1273">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="91d32-1274">После успешного выполнения идентификатор элемента будет предоставлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="91d32-1274">On success, the item identifier is provided in the `asyncResult.value` property.|</span></span>|

##### <a name="requirements"></a><span data-ttu-id="91d32-1275">Требования</span><span class="sxs-lookup"><span data-stu-id="91d32-1275">Requirements</span></span>

|<span data-ttu-id="91d32-1276">Требование</span><span class="sxs-lookup"><span data-stu-id="91d32-1276">Requirement</span></span>|<span data-ttu-id="91d32-1277">Значение</span><span class="sxs-lookup"><span data-stu-id="91d32-1277">Value</span></span>|
|---|---|
|[<span data-ttu-id="91d32-1278">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="91d32-1278">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="91d32-1279">1.3</span><span class="sxs-lookup"><span data-stu-id="91d32-1279">1.3</span></span>|
|[<span data-ttu-id="91d32-1280">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="91d32-1280">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="91d32-1281">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="91d32-1281">ReadWriteItem</span></span>|
|[<span data-ttu-id="91d32-1282">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="91d32-1282">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="91d32-1283">Создание</span><span class="sxs-lookup"><span data-stu-id="91d32-1283">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="91d32-1284">Примеры</span><span class="sxs-lookup"><span data-stu-id="91d32-1284">Examples</span></span>

```
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result
  });
```

<span data-ttu-id="91d32-p182">Ниже приведен пример параметра `result`, переданного функции обратного вызова. Свойство `value` содержит идентификатор элемента.</span><span class="sxs-lookup"><span data-stu-id="91d32-p182">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

####  <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="91d32-1287">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="91d32-1287">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="91d32-1288">Асинхронно вставляет данные в текст или тему сообщения.</span><span class="sxs-lookup"><span data-stu-id="91d32-1288">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="91d32-p183">Метод `setSelectedDataAsync` вставляет указанную строку в местоположение курсора в теме или тексте элемента либо, если текст выделен в редакторе, он заменяет выделенный текст. Если курсор находится вне текста или темы элемента, возвращается ошибка. После вставки курсор помещается в конец вставленного содержимого.</span><span class="sxs-lookup"><span data-stu-id="91d32-p183">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="91d32-1292">Параметры:</span><span class="sxs-lookup"><span data-stu-id="91d32-1292">Parameters:</span></span>

|<span data-ttu-id="91d32-1293">Имя</span><span class="sxs-lookup"><span data-stu-id="91d32-1293">Name</span></span>|<span data-ttu-id="91d32-1294">Тип</span><span class="sxs-lookup"><span data-stu-id="91d32-1294">Type</span></span>|<span data-ttu-id="91d32-1295">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="91d32-1295">Attributes</span></span>|<span data-ttu-id="91d32-1296">Описание</span><span class="sxs-lookup"><span data-stu-id="91d32-1296">Description</span></span>|
|---|---|---|---|
|`data`|<span data-ttu-id="91d32-1297">String</span><span class="sxs-lookup"><span data-stu-id="91d32-1297">String</span></span>||<span data-ttu-id="91d32-p184">Вставляемые данные. Объем данных не должен превышать 1 000 000 символов. Если передано больше 1 000 000 символов, возвращается исключение `ArgumentOutOfRange`.</span><span class="sxs-lookup"><span data-stu-id="91d32-p184">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`|<span data-ttu-id="91d32-1301">Объект</span><span class="sxs-lookup"><span data-stu-id="91d32-1301">Object</span></span>|<span data-ttu-id="91d32-1302">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="91d32-1302">&lt;optional&gt;</span></span>|<span data-ttu-id="91d32-1303">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="91d32-1303">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="91d32-1304">Объект</span><span class="sxs-lookup"><span data-stu-id="91d32-1304">Object</span></span>|<span data-ttu-id="91d32-1305">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="91d32-1305">&lt;optional&gt;</span></span>|<span data-ttu-id="91d32-1306">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="91d32-1306">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="91d32-1307">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="91d32-1307">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="91d32-1308">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="91d32-1308">&lt;optional&gt;</span></span>|<span data-ttu-id="91d32-p185">Если задано значение `text`, текущий стиль применяется в Outlook и веб-приложении Outlook. Если поле представляет собой редактор HTML, вставляются только текстовые данные, даже если они имеют формат HTML.</span><span class="sxs-lookup"><span data-stu-id="91d32-p185">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="91d32-p186">Если `html` и поле поддерживают HTML (а тема не поддерживает), в веб-приложении Outlook применяется текущий стиль, а в Outlook — стиль по умолчанию. Если поле является текстовым, возвращается ошибка `InvalidDataFormat`.</span><span class="sxs-lookup"><span data-stu-id="91d32-p186">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="91d32-1313">Если тип `coercionType` не установлен, результат зависит от поля: если поле имеет формат HTML, то используется HTML; если поле является текстовым, то используется обычный текст.</span><span class="sxs-lookup"><span data-stu-id="91d32-1313">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`|<span data-ttu-id="91d32-1314">function</span><span class="sxs-lookup"><span data-stu-id="91d32-1314">function</span></span>||<span data-ttu-id="91d32-1315">Когда метод завершает выполнение, переданная в параметре `callback` функция вызывается с единственным параметром `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="91d32-1315">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="91d32-1316">Требования</span><span class="sxs-lookup"><span data-stu-id="91d32-1316">Requirements</span></span>

|<span data-ttu-id="91d32-1317">Требование</span><span class="sxs-lookup"><span data-stu-id="91d32-1317">Requirement</span></span>|<span data-ttu-id="91d32-1318">Значение</span><span class="sxs-lookup"><span data-stu-id="91d32-1318">Value</span></span>|
|---|---|
|[<span data-ttu-id="91d32-1319">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="91d32-1319">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="91d32-1320">1.2</span><span class="sxs-lookup"><span data-stu-id="91d32-1320">1.2</span></span>|
|[<span data-ttu-id="91d32-1321">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="91d32-1321">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="91d32-1322">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="91d32-1322">ReadWriteItem</span></span>|
|[<span data-ttu-id="91d32-1323">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="91d32-1323">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="91d32-1324">Создание</span><span class="sxs-lookup"><span data-stu-id="91d32-1324">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="91d32-1325">Пример</span><span class="sxs-lookup"><span data-stu-id="91d32-1325">Example</span></span>

```
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```