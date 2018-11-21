
# <a name="item"></a><span data-ttu-id="542eb-101">item</span><span class="sxs-lookup"><span data-stu-id="542eb-101">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="542eb-102">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="542eb-102">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="542eb-p101">Пространство имен `item` используется для доступа к выбранному в данный момент сообщению, приглашению на собрание или описанию встречи. Вы можете определить тип пространства имен `item` с помощью свойства [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="542eb-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="542eb-105">Requirements</span><span class="sxs-lookup"><span data-stu-id="542eb-105">Requirements</span></span>

|<span data-ttu-id="542eb-106">Requirement</span><span class="sxs-lookup"><span data-stu-id="542eb-106">Requirement</span></span>|<span data-ttu-id="542eb-107">Значение</span><span class="sxs-lookup"><span data-stu-id="542eb-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="542eb-108">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="542eb-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="542eb-109">1.0</span><span class="sxs-lookup"><span data-stu-id="542eb-109">1.0</span></span>|
|[<span data-ttu-id="542eb-110">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="542eb-110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="542eb-111">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="542eb-111">Restricted</span></span>|
|[<span data-ttu-id="542eb-112">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="542eb-112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="542eb-113">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="542eb-113">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="542eb-114">Элементы и методы</span><span class="sxs-lookup"><span data-stu-id="542eb-114">Members and methods</span></span>

| <span data-ttu-id="542eb-115">Элемент</span><span class="sxs-lookup"><span data-stu-id="542eb-115">Member</span></span> | <span data-ttu-id="542eb-116">Тип</span><span class="sxs-lookup"><span data-stu-id="542eb-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="542eb-117">attachments</span><span class="sxs-lookup"><span data-stu-id="542eb-117">attachments</span></span>](#attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails) | <span data-ttu-id="542eb-118">Элемент</span><span class="sxs-lookup"><span data-stu-id="542eb-118">Member</span></span> |
| [<span data-ttu-id="542eb-119">bcc</span><span class="sxs-lookup"><span data-stu-id="542eb-119">bcc</span></span>](#bcc-recipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="542eb-120">Элемент</span><span class="sxs-lookup"><span data-stu-id="542eb-120">Member</span></span> |
| [<span data-ttu-id="542eb-121">body</span><span class="sxs-lookup"><span data-stu-id="542eb-121">body</span></span>](#body-bodyjavascriptapioutlookofficebody) | <span data-ttu-id="542eb-122">Элемент</span><span class="sxs-lookup"><span data-stu-id="542eb-122">Member</span></span> |
| [<span data-ttu-id="542eb-123">cc</span><span class="sxs-lookup"><span data-stu-id="542eb-123">cc</span></span>](#cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="542eb-124">Элемент</span><span class="sxs-lookup"><span data-stu-id="542eb-124">Member</span></span> |
| [<span data-ttu-id="542eb-125">conversationId</span><span class="sxs-lookup"><span data-stu-id="542eb-125">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="542eb-126">Элемент</span><span class="sxs-lookup"><span data-stu-id="542eb-126">Member</span></span> |
| [<span data-ttu-id="542eb-127">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="542eb-127">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="542eb-128">Элемент</span><span class="sxs-lookup"><span data-stu-id="542eb-128">Member</span></span> |
| [<span data-ttu-id="542eb-129">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="542eb-129">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="542eb-130">Элемент</span><span class="sxs-lookup"><span data-stu-id="542eb-130">Member</span></span> |
| [<span data-ttu-id="542eb-131">end</span><span class="sxs-lookup"><span data-stu-id="542eb-131">end</span></span>](#end-datetimejavascriptapioutlookofficetime) | <span data-ttu-id="542eb-132">Элемент</span><span class="sxs-lookup"><span data-stu-id="542eb-132">Member</span></span> |
| [<span data-ttu-id="542eb-133">from</span><span class="sxs-lookup"><span data-stu-id="542eb-133">from</span></span>](#from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom) | <span data-ttu-id="542eb-134">Элемент</span><span class="sxs-lookup"><span data-stu-id="542eb-134">Member</span></span> |
| [<span data-ttu-id="542eb-135">internetHeaders</span><span class="sxs-lookup"><span data-stu-id="542eb-135">internetHeaders</span></span>](#internetheaders-internetheadersjavascriptapioutlookofficeinternetheaders) | <span data-ttu-id="542eb-136">Элемент</span><span class="sxs-lookup"><span data-stu-id="542eb-136">Member</span></span> |
| [<span data-ttu-id="542eb-137">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="542eb-137">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="542eb-138">Элемент</span><span class="sxs-lookup"><span data-stu-id="542eb-138">Member</span></span> |
| [<span data-ttu-id="542eb-139">itemClass</span><span class="sxs-lookup"><span data-stu-id="542eb-139">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="542eb-140">Элемент</span><span class="sxs-lookup"><span data-stu-id="542eb-140">Member</span></span> |
| [<span data-ttu-id="542eb-141">itemId</span><span class="sxs-lookup"><span data-stu-id="542eb-141">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="542eb-142">Элемент</span><span class="sxs-lookup"><span data-stu-id="542eb-142">Member</span></span> |
| [<span data-ttu-id="542eb-143">itemType</span><span class="sxs-lookup"><span data-stu-id="542eb-143">itemType</span></span>](#itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype) | <span data-ttu-id="542eb-144">Элемент</span><span class="sxs-lookup"><span data-stu-id="542eb-144">Member</span></span> |
| [<span data-ttu-id="542eb-145">location</span><span class="sxs-lookup"><span data-stu-id="542eb-145">location</span></span>](#location-stringlocationjavascriptapioutlookofficelocation) | <span data-ttu-id="542eb-146">Элемент</span><span class="sxs-lookup"><span data-stu-id="542eb-146">Member</span></span> |
| [<span data-ttu-id="542eb-147">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="542eb-147">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="542eb-148">Элемент</span><span class="sxs-lookup"><span data-stu-id="542eb-148">Member</span></span> |
| [<span data-ttu-id="542eb-149">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="542eb-149">notificationMessages</span></span>](#notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessages) | <span data-ttu-id="542eb-150">Элемент</span><span class="sxs-lookup"><span data-stu-id="542eb-150">Member</span></span> |
| [<span data-ttu-id="542eb-151">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="542eb-151">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="542eb-152">Элемент</span><span class="sxs-lookup"><span data-stu-id="542eb-152">Member</span></span> |
| [<span data-ttu-id="542eb-153">organizer</span><span class="sxs-lookup"><span data-stu-id="542eb-153">organizer</span></span>](#organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsorganizerjavascriptapioutlookofficeorganizer) | <span data-ttu-id="542eb-154">Элемент</span><span class="sxs-lookup"><span data-stu-id="542eb-154">Member</span></span> |
| [<span data-ttu-id="542eb-155">recurrence</span><span class="sxs-lookup"><span data-stu-id="542eb-155">recurrence</span></span>](#nullable-recurrence-recurrencejavascriptapioutlookofficerecurrence) | <span data-ttu-id="542eb-156">Элемент</span><span class="sxs-lookup"><span data-stu-id="542eb-156">Member</span></span> |
| [<span data-ttu-id="542eb-157">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="542eb-157">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="542eb-158">Member</span><span class="sxs-lookup"><span data-stu-id="542eb-158">Member</span></span> |
| [<span data-ttu-id="542eb-159">sender</span><span class="sxs-lookup"><span data-stu-id="542eb-159">sender</span></span>](#sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails) | <span data-ttu-id="542eb-160">Элемент</span><span class="sxs-lookup"><span data-stu-id="542eb-160">Member</span></span> |
| [<span data-ttu-id="542eb-161">seriesId</span><span class="sxs-lookup"><span data-stu-id="542eb-161">seriesId</span></span>](#nullable-seriesid-string) | <span data-ttu-id="542eb-162">Элемент</span><span class="sxs-lookup"><span data-stu-id="542eb-162">Member</span></span> |
| [<span data-ttu-id="542eb-163">start</span><span class="sxs-lookup"><span data-stu-id="542eb-163">start</span></span>](#start-datetimejavascriptapioutlookofficetime) | <span data-ttu-id="542eb-164">Элемент</span><span class="sxs-lookup"><span data-stu-id="542eb-164">Member</span></span> |
| [<span data-ttu-id="542eb-165">subject</span><span class="sxs-lookup"><span data-stu-id="542eb-165">subject</span></span>](#subject-stringsubjectjavascriptapioutlookofficesubject) | <span data-ttu-id="542eb-166">Элемент</span><span class="sxs-lookup"><span data-stu-id="542eb-166">Member</span></span> |
| [<span data-ttu-id="542eb-167">to</span><span class="sxs-lookup"><span data-stu-id="542eb-167">to</span></span>](#to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="542eb-168">Элемент</span><span class="sxs-lookup"><span data-stu-id="542eb-168">Member</span></span> |
| [<span data-ttu-id="542eb-169">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="542eb-169">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="542eb-170">Метод</span><span class="sxs-lookup"><span data-stu-id="542eb-170">Method</span></span> |
| [<span data-ttu-id="542eb-171">addFileAttachmentFromBase64Async</span><span class="sxs-lookup"><span data-stu-id="542eb-171">addFileAttachmentFromBase64Async</span></span>](#addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback) | <span data-ttu-id="542eb-172">Метод</span><span class="sxs-lookup"><span data-stu-id="542eb-172">Method</span></span> |
| [<span data-ttu-id="542eb-173">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="542eb-173">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="542eb-174">Метод</span><span class="sxs-lookup"><span data-stu-id="542eb-174">Method</span></span> |
| [<span data-ttu-id="542eb-175">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="542eb-175">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="542eb-176">Метод</span><span class="sxs-lookup"><span data-stu-id="542eb-176">Method</span></span> |
| [<span data-ttu-id="542eb-177">close</span><span class="sxs-lookup"><span data-stu-id="542eb-177">close</span></span>](#close) | <span data-ttu-id="542eb-178">Метод</span><span class="sxs-lookup"><span data-stu-id="542eb-178">Method</span></span> |
| [<span data-ttu-id="542eb-179">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="542eb-179">displayReplyAllForm</span></span>](#displayreplyallformformdata) | <span data-ttu-id="542eb-180">Метод</span><span class="sxs-lookup"><span data-stu-id="542eb-180">Method</span></span> |
| [<span data-ttu-id="542eb-181">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="542eb-181">displayReplyForm</span></span>](#displayreplyformformdata) | <span data-ttu-id="542eb-182">Метод</span><span class="sxs-lookup"><span data-stu-id="542eb-182">Method</span></span> |
| [<span data-ttu-id="542eb-183">getAttachmentContentAsync</span><span class="sxs-lookup"><span data-stu-id="542eb-183">getAttachmentContentAsync</span></span>](#getattachmentcontentasyncattachmentid-options-callback--attachmentcontentjavascriptapioutlookofficeattachmentcontent) | <span data-ttu-id="542eb-184">Метод</span><span class="sxs-lookup"><span data-stu-id="542eb-184">Method</span></span> |
| [<span data-ttu-id="542eb-185">getAttachmentsAsync</span><span class="sxs-lookup"><span data-stu-id="542eb-185">getAttachmentsAsync</span></span>](#getattachmentsasyncoptions-callback--arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails) | <span data-ttu-id="542eb-186">Метод</span><span class="sxs-lookup"><span data-stu-id="542eb-186">Method</span></span> |
| [<span data-ttu-id="542eb-187">getEntities</span><span class="sxs-lookup"><span data-stu-id="542eb-187">getEntities</span></span>](#getentities--entitiesjavascriptapioutlookofficeentities) | <span data-ttu-id="542eb-188">Метод</span><span class="sxs-lookup"><span data-stu-id="542eb-188">Method</span></span> |
| [<span data-ttu-id="542eb-189">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="542eb-189">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion) | <span data-ttu-id="542eb-190">Метод</span><span class="sxs-lookup"><span data-stu-id="542eb-190">Method</span></span> |
| [<span data-ttu-id="542eb-191">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="542eb-191">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion) | <span data-ttu-id="542eb-192">Метод</span><span class="sxs-lookup"><span data-stu-id="542eb-192">Method</span></span> |
| [<span data-ttu-id="542eb-193">getInitializationContextAsync</span><span class="sxs-lookup"><span data-stu-id="542eb-193">getInitializationContextAsync</span></span>](#getinitializationcontextasyncoptions-callback) | <span data-ttu-id="542eb-194">Метод</span><span class="sxs-lookup"><span data-stu-id="542eb-194">Method</span></span> |
| [<span data-ttu-id="542eb-195">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="542eb-195">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="542eb-196">Метод</span><span class="sxs-lookup"><span data-stu-id="542eb-196">Method</span></span> |
| [<span data-ttu-id="542eb-197">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="542eb-197">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="542eb-198">Метод</span><span class="sxs-lookup"><span data-stu-id="542eb-198">Method</span></span> |
| [<span data-ttu-id="542eb-199">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="542eb-199">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="542eb-200">Метод</span><span class="sxs-lookup"><span data-stu-id="542eb-200">Method</span></span> |
| [<span data-ttu-id="542eb-201">getSelectedEntities</span><span class="sxs-lookup"><span data-stu-id="542eb-201">getSelectedEntities</span></span>](#getselectedentities--entitiesjavascriptapioutlookofficeentities) | <span data-ttu-id="542eb-202">Метод</span><span class="sxs-lookup"><span data-stu-id="542eb-202">Method</span></span> |
| [<span data-ttu-id="542eb-203">getSelectedRegExMatches</span><span class="sxs-lookup"><span data-stu-id="542eb-203">getSelectedRegExMatches</span></span>](#getselectedregexmatches--object) | <span data-ttu-id="542eb-204">Метод</span><span class="sxs-lookup"><span data-stu-id="542eb-204">Method</span></span> |
| [<span data-ttu-id="542eb-205">getSharedPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="542eb-205">getSharedPropertiesAsync</span></span>](#getsharedpropertiesasyncoptions-callback) | <span data-ttu-id="542eb-206">Метод</span><span class="sxs-lookup"><span data-stu-id="542eb-206">Method</span></span> |
| [<span data-ttu-id="542eb-207">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="542eb-207">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="542eb-208">Метод</span><span class="sxs-lookup"><span data-stu-id="542eb-208">Method</span></span> |
| [<span data-ttu-id="542eb-209">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="542eb-209">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="542eb-210">Метод</span><span class="sxs-lookup"><span data-stu-id="542eb-210">Method</span></span> |
| [<span data-ttu-id="542eb-211">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="542eb-211">removeHandlerAsync</span></span>](#removehandlerasynceventtype-handler-options-callback) | <span data-ttu-id="542eb-212">Метод</span><span class="sxs-lookup"><span data-stu-id="542eb-212">Method</span></span> |
| [<span data-ttu-id="542eb-213">saveAsync</span><span class="sxs-lookup"><span data-stu-id="542eb-213">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="542eb-214">Метод</span><span class="sxs-lookup"><span data-stu-id="542eb-214">Method</span></span> |
| [<span data-ttu-id="542eb-215">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="542eb-215">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="542eb-216">Метод</span><span class="sxs-lookup"><span data-stu-id="542eb-216">Method</span></span> |

### <a name="example"></a><span data-ttu-id="542eb-217">Пример</span><span class="sxs-lookup"><span data-stu-id="542eb-217">Example</span></span>

<span data-ttu-id="542eb-218">В примере кода JavaScript, приведенном ниже, показано, как получить доступ к свойству `subject` текущего элемента в Outlook.</span><span class="sxs-lookup"><span data-stu-id="542eb-218">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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
}
```

### <a name="members"></a><span data-ttu-id="542eb-219">Элементы</span><span class="sxs-lookup"><span data-stu-id="542eb-219">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails"></a><span data-ttu-id="542eb-220">attachments :Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="542eb-220">attachments :Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

<span data-ttu-id="542eb-221">Получает вложения элемента в качестве массива.</span><span class="sxs-lookup"><span data-stu-id="542eb-221">Gets the item's attachments as an array.</span></span> <span data-ttu-id="542eb-222">Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="542eb-222">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="542eb-223">Outlook блокирует определенные типы файлов из-за потенциальных проблем с безопасностью, поэтому они не возвращаются.</span><span class="sxs-lookup"><span data-stu-id="542eb-223">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="542eb-224">Дополнительные сведения см. в статье [Блокировка вложений в Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="542eb-224">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="542eb-225">Тип:</span><span class="sxs-lookup"><span data-stu-id="542eb-225">Type:</span></span>

*   <span data-ttu-id="542eb-226">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="542eb-226">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="542eb-227">Требования</span><span class="sxs-lookup"><span data-stu-id="542eb-227">Requirements</span></span>

|<span data-ttu-id="542eb-228">Requirement</span><span class="sxs-lookup"><span data-stu-id="542eb-228">Requirement</span></span>|<span data-ttu-id="542eb-229">Значение</span><span class="sxs-lookup"><span data-stu-id="542eb-229">Value</span></span>|
|---|---|
|[<span data-ttu-id="542eb-230">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="542eb-230">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="542eb-231">1.0</span><span class="sxs-lookup"><span data-stu-id="542eb-231">1.0</span></span>|
|[<span data-ttu-id="542eb-232">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="542eb-232">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="542eb-233">ReadItem</span><span class="sxs-lookup"><span data-stu-id="542eb-233">ReadItem</span></span>|
|[<span data-ttu-id="542eb-234">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="542eb-234">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="542eb-235">Чтение</span><span class="sxs-lookup"><span data-stu-id="542eb-235">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="542eb-236">Пример</span><span class="sxs-lookup"><span data-stu-id="542eb-236">Example</span></span>

<span data-ttu-id="542eb-237">С помощью приведенного ниже кода можно создать HTML-строку с подробными сведениями обо всех вложениях для текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="542eb-237">The following code builds an HTML string with details of all attachments on the current item.</span></span>

```javascript
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

####  <a name="bcc-recipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="542eb-238">bcc :[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="542eb-238">bcc :[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="542eb-239">Получает объект, который предоставляет методы для получения или обновления получателей скрытой копии сообщения.</span><span class="sxs-lookup"><span data-stu-id="542eb-239">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="542eb-240">Только в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="542eb-240">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="542eb-241">Тип:</span><span class="sxs-lookup"><span data-stu-id="542eb-241">Type:</span></span>

*   [<span data-ttu-id="542eb-242">Recipients</span><span class="sxs-lookup"><span data-stu-id="542eb-242">Recipients</span></span>](/javascript/api/outlook/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="542eb-243">Требования</span><span class="sxs-lookup"><span data-stu-id="542eb-243">Requirements</span></span>

|<span data-ttu-id="542eb-244">Requirement</span><span class="sxs-lookup"><span data-stu-id="542eb-244">Requirement</span></span>|<span data-ttu-id="542eb-245">Значение</span><span class="sxs-lookup"><span data-stu-id="542eb-245">Value</span></span>|
|---|---|
|[<span data-ttu-id="542eb-246">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="542eb-246">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="542eb-247">1.1</span><span class="sxs-lookup"><span data-stu-id="542eb-247">1.1</span></span>|
|[<span data-ttu-id="542eb-248">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="542eb-248">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="542eb-249">ReadItem</span><span class="sxs-lookup"><span data-stu-id="542eb-249">ReadItem</span></span>|
|[<span data-ttu-id="542eb-250">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="542eb-250">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="542eb-251">Создание</span><span class="sxs-lookup"><span data-stu-id="542eb-251">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="542eb-252">Пример</span><span class="sxs-lookup"><span data-stu-id="542eb-252">Example</span></span>

```javascript
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlookofficebody"></a><span data-ttu-id="542eb-253">body :[Body](/javascript/api/outlook/office.body)</span><span class="sxs-lookup"><span data-stu-id="542eb-253">body :[Body](/javascript/api/outlook/office.body)</span></span>

<span data-ttu-id="542eb-254">Получает объект, предоставляющий методы для работы с основным текстом элемента.</span><span class="sxs-lookup"><span data-stu-id="542eb-254">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="542eb-255">Тип:</span><span class="sxs-lookup"><span data-stu-id="542eb-255">Type:</span></span>

*   [<span data-ttu-id="542eb-256">Body</span><span class="sxs-lookup"><span data-stu-id="542eb-256">Body</span></span>](/javascript/api/outlook/office.body)

##### <a name="requirements"></a><span data-ttu-id="542eb-257">Требования</span><span class="sxs-lookup"><span data-stu-id="542eb-257">Requirements</span></span>

|<span data-ttu-id="542eb-258">Requirement</span><span class="sxs-lookup"><span data-stu-id="542eb-258">Requirement</span></span>|<span data-ttu-id="542eb-259">Значение</span><span class="sxs-lookup"><span data-stu-id="542eb-259">Value</span></span>|
|---|---|
|[<span data-ttu-id="542eb-260">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="542eb-260">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="542eb-261">1.1</span><span class="sxs-lookup"><span data-stu-id="542eb-261">1.1</span></span>|
|[<span data-ttu-id="542eb-262">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="542eb-262">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="542eb-263">ReadItem</span><span class="sxs-lookup"><span data-stu-id="542eb-263">ReadItem</span></span>|
|[<span data-ttu-id="542eb-264">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="542eb-264">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="542eb-265">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="542eb-265">Compose or read</span></span>|

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="542eb-266">cc :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="542eb-266">cc :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="542eb-267">Предоставляет доступ к получателям копии сообщения.</span><span class="sxs-lookup"><span data-stu-id="542eb-267">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="542eb-268">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="542eb-268">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="542eb-269">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="542eb-269">Read mode</span></span>

<span data-ttu-id="542eb-p106">Свойство `cc` возвращает массив, который содержит объект `EmailAddressDetails` для каждого получателя, указанного в строке **Копия** сообщения. Коллекция может включать не более 100 элементов.</span><span class="sxs-lookup"><span data-stu-id="542eb-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="542eb-272">Режим создания</span><span class="sxs-lookup"><span data-stu-id="542eb-272">Compose mode</span></span>

<span data-ttu-id="542eb-273">Свойство `cc` возвращает объект `Recipients`, предоставляющий методы для получения или обновления получателей, которые указаны в строке **Копия** сообщения.</span><span class="sxs-lookup"><span data-stu-id="542eb-273">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="542eb-274">Тип:</span><span class="sxs-lookup"><span data-stu-id="542eb-274">Type:</span></span>

*   <span data-ttu-id="542eb-275">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="542eb-275">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="542eb-276">Требования</span><span class="sxs-lookup"><span data-stu-id="542eb-276">Requirements</span></span>

|<span data-ttu-id="542eb-277">Requirement</span><span class="sxs-lookup"><span data-stu-id="542eb-277">Requirement</span></span>|<span data-ttu-id="542eb-278">Значение</span><span class="sxs-lookup"><span data-stu-id="542eb-278">Value</span></span>|
|---|---|
|[<span data-ttu-id="542eb-279">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="542eb-279">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="542eb-280">1.0</span><span class="sxs-lookup"><span data-stu-id="542eb-280">1.0</span></span>|
|[<span data-ttu-id="542eb-281">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="542eb-281">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="542eb-282">ReadItem</span><span class="sxs-lookup"><span data-stu-id="542eb-282">ReadItem</span></span>|
|[<span data-ttu-id="542eb-283">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="542eb-283">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="542eb-284">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="542eb-284">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="542eb-285">Пример</span><span class="sxs-lookup"><span data-stu-id="542eb-285">Example</span></span>

```javascript
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="542eb-286">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="542eb-286">(nullable) conversationId :String</span></span>

<span data-ttu-id="542eb-287">Получает идентификатор разговора по электронной почте, содержащего конкретное сообщение.</span><span class="sxs-lookup"><span data-stu-id="542eb-287">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="542eb-p107">Вы можете получить целочисленное значение этого свойства, если ваше почтовое приложение активируется в формах просмотра или формах создания ответов. Если пользователь изменит тему ответа, после его отправки идентификатор беседы будет изменен, и полученное ранее значение будет недействительным.</span><span class="sxs-lookup"><span data-stu-id="542eb-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="542eb-p108">Это свойство имеет значение NULL для нового элемента в форме создания. Свойство `conversationId` вернет значение, если пользователь задаст тему и сохранит элемент.</span><span class="sxs-lookup"><span data-stu-id="542eb-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="542eb-292">Тип:</span><span class="sxs-lookup"><span data-stu-id="542eb-292">Type:</span></span>

*   <span data-ttu-id="542eb-293">String</span><span class="sxs-lookup"><span data-stu-id="542eb-293">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="542eb-294">Требования</span><span class="sxs-lookup"><span data-stu-id="542eb-294">Requirements</span></span>

|<span data-ttu-id="542eb-295">Requirement</span><span class="sxs-lookup"><span data-stu-id="542eb-295">Requirement</span></span>|<span data-ttu-id="542eb-296">Значение</span><span class="sxs-lookup"><span data-stu-id="542eb-296">Value</span></span>|
|---|---|
|[<span data-ttu-id="542eb-297">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="542eb-297">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="542eb-298">1.0</span><span class="sxs-lookup"><span data-stu-id="542eb-298">1.0</span></span>|
|[<span data-ttu-id="542eb-299">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="542eb-299">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="542eb-300">ReadItem</span><span class="sxs-lookup"><span data-stu-id="542eb-300">ReadItem</span></span>|
|[<span data-ttu-id="542eb-301">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="542eb-301">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="542eb-302">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="542eb-302">Compose or read</span></span>|

#### <a name="datetimecreated-date"></a><span data-ttu-id="542eb-303">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="542eb-303">dateTimeCreated :Date</span></span>

<span data-ttu-id="542eb-p109">Получает дату и время создания элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="542eb-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="542eb-306">Тип:</span><span class="sxs-lookup"><span data-stu-id="542eb-306">Type:</span></span>

*   <span data-ttu-id="542eb-307">Date</span><span class="sxs-lookup"><span data-stu-id="542eb-307">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="542eb-308">Требования</span><span class="sxs-lookup"><span data-stu-id="542eb-308">Requirements</span></span>

|<span data-ttu-id="542eb-309">Requirement</span><span class="sxs-lookup"><span data-stu-id="542eb-309">Requirement</span></span>|<span data-ttu-id="542eb-310">Значение</span><span class="sxs-lookup"><span data-stu-id="542eb-310">Value</span></span>|
|---|---|
|[<span data-ttu-id="542eb-311">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="542eb-311">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="542eb-312">1.0</span><span class="sxs-lookup"><span data-stu-id="542eb-312">1.0</span></span>|
|[<span data-ttu-id="542eb-313">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="542eb-313">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="542eb-314">ReadItem</span><span class="sxs-lookup"><span data-stu-id="542eb-314">ReadItem</span></span>|
|[<span data-ttu-id="542eb-315">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="542eb-315">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="542eb-316">Чтение</span><span class="sxs-lookup"><span data-stu-id="542eb-316">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="542eb-317">Пример</span><span class="sxs-lookup"><span data-stu-id="542eb-317">Example</span></span>

```javascript
var created = Office.context.mailbox.item.dateTimeCreated;
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="542eb-318">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="542eb-318">dateTimeModified :Date</span></span>

<span data-ttu-id="542eb-p110">Получает дату и время последнего изменения элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="542eb-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="542eb-321">Этот элемент не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="542eb-321">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="542eb-322">Тип:</span><span class="sxs-lookup"><span data-stu-id="542eb-322">Type:</span></span>

*   <span data-ttu-id="542eb-323">Date</span><span class="sxs-lookup"><span data-stu-id="542eb-323">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="542eb-324">Требования</span><span class="sxs-lookup"><span data-stu-id="542eb-324">Requirements</span></span>

|<span data-ttu-id="542eb-325">Requirement</span><span class="sxs-lookup"><span data-stu-id="542eb-325">Requirement</span></span>|<span data-ttu-id="542eb-326">Значение</span><span class="sxs-lookup"><span data-stu-id="542eb-326">Value</span></span>|
|---|---|
|[<span data-ttu-id="542eb-327">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="542eb-327">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="542eb-328">1.0</span><span class="sxs-lookup"><span data-stu-id="542eb-328">1.0</span></span>|
|[<span data-ttu-id="542eb-329">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="542eb-329">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="542eb-330">ReadItem</span><span class="sxs-lookup"><span data-stu-id="542eb-330">ReadItem</span></span>|
|[<span data-ttu-id="542eb-331">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="542eb-331">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="542eb-332">Чтение</span><span class="sxs-lookup"><span data-stu-id="542eb-332">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="542eb-333">Пример</span><span class="sxs-lookup"><span data-stu-id="542eb-333">Example</span></span>

```javascript
var modified = Office.context.mailbox.item.dateTimeModified;
```

####  <a name="end-datetimejavascriptapioutlookofficetime"></a><span data-ttu-id="542eb-334">end :Date|[Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="542eb-334">end :Date|[Time](/javascript/api/outlook/office.time)</span></span>

<span data-ttu-id="542eb-335">Получает или задает дату и время окончания встречи.</span><span class="sxs-lookup"><span data-stu-id="542eb-335">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="542eb-p111">Свойство `end` представлено в виде значения даты и времени в формате UTC. Преобразовать значение свойства end в местные значения даты и времени клиента можно с помощью метода [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime).</span><span class="sxs-lookup"><span data-stu-id="542eb-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="542eb-338">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="542eb-338">Read mode</span></span>

<span data-ttu-id="542eb-339">Свойство `end` возвращает объект `Date`.</span><span class="sxs-lookup"><span data-stu-id="542eb-339">The `end` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="542eb-340">Режим создания</span><span class="sxs-lookup"><span data-stu-id="542eb-340">Compose mode</span></span>

<span data-ttu-id="542eb-341">Свойство `end` возвращает объект `Time`.</span><span class="sxs-lookup"><span data-stu-id="542eb-341">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="542eb-342">Если вы задаете время окончания с помощью метода [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-), необходимо использовать метод [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) для преобразования местного времени на клиенте в формат UTC для сервера.</span><span class="sxs-lookup"><span data-stu-id="542eb-342">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="542eb-343">Тип:</span><span class="sxs-lookup"><span data-stu-id="542eb-343">Type:</span></span>

*   <span data-ttu-id="542eb-344">Date | [Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="542eb-344">Date | [Time](/javascript/api/outlook/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="542eb-345">Требования</span><span class="sxs-lookup"><span data-stu-id="542eb-345">Requirements</span></span>

|<span data-ttu-id="542eb-346">Requirement</span><span class="sxs-lookup"><span data-stu-id="542eb-346">Requirement</span></span>|<span data-ttu-id="542eb-347">Значение</span><span class="sxs-lookup"><span data-stu-id="542eb-347">Value</span></span>|
|---|---|
|[<span data-ttu-id="542eb-348">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="542eb-348">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="542eb-349">1.0</span><span class="sxs-lookup"><span data-stu-id="542eb-349">1.0</span></span>|
|[<span data-ttu-id="542eb-350">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="542eb-350">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="542eb-351">ReadItem</span><span class="sxs-lookup"><span data-stu-id="542eb-351">ReadItem</span></span>|
|[<span data-ttu-id="542eb-352">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="542eb-352">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="542eb-353">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="542eb-353">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="542eb-354">Пример</span><span class="sxs-lookup"><span data-stu-id="542eb-354">Example</span></span>

<span data-ttu-id="542eb-355">В примере ниже показано, как с помощью метода [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) объекта `Time` задать время окончания встречи в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="542eb-355">The following example sets the end time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

```javascript
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

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom"></a><span data-ttu-id="542eb-356">from :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[From](/javascript/api/outlook/office.from)</span><span class="sxs-lookup"><span data-stu-id="542eb-356">from :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[From](/javascript/api/outlook/office.from)</span></span>

<span data-ttu-id="542eb-357">Получает адрес электронной почты отправителя сообщения.</span><span class="sxs-lookup"><span data-stu-id="542eb-357">Gets the email address of the sender of a message.</span></span>

<span data-ttu-id="542eb-p112">Свойства `from` и [`sender`](#sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails) представляют одно лицо, если сообщение не отправлено представителем. В противном случае свойство `from` представляет лицо, делегировавшее полномочия, а свойство sender — представителя.</span><span class="sxs-lookup"><span data-stu-id="542eb-p112">The `from` and [`sender`](#sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="542eb-360">Свойству `recipientType`, принадлежащему объекту `EmailAddressDetails` в свойстве `from`, задано значение `undefined`.</span><span class="sxs-lookup"><span data-stu-id="542eb-360">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="542eb-361">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="542eb-361">Read mode</span></span>

<span data-ttu-id="542eb-362">Свойство `from` возвращает объект `EmailAddressDetails`.</span><span class="sxs-lookup"><span data-stu-id="542eb-362">The `from` property returns an `EmailAddressDetails` object.</span></span>

```javascript
var subject = Office.context.mailbox.item.from;
```

##### <a name="compose-mode"></a><span data-ttu-id="542eb-363">Режим создания</span><span class="sxs-lookup"><span data-stu-id="542eb-363">Compose mode</span></span>

<span data-ttu-id="542eb-364">Свойство `from` возвращает объект `From`, который предоставляет метод для получения значения отправителя.</span><span class="sxs-lookup"><span data-stu-id="542eb-364">The `from` property returns a `From` object that provides a method to get the from value.</span></span>

```javascript
Office.context.mailbox.item.from.getAsync(callback);

function callback(asyncResult) {
  var from = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="542eb-365">Тип:</span><span class="sxs-lookup"><span data-stu-id="542eb-365">Type:</span></span>

*   <span data-ttu-id="542eb-366">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [From](/javascript/api/outlook/office.from)</span><span class="sxs-lookup"><span data-stu-id="542eb-366">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [From](/javascript/api/outlook/office.from)</span></span>

##### <a name="requirements"></a><span data-ttu-id="542eb-367">Требования</span><span class="sxs-lookup"><span data-stu-id="542eb-367">Requirements</span></span>

|<span data-ttu-id="542eb-368">Требование</span><span class="sxs-lookup"><span data-stu-id="542eb-368">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="542eb-369">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="542eb-369">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="542eb-370">1.0</span><span class="sxs-lookup"><span data-stu-id="542eb-370">1.0</span></span>|<span data-ttu-id="542eb-371">1.7</span><span class="sxs-lookup"><span data-stu-id="542eb-371">1.7</span></span>|
|[<span data-ttu-id="542eb-372">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="542eb-372">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="542eb-373">ReadItem</span><span class="sxs-lookup"><span data-stu-id="542eb-373">ReadItem</span></span>|<span data-ttu-id="542eb-374">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="542eb-374">ReadWriteItem</span></span>|
|[<span data-ttu-id="542eb-375">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="542eb-375">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="542eb-376">Чтение</span><span class="sxs-lookup"><span data-stu-id="542eb-376">Read</span></span>|<span data-ttu-id="542eb-377">Создание</span><span class="sxs-lookup"><span data-stu-id="542eb-377">Compose</span></span>|

#### <a name="internetheaders-internetheadersjavascriptapioutlookofficeinternetheaders"></a><span data-ttu-id="542eb-378">internetHeaders :[InternetHeaders](/javascript/api/outlook/office.internetheaders)</span><span class="sxs-lookup"><span data-stu-id="542eb-378">internetHeaders :[InternetHeaders](/javascript/api/outlook/office.internetheaders)</span></span>

<span data-ttu-id="542eb-379">Получает или задает заголовки Интернета в сообщении.</span><span class="sxs-lookup"><span data-stu-id="542eb-379">Gets or sets the internet headers of a message.</span></span>

##### <a name="type"></a><span data-ttu-id="542eb-380">Тип:</span><span class="sxs-lookup"><span data-stu-id="542eb-380">Type:</span></span>

*   [<span data-ttu-id="542eb-381">InternetHeaders</span><span class="sxs-lookup"><span data-stu-id="542eb-381">InternetHeaders</span></span>](/javascript/api/outlook/office.internetheaders)

##### <a name="requirements"></a><span data-ttu-id="542eb-382">Требования</span><span class="sxs-lookup"><span data-stu-id="542eb-382">Requirements</span></span>

|<span data-ttu-id="542eb-383">Requirement</span><span class="sxs-lookup"><span data-stu-id="542eb-383">Requirement</span></span>|<span data-ttu-id="542eb-384">Значение</span><span class="sxs-lookup"><span data-stu-id="542eb-384">Value</span></span>|
|---|---|
|[<span data-ttu-id="542eb-385">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="542eb-385">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="542eb-386">Предварительная версия</span><span class="sxs-lookup"><span data-stu-id="542eb-386">Preview</span></span>|
|[<span data-ttu-id="542eb-387">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="542eb-387">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="542eb-388">ReadItem</span><span class="sxs-lookup"><span data-stu-id="542eb-388">ReadItem</span></span>|
|[<span data-ttu-id="542eb-389">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="542eb-389">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="542eb-390">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="542eb-390">Compose or read</span></span>|

#### <a name="internetmessageid-string"></a><span data-ttu-id="542eb-391">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="542eb-391">internetMessageId :String</span></span>

<span data-ttu-id="542eb-p113">Получает идентификатор интернет-сообщения для электронного сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="542eb-p113">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="542eb-394">Тип:</span><span class="sxs-lookup"><span data-stu-id="542eb-394">Type:</span></span>

*   <span data-ttu-id="542eb-395">String</span><span class="sxs-lookup"><span data-stu-id="542eb-395">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="542eb-396">Требования</span><span class="sxs-lookup"><span data-stu-id="542eb-396">Requirements</span></span>

|<span data-ttu-id="542eb-397">Requirement</span><span class="sxs-lookup"><span data-stu-id="542eb-397">Requirement</span></span>|<span data-ttu-id="542eb-398">Значение</span><span class="sxs-lookup"><span data-stu-id="542eb-398">Value</span></span>|
|---|---|
|[<span data-ttu-id="542eb-399">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="542eb-399">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="542eb-400">1.0</span><span class="sxs-lookup"><span data-stu-id="542eb-400">1.0</span></span>|
|[<span data-ttu-id="542eb-401">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="542eb-401">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="542eb-402">ReadItem</span><span class="sxs-lookup"><span data-stu-id="542eb-402">ReadItem</span></span>|
|[<span data-ttu-id="542eb-403">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="542eb-403">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="542eb-404">Чтение</span><span class="sxs-lookup"><span data-stu-id="542eb-404">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="542eb-405">Пример</span><span class="sxs-lookup"><span data-stu-id="542eb-405">Example</span></span>

```javascript
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="542eb-406">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="542eb-406">itemClass :String</span></span>

<span data-ttu-id="542eb-p114">Получает класс элемента веб-служб Exchange для выбранного элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="542eb-p114">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="542eb-p115">Свойство `itemClass` указывает класс сообщения выбранного элемента. Ниже приводятся классы сообщения по умолчанию для элемента сообщения или встречи.</span><span class="sxs-lookup"><span data-stu-id="542eb-p115">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

|<span data-ttu-id="542eb-411">Тип</span><span class="sxs-lookup"><span data-stu-id="542eb-411">Type</span></span>|<span data-ttu-id="542eb-412">Описание</span><span class="sxs-lookup"><span data-stu-id="542eb-412">Description</span></span>|<span data-ttu-id="542eb-413">Класс элемента</span><span class="sxs-lookup"><span data-stu-id="542eb-413">item class</span></span>|
|---|---|---|
|<span data-ttu-id="542eb-414">Элементы встречи</span><span class="sxs-lookup"><span data-stu-id="542eb-414">Appointment items</span></span>|<span data-ttu-id="542eb-415">Это элементы календаря для класса элемента `IPM.Appointment` или `IPM.Appointment.Occurence`.</span><span class="sxs-lookup"><span data-stu-id="542eb-415">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurence`.</span></span>|`IPM.Appointment`<br />`IPM.Appointment.Occurence`|
|<span data-ttu-id="542eb-416">Элементы сообщения</span><span class="sxs-lookup"><span data-stu-id="542eb-416">Message items</span></span>|<span data-ttu-id="542eb-417">Сюда входят электронные сообщения, для которых по умолчанию задан класс сообщения `IPM.Note`, а также приглашения на собрания, ответы на них и уведомления об их отмене, использующие `IPM.Schedule.Meeting` в качестве базового класса сообщения.</span><span class="sxs-lookup"><span data-stu-id="542eb-417">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span>|`IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled`|

<span data-ttu-id="542eb-418">Можно создавать настраиваемые классы сообщения, расширяющие классы сообщения по умолчанию, например настраиваемый класс сообщения о встрече `IPM.Appointment.Contoso`.</span><span class="sxs-lookup"><span data-stu-id="542eb-418">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="542eb-419">Тип:</span><span class="sxs-lookup"><span data-stu-id="542eb-419">Type:</span></span>

*   <span data-ttu-id="542eb-420">String</span><span class="sxs-lookup"><span data-stu-id="542eb-420">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="542eb-421">Требования</span><span class="sxs-lookup"><span data-stu-id="542eb-421">Requirements</span></span>

|<span data-ttu-id="542eb-422">Requirement</span><span class="sxs-lookup"><span data-stu-id="542eb-422">Requirement</span></span>|<span data-ttu-id="542eb-423">Значение</span><span class="sxs-lookup"><span data-stu-id="542eb-423">Value</span></span>|
|---|---|
|[<span data-ttu-id="542eb-424">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="542eb-424">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="542eb-425">1.0</span><span class="sxs-lookup"><span data-stu-id="542eb-425">1.0</span></span>|
|[<span data-ttu-id="542eb-426">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="542eb-426">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="542eb-427">ReadItem</span><span class="sxs-lookup"><span data-stu-id="542eb-427">ReadItem</span></span>|
|[<span data-ttu-id="542eb-428">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="542eb-428">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="542eb-429">Чтение</span><span class="sxs-lookup"><span data-stu-id="542eb-429">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="542eb-430">Пример</span><span class="sxs-lookup"><span data-stu-id="542eb-430">Example</span></span>

```javascript
var itemClass = Office.context.mailbox.item.itemClass;
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="542eb-431">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="542eb-431">(nullable) itemId :String</span></span>

<span data-ttu-id="542eb-p116">Получает идентификатор элемента веб-служб Exchange для текущего элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="542eb-p116">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="542eb-434">Идентификатор, возвращаемый свойством `itemId`, совпадает с идентификатором элемента веб-служб Exchange.</span><span class="sxs-lookup"><span data-stu-id="542eb-434">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="542eb-435">Свойство `itemId` не совпадает с идентификатором записи Outlook, а также идентификатором, который используется REST API Outlook.</span><span class="sxs-lookup"><span data-stu-id="542eb-435">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="542eb-436">Прежде чем совершать вызовы REST API, используя это значение, его необходимо преобразовать с помощью [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="542eb-436">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="542eb-437">Дополнительные сведения см. в статье [Использование REST API Outlook из надстройки Outlook](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="542eb-437">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="542eb-p118">Свойство `itemId` недоступно в режиме создания. Если требуется идентификатор элемента, с помощью метода [`saveAsync`](#saveasyncoptions-callback) можно сохранить элемент в хранилище. При этом в параметре [`AsyncResult.value`](/javascript/api/office/office.asyncresult) функции обратного вызова возвращается идентификатор элемента.</span><span class="sxs-lookup"><span data-stu-id="542eb-p118">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="542eb-440">Тип:</span><span class="sxs-lookup"><span data-stu-id="542eb-440">Type:</span></span>

*   <span data-ttu-id="542eb-441">String</span><span class="sxs-lookup"><span data-stu-id="542eb-441">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="542eb-442">Требования</span><span class="sxs-lookup"><span data-stu-id="542eb-442">Requirements</span></span>

|<span data-ttu-id="542eb-443">Requirement</span><span class="sxs-lookup"><span data-stu-id="542eb-443">Requirement</span></span>|<span data-ttu-id="542eb-444">Значение</span><span class="sxs-lookup"><span data-stu-id="542eb-444">Value</span></span>|
|---|---|
|[<span data-ttu-id="542eb-445">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="542eb-445">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="542eb-446">1.0</span><span class="sxs-lookup"><span data-stu-id="542eb-446">1.0</span></span>|
|[<span data-ttu-id="542eb-447">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="542eb-447">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="542eb-448">ReadItem</span><span class="sxs-lookup"><span data-stu-id="542eb-448">ReadItem</span></span>|
|[<span data-ttu-id="542eb-449">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="542eb-449">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="542eb-450">Чтение</span><span class="sxs-lookup"><span data-stu-id="542eb-450">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="542eb-451">Пример</span><span class="sxs-lookup"><span data-stu-id="542eb-451">Example</span></span>

<span data-ttu-id="542eb-p119">Указанный ниже код проверяет наличие идентификатора элемента. Если свойство `itemId` возвращает значение `null` или `undefined`, элемент будет сохранен в хранилище, а из асинхронного результата будет получен идентификатор элемента.</span><span class="sxs-lookup"><span data-stu-id="542eb-p119">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```javascript
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result){
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype"></a><span data-ttu-id="542eb-454">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="542eb-454">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="542eb-455">Получает тип элемента, который представляет экземпляр.</span><span class="sxs-lookup"><span data-stu-id="542eb-455">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="542eb-456">Свойство `itemType` возвращает одно из значений перечисления `ItemType`, которое указывает, является ли экземпляр объекта `item` сообщением или встречей.</span><span class="sxs-lookup"><span data-stu-id="542eb-456">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="542eb-457">Тип:</span><span class="sxs-lookup"><span data-stu-id="542eb-457">Type:</span></span>

*   [<span data-ttu-id="542eb-458">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="542eb-458">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="542eb-459">Требования</span><span class="sxs-lookup"><span data-stu-id="542eb-459">Requirements</span></span>

|<span data-ttu-id="542eb-460">Requirement</span><span class="sxs-lookup"><span data-stu-id="542eb-460">Requirement</span></span>|<span data-ttu-id="542eb-461">Значение</span><span class="sxs-lookup"><span data-stu-id="542eb-461">Value</span></span>|
|---|---|
|[<span data-ttu-id="542eb-462">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="542eb-462">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="542eb-463">1.0</span><span class="sxs-lookup"><span data-stu-id="542eb-463">1.0</span></span>|
|[<span data-ttu-id="542eb-464">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="542eb-464">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="542eb-465">ReadItem</span><span class="sxs-lookup"><span data-stu-id="542eb-465">ReadItem</span></span>|
|[<span data-ttu-id="542eb-466">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="542eb-466">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="542eb-467">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="542eb-467">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="542eb-468">Пример</span><span class="sxs-lookup"><span data-stu-id="542eb-468">Example</span></span>

```javascript
if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Message)
  // do something
else
  // do something else
```

####  <a name="location-stringlocationjavascriptapioutlookofficelocation"></a><span data-ttu-id="542eb-469">location :String|[Location](/javascript/api/outlook/office.location)</span><span class="sxs-lookup"><span data-stu-id="542eb-469">location :String|[Location](/javascript/api/outlook/office.location)</span></span>

<span data-ttu-id="542eb-470">Получает или задает место встречи.</span><span class="sxs-lookup"><span data-stu-id="542eb-470">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="542eb-471">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="542eb-471">Read mode</span></span>

<span data-ttu-id="542eb-472">Свойство `location` возвращает строку, содержащую сведения о месте встречи.</span><span class="sxs-lookup"><span data-stu-id="542eb-472">The `location` property returns a string that contains the location of the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="542eb-473">Режим создания</span><span class="sxs-lookup"><span data-stu-id="542eb-473">Compose mode</span></span>

<span data-ttu-id="542eb-474">Свойство `location` возвращает объект `Location`, предоставляющий методы, которые используются для получения и задания места встречи.</span><span class="sxs-lookup"><span data-stu-id="542eb-474">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="542eb-475">Тип:</span><span class="sxs-lookup"><span data-stu-id="542eb-475">Type:</span></span>

*   <span data-ttu-id="542eb-476">String | [Location](/javascript/api/outlook/office.location)</span><span class="sxs-lookup"><span data-stu-id="542eb-476">String | [Location](/javascript/api/outlook/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="542eb-477">Требования</span><span class="sxs-lookup"><span data-stu-id="542eb-477">Requirements</span></span>

|<span data-ttu-id="542eb-478">Requirement</span><span class="sxs-lookup"><span data-stu-id="542eb-478">Requirement</span></span>|<span data-ttu-id="542eb-479">Значение</span><span class="sxs-lookup"><span data-stu-id="542eb-479">Value</span></span>|
|---|---|
|[<span data-ttu-id="542eb-480">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="542eb-480">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="542eb-481">1.0</span><span class="sxs-lookup"><span data-stu-id="542eb-481">1.0</span></span>|
|[<span data-ttu-id="542eb-482">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="542eb-482">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="542eb-483">ReadItem</span><span class="sxs-lookup"><span data-stu-id="542eb-483">ReadItem</span></span>|
|[<span data-ttu-id="542eb-484">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="542eb-484">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="542eb-485">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="542eb-485">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="542eb-486">Пример</span><span class="sxs-lookup"><span data-stu-id="542eb-486">Example</span></span>

```javascript
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

#### <a name="normalizedsubject-string"></a><span data-ttu-id="542eb-487">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="542eb-487">normalizedSubject :String</span></span>

<span data-ttu-id="542eb-p120">Получает тему элемента со всеми удаленными префиксами (включая `RE:` и `FWD:`). Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="542eb-p120">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="542eb-p121">Свойство normalizedSubject получает тему элемента со стандартными префиксами (такими как `RE:` и `FW:`), добавляемыми почтовыми программами. Для получения темы элемента с неизмененными префиксами используйте свойство [`subject`](#subject-stringsubjectjavascriptapioutlookofficesubject).</span><span class="sxs-lookup"><span data-stu-id="542eb-p121">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubjectjavascriptapioutlookofficesubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="542eb-492">Тип:</span><span class="sxs-lookup"><span data-stu-id="542eb-492">Type:</span></span>

*   <span data-ttu-id="542eb-493">String</span><span class="sxs-lookup"><span data-stu-id="542eb-493">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="542eb-494">Требования</span><span class="sxs-lookup"><span data-stu-id="542eb-494">Requirements</span></span>

|<span data-ttu-id="542eb-495">Requirement</span><span class="sxs-lookup"><span data-stu-id="542eb-495">Requirement</span></span>|<span data-ttu-id="542eb-496">Значение</span><span class="sxs-lookup"><span data-stu-id="542eb-496">Value</span></span>|
|---|---|
|[<span data-ttu-id="542eb-497">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="542eb-497">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="542eb-498">1.0</span><span class="sxs-lookup"><span data-stu-id="542eb-498">1.0</span></span>|
|[<span data-ttu-id="542eb-499">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="542eb-499">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="542eb-500">ReadItem</span><span class="sxs-lookup"><span data-stu-id="542eb-500">ReadItem</span></span>|
|[<span data-ttu-id="542eb-501">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="542eb-501">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="542eb-502">Чтение</span><span class="sxs-lookup"><span data-stu-id="542eb-502">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="542eb-503">Пример</span><span class="sxs-lookup"><span data-stu-id="542eb-503">Example</span></span>

```javascript
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
```

####  <a name="notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessages"></a><span data-ttu-id="542eb-504">notificationMessages :[NotificationMessages](/javascript/api/outlook/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="542eb-504">notificationMessages :[NotificationMessages](/javascript/api/outlook/office.notificationmessages)</span></span>

<span data-ttu-id="542eb-505">Получает сообщения уведомления для элемента.</span><span class="sxs-lookup"><span data-stu-id="542eb-505">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="542eb-506">Тип:</span><span class="sxs-lookup"><span data-stu-id="542eb-506">Type:</span></span>

*   [<span data-ttu-id="542eb-507">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="542eb-507">NotificationMessages</span></span>](/javascript/api/outlook/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="542eb-508">Требования</span><span class="sxs-lookup"><span data-stu-id="542eb-508">Requirements</span></span>

|<span data-ttu-id="542eb-509">Requirement</span><span class="sxs-lookup"><span data-stu-id="542eb-509">Requirement</span></span>|<span data-ttu-id="542eb-510">Значение</span><span class="sxs-lookup"><span data-stu-id="542eb-510">Value</span></span>|
|---|---|
|[<span data-ttu-id="542eb-511">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="542eb-511">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="542eb-512">1.3</span><span class="sxs-lookup"><span data-stu-id="542eb-512">1.3</span></span>|
|[<span data-ttu-id="542eb-513">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="542eb-513">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="542eb-514">ReadItem</span><span class="sxs-lookup"><span data-stu-id="542eb-514">ReadItem</span></span>|
|[<span data-ttu-id="542eb-515">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="542eb-515">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="542eb-516">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="542eb-516">Compose or read</span></span>|

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="542eb-517">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="542eb-517">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="542eb-518">Предоставляет доступ к необязательным участникам события.</span><span class="sxs-lookup"><span data-stu-id="542eb-518">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="542eb-519">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="542eb-519">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="542eb-520">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="542eb-520">Read mode</span></span>

<span data-ttu-id="542eb-521">Свойство `optionalAttendees` возвращает массив, содержащий объект `EmailAddressDetails` для каждого необязательного участника собрания.</span><span class="sxs-lookup"><span data-stu-id="542eb-521">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="542eb-522">Режим создания</span><span class="sxs-lookup"><span data-stu-id="542eb-522">Compose mode</span></span>

<span data-ttu-id="542eb-523">Свойство `optionalAttendees` возвращает объект `Recipients`, который предоставляет методы для получения или обновления необязательных участников собрания.</span><span class="sxs-lookup"><span data-stu-id="542eb-523">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="542eb-524">Тип:</span><span class="sxs-lookup"><span data-stu-id="542eb-524">Type:</span></span>

*   <span data-ttu-id="542eb-525">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="542eb-525">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="542eb-526">Требования</span><span class="sxs-lookup"><span data-stu-id="542eb-526">Requirements</span></span>

|<span data-ttu-id="542eb-527">Requirement</span><span class="sxs-lookup"><span data-stu-id="542eb-527">Requirement</span></span>|<span data-ttu-id="542eb-528">Значение</span><span class="sxs-lookup"><span data-stu-id="542eb-528">Value</span></span>|
|---|---|
|[<span data-ttu-id="542eb-529">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="542eb-529">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="542eb-530">1.0</span><span class="sxs-lookup"><span data-stu-id="542eb-530">1.0</span></span>|
|[<span data-ttu-id="542eb-531">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="542eb-531">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="542eb-532">ReadItem</span><span class="sxs-lookup"><span data-stu-id="542eb-532">ReadItem</span></span>|
|[<span data-ttu-id="542eb-533">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="542eb-533">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="542eb-534">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="542eb-534">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="542eb-535">Пример</span><span class="sxs-lookup"><span data-stu-id="542eb-535">Example</span></span>

```javascript
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsorganizerjavascriptapioutlookofficeorganizer"></a><span data-ttu-id="542eb-536">organizer :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[Organizer](/javascript/api/outlook/office.organizer)</span><span class="sxs-lookup"><span data-stu-id="542eb-536">organizer :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[Organizer](/javascript/api/outlook/office.organizer)</span></span>

<span data-ttu-id="542eb-537">Получает адрес электронной почты организатора указанного собрания.</span><span class="sxs-lookup"><span data-stu-id="542eb-537">Gets the email address of the organizer for a specified meeting.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="542eb-538">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="542eb-538">Read mode</span></span>

<span data-ttu-id="542eb-539">Свойство `organizer` возвращает объект [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails), представляющий организатора собрания.</span><span class="sxs-lookup"><span data-stu-id="542eb-539">The `organizer` property returns an [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) object that represents the meeting organizer.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="542eb-540">Режим создания</span><span class="sxs-lookup"><span data-stu-id="542eb-540">Compose mode</span></span>

<span data-ttu-id="542eb-541">Свойство `organizer` возвращает объект [Organizer](/javascript/api/outlook/office.organizer), который предоставляет метод для получения значения организатора.</span><span class="sxs-lookup"><span data-stu-id="542eb-541">The `organizer` property returns an [Organizer](/javascript/api/outlook/office.organizer) object that provides a method to get the organizer value.</span></span>

##### <a name="type"></a><span data-ttu-id="542eb-542">Тип:</span><span class="sxs-lookup"><span data-stu-id="542eb-542">Type:</span></span>

*   <span data-ttu-id="542eb-543">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [Organizer](/javascript/api/outlook/office.organizer)</span><span class="sxs-lookup"><span data-stu-id="542eb-543">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [Organizer](/javascript/api/outlook/office.organizer)</span></span>

##### <a name="requirements"></a><span data-ttu-id="542eb-544">Требования</span><span class="sxs-lookup"><span data-stu-id="542eb-544">Requirements</span></span>

|<span data-ttu-id="542eb-545">Требование</span><span class="sxs-lookup"><span data-stu-id="542eb-545">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="542eb-546">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="542eb-546">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="542eb-547">1.0</span><span class="sxs-lookup"><span data-stu-id="542eb-547">1.0</span></span>|<span data-ttu-id="542eb-548">1.7</span><span class="sxs-lookup"><span data-stu-id="542eb-548">1.7</span></span>|
|[<span data-ttu-id="542eb-549">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="542eb-549">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="542eb-550">ReadItem</span><span class="sxs-lookup"><span data-stu-id="542eb-550">ReadItem</span></span>|<span data-ttu-id="542eb-551">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="542eb-551">ReadWriteItem</span></span>|
|[<span data-ttu-id="542eb-552">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="542eb-552">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="542eb-553">Чтение</span><span class="sxs-lookup"><span data-stu-id="542eb-553">Read</span></span>|<span data-ttu-id="542eb-554">Создание</span><span class="sxs-lookup"><span data-stu-id="542eb-554">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="542eb-555">Пример</span><span class="sxs-lookup"><span data-stu-id="542eb-555">Example</span></span>

```javascript
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
```

#### <a name="nullable-recurrence-recurrencejavascriptapioutlookofficerecurrence"></a><span data-ttu-id="542eb-556">(nullable) recurrence :[Recurrence](/javascript/api/outlook/office.recurrence)</span><span class="sxs-lookup"><span data-stu-id="542eb-556">(nullable) recurrence :[Recurrence](/javascript/api/outlook/office.recurrence)</span></span>

<span data-ttu-id="542eb-557">Получает или задает расписание повторения для встречи.</span><span class="sxs-lookup"><span data-stu-id="542eb-557">Gets or sets the recurrence pattern of an appointment.</span></span> <span data-ttu-id="542eb-558">Получает расписание повторения для приглашения на собрание.</span><span class="sxs-lookup"><span data-stu-id="542eb-558">Gets the recurrence pattern of a meeting request.</span></span> <span data-ttu-id="542eb-559">Доступно в режимах чтения и создания для элементов встречи.</span><span class="sxs-lookup"><span data-stu-id="542eb-559">Read and compose modes for appointment items.</span></span> <span data-ttu-id="542eb-560">Доступно в режиме чтения для элементов приглашения на собрание.</span><span class="sxs-lookup"><span data-stu-id="542eb-560">Read mode for meeting request items.</span></span>

<span data-ttu-id="542eb-561">Свойство `recurrence` возвращает объект [recurrence](/javascript/api/outlook/office.recurrence) для повторяющихся встреч или приглашений на собрание, если элемент представляет собой серию или экземпляр в пределах серии.</span><span class="sxs-lookup"><span data-stu-id="542eb-561">The `recurrence` property returns a [recurrence](/javascript/api/outlook/office.recurrence) object for recurring appointments or meetings requests if an item is a series or an instance in a series.</span></span> <span data-ttu-id="542eb-562">Значение `null` возвращается для отдельных встреч и приглашений на собрания, связанных с одной встречей.</span><span class="sxs-lookup"><span data-stu-id="542eb-562">`null` is returned for single appointments and meeting requests of single appointments.</span></span> <span data-ttu-id="542eb-563">Значение `undefined` возвращается для сообщений, которые не являются приглашениями на собрания.</span><span class="sxs-lookup"><span data-stu-id="542eb-563">`undefined` is returned for messages that are not meeting requests.</span></span>

> <span data-ttu-id="542eb-564">Примечание. Приглашения на собрания имеют значение `itemClass` для класса IPM.Schedule.Meeting.Request.</span><span class="sxs-lookup"><span data-stu-id="542eb-564">Note: Meeting requests have an `itemClass` value of IPM.Schedule.Meeting.Request.</span></span>

> <span data-ttu-id="542eb-565">Примечание. Если объект recurrence имеет значение `null`, он представляет собой отдельную встречу или приглашение на собрание, связанное с одной встречей, и НЕ входит в серию.</span><span class="sxs-lookup"><span data-stu-id="542eb-565">Note: If the recurrence object is `null`, this indicates that the object is a single appointment or a meeting request of a single appointment and NOT a part of a series.</span></span>

##### <a name="type"></a><span data-ttu-id="542eb-566">Тип:</span><span class="sxs-lookup"><span data-stu-id="542eb-566">Type:</span></span>

* [<span data-ttu-id="542eb-567">Recurrence</span><span class="sxs-lookup"><span data-stu-id="542eb-567">Recurrence</span></span>](/javascript/api/outlook/office.recurrence)

|<span data-ttu-id="542eb-568">Требование</span><span class="sxs-lookup"><span data-stu-id="542eb-568">Requirement</span></span>|<span data-ttu-id="542eb-569">Значение</span><span class="sxs-lookup"><span data-stu-id="542eb-569">Value</span></span>|
|---|---|
|[<span data-ttu-id="542eb-570">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="542eb-570">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="542eb-571">1.7</span><span class="sxs-lookup"><span data-stu-id="542eb-571">1.7</span></span>|
|[<span data-ttu-id="542eb-572">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="542eb-572">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="542eb-573">ReadItem</span><span class="sxs-lookup"><span data-stu-id="542eb-573">ReadItem</span></span>|
|[<span data-ttu-id="542eb-574">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="542eb-574">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="542eb-575">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="542eb-575">Compose or read</span></span>|

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="542eb-576">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="542eb-576">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="542eb-577">Предоставляет доступ к обязательным участникам события.</span><span class="sxs-lookup"><span data-stu-id="542eb-577">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="542eb-578">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="542eb-578">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="542eb-579">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="542eb-579">Read mode</span></span>

<span data-ttu-id="542eb-580">Свойство `requiredAttendees` возвращает массив, содержащий объект `EmailAddressDetails` для каждого обязательного участника собрания.</span><span class="sxs-lookup"><span data-stu-id="542eb-580">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="542eb-581">Режим создания</span><span class="sxs-lookup"><span data-stu-id="542eb-581">Compose mode</span></span>

<span data-ttu-id="542eb-582">Свойство `requiredAttendees` возвращает объект `Recipients`, предоставляющий методы, с помощью которых можно получить или обновить сведения об обязательных участниках собрания.</span><span class="sxs-lookup"><span data-stu-id="542eb-582">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="542eb-583">Тип:</span><span class="sxs-lookup"><span data-stu-id="542eb-583">Type:</span></span>

*   <span data-ttu-id="542eb-584">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="542eb-584">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="542eb-585">Требования</span><span class="sxs-lookup"><span data-stu-id="542eb-585">Requirements</span></span>

|<span data-ttu-id="542eb-586">Requirement</span><span class="sxs-lookup"><span data-stu-id="542eb-586">Requirement</span></span>|<span data-ttu-id="542eb-587">Значение</span><span class="sxs-lookup"><span data-stu-id="542eb-587">Value</span></span>|
|---|---|
|[<span data-ttu-id="542eb-588">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="542eb-588">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="542eb-589">1.0</span><span class="sxs-lookup"><span data-stu-id="542eb-589">1.0</span></span>|
|[<span data-ttu-id="542eb-590">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="542eb-590">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="542eb-591">ReadItem</span><span class="sxs-lookup"><span data-stu-id="542eb-591">ReadItem</span></span>|
|[<span data-ttu-id="542eb-592">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="542eb-592">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="542eb-593">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="542eb-593">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="542eb-594">Пример</span><span class="sxs-lookup"><span data-stu-id="542eb-594">Example</span></span>

```javascript
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
}
```

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails"></a><span data-ttu-id="542eb-595">sender :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="542eb-595">sender :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)</span></span>

<span data-ttu-id="542eb-p126">Получает электронный адрес отправителя электронного сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="542eb-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="542eb-p127">Свойства [`from`](#from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom) и `sender` представляют одно лицо, если сообщение не отправлено представителем. В противном случае свойство `from` представляет лицо, делегировавшее полномочия, а свойство sender — представителя.</span><span class="sxs-lookup"><span data-stu-id="542eb-p127">The [`from`](#from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="542eb-600">Свойству `recipientType`, принадлежащему объекту `EmailAddressDetails` в свойстве `sender`, задано значение `undefined`.</span><span class="sxs-lookup"><span data-stu-id="542eb-600">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="542eb-601">Тип:</span><span class="sxs-lookup"><span data-stu-id="542eb-601">Type:</span></span>

*   [<span data-ttu-id="542eb-602">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="542eb-602">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="542eb-603">Требования</span><span class="sxs-lookup"><span data-stu-id="542eb-603">Requirements</span></span>

|<span data-ttu-id="542eb-604">Requirement</span><span class="sxs-lookup"><span data-stu-id="542eb-604">Requirement</span></span>|<span data-ttu-id="542eb-605">Значение</span><span class="sxs-lookup"><span data-stu-id="542eb-605">Value</span></span>|
|---|---|
|[<span data-ttu-id="542eb-606">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="542eb-606">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="542eb-607">1.0</span><span class="sxs-lookup"><span data-stu-id="542eb-607">1.0</span></span>|
|[<span data-ttu-id="542eb-608">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="542eb-608">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="542eb-609">ReadItem</span><span class="sxs-lookup"><span data-stu-id="542eb-609">ReadItem</span></span>|
|[<span data-ttu-id="542eb-610">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="542eb-610">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="542eb-611">Чтение</span><span class="sxs-lookup"><span data-stu-id="542eb-611">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="542eb-612">Пример</span><span class="sxs-lookup"><span data-stu-id="542eb-612">Example</span></span>

```javascript
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
```

#### <a name="nullable-seriesid-string"></a><span data-ttu-id="542eb-613">(nullable) seriesId :String</span><span class="sxs-lookup"><span data-stu-id="542eb-613">(nullable) seriesId :String</span></span>

<span data-ttu-id="542eb-614">Получает идентификатор серии, к которой относится экземпляр.</span><span class="sxs-lookup"><span data-stu-id="542eb-614">Gets the id of the series that an instance belongs to.</span></span>

<span data-ttu-id="542eb-615">В Outlook Web App и Outlook свойство `seriesId` возвращает идентификатор веб-служб Exchange (EWS) родительского элемента (серии), к которому относится этот элемент.</span><span class="sxs-lookup"><span data-stu-id="542eb-615">In OWA and Outlook, the `seriesId` returns the Exchange Web Services (EWS) ID of the parent (series) item that this item belongs to.</span></span> <span data-ttu-id="542eb-616">Однако в iOS и Android свойство `seriesId` возвращает идентификатор REST родительского элемента.</span><span class="sxs-lookup"><span data-stu-id="542eb-616">However, in iOS and Android, the `seriesId` returns the REST ID of the parent item.</span></span>

> [!NOTE]
> <span data-ttu-id="542eb-617">Идентификатор, возвращаемый свойством `seriesId`, совпадает с идентификатором элемента веб-служб Exchange.</span><span class="sxs-lookup"><span data-stu-id="542eb-617">The identifier returned by the `seriesId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="542eb-618">Свойство `seriesId` не совпадает с идентификаторами Outlook, которые используются в REST API Outlook.</span><span class="sxs-lookup"><span data-stu-id="542eb-618">The `seriesId` property is not identical to the Outlook IDs used by the Outlook REST API.</span></span> <span data-ttu-id="542eb-619">Прежде чем совершать вызовы REST API, используя это значение, его необходимо преобразовать с помощью [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="542eb-619">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="542eb-620">Дополнительные сведения см. в статье [Использование REST API Outlook из надстройки Outlook](https://docs.microsoft.com/outlook/add-ins/use-rest-api).</span><span class="sxs-lookup"><span data-stu-id="542eb-620">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api).</span></span>

<span data-ttu-id="542eb-621">Свойство `seriesId` возвращает значение `null` для элементов, у которых нет родительских элементов, например отдельных встреч, элементов серий или приглашений на собрания, и возвращает значение `undefined` для всех других элементов, которые не представляют собой приглашения на собрания.</span><span class="sxs-lookup"><span data-stu-id="542eb-621">The `seriesId` property returns `null` for items that do not have parent items such as single appointments, series items, or meeting requests and returns `undefined` for any other items that are not meeting requests.</span></span>

##### <a name="type"></a><span data-ttu-id="542eb-622">Тип:</span><span class="sxs-lookup"><span data-stu-id="542eb-622">Type:</span></span>

* <span data-ttu-id="542eb-623">String</span><span class="sxs-lookup"><span data-stu-id="542eb-623">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="542eb-624">Требования</span><span class="sxs-lookup"><span data-stu-id="542eb-624">Requirements</span></span>

|<span data-ttu-id="542eb-625">Requirement</span><span class="sxs-lookup"><span data-stu-id="542eb-625">Requirement</span></span>|<span data-ttu-id="542eb-626">Значение</span><span class="sxs-lookup"><span data-stu-id="542eb-626">Value</span></span>|
|---|---|
|[<span data-ttu-id="542eb-627">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="542eb-627">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="542eb-628">1.7</span><span class="sxs-lookup"><span data-stu-id="542eb-628">1.7</span></span>|
|[<span data-ttu-id="542eb-629">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="542eb-629">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="542eb-630">ReadItem</span><span class="sxs-lookup"><span data-stu-id="542eb-630">ReadItem</span></span>|
|[<span data-ttu-id="542eb-631">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="542eb-631">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="542eb-632">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="542eb-632">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="542eb-633">Пример</span><span class="sxs-lookup"><span data-stu-id="542eb-633">Example</span></span>

```javascript
var seriesId = Office.context.mailbox.item.seriesId;
var isSeries = (seriesId == null);
```

####  <a name="start-datetimejavascriptapioutlookofficetime"></a><span data-ttu-id="542eb-634">start :Date|[Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="542eb-634">start :Date|[Time](/javascript/api/outlook/office.time)</span></span>

<span data-ttu-id="542eb-635">Получает или задает дату и время начала встречи.</span><span class="sxs-lookup"><span data-stu-id="542eb-635">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="542eb-p130">Свойство `start` представлено в виде значения даты и времени в формате UTC. Это значение можно преобразовать в местные значения даты и времени клиента с помощью метода [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime).</span><span class="sxs-lookup"><span data-stu-id="542eb-p130">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="542eb-638">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="542eb-638">Read mode</span></span>

<span data-ttu-id="542eb-639">Свойство `start` возвращает объект `Date`.</span><span class="sxs-lookup"><span data-stu-id="542eb-639">The `start` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="542eb-640">Режим создания</span><span class="sxs-lookup"><span data-stu-id="542eb-640">Compose mode</span></span>

<span data-ttu-id="542eb-641">Свойство `start` возвращает объект `Time`.</span><span class="sxs-lookup"><span data-stu-id="542eb-641">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="542eb-642">Если вы задаете время начала с помощью метода [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-), необходимо использовать метод [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) для преобразования местного времени на клиенте в формат UTC для сервера.</span><span class="sxs-lookup"><span data-stu-id="542eb-642">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="542eb-643">Тип:</span><span class="sxs-lookup"><span data-stu-id="542eb-643">Type:</span></span>

*   <span data-ttu-id="542eb-644">Date | [Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="542eb-644">Date | [Time](/javascript/api/outlook/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="542eb-645">Требования</span><span class="sxs-lookup"><span data-stu-id="542eb-645">Requirements</span></span>

|<span data-ttu-id="542eb-646">Requirement</span><span class="sxs-lookup"><span data-stu-id="542eb-646">Requirement</span></span>|<span data-ttu-id="542eb-647">Значение</span><span class="sxs-lookup"><span data-stu-id="542eb-647">Value</span></span>|
|---|---|
|[<span data-ttu-id="542eb-648">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="542eb-648">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="542eb-649">1.0</span><span class="sxs-lookup"><span data-stu-id="542eb-649">1.0</span></span>|
|[<span data-ttu-id="542eb-650">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="542eb-650">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="542eb-651">ReadItem</span><span class="sxs-lookup"><span data-stu-id="542eb-651">ReadItem</span></span>|
|[<span data-ttu-id="542eb-652">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="542eb-652">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="542eb-653">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="542eb-653">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="542eb-654">Пример</span><span class="sxs-lookup"><span data-stu-id="542eb-654">Example</span></span>

<span data-ttu-id="542eb-655">В примере ниже с помощью метода [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) объекта `Time` задается время начала встречи в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="542eb-655">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

```javascript
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

####  <a name="subject-stringsubjectjavascriptapioutlookofficesubject"></a><span data-ttu-id="542eb-656">subject :String|[Subject](/javascript/api/outlook/office.subject)</span><span class="sxs-lookup"><span data-stu-id="542eb-656">subject :String|[Subject](/javascript/api/outlook/office.subject)</span></span>

<span data-ttu-id="542eb-657">Получает или задает описание, которое отображается в поле темы элемента.</span><span class="sxs-lookup"><span data-stu-id="542eb-657">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="542eb-658">Свойство `subject` получает или задает всю тему элемента для отправки с почтового сервера.</span><span class="sxs-lookup"><span data-stu-id="542eb-658">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="542eb-659">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="542eb-659">Read mode</span></span>

<span data-ttu-id="542eb-p131">Свойство `subject` возвращает строку. С помощью свойства [`normalizedSubject`](#normalizedsubject-string) можно получить тему без начальных префиксов, таких как `RE:` и `FW:`.</span><span class="sxs-lookup"><span data-stu-id="542eb-p131">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```javascript
var subject = Office.context.mailbox.item.subject;
```

##### <a name="compose-mode"></a><span data-ttu-id="542eb-662">Режим создания</span><span class="sxs-lookup"><span data-stu-id="542eb-662">Compose mode</span></span>

<span data-ttu-id="542eb-663">Свойство `subject` возвращает объект `Subject`, который предоставляет методы для получения и задания темы.</span><span class="sxs-lookup"><span data-stu-id="542eb-663">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```javascript
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="542eb-664">Тип:</span><span class="sxs-lookup"><span data-stu-id="542eb-664">Type:</span></span>

*   <span data-ttu-id="542eb-665">String | [Subject](/javascript/api/outlook/office.subject)</span><span class="sxs-lookup"><span data-stu-id="542eb-665">String | [Subject](/javascript/api/outlook/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="542eb-666">Требования</span><span class="sxs-lookup"><span data-stu-id="542eb-666">Requirements</span></span>

|<span data-ttu-id="542eb-667">Requirement</span><span class="sxs-lookup"><span data-stu-id="542eb-667">Requirement</span></span>|<span data-ttu-id="542eb-668">Значение</span><span class="sxs-lookup"><span data-stu-id="542eb-668">Value</span></span>|
|---|---|
|[<span data-ttu-id="542eb-669">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="542eb-669">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="542eb-670">1.0</span><span class="sxs-lookup"><span data-stu-id="542eb-670">1.0</span></span>|
|[<span data-ttu-id="542eb-671">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="542eb-671">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="542eb-672">ReadItem</span><span class="sxs-lookup"><span data-stu-id="542eb-672">ReadItem</span></span>|
|[<span data-ttu-id="542eb-673">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="542eb-673">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="542eb-674">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="542eb-674">Compose or read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="542eb-675">to :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="542eb-675">to :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="542eb-676">Предоставляет доступ к получателям, указанным в строке **Кому** сообщения.</span><span class="sxs-lookup"><span data-stu-id="542eb-676">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="542eb-677">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="542eb-677">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="542eb-678">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="542eb-678">Read mode</span></span>

<span data-ttu-id="542eb-p133">Свойство `to` возвращает массив, содержащий объект `EmailAddressDetails` для каждого получателя в строке **Кому** сообщения. Коллекция может включать не более 100 элементов.</span><span class="sxs-lookup"><span data-stu-id="542eb-p133">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="542eb-681">Режим создания</span><span class="sxs-lookup"><span data-stu-id="542eb-681">Compose mode</span></span>

<span data-ttu-id="542eb-682">Свойство `to` возвращает объект `Recipients`, предоставляющий методы для получения или обновления получателей, которые указаны в строке **Кому** сообщения.</span><span class="sxs-lookup"><span data-stu-id="542eb-682">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="542eb-683">Тип:</span><span class="sxs-lookup"><span data-stu-id="542eb-683">Type:</span></span>

*   <span data-ttu-id="542eb-684">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="542eb-684">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="542eb-685">Требования</span><span class="sxs-lookup"><span data-stu-id="542eb-685">Requirements</span></span>

|<span data-ttu-id="542eb-686">Requirement</span><span class="sxs-lookup"><span data-stu-id="542eb-686">Requirement</span></span>|<span data-ttu-id="542eb-687">Значение</span><span class="sxs-lookup"><span data-stu-id="542eb-687">Value</span></span>|
|---|---|
|[<span data-ttu-id="542eb-688">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="542eb-688">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="542eb-689">1.0</span><span class="sxs-lookup"><span data-stu-id="542eb-689">1.0</span></span>|
|[<span data-ttu-id="542eb-690">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="542eb-690">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="542eb-691">ReadItem</span><span class="sxs-lookup"><span data-stu-id="542eb-691">ReadItem</span></span>|
|[<span data-ttu-id="542eb-692">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="542eb-692">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="542eb-693">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="542eb-693">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="542eb-694">Пример</span><span class="sxs-lookup"><span data-stu-id="542eb-694">Example</span></span>

```javascript
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

### <a name="methods"></a><span data-ttu-id="542eb-695">Методы</span><span class="sxs-lookup"><span data-stu-id="542eb-695">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="542eb-696">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="542eb-696">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="542eb-697">Добавляет файл в сообщение или встречу в качестве вложения.</span><span class="sxs-lookup"><span data-stu-id="542eb-697">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="542eb-698">Метод `addFileAttachmentAsync` передает файл по указанному универсальному коду ресурса (URI) и вкладывает его в элемент в форме создания.</span><span class="sxs-lookup"><span data-stu-id="542eb-698">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="542eb-699">Идентификатор можно последовательно использовать с методом [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback), чтобы удалить вложение, добавленное во время текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="542eb-699">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="542eb-700">Параметры</span><span class="sxs-lookup"><span data-stu-id="542eb-700">Parameters:</span></span>
|<span data-ttu-id="542eb-701">Имя</span><span class="sxs-lookup"><span data-stu-id="542eb-701">Name</span></span>|<span data-ttu-id="542eb-702">Тип</span><span class="sxs-lookup"><span data-stu-id="542eb-702">Type</span></span>|<span data-ttu-id="542eb-703">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="542eb-703">Attributes</span></span>|<span data-ttu-id="542eb-704">Описание</span><span class="sxs-lookup"><span data-stu-id="542eb-704">Description</span></span>|
|---|---|---|---|
|`uri`|<span data-ttu-id="542eb-705">String</span><span class="sxs-lookup"><span data-stu-id="542eb-705">String</span></span>||<span data-ttu-id="542eb-p134">Универсальный код ресурса (URI), представляющий расположение файла, который нужно вложить в сообщение или встречу. Максимальная длина — 2048 символов.</span><span class="sxs-lookup"><span data-stu-id="542eb-p134">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="542eb-708">String</span><span class="sxs-lookup"><span data-stu-id="542eb-708">String</span></span>||<span data-ttu-id="542eb-p135">Имя вложения, которое отображается при передаче вложения. Максимальная длина — 255 символов.</span><span class="sxs-lookup"><span data-stu-id="542eb-p135">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="542eb-711">Object</span><span class="sxs-lookup"><span data-stu-id="542eb-711">Object</span></span>|<span data-ttu-id="542eb-712">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="542eb-712">&lt;optional&gt;</span></span>|<span data-ttu-id="542eb-713">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="542eb-713">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="542eb-714">Object</span><span class="sxs-lookup"><span data-stu-id="542eb-714">Object</span></span>|<span data-ttu-id="542eb-715">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="542eb-715">&lt;optional&gt;</span></span>|<span data-ttu-id="542eb-716">В методе обратного вызова разработчики могут указать любой объект, к которому необходимо получить доступ.</span><span class="sxs-lookup"><span data-stu-id="542eb-716">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="542eb-717">Boolean</span><span class="sxs-lookup"><span data-stu-id="542eb-717">Boolean</span></span>|<span data-ttu-id="542eb-718">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="542eb-718">&lt;optional&gt;</span></span>|<span data-ttu-id="542eb-719">Значение `true` указывает на то, что вложение будет встроено в текст сообщения и не должно отображаться в списке вложений.</span><span class="sxs-lookup"><span data-stu-id="542eb-719">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="542eb-720">function</span><span class="sxs-lookup"><span data-stu-id="542eb-720">function</span></span>|<span data-ttu-id="542eb-721">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="542eb-721">&lt;optional&gt;</span></span>|<span data-ttu-id="542eb-722">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="542eb-722">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="542eb-723">После успешного выполнения идентификатор вложения будет представлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="542eb-723">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="542eb-724">Если добавить вложение не удастся, объект `asyncResult` будет содержать объект `Error` с описанием ошибки.</span><span class="sxs-lookup"><span data-stu-id="542eb-724">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="542eb-725">Ошибки</span><span class="sxs-lookup"><span data-stu-id="542eb-725">Errors</span></span>

|<span data-ttu-id="542eb-726">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="542eb-726">Error code</span></span>|<span data-ttu-id="542eb-727">Описание</span><span class="sxs-lookup"><span data-stu-id="542eb-727">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="542eb-728">Вложение превышает максимальный размер.</span><span class="sxs-lookup"><span data-stu-id="542eb-728">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="542eb-729">Расширение вложения не поддерживается.</span><span class="sxs-lookup"><span data-stu-id="542eb-729">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="542eb-730">Сообщение или встреча содержат слишком много вложений.</span><span class="sxs-lookup"><span data-stu-id="542eb-730">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="542eb-731">Требования</span><span class="sxs-lookup"><span data-stu-id="542eb-731">Requirements</span></span>

|<span data-ttu-id="542eb-732">Requirement</span><span class="sxs-lookup"><span data-stu-id="542eb-732">Requirement</span></span>|<span data-ttu-id="542eb-733">Значение</span><span class="sxs-lookup"><span data-stu-id="542eb-733">Value</span></span>|
|---|---|
|[<span data-ttu-id="542eb-734">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="542eb-734">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="542eb-735">1.1</span><span class="sxs-lookup"><span data-stu-id="542eb-735">1.1</span></span>|
|[<span data-ttu-id="542eb-736">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="542eb-736">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="542eb-737">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="542eb-737">ReadWriteItem</span></span>|
|[<span data-ttu-id="542eb-738">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="542eb-738">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="542eb-739">Создание</span><span class="sxs-lookup"><span data-stu-id="542eb-739">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="542eb-740">Примеры</span><span class="sxs-lookup"><span data-stu-id="542eb-740">Examples</span></span>

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

<span data-ttu-id="542eb-741">В приведенном ниже примере файл изображения добавляется в качестве встроенного вложения, а ссылка на вложение добавляется в текст сообщения.</span><span class="sxs-lookup"><span data-stu-id="542eb-741">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

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

#### <a name="addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback"></a><span data-ttu-id="542eb-742">addFileAttachmentFromBase64Async(base64File, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="542eb-742">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="542eb-743">Добавляет файл из кодирования base64 в сообщение или встречу в качестве вложения.</span><span class="sxs-lookup"><span data-stu-id="542eb-743">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="542eb-744">Метод `addFileAttachmentFromBase64Async` передает файл из кодировки base64 и вкладывает его в элемент в форме создания.</span><span class="sxs-lookup"><span data-stu-id="542eb-744">The `addFileAttachmentFromBase64Async` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span> <span data-ttu-id="542eb-745">Этот способ возвращает идентификатор вложения в объекте AsyncResult.value.</span><span class="sxs-lookup"><span data-stu-id="542eb-745">This method returns the attachment identifier in the AsyncResult.value object.</span></span>

<span data-ttu-id="542eb-746">Идентификатор можно использовать с методом [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback), чтобы удалить вложение, добавленное во время текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="542eb-746">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="542eb-747">Параметры</span><span class="sxs-lookup"><span data-stu-id="542eb-747">Parameters:</span></span>
|<span data-ttu-id="542eb-748">Имя</span><span class="sxs-lookup"><span data-stu-id="542eb-748">Name</span></span>|<span data-ttu-id="542eb-749">Тип</span><span class="sxs-lookup"><span data-stu-id="542eb-749">Type</span></span>|<span data-ttu-id="542eb-750">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="542eb-750">Attributes</span></span>|<span data-ttu-id="542eb-751">Описание</span><span class="sxs-lookup"><span data-stu-id="542eb-751">Description</span></span>|
|---|---|---|---|
|`base64File`|<span data-ttu-id="542eb-752">Строка</span><span class="sxs-lookup"><span data-stu-id="542eb-752">String</span></span>||<span data-ttu-id="542eb-753">Закодированное содержимое base64 изображения или файла, которое следует добавить в сообщение электронной почты или событие.</span><span class="sxs-lookup"><span data-stu-id="542eb-753">The base64 encoded content of an image or file to be added to an email or event.</span></span>|
|`attachmentName`|<span data-ttu-id="542eb-754">Строка</span><span class="sxs-lookup"><span data-stu-id="542eb-754">String</span></span>||<span data-ttu-id="542eb-p137">Имя вложения, которое отображается при передаче вложения. Максимальная длина — 255 символов.</span><span class="sxs-lookup"><span data-stu-id="542eb-p137">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="542eb-757">Object</span><span class="sxs-lookup"><span data-stu-id="542eb-757">Object</span></span>|<span data-ttu-id="542eb-758">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="542eb-758">&lt;optional&gt;</span></span>|<span data-ttu-id="542eb-759">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="542eb-759">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="542eb-760">Object</span><span class="sxs-lookup"><span data-stu-id="542eb-760">Object</span></span>|<span data-ttu-id="542eb-761">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="542eb-761">&lt;optional&gt;</span></span>|<span data-ttu-id="542eb-762">В методе обратного вызова разработчики могут указать любой объект, к которому необходимо получить доступ.</span><span class="sxs-lookup"><span data-stu-id="542eb-762">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="542eb-763">Boolean</span><span class="sxs-lookup"><span data-stu-id="542eb-763">Boolean</span></span>|<span data-ttu-id="542eb-764">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="542eb-764">&lt;optional&gt;</span></span>|<span data-ttu-id="542eb-765">Значение `true` указывает на то, что вложение будет встроено в текст сообщения и не должно отображаться в списке вложений.</span><span class="sxs-lookup"><span data-stu-id="542eb-765">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="542eb-766">function</span><span class="sxs-lookup"><span data-stu-id="542eb-766">function</span></span>|<span data-ttu-id="542eb-767">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="542eb-767">&lt;optional&gt;</span></span>|<span data-ttu-id="542eb-768">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="542eb-768">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="542eb-769">После успешного выполнения идентификатор вложения будет представлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="542eb-769">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="542eb-770">Если добавить вложение не удастся, объект `asyncResult` будет содержать объект `Error` с описанием ошибки.</span><span class="sxs-lookup"><span data-stu-id="542eb-770">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="542eb-771">Ошибки</span><span class="sxs-lookup"><span data-stu-id="542eb-771">Errors</span></span>

|<span data-ttu-id="542eb-772">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="542eb-772">Error code</span></span>|<span data-ttu-id="542eb-773">Описание</span><span class="sxs-lookup"><span data-stu-id="542eb-773">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="542eb-774">Вложение превышает максимальный размер.</span><span class="sxs-lookup"><span data-stu-id="542eb-774">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="542eb-775">Расширение вложения не поддерживается.</span><span class="sxs-lookup"><span data-stu-id="542eb-775">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="542eb-776">Сообщение или встреча содержат слишком много вложений.</span><span class="sxs-lookup"><span data-stu-id="542eb-776">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="542eb-777">Требования</span><span class="sxs-lookup"><span data-stu-id="542eb-777">Requirements</span></span>

|<span data-ttu-id="542eb-778">Requirement</span><span class="sxs-lookup"><span data-stu-id="542eb-778">Requirement</span></span>|<span data-ttu-id="542eb-779">Значение</span><span class="sxs-lookup"><span data-stu-id="542eb-779">Value</span></span>|
|---|---|
|[<span data-ttu-id="542eb-780">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="542eb-780">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="542eb-781">Предварительная версия</span><span class="sxs-lookup"><span data-stu-id="542eb-781">Preview</span></span>|
|[<span data-ttu-id="542eb-782">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="542eb-782">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="542eb-783">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="542eb-783">ReadWriteItem</span></span>|
|[<span data-ttu-id="542eb-784">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="542eb-784">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="542eb-785">Создание</span><span class="sxs-lookup"><span data-stu-id="542eb-785">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="542eb-786">Примеры</span><span class="sxs-lookup"><span data-stu-id="542eb-786">Examples</span></span>

```js
Office.context.mailbox.item.addFileAttachmentFromBase64Async(
  base64String,
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

####  <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="542eb-787">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="542eb-787">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="542eb-788">Добавляет обработчик для поддерживаемого события.</span><span class="sxs-lookup"><span data-stu-id="542eb-788">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="542eb-789">Сейчас поддерживаются следующие типы событий: `Office.EventType.AttachmentsChanged`, `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged` и `Office.EventType.RecurrenceChanged`.</span><span class="sxs-lookup"><span data-stu-id="542eb-789">Currently the supported event types are `Office.EventType.AttachmentsChanged`, `Office.EventType.AppointmentTimeChanged`, and `Office.EventType.RecipientsChanged`</span></span>

##### <a name="parameters"></a><span data-ttu-id="542eb-790">Параметры:</span><span class="sxs-lookup"><span data-stu-id="542eb-790">Parameters:</span></span>

| <span data-ttu-id="542eb-791">Имя</span><span class="sxs-lookup"><span data-stu-id="542eb-791">Name</span></span> | <span data-ttu-id="542eb-792">Тип</span><span class="sxs-lookup"><span data-stu-id="542eb-792">Type</span></span> | <span data-ttu-id="542eb-793">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="542eb-793">Attributes</span></span> | <span data-ttu-id="542eb-794">Описание</span><span class="sxs-lookup"><span data-stu-id="542eb-794">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="542eb-795">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="542eb-795">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="542eb-796">Событие, которое должно вызвать обработчик.</span><span class="sxs-lookup"><span data-stu-id="542eb-796">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="542eb-797">Function</span><span class="sxs-lookup"><span data-stu-id="542eb-797">Function</span></span> || <span data-ttu-id="542eb-p138">Функция для обработки события. Функция должна принимать один параметр, представляющий собой объектный литерал. Значение свойства `type` параметра совпадет со значением параметра `eventType`, переданного методу `addHandlerAsync`.</span><span class="sxs-lookup"><span data-stu-id="542eb-p138">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="542eb-801">Объект</span><span class="sxs-lookup"><span data-stu-id="542eb-801">Object</span></span> | <span data-ttu-id="542eb-802">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="542eb-802">&lt;optional&gt;</span></span> | <span data-ttu-id="542eb-803">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="542eb-803">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="542eb-804">Object</span><span class="sxs-lookup"><span data-stu-id="542eb-804">Object</span></span> | <span data-ttu-id="542eb-805">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="542eb-805">&lt;optional&gt;</span></span> | <span data-ttu-id="542eb-806">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="542eb-806">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="542eb-807">функция</span><span class="sxs-lookup"><span data-stu-id="542eb-807">function</span></span>| <span data-ttu-id="542eb-808">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="542eb-808">&lt;optional&gt;</span></span>|<span data-ttu-id="542eb-809">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="542eb-809">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="542eb-810">Требования</span><span class="sxs-lookup"><span data-stu-id="542eb-810">Requirements</span></span>

|<span data-ttu-id="542eb-811">Requirement</span><span class="sxs-lookup"><span data-stu-id="542eb-811">Requirement</span></span>| <span data-ttu-id="542eb-812">Значение</span><span class="sxs-lookup"><span data-stu-id="542eb-812">Value</span></span>|
|---|---|
|[<span data-ttu-id="542eb-813">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="542eb-813">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="542eb-814">1.7</span><span class="sxs-lookup"><span data-stu-id="542eb-814">1.7</span></span> |
|[<span data-ttu-id="542eb-815">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="542eb-815">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="542eb-816">ReadItem</span><span class="sxs-lookup"><span data-stu-id="542eb-816">ReadItem</span></span> |
|[<span data-ttu-id="542eb-817">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="542eb-817">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="542eb-818">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="542eb-818">Compose or read</span></span> |

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="542eb-819">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="542eb-819">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="542eb-820">Добавляет к сообщению элемент Exchange, например сообщение, в виде вложения.</span><span class="sxs-lookup"><span data-stu-id="542eb-820">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="542eb-p139">С помощью метода `addItemAttachmentAsync` можно в элемент формы создания вложить элемент с указанным идентификатором Exchange. Если указать метод обратного вызова, то этот метод вызывается с помощью параметра `asyncResult`, который содержит идентификатор вложения или код, указывающий на ошибки, которые произошли при вложении элемента. При необходимости можно использовать параметр `options` для передачи сведений о состоянии методу обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="542eb-p139">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="542eb-824">Идентификатор можно использовать с методом [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback), чтобы удалить вложение, добавленное во время текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="542eb-824">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="542eb-825">Если ваша надстройка Office выполняется в Outlook Web App, метод `addItemAttachmentAsync` обеспечивает вложение элементов в элементы, отличные от редактируемого. Однако это действие не рекомендуем выполнять, так как оно не поддерживается.</span><span class="sxs-lookup"><span data-stu-id="542eb-825">If your Office Add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="542eb-826">Параметры:</span><span class="sxs-lookup"><span data-stu-id="542eb-826">Parameters:</span></span>

|<span data-ttu-id="542eb-827">Имя</span><span class="sxs-lookup"><span data-stu-id="542eb-827">Name</span></span>|<span data-ttu-id="542eb-828">Тип</span><span class="sxs-lookup"><span data-stu-id="542eb-828">Type</span></span>|<span data-ttu-id="542eb-829">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="542eb-829">Attributes</span></span>|<span data-ttu-id="542eb-830">Описание</span><span class="sxs-lookup"><span data-stu-id="542eb-830">Description</span></span>|
|---|---|---|---|
|`itemId`|<span data-ttu-id="542eb-831">String</span><span class="sxs-lookup"><span data-stu-id="542eb-831">String</span></span>||<span data-ttu-id="542eb-p140">Идентификатор Exchange для вкладываемого элемента. Максимальная длина — 100 символов.</span><span class="sxs-lookup"><span data-stu-id="542eb-p140">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="542eb-834">String</span><span class="sxs-lookup"><span data-stu-id="542eb-834">String</span></span>||<span data-ttu-id="542eb-p141">Тема вкладываемого элемента. Максимальная длина — 255 символов.</span><span class="sxs-lookup"><span data-stu-id="542eb-p141">The sujbect of the item to be attached. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="542eb-837">Object</span><span class="sxs-lookup"><span data-stu-id="542eb-837">Object</span></span>|<span data-ttu-id="542eb-838">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="542eb-838">&lt;optional&gt;</span></span>|<span data-ttu-id="542eb-839">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="542eb-839">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="542eb-840">Object</span><span class="sxs-lookup"><span data-stu-id="542eb-840">Object</span></span>|<span data-ttu-id="542eb-841">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="542eb-841">&lt;optional&gt;</span></span>|<span data-ttu-id="542eb-842">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="542eb-842">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="542eb-843">функция</span><span class="sxs-lookup"><span data-stu-id="542eb-843">function</span></span>|<span data-ttu-id="542eb-844">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="542eb-844">&lt;optional&gt;</span></span>|<span data-ttu-id="542eb-845">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="542eb-845">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="542eb-846">После успешного выполнения идентификатор вложения будет представлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="542eb-846">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="542eb-847">Если добавить вложение не удастся, объект `asyncResult` будет содержать объект `Error` с описанием ошибки.</span><span class="sxs-lookup"><span data-stu-id="542eb-847">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="542eb-848">Ошибки</span><span class="sxs-lookup"><span data-stu-id="542eb-848">Errors</span></span>

|<span data-ttu-id="542eb-849">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="542eb-849">Error code</span></span>|<span data-ttu-id="542eb-850">Описание</span><span class="sxs-lookup"><span data-stu-id="542eb-850">Description</span></span>|
|------------|-------------|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="542eb-851">Сообщение или встреча содержат слишком много вложений.</span><span class="sxs-lookup"><span data-stu-id="542eb-851">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="542eb-852">Требования</span><span class="sxs-lookup"><span data-stu-id="542eb-852">Requirements</span></span>

|<span data-ttu-id="542eb-853">Requirement</span><span class="sxs-lookup"><span data-stu-id="542eb-853">Requirement</span></span>|<span data-ttu-id="542eb-854">Значение</span><span class="sxs-lookup"><span data-stu-id="542eb-854">Value</span></span>|
|---|---|
|[<span data-ttu-id="542eb-855">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="542eb-855">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="542eb-856">1.1</span><span class="sxs-lookup"><span data-stu-id="542eb-856">1.1</span></span>|
|[<span data-ttu-id="542eb-857">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="542eb-857">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="542eb-858">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="542eb-858">ReadWriteItem</span></span>|
|[<span data-ttu-id="542eb-859">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="542eb-859">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="542eb-860">Создание</span><span class="sxs-lookup"><span data-stu-id="542eb-860">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="542eb-861">Пример</span><span class="sxs-lookup"><span data-stu-id="542eb-861">Example</span></span>

<span data-ttu-id="542eb-862">В следующем примере существующий элемент Outlook добавляется в виде вложения с именем `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="542eb-862">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

```javascript
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

####  <a name="close"></a><span data-ttu-id="542eb-863">close()</span><span class="sxs-lookup"><span data-stu-id="542eb-863">close()</span></span>

<span data-ttu-id="542eb-864">Закрывает текущий создаваемый элемент.</span><span class="sxs-lookup"><span data-stu-id="542eb-864">Closes the current item that is being composed.</span></span>

<span data-ttu-id="542eb-p142">Работа метода `close` зависит от текущего состояния создаваемого элемента. Если элемент содержит несохраненные изменения, клиент предложит пользователю сохранить или отклонить их либо отменить действие закрытия.</span><span class="sxs-lookup"><span data-stu-id="542eb-p142">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="542eb-867">Если в Outlook в Интернете элемент представляет собой встречу, ранее сохраненную с помощью метода `saveAsync`, пользователю предлагается сохранить, отклонить или отменить действие, даже если с момента последнего сохранения элемента не вносились какие-либо изменения.</span><span class="sxs-lookup"><span data-stu-id="542eb-867">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="542eb-868">Если в клиенте Outlook для настольных ПК сообщение представляет собой ответ в тексте, метод `close` не работает.</span><span class="sxs-lookup"><span data-stu-id="542eb-868">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="542eb-869">Требования</span><span class="sxs-lookup"><span data-stu-id="542eb-869">Requirements</span></span>

|<span data-ttu-id="542eb-870">Requirement</span><span class="sxs-lookup"><span data-stu-id="542eb-870">Requirement</span></span>|<span data-ttu-id="542eb-871">Значение</span><span class="sxs-lookup"><span data-stu-id="542eb-871">Value</span></span>|
|---|---|
|[<span data-ttu-id="542eb-872">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="542eb-872">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="542eb-873">1.3</span><span class="sxs-lookup"><span data-stu-id="542eb-873">1.3</span></span>|
|[<span data-ttu-id="542eb-874">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="542eb-874">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="542eb-875">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="542eb-875">Restricted</span></span>|
|[<span data-ttu-id="542eb-876">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="542eb-876">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="542eb-877">Создание</span><span class="sxs-lookup"><span data-stu-id="542eb-877">Compose</span></span>|

#### <a name="displayreplyallformformdata"></a><span data-ttu-id="542eb-878">displayReplyAllForm(formData)</span><span class="sxs-lookup"><span data-stu-id="542eb-878">displayReplyAllForm(formData)</span></span>

<span data-ttu-id="542eb-879">Отображает форму ответа, включающую отправителя и всех получателей выбранного сообщения или организатора и всех участников выбранной встречи.</span><span class="sxs-lookup"><span data-stu-id="542eb-879">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="542eb-880">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="542eb-880">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="542eb-881">В Outlook Web App форма ответа отображается в виде всплывающей формы в представлении с 3 либо 1 или 2 колонками.</span><span class="sxs-lookup"><span data-stu-id="542eb-881">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="542eb-882">Если любой строковый параметр превышает указанные для него ограничения, `displayReplyAllForm` возвращает исключение.</span><span class="sxs-lookup"><span data-stu-id="542eb-882">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="542eb-p143">Если в параметре `formData.attachments` указаны вложения, Outlook и Outlook Web App пытаются скачать их и вложить в форму ответа. Если какие-либо вложения добавить не удается, в форме отображается сообщение об ошибке. Если сообщения об ошибках не предусмотрены, то они не отображаются.</span><span class="sxs-lookup"><span data-stu-id="542eb-p143">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="542eb-886">Параметры</span><span class="sxs-lookup"><span data-stu-id="542eb-886">Parameters:</span></span>

|<span data-ttu-id="542eb-887">Имя</span><span class="sxs-lookup"><span data-stu-id="542eb-887">Name</span></span>|<span data-ttu-id="542eb-888">Тип</span><span class="sxs-lookup"><span data-stu-id="542eb-888">Type</span></span>|<span data-ttu-id="542eb-889">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="542eb-889">Attributes</span></span>|<span data-ttu-id="542eb-890">Описание</span><span class="sxs-lookup"><span data-stu-id="542eb-890">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="542eb-891">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="542eb-891">String &#124; Object</span></span>||<span data-ttu-id="542eb-p144">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="542eb-p144">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="542eb-894">**ИЛИ**</span><span class="sxs-lookup"><span data-stu-id="542eb-894">**OR**</span></span><br/><span data-ttu-id="542eb-p145">Объект, который содержит текст или данные вложения и функцию обратного вызова. Ниже представлено определение этого объекта.</span><span class="sxs-lookup"><span data-stu-id="542eb-p145">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="542eb-897">String</span><span class="sxs-lookup"><span data-stu-id="542eb-897">String</span></span>|<span data-ttu-id="542eb-898">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="542eb-898">&lt;optional&gt;</span></span>|<span data-ttu-id="542eb-p146">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="542eb-p146">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="542eb-901">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="542eb-901">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="542eb-902">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="542eb-902">&lt;optional&gt;</span></span>|<span data-ttu-id="542eb-903">Массив объектов JSON, представляющих собой вложенные файлы или элементы.</span><span class="sxs-lookup"><span data-stu-id="542eb-903">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="542eb-904">String</span><span class="sxs-lookup"><span data-stu-id="542eb-904">String</span></span>||<span data-ttu-id="542eb-p147">Указывает тип вложения. Допустимые значения: `file` для вложенного файла и `item` для вложенного элемента.</span><span class="sxs-lookup"><span data-stu-id="542eb-p147">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="542eb-907">String</span><span class="sxs-lookup"><span data-stu-id="542eb-907">String</span></span>||<span data-ttu-id="542eb-908">Строка, содержащая имя вложения, длиной до 255 символов.</span><span class="sxs-lookup"><span data-stu-id="542eb-908">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="542eb-909">String</span><span class="sxs-lookup"><span data-stu-id="542eb-909">String</span></span>||<span data-ttu-id="542eb-p148">Используется, только если свойству `type` задано значение `file`. URI расположения файла.</span><span class="sxs-lookup"><span data-stu-id="542eb-p148">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="542eb-912">Логический</span><span class="sxs-lookup"><span data-stu-id="542eb-912">Boolean</span></span>||<span data-ttu-id="542eb-p149">Используется, только если свойству `type` задано значение `file`. Значение `true` указывает на то, что вложение будет встроено в текст сообщения и не должно отображаться в списке вложений.</span><span class="sxs-lookup"><span data-stu-id="542eb-p149">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="542eb-915">String</span><span class="sxs-lookup"><span data-stu-id="542eb-915">String</span></span>||<span data-ttu-id="542eb-p150">Используется, только если свойству `type` задано значение `item`. Идентификатор вложения EWS. Это строка длиной до 100 символов.</span><span class="sxs-lookup"><span data-stu-id="542eb-p150">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="542eb-919">function</span><span class="sxs-lookup"><span data-stu-id="542eb-919">function</span></span>|<span data-ttu-id="542eb-920">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="542eb-920">&lt;optional&gt;</span></span>|<span data-ttu-id="542eb-921">По завершении работы метода функция, переданная параметру `callback`, вызывается с помощью одного параметра `asyncResult`, представляющего собой объект [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="542eb-921">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="542eb-922">Требования</span><span class="sxs-lookup"><span data-stu-id="542eb-922">Requirements</span></span>

|<span data-ttu-id="542eb-923">Requirement</span><span class="sxs-lookup"><span data-stu-id="542eb-923">Requirement</span></span>|<span data-ttu-id="542eb-924">Значение</span><span class="sxs-lookup"><span data-stu-id="542eb-924">Value</span></span>|
|---|---|
|[<span data-ttu-id="542eb-925">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="542eb-925">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="542eb-926">1.0</span><span class="sxs-lookup"><span data-stu-id="542eb-926">1.0</span></span>|
|[<span data-ttu-id="542eb-927">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="542eb-927">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="542eb-928">ReadItem</span><span class="sxs-lookup"><span data-stu-id="542eb-928">ReadItem</span></span>|
|[<span data-ttu-id="542eb-929">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="542eb-929">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="542eb-930">Чтение</span><span class="sxs-lookup"><span data-stu-id="542eb-930">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="542eb-931">Примеры</span><span class="sxs-lookup"><span data-stu-id="542eb-931">Examples</span></span>

<span data-ttu-id="542eb-932">Приведенный ниже код передает строку в функцию `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="542eb-932">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="542eb-933">Ответ с пустым текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="542eb-933">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="542eb-934">Ответ только с текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="542eb-934">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="542eb-935">Ответ с текстом сообщения и вложенным файлом.</span><span class="sxs-lookup"><span data-stu-id="542eb-935">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="542eb-936">Ответ с текстом сообщения и вложенным элементом.</span><span class="sxs-lookup"><span data-stu-id="542eb-936">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="542eb-937">Ответ с текстом сообщения, вложенным файлом, вложенным элементом и обратным вызовом.</span><span class="sxs-lookup"><span data-stu-id="542eb-937">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata"></a><span data-ttu-id="542eb-938">displayReplyForm(formData)</span><span class="sxs-lookup"><span data-stu-id="542eb-938">displayReplyForm(formData)</span></span>

<span data-ttu-id="542eb-939">Отображает форму ответа, включающую только отправителя выбранного сообщения или организатора выбранной встречи.</span><span class="sxs-lookup"><span data-stu-id="542eb-939">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="542eb-940">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="542eb-940">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="542eb-941">В Outlook Web App форма ответа отображается в виде всплывающей формы в представлении с 3 либо 1 или 2 колонками.</span><span class="sxs-lookup"><span data-stu-id="542eb-941">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="542eb-942">Если любой строковый параметр превышает указанные для него ограничения, `displayReplyForm` возвращает исключение.</span><span class="sxs-lookup"><span data-stu-id="542eb-942">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="542eb-p151">Если в параметре `formData.attachments` указаны вложения, Outlook и Outlook Web App пытаются скачать их и вложить в форму ответа. Если какие-либо вложения добавить не удается, в форме отображается сообщение об ошибке. Если сообщения об ошибках не предусмотрены, то они не отображаются.</span><span class="sxs-lookup"><span data-stu-id="542eb-p151">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="542eb-946">Параметры</span><span class="sxs-lookup"><span data-stu-id="542eb-946">Parameters:</span></span>

|<span data-ttu-id="542eb-947">Имя</span><span class="sxs-lookup"><span data-stu-id="542eb-947">Name</span></span>|<span data-ttu-id="542eb-948">Тип</span><span class="sxs-lookup"><span data-stu-id="542eb-948">Type</span></span>|<span data-ttu-id="542eb-949">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="542eb-949">Attributes</span></span>|<span data-ttu-id="542eb-950">Описание</span><span class="sxs-lookup"><span data-stu-id="542eb-950">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="542eb-951">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="542eb-951">String &#124; Object</span></span>||<span data-ttu-id="542eb-p152">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="542eb-p152">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="542eb-954">**ИЛИ**</span><span class="sxs-lookup"><span data-stu-id="542eb-954">**OR**</span></span><br/><span data-ttu-id="542eb-p153">Объект, который содержит текст или данные вложения и функцию обратного вызова. Ниже представлено определение этого объекта.</span><span class="sxs-lookup"><span data-stu-id="542eb-p153">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="542eb-957">String</span><span class="sxs-lookup"><span data-stu-id="542eb-957">String</span></span>|<span data-ttu-id="542eb-958">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="542eb-958">&lt;optional&gt;</span></span>|<span data-ttu-id="542eb-p154">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="542eb-p154">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="542eb-961">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="542eb-961">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="542eb-962">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="542eb-962">&lt;optional&gt;</span></span>|<span data-ttu-id="542eb-963">Массив объектов JSON, представляющих собой вложенные файлы или элементы.</span><span class="sxs-lookup"><span data-stu-id="542eb-963">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="542eb-964">String</span><span class="sxs-lookup"><span data-stu-id="542eb-964">String</span></span>||<span data-ttu-id="542eb-p155">Указывает тип вложения. Допустимые значения: `file` для вложенного файла и `item` для вложенного элемента.</span><span class="sxs-lookup"><span data-stu-id="542eb-p155">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="542eb-967">String</span><span class="sxs-lookup"><span data-stu-id="542eb-967">String</span></span>||<span data-ttu-id="542eb-968">Строка, содержащая имя вложения, длиной до 255 символов.</span><span class="sxs-lookup"><span data-stu-id="542eb-968">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="542eb-969">String</span><span class="sxs-lookup"><span data-stu-id="542eb-969">String</span></span>||<span data-ttu-id="542eb-p156">Используется, только если свойству `type` задано значение `file`. URI расположения файла.</span><span class="sxs-lookup"><span data-stu-id="542eb-p156">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="542eb-972">Логический</span><span class="sxs-lookup"><span data-stu-id="542eb-972">Boolean</span></span>||<span data-ttu-id="542eb-p157">Используется, только если свойству `type` задано значение `file`. Значение `true` указывает на то, что вложение будет встроено в текст сообщения и не должно отображаться в списке вложений.</span><span class="sxs-lookup"><span data-stu-id="542eb-p157">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="542eb-975">String</span><span class="sxs-lookup"><span data-stu-id="542eb-975">String</span></span>||<span data-ttu-id="542eb-p158">Используется, только если свойству `type` задано значение `item`. Идентификатор вложения EWS. Это строка длиной до 100 символов.</span><span class="sxs-lookup"><span data-stu-id="542eb-p158">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="542eb-979">function</span><span class="sxs-lookup"><span data-stu-id="542eb-979">function</span></span>|<span data-ttu-id="542eb-980">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="542eb-980">&lt;optional&gt;</span></span>|<span data-ttu-id="542eb-981">По завершении работы метода функция, переданная параметру `callback`, вызывается с помощью одного параметра `asyncResult`, представляющего собой объект [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="542eb-981">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="542eb-982">Требования</span><span class="sxs-lookup"><span data-stu-id="542eb-982">Requirements</span></span>

|<span data-ttu-id="542eb-983">Requirement</span><span class="sxs-lookup"><span data-stu-id="542eb-983">Requirement</span></span>|<span data-ttu-id="542eb-984">Значение</span><span class="sxs-lookup"><span data-stu-id="542eb-984">Value</span></span>|
|---|---|
|[<span data-ttu-id="542eb-985">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="542eb-985">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="542eb-986">1.0</span><span class="sxs-lookup"><span data-stu-id="542eb-986">1.0</span></span>|
|[<span data-ttu-id="542eb-987">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="542eb-987">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="542eb-988">ReadItem</span><span class="sxs-lookup"><span data-stu-id="542eb-988">ReadItem</span></span>|
|[<span data-ttu-id="542eb-989">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="542eb-989">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="542eb-990">Чтение</span><span class="sxs-lookup"><span data-stu-id="542eb-990">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="542eb-991">Примеры</span><span class="sxs-lookup"><span data-stu-id="542eb-991">Examples</span></span>

<span data-ttu-id="542eb-992">Приведенный ниже код передает строку в функцию `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="542eb-992">The following code passes a string to the `displayReplyForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="542eb-993">Ответ с пустым текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="542eb-993">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="542eb-994">Ответ только с текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="542eb-994">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="542eb-995">Ответ с текстом сообщения и вложенным файлом.</span><span class="sxs-lookup"><span data-stu-id="542eb-995">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="542eb-996">Ответ с текстом сообщения и вложенным элементом.</span><span class="sxs-lookup"><span data-stu-id="542eb-996">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="542eb-997">Ответ с текстом сообщения, вложенным файлом, вложенным элементом и обратным вызовом.</span><span class="sxs-lookup"><span data-stu-id="542eb-997">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getattachmentcontentasyncattachmentid-options-callback--attachmentcontentjavascriptapioutlookofficeattachmentcontent"></a><span data-ttu-id="542eb-998">getAttachmentContentAsync(attachmentId, [options], callback) → [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span><span class="sxs-lookup"><span data-stu-id="542eb-998">getAttachmentContentAsync(attachmentId, [options], callback) → [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span></span>

<span data-ttu-id="542eb-999">Получает указанное вложение из сообщения или встречи и возвращает в качестве объекта `AttachmentContent`.</span><span class="sxs-lookup"><span data-stu-id="542eb-999">Gets the specified attachment from a message or appointment and returns it as an `AttachmentContent` object.</span></span>

<span data-ttu-id="542eb-1000">Метод `getAttachmentContentAsync` получает вложение с указанным идентификатором из элемента.</span><span class="sxs-lookup"><span data-stu-id="542eb-1000">The `getAttachmentContentAsync` method removes the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="542eb-1001">Рекомендуется использовать идентификатор для получения вложения в том же сеансе, в котором были получены идентификаторы вложений attachmentIds посредством вызова `getAttachmentsAsync` или `item.attachments`.</span><span class="sxs-lookup"><span data-stu-id="542eb-1001">As a best practice, you should use the identifier to retrieve an attachment in the same session that the attachmentIds were retrieved with the `getAttachmentsAsync` or `item.attachments` call.</span></span> <span data-ttu-id="542eb-1002">В Outlook Web App и Outlook Web App для устройств идентификатор вложения действителен только в рамках одного сеанса.</span><span class="sxs-lookup"><span data-stu-id="542eb-1002">In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="542eb-1003">Сеанс завершается, когда пользователь закрывает приложение или начинает создавать элемент во встроенной форме, а затем переходит из формы в отдельное окно.</span><span class="sxs-lookup"><span data-stu-id="542eb-1003">A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="542eb-1004">Параметры:</span><span class="sxs-lookup"><span data-stu-id="542eb-1004">Parameters:</span></span>

|<span data-ttu-id="542eb-1005">Имя</span><span class="sxs-lookup"><span data-stu-id="542eb-1005">Name</span></span>|<span data-ttu-id="542eb-1006">Тип</span><span class="sxs-lookup"><span data-stu-id="542eb-1006">Type</span></span>|<span data-ttu-id="542eb-1007">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="542eb-1007">Attributes</span></span>|<span data-ttu-id="542eb-1008">Описание</span><span class="sxs-lookup"><span data-stu-id="542eb-1008">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="542eb-1009">Строка</span><span class="sxs-lookup"><span data-stu-id="542eb-1009">String</span></span>||<span data-ttu-id="542eb-1010">Идентификатор вложения, который необходимо получить.</span><span class="sxs-lookup"><span data-stu-id="542eb-1010">The identifier of the attachment you want to get.</span></span> <span data-ttu-id="542eb-1011">Максимальная длина строки — 100 символов.</span><span class="sxs-lookup"><span data-stu-id="542eb-1011">The maximum length of the string is 100 characters.</span></span>|
|`options`|<span data-ttu-id="542eb-1012">Object</span><span class="sxs-lookup"><span data-stu-id="542eb-1012">Object</span></span>|<span data-ttu-id="542eb-1013">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="542eb-1013">&lt;optional&gt;</span></span>|<span data-ttu-id="542eb-1014">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="542eb-1014">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="542eb-1015">Object</span><span class="sxs-lookup"><span data-stu-id="542eb-1015">Object</span></span>|<span data-ttu-id="542eb-1016">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="542eb-1016">&lt;optional&gt;</span></span>|<span data-ttu-id="542eb-1017">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="542eb-1017">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="542eb-1018">функция</span><span class="sxs-lookup"><span data-stu-id="542eb-1018">function</span></span>|<span data-ttu-id="542eb-1019">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="542eb-1019">&lt;optional&gt;</span></span>|<span data-ttu-id="542eb-1020">По завершении работы метода функция, переданная параметру `callback`, вызывается с помощью одного параметра `asyncResult`, представляющего собой объект [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="542eb-1020">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="542eb-1021">Требования</span><span class="sxs-lookup"><span data-stu-id="542eb-1021">Requirements</span></span>

|<span data-ttu-id="542eb-1022">Requirement</span><span class="sxs-lookup"><span data-stu-id="542eb-1022">Requirement</span></span>|<span data-ttu-id="542eb-1023">Значение</span><span class="sxs-lookup"><span data-stu-id="542eb-1023">Value</span></span>|
|---|---|
|[<span data-ttu-id="542eb-1024">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="542eb-1024">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="542eb-1025">Предварительная версия</span><span class="sxs-lookup"><span data-stu-id="542eb-1025">Preview</span></span>|
|[<span data-ttu-id="542eb-1026">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="542eb-1026">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="542eb-1027">ReadItem</span><span class="sxs-lookup"><span data-stu-id="542eb-1027">ReadItem</span></span>|
|[<span data-ttu-id="542eb-1028">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="542eb-1028">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="542eb-1029">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="542eb-1029">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="542eb-1030">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="542eb-1030">Returns:</span></span>

<span data-ttu-id="542eb-1031">Тип: [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span><span class="sxs-lookup"><span data-stu-id="542eb-1031">Type: [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span></span>

##### <a name="example"></a><span data-ttu-id="542eb-1032">Пример</span><span class="sxs-lookup"><span data-stu-id="542eb-1032">Example</span></span>

```javascript
var item = Office.context.mailbox.item;
var listOfAttachments = [];
item.getAttachmentsAsync(callback);
function callback(result) {
    if (result.value.length > 0) {
        for (i = 0 ; i < result.value.length ; i++) {
            var options = {asyncContext: {type: result.value[i].attachmentType}};
            getAttachmentContentAsync(result.value[i].id, options, handleAttachmentsCallback);  
        }
    }
}

function handleAttachmentsCallback(result) {
    // parse string to be a url, an .eml file, a base64-encoded string, or an .icalendar file
    if (result.format == Office.MailboxEnums.AttachmentContentFormat.Base64) {
        // handle file attachment
    }
    else if (result.format == Office.MailboxEnums.AttachmentContentFormat.Eml) {
        // handle item attachment
    }
    else if (result.format == Office.MailboxEnums.AttachmentContentFormat.ICalendar) {
        // handle .icalender attachment
    }
    else {
        // handle cloud attachment  
    }
}
```

#### <a name="getattachmentsasyncoptions-callback--arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails"></a><span data-ttu-id="542eb-1033">getAttachmentsAsync([options], callback) → Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="542eb-1033">getAttachmentsAsync([options], callback) → Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

<span data-ttu-id="542eb-1034">Получает вложения элемента в качестве массива.</span><span class="sxs-lookup"><span data-stu-id="542eb-1034">Gets the item's attachments as an array.</span></span> <span data-ttu-id="542eb-1035">Только в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="542eb-1035">Compose mode only.</span></span>

##### <a name="parameters"></a><span data-ttu-id="542eb-1036">Параметры:</span><span class="sxs-lookup"><span data-stu-id="542eb-1036">Parameters:</span></span>

|<span data-ttu-id="542eb-1037">Имя</span><span class="sxs-lookup"><span data-stu-id="542eb-1037">Name</span></span>|<span data-ttu-id="542eb-1038">Тип</span><span class="sxs-lookup"><span data-stu-id="542eb-1038">Type</span></span>|<span data-ttu-id="542eb-1039">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="542eb-1039">Attributes</span></span>|<span data-ttu-id="542eb-1040">Описание</span><span class="sxs-lookup"><span data-stu-id="542eb-1040">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="542eb-1041">Object</span><span class="sxs-lookup"><span data-stu-id="542eb-1041">Object</span></span>|<span data-ttu-id="542eb-1042">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="542eb-1042">&lt;optional&gt;</span></span>|<span data-ttu-id="542eb-1043">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="542eb-1043">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="542eb-1044">Object</span><span class="sxs-lookup"><span data-stu-id="542eb-1044">Object</span></span>|<span data-ttu-id="542eb-1045">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="542eb-1045">&lt;optional&gt;</span></span>|<span data-ttu-id="542eb-1046">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="542eb-1046">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="542eb-1047">функция</span><span class="sxs-lookup"><span data-stu-id="542eb-1047">function</span></span>|<span data-ttu-id="542eb-1048">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="542eb-1048">&lt;optional&gt;</span></span>|<span data-ttu-id="542eb-1049">По завершении работы метода функция, переданная параметру `callback`, вызывается с помощью одного параметра `asyncResult`, представляющего собой объект [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="542eb-1049">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="542eb-1050">Требования</span><span class="sxs-lookup"><span data-stu-id="542eb-1050">Requirements</span></span>

|<span data-ttu-id="542eb-1051">Requirement</span><span class="sxs-lookup"><span data-stu-id="542eb-1051">Requirement</span></span>|<span data-ttu-id="542eb-1052">Значение</span><span class="sxs-lookup"><span data-stu-id="542eb-1052">Value</span></span>|
|---|---|
|[<span data-ttu-id="542eb-1053">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="542eb-1053">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="542eb-1054">Предварительная версия</span><span class="sxs-lookup"><span data-stu-id="542eb-1054">Preview</span></span>|
|[<span data-ttu-id="542eb-1055">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="542eb-1055">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="542eb-1056">ReadItem</span><span class="sxs-lookup"><span data-stu-id="542eb-1056">ReadItem</span></span>|
|[<span data-ttu-id="542eb-1057">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="542eb-1057">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="542eb-1058">Создание</span><span class="sxs-lookup"><span data-stu-id="542eb-1058">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="542eb-1059">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="542eb-1059">Returns:</span></span>

<span data-ttu-id="542eb-1060">Тип: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="542eb-1060">Type: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

##### <a name="example"></a><span data-ttu-id="542eb-1061">Пример</span><span class="sxs-lookup"><span data-stu-id="542eb-1061">Example</span></span>

<span data-ttu-id="542eb-1062">В приведенном ниже примере создается HTML-строка с подробными сведениями обо всех вложениях для текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="542eb-1062">The following code builds an HTML string with details of all attachments on the current item.</span></span>

```javascript
var item = Office.context.mailbox.item;
var outputString = "";
item.getAttachmentsAsync(callback);  
function callback(result) {
    if (result.value.length > 0) {
        for (i = 0 ; i < result.value.length ; i++) {
            var _att = result.value [i];
            outputString += "<BR>" + i + ". Name: ";
            outputString += _att.name;
            outputString += "<BR>ID: " + _att.id;
            outputString += "<BR>contentType: " + _att.contentType;
            outputString += "<BR>size: " + _att.size;
            outputString += "<BR>attachmentType: " + _att.attachmentType;
            outputString += "<BR>isInline: " + _att.isInline;
        }
    }
}
```

#### <a name="getentities--entitiesjavascriptapioutlookofficeentities"></a><span data-ttu-id="542eb-1063">getEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="542eb-1063">getEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span></span>

<span data-ttu-id="542eb-1064">Получает сущности, обнаруженные в теле выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="542eb-1064">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="542eb-1065">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="542eb-1065">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="542eb-1066">Требования</span><span class="sxs-lookup"><span data-stu-id="542eb-1066">Requirements</span></span>

|<span data-ttu-id="542eb-1067">Requirement</span><span class="sxs-lookup"><span data-stu-id="542eb-1067">Requirement</span></span>|<span data-ttu-id="542eb-1068">Значение</span><span class="sxs-lookup"><span data-stu-id="542eb-1068">Value</span></span>|
|---|---|
|[<span data-ttu-id="542eb-1069">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="542eb-1069">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="542eb-1070">1.0</span><span class="sxs-lookup"><span data-stu-id="542eb-1070">1.0</span></span>|
|[<span data-ttu-id="542eb-1071">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="542eb-1071">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="542eb-1072">ReadItem</span><span class="sxs-lookup"><span data-stu-id="542eb-1072">ReadItem</span></span>|
|[<span data-ttu-id="542eb-1073">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="542eb-1073">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="542eb-1074">Чтение</span><span class="sxs-lookup"><span data-stu-id="542eb-1074">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="542eb-1075">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="542eb-1075">Returns:</span></span>

<span data-ttu-id="542eb-1076">Тип: [Entities](/javascript/api/outlook/office.entities)</span><span class="sxs-lookup"><span data-stu-id="542eb-1076">Type: [Entities](/javascript/api/outlook/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="542eb-1077">Пример</span><span class="sxs-lookup"><span data-stu-id="542eb-1077">Example</span></span>

<span data-ttu-id="542eb-1078">Ниже приведен пример получения доступа к сущностям контактов в теле текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="542eb-1078">The following example accesses the contacts entities in the current item's body.</span></span>

```javascript
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion"></a><span data-ttu-id="542eb-1079">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="542eb-1079">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span></span>

<span data-ttu-id="542eb-1080">Получает массив всех сущностей указанного типа, обнаруженных в теле выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="542eb-1080">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="542eb-1081">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="542eb-1081">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="542eb-1082">Параметры</span><span class="sxs-lookup"><span data-stu-id="542eb-1082">Parameters:</span></span>

|<span data-ttu-id="542eb-1083">Имя</span><span class="sxs-lookup"><span data-stu-id="542eb-1083">Name</span></span>|<span data-ttu-id="542eb-1084">Тип</span><span class="sxs-lookup"><span data-stu-id="542eb-1084">Type</span></span>|<span data-ttu-id="542eb-1085">Описание</span><span class="sxs-lookup"><span data-stu-id="542eb-1085">Description</span></span>|
|---|---|---|
|`entityType`|[<span data-ttu-id="542eb-1086">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="542eb-1086">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook/office.mailboxenums.entitytype)|<span data-ttu-id="542eb-1087">Одно из значений перечисления EntityType.</span><span class="sxs-lookup"><span data-stu-id="542eb-1087">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="542eb-1088">Требования</span><span class="sxs-lookup"><span data-stu-id="542eb-1088">Requirements</span></span>

|<span data-ttu-id="542eb-1089">Requirement</span><span class="sxs-lookup"><span data-stu-id="542eb-1089">Requirement</span></span>|<span data-ttu-id="542eb-1090">Значение</span><span class="sxs-lookup"><span data-stu-id="542eb-1090">Value</span></span>|
|---|---|
|[<span data-ttu-id="542eb-1091">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="542eb-1091">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="542eb-1092">1.0</span><span class="sxs-lookup"><span data-stu-id="542eb-1092">1.0</span></span>|
|[<span data-ttu-id="542eb-1093">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="542eb-1093">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="542eb-1094">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="542eb-1094">Restricted</span></span>|
|[<span data-ttu-id="542eb-1095">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="542eb-1095">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="542eb-1096">Чтение</span><span class="sxs-lookup"><span data-stu-id="542eb-1096">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="542eb-1097">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="542eb-1097">Returns:</span></span>

<span data-ttu-id="542eb-1098">Если значение, переданное в `entityType`, не является допустимым членом перечисления `EntityType`, метод возвращает значение NULL.</span><span class="sxs-lookup"><span data-stu-id="542eb-1098">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="542eb-1099">Если в теле элемента отсутствуют сущности указанного типа, метод возвращает пустой массив.</span><span class="sxs-lookup"><span data-stu-id="542eb-1099">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="542eb-1100">В противном случае тип объектов в возвращаемом массиве зависит от типа сущности, запрошенной в параметре `entityType`.</span><span class="sxs-lookup"><span data-stu-id="542eb-1100">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="542eb-1101">Хотя минимальный уровень разрешений для использования этого метода — **Restricted**, для некоторых типов сущностей требуется доступ на уровне **ReadItem**, как указано в приведенной ниже таблице.</span><span class="sxs-lookup"><span data-stu-id="542eb-1101">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

|<span data-ttu-id="542eb-1102">Значение параметра `entityType`</span><span class="sxs-lookup"><span data-stu-id="542eb-1102">Value of `entityType`</span></span>|<span data-ttu-id="542eb-1103">Тип объектов в возвращаемом массиве</span><span class="sxs-lookup"><span data-stu-id="542eb-1103">Type of objects in returned array</span></span>|<span data-ttu-id="542eb-1104">Необходимый уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="542eb-1104">Required Permission Level</span></span>|
|---|---|---|
|`Address`|<span data-ttu-id="542eb-1105">String</span><span class="sxs-lookup"><span data-stu-id="542eb-1105">String</span></span>|<span data-ttu-id="542eb-1106">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="542eb-1106">**Restricted**</span></span>|
|`Contact`|<span data-ttu-id="542eb-1107">Contact</span><span class="sxs-lookup"><span data-stu-id="542eb-1107">Contact</span></span>|<span data-ttu-id="542eb-1108">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="542eb-1108">**ReadItem**</span></span>|
|`EmailAddress`|<span data-ttu-id="542eb-1109">String</span><span class="sxs-lookup"><span data-stu-id="542eb-1109">String</span></span>|<span data-ttu-id="542eb-1110">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="542eb-1110">**ReadItem**</span></span>|
|`MeetingSuggestion`|<span data-ttu-id="542eb-1111">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="542eb-1111">MeetingSuggestion</span></span>|<span data-ttu-id="542eb-1112">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="542eb-1112">**ReadItem**</span></span>|
|`PhoneNumber`|<span data-ttu-id="542eb-1113">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="542eb-1113">PhoneNumber</span></span>|<span data-ttu-id="542eb-1114">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="542eb-1114">**Restricted**</span></span>|
|`TaskSuggestion`|<span data-ttu-id="542eb-1115">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="542eb-1115">TaskSuggestion</span></span>|<span data-ttu-id="542eb-1116">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="542eb-1116">**ReadItem**</span></span>|
|`URL`|<span data-ttu-id="542eb-1117">String</span><span class="sxs-lookup"><span data-stu-id="542eb-1117">String</span></span>|<span data-ttu-id="542eb-1118">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="542eb-1118">**Restricted**</span></span>|

<span data-ttu-id="542eb-1119">Тип: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="542eb-1119">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="542eb-1120">Пример</span><span class="sxs-lookup"><span data-stu-id="542eb-1120">Example</span></span>

<span data-ttu-id="542eb-1121">В примере ниже показано, как получить доступ к массиву строк, которые представляют собой почтовые адреса в теле текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="542eb-1121">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion"></a><span data-ttu-id="542eb-1122">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="542eb-1122">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span></span>

<span data-ttu-id="542eb-1123">Возвращает известные сущности в выбранном элементе, которые проходят через именованный фильтр, определяемый в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="542eb-1123">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="542eb-1124">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="542eb-1124">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="542eb-1125">Метод `getFilteredEntitiesByName` возвращает сущности, соответствующие регулярному выражению, которое определяется в элементе правила [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) в XML-файле манифеста, с использованием указанного значения элемента `FilterName`.</span><span class="sxs-lookup"><span data-stu-id="542eb-1125">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="542eb-1126">Параметры</span><span class="sxs-lookup"><span data-stu-id="542eb-1126">Parameters:</span></span>

|<span data-ttu-id="542eb-1127">Имя</span><span class="sxs-lookup"><span data-stu-id="542eb-1127">Name</span></span>|<span data-ttu-id="542eb-1128">Тип</span><span class="sxs-lookup"><span data-stu-id="542eb-1128">Type</span></span>|<span data-ttu-id="542eb-1129">Описание</span><span class="sxs-lookup"><span data-stu-id="542eb-1129">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="542eb-1130">String</span><span class="sxs-lookup"><span data-stu-id="542eb-1130">String</span></span>|<span data-ttu-id="542eb-1131">Имя элемента правила `ItemHasKnownEntity`, определяющее соответствующий фильтр.</span><span class="sxs-lookup"><span data-stu-id="542eb-1131">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="542eb-1132">Требования</span><span class="sxs-lookup"><span data-stu-id="542eb-1132">Requirements</span></span>

|<span data-ttu-id="542eb-1133">Requirement</span><span class="sxs-lookup"><span data-stu-id="542eb-1133">Requirement</span></span>|<span data-ttu-id="542eb-1134">Значение</span><span class="sxs-lookup"><span data-stu-id="542eb-1134">Value</span></span>|
|---|---|
|[<span data-ttu-id="542eb-1135">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="542eb-1135">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="542eb-1136">1.0</span><span class="sxs-lookup"><span data-stu-id="542eb-1136">1.0</span></span>|
|[<span data-ttu-id="542eb-1137">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="542eb-1137">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="542eb-1138">ReadItem</span><span class="sxs-lookup"><span data-stu-id="542eb-1138">ReadItem</span></span>|
|[<span data-ttu-id="542eb-1139">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="542eb-1139">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="542eb-1140">Чтение</span><span class="sxs-lookup"><span data-stu-id="542eb-1140">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="542eb-1141">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="542eb-1141">Returns:</span></span>

<span data-ttu-id="542eb-p163">Если в манифесте нет элемента `ItemHasKnownEntity` со значением `FilterName`, соответствующим параметру `name`, метод возвращает `null`. Если параметр `name` соответствует элементу `ItemHasKnownEntity` в манифесте, но при этом в текущем элементе нет соответствующих сущностей, метод возвращает пустой массив.</span><span class="sxs-lookup"><span data-stu-id="542eb-p163">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="542eb-1144">Тип: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="542eb-1144">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span></span>

#### <a name="getinitializationcontextasyncoptions-callback"></a><span data-ttu-id="542eb-1145">getInitializationContextAsync([options], [callback])</span><span class="sxs-lookup"><span data-stu-id="542eb-1145">getInitializationContextAsync([options], [callback])</span></span>

<span data-ttu-id="542eb-1146">Получает данные инициализации, передаваемые при [активации надстройки интерактивным сообщением](https://docs.microsoft.com/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span><span class="sxs-lookup"><span data-stu-id="542eb-1146">Gets initialization data passed when the add-in is [activated by an actionable message](https://docs.microsoft.com/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span></span>

> [!NOTE]
> <span data-ttu-id="542eb-1147">Этот метод поддерживается только версией Outlook 2016 для Windows или более поздней (версии "нажми и работай" с номером больше 16.0.8413.1000) и Outlook в Интернете для Office 365.</span><span class="sxs-lookup"><span data-stu-id="542eb-1147">Note: This method is only supported by Outlook 2016 for Windows (Click-to-Run versions greater than 16.0.8413.1000) and Outlook on the web for Office 365.</span></span>

##### <a name="parameters"></a><span data-ttu-id="542eb-1148">Параметры:</span><span class="sxs-lookup"><span data-stu-id="542eb-1148">Parameters:</span></span>
|<span data-ttu-id="542eb-1149">Имя</span><span class="sxs-lookup"><span data-stu-id="542eb-1149">Name</span></span>|<span data-ttu-id="542eb-1150">Тип</span><span class="sxs-lookup"><span data-stu-id="542eb-1150">Type</span></span>|<span data-ttu-id="542eb-1151">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="542eb-1151">Attributes</span></span>|<span data-ttu-id="542eb-1152">Описание</span><span class="sxs-lookup"><span data-stu-id="542eb-1152">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="542eb-1153">Object</span><span class="sxs-lookup"><span data-stu-id="542eb-1153">Object</span></span>|<span data-ttu-id="542eb-1154">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="542eb-1154">&lt;optional&gt;</span></span>|<span data-ttu-id="542eb-1155">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="542eb-1155">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="542eb-1156">Object</span><span class="sxs-lookup"><span data-stu-id="542eb-1156">Object</span></span>|<span data-ttu-id="542eb-1157">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="542eb-1157">&lt;optional&gt;</span></span>|<span data-ttu-id="542eb-1158">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="542eb-1158">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="542eb-1159">функция</span><span class="sxs-lookup"><span data-stu-id="542eb-1159">function</span></span>|<span data-ttu-id="542eb-1160">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="542eb-1160">&lt;optional&gt;</span></span>|<span data-ttu-id="542eb-1161">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="542eb-1161">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="542eb-1162">В случае успешного выполнения данные инициализации предоставляются в свойстве `asyncResult.value` как строка.</span><span class="sxs-lookup"><span data-stu-id="542eb-1162">On success, the intialization data is provided in the `asyncResult.value` property as a string.</span></span><br/><span data-ttu-id="542eb-1163">Если контекст инициализации отсутствует, объект `asyncResult` будет содержать объект `Error`, одному свойству которого (`code`) будет присвоено значение `9020`, а другому (`name`) — значение `GenericResponseError`.</span><span class="sxs-lookup"><span data-stu-id="542eb-1163">If there is no initialization context, the `asyncResult` object will contain an `Error` object with its `code` property set to `9020` and its `name` property set to `GenericResponseError`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="542eb-1164">Требования</span><span class="sxs-lookup"><span data-stu-id="542eb-1164">Requirements</span></span>

|<span data-ttu-id="542eb-1165">Requirement</span><span class="sxs-lookup"><span data-stu-id="542eb-1165">Requirement</span></span>|<span data-ttu-id="542eb-1166">Значение</span><span class="sxs-lookup"><span data-stu-id="542eb-1166">Value</span></span>|
|---|---|
|[<span data-ttu-id="542eb-1167">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="542eb-1167">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="542eb-1168">Предварительная версия</span><span class="sxs-lookup"><span data-stu-id="542eb-1168">Preview</span></span>|
|[<span data-ttu-id="542eb-1169">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="542eb-1169">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="542eb-1170">ReadItem</span><span class="sxs-lookup"><span data-stu-id="542eb-1170">ReadItem</span></span>|
|[<span data-ttu-id="542eb-1171">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="542eb-1171">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="542eb-1172">Чтение</span><span class="sxs-lookup"><span data-stu-id="542eb-1172">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="542eb-1173">Пример</span><span class="sxs-lookup"><span data-stu-id="542eb-1173">Example</span></span>

```javascript
// Get the initialization context (if present)
Office.context.mailbox.item.getInitializationContextAsync(
  function(asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
      if (asyncResult.value != null && asyncResult.value.length > 0) {
        // The value is a string, parse to an object
        var context = JSON.parse(asyncResult.value);
        // Do something with context
      } else {
        // Empty context, treat as no context
      }
    } else {
      if (asyncResult.error.code == 9020) {
        // GenericResponseError returned when there is
        // no context
        // Treat as no context
      } else {
        // Handle the error
      }
    }
  }
);
```

#### <a name="getregexmatches--object"></a><span data-ttu-id="542eb-1174">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="542eb-1174">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="542eb-1175">Возвращает строковые значения в выбранном элементе, которые соответствуют регулярным выражениям, определенным в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="542eb-1175">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="542eb-1176">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="542eb-1176">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="542eb-p164">Метод `getRegExMatches` возвращает строки, соответствующие регулярному выражению, которое определяется в каждом элементе правила `ItemHasRegularExpressionMatch` или `ItemHasKnownEntity` в XML-файле манифеста. Для правила `ItemHasRegularExpressionMatch` соответствующую строку должно содержать свойство элемента, указанного этим правилом. Простой тип `PropertyName` определяет поддерживаемые свойства.</span><span class="sxs-lookup"><span data-stu-id="542eb-p164">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="542eb-1180">Например, рассмотрим манифест надстройки, который содержит указанный ниже элемент `Rule`.</span><span class="sxs-lookup"><span data-stu-id="542eb-1180">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="542eb-1181">Объект, возвращаемый методом `getRegExMatches`, будет содержать два свойства: `fruits` и `veggies`.</span><span class="sxs-lookup"><span data-stu-id="542eb-1181">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="542eb-p165">Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты. Лучше используйте метод [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) для этого.</span><span class="sxs-lookup"><span data-stu-id="542eb-p165">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="542eb-1185">Requirements</span><span class="sxs-lookup"><span data-stu-id="542eb-1185">Requirements</span></span>

|<span data-ttu-id="542eb-1186">Requirement</span><span class="sxs-lookup"><span data-stu-id="542eb-1186">Requirement</span></span>|<span data-ttu-id="542eb-1187">Значение</span><span class="sxs-lookup"><span data-stu-id="542eb-1187">Value</span></span>|
|---|---|
|[<span data-ttu-id="542eb-1188">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="542eb-1188">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="542eb-1189">1.0</span><span class="sxs-lookup"><span data-stu-id="542eb-1189">1.0</span></span>|
|[<span data-ttu-id="542eb-1190">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="542eb-1190">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="542eb-1191">ReadItem</span><span class="sxs-lookup"><span data-stu-id="542eb-1191">ReadItem</span></span>|
|[<span data-ttu-id="542eb-1192">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="542eb-1192">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="542eb-1193">Чтение</span><span class="sxs-lookup"><span data-stu-id="542eb-1193">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="542eb-1194">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="542eb-1194">Returns:</span></span>

<span data-ttu-id="542eb-p166">Объект, содержащий массив строк, которые соответствуют регулярным выражениям, определяемым в XML-файле манифеста. Имя каждого массива равно соответствующему значению атрибута `RegExName` подходящего правила `ItemHasRegularExpressionMatch` или атрибута `FilterName` соответствующего правила `ItemHasKnownEntity`.</span><span class="sxs-lookup"><span data-stu-id="542eb-p166">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="542eb-1197">

<dt>Тип</dt>

</span><span class="sxs-lookup"><span data-stu-id="542eb-1197">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="542eb-1198">Object</span><span class="sxs-lookup"><span data-stu-id="542eb-1198">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="542eb-1199">Пример</span><span class="sxs-lookup"><span data-stu-id="542eb-1199">Example</span></span>

<span data-ttu-id="542eb-1200">В приведенном ниже примере показано, как получить доступ к массиву совпадений с элементами `fruits` и `veggies` правил активации регулярных выражений, указанными в манифесте.</span><span class="sxs-lookup"><span data-stu-id="542eb-1200">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```javascript
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veges = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="542eb-1201">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="542eb-1201">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="542eb-1202">Возвращает строковые значения в выбранном элементе, которые соответствуют именованному регулярному выражению, определенному в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="542eb-1202">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="542eb-1203">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="542eb-1203">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="542eb-1204">Метод `getRegExMatchesByName` возвращает строки, соответствующие регулярному выражению, которое определяется в элементе правила `ItemHasRegularExpressionMatch` в XML-файле манифеста, с использованием указанного значения элемента `RegExName`.</span><span class="sxs-lookup"><span data-stu-id="542eb-1204">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="542eb-p167">Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты.</span><span class="sxs-lookup"><span data-stu-id="542eb-p167">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="542eb-1207">Параметры</span><span class="sxs-lookup"><span data-stu-id="542eb-1207">Parameters:</span></span>

|<span data-ttu-id="542eb-1208">Имя</span><span class="sxs-lookup"><span data-stu-id="542eb-1208">Name</span></span>|<span data-ttu-id="542eb-1209">Тип</span><span class="sxs-lookup"><span data-stu-id="542eb-1209">Type</span></span>|<span data-ttu-id="542eb-1210">Описание</span><span class="sxs-lookup"><span data-stu-id="542eb-1210">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="542eb-1211">String</span><span class="sxs-lookup"><span data-stu-id="542eb-1211">String</span></span>|<span data-ttu-id="542eb-1212">Имя элемента правила `ItemHasRegularExpressionMatch`, определяющее соответствующий фильтр.</span><span class="sxs-lookup"><span data-stu-id="542eb-1212">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="542eb-1213">Требования</span><span class="sxs-lookup"><span data-stu-id="542eb-1213">Requirements</span></span>

|<span data-ttu-id="542eb-1214">Requirement</span><span class="sxs-lookup"><span data-stu-id="542eb-1214">Requirement</span></span>|<span data-ttu-id="542eb-1215">Значение</span><span class="sxs-lookup"><span data-stu-id="542eb-1215">Value</span></span>|
|---|---|
|[<span data-ttu-id="542eb-1216">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="542eb-1216">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="542eb-1217">1.0</span><span class="sxs-lookup"><span data-stu-id="542eb-1217">1.0</span></span>|
|[<span data-ttu-id="542eb-1218">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="542eb-1218">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="542eb-1219">ReadItem</span><span class="sxs-lookup"><span data-stu-id="542eb-1219">ReadItem</span></span>|
|[<span data-ttu-id="542eb-1220">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="542eb-1220">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="542eb-1221">Чтение</span><span class="sxs-lookup"><span data-stu-id="542eb-1221">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="542eb-1222">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="542eb-1222">Returns:</span></span>

<span data-ttu-id="542eb-1223">Массив строк, соответствующих регулярному выражению, определяемому в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="542eb-1223">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="542eb-1224">

<dt>Тип</dt>

</span><span class="sxs-lookup"><span data-stu-id="542eb-1224">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="542eb-1225">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="542eb-1225">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="542eb-1226">Пример</span><span class="sxs-lookup"><span data-stu-id="542eb-1226">Example</span></span>

```javascript
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="542eb-1227">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="542eb-1227">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="542eb-1228">Асинхронно возвращает данные, выбранные в теме или тексте сообщения.</span><span class="sxs-lookup"><span data-stu-id="542eb-1228">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="542eb-p168">Если выделенный фрагмент отсутствует, но курсор находится в тексте или теме, метод возвращает значение NULL для выбранных данных. Если выбраны не текст и не тема, метод возвращает ошибку `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="542eb-p168">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="542eb-1231">Параметры</span><span class="sxs-lookup"><span data-stu-id="542eb-1231">Parameters:</span></span>

|<span data-ttu-id="542eb-1232">Имя</span><span class="sxs-lookup"><span data-stu-id="542eb-1232">Name</span></span>|<span data-ttu-id="542eb-1233">Тип</span><span class="sxs-lookup"><span data-stu-id="542eb-1233">Type</span></span>|<span data-ttu-id="542eb-1234">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="542eb-1234">Attributes</span></span>|<span data-ttu-id="542eb-1235">Описание</span><span class="sxs-lookup"><span data-stu-id="542eb-1235">Description</span></span>|
|---|---|---|---|
|`coercionType`|[<span data-ttu-id="542eb-1236">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="542eb-1236">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="542eb-p169">Запрашивает формат данных. Если задано значение Text, метод возвращает обычный текст как строку, удаляя все имеющиеся HTML-теги. Если задано значение HTML, метод возвращает выделенный текст (обычный текст или HTML).</span><span class="sxs-lookup"><span data-stu-id="542eb-p169">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`|<span data-ttu-id="542eb-1240">Object</span><span class="sxs-lookup"><span data-stu-id="542eb-1240">Object</span></span>|<span data-ttu-id="542eb-1241">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="542eb-1241">&lt;optional&gt;</span></span>|<span data-ttu-id="542eb-1242">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="542eb-1242">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="542eb-1243">Object</span><span class="sxs-lookup"><span data-stu-id="542eb-1243">Object</span></span>|<span data-ttu-id="542eb-1244">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="542eb-1244">&lt;optional&gt;</span></span>|<span data-ttu-id="542eb-1245">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="542eb-1245">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="542eb-1246">функция</span><span class="sxs-lookup"><span data-stu-id="542eb-1246">function</span></span>||<span data-ttu-id="542eb-1247">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="542eb-1247">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="542eb-1248">Чтобы получить доступ к выбранным данным из метода обратного вызова, вызовите `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="542eb-1248">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="542eb-1249">Для доступа к исходному свойству, представляющему собой источник выбранных данных, вызовите параметр `asyncResult.value.sourceProperty`, который может иметь значение `body` или `subject`.</span><span class="sxs-lookup"><span data-stu-id="542eb-1249">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="542eb-1250">Требования</span><span class="sxs-lookup"><span data-stu-id="542eb-1250">Requirements</span></span>

|<span data-ttu-id="542eb-1251">Requirement</span><span class="sxs-lookup"><span data-stu-id="542eb-1251">Requirement</span></span>|<span data-ttu-id="542eb-1252">Значение</span><span class="sxs-lookup"><span data-stu-id="542eb-1252">Value</span></span>|
|---|---|
|[<span data-ttu-id="542eb-1253">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="542eb-1253">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="542eb-1254">1.2</span><span class="sxs-lookup"><span data-stu-id="542eb-1254">1.2</span></span>|
|[<span data-ttu-id="542eb-1255">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="542eb-1255">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="542eb-1256">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="542eb-1256">ReadWriteItem</span></span>|
|[<span data-ttu-id="542eb-1257">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="542eb-1257">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="542eb-1258">Создание</span><span class="sxs-lookup"><span data-stu-id="542eb-1258">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="542eb-1259">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="542eb-1259">Returns:</span></span>

<span data-ttu-id="542eb-1260">Выбранные данные в виде строки с форматом, определенным в параметре `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="542eb-1260">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="542eb-1261">

<dt>Тип</dt>

</span><span class="sxs-lookup"><span data-stu-id="542eb-1261">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="542eb-1262">String</span><span class="sxs-lookup"><span data-stu-id="542eb-1262">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="542eb-1263">Пример</span><span class="sxs-lookup"><span data-stu-id="542eb-1263">Example</span></span>

```javascript
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

#### <a name="getselectedentities--entitiesjavascriptapioutlookofficeentities"></a><span data-ttu-id="542eb-1264">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="542eb-1264">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span></span>

<span data-ttu-id="542eb-p171">Возвращает сущности, найденные в выделенном совпадении, выбранном пользователем. Выделенные совпадения применяются к [контекстным надстройкам](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="542eb-p171">Gets the entities found in a highlighted match a user has selected. Highlighted matches apply to [contextual add-ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="542eb-1267">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="542eb-1267">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="542eb-1268">Требования</span><span class="sxs-lookup"><span data-stu-id="542eb-1268">Requirements</span></span>

|<span data-ttu-id="542eb-1269">Requirement</span><span class="sxs-lookup"><span data-stu-id="542eb-1269">Requirement</span></span>|<span data-ttu-id="542eb-1270">Значение</span><span class="sxs-lookup"><span data-stu-id="542eb-1270">Value</span></span>|
|---|---|
|[<span data-ttu-id="542eb-1271">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="542eb-1271">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="542eb-1272">1.6</span><span class="sxs-lookup"><span data-stu-id="542eb-1272">1.6</span></span>|
|[<span data-ttu-id="542eb-1273">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="542eb-1273">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="542eb-1274">ReadItem</span><span class="sxs-lookup"><span data-stu-id="542eb-1274">ReadItem</span></span>|
|[<span data-ttu-id="542eb-1275">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="542eb-1275">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="542eb-1276">Чтение</span><span class="sxs-lookup"><span data-stu-id="542eb-1276">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="542eb-1277">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="542eb-1277">Returns:</span></span>

<span data-ttu-id="542eb-1278">Тип: [Entities](/javascript/api/outlook/office.entities)</span><span class="sxs-lookup"><span data-stu-id="542eb-1278">Type: [Entities](/javascript/api/outlook/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="542eb-1279">Пример</span><span class="sxs-lookup"><span data-stu-id="542eb-1279">Example</span></span>

<span data-ttu-id="542eb-1280">В приведенном ниже примере показано, как получить доступ к сущностям адресов в выделенном совпадении, выбранном пользователем.</span><span class="sxs-lookup"><span data-stu-id="542eb-1280">The following example accesses the addresses entities in the highlighted match selected by the user.</span></span>

```javascript
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

#### <a name="getselectedregexmatches--object"></a><span data-ttu-id="542eb-1281">getSelectedRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="542eb-1281">getSelectedRegExMatches() → {Object}</span></span>

<span data-ttu-id="542eb-p172">Возвращает строковые значения в выделенном совпадении, которые соответствуют регулярным выражениям, определенным в XML-файле манифеста. Выделенные совпадения применяются к [контекстным надстройкам](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="542eb-p172">Returns string values in a highlighted match that match the regular expressions defined in the manifest XML file. Highlighted matches apply to [contextual add-ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="542eb-1284">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="542eb-1284">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="542eb-p173">Метод `getSelectedRegExMatches` возвращает строки, соответствующие регулярному выражению, которое определяется в каждом элементе правила `ItemHasRegularExpressionMatch` или `ItemHasKnownEntity` в XML-файле манифеста. Для правила `ItemHasRegularExpressionMatch` соответствующую строку должно содержать свойство элемента, указанного этим правилом. Простой тип `PropertyName` определяет поддерживаемые свойства.</span><span class="sxs-lookup"><span data-stu-id="542eb-p173">The `getSelectedRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="542eb-1288">Например, рассмотрим манифест надстройки, который содержит указанный ниже элемент `Rule`.</span><span class="sxs-lookup"><span data-stu-id="542eb-1288">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="542eb-1289">Объект, возвращаемый методом `getRegExMatches`, будет содержать два свойства: `fruits` и `veggies`.</span><span class="sxs-lookup"><span data-stu-id="542eb-1289">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="542eb-p174">Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты. Лучше используйте метод [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) для этого.</span><span class="sxs-lookup"><span data-stu-id="542eb-p174">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="542eb-1293">Requirements</span><span class="sxs-lookup"><span data-stu-id="542eb-1293">Requirements</span></span>

|<span data-ttu-id="542eb-1294">Requirement</span><span class="sxs-lookup"><span data-stu-id="542eb-1294">Requirement</span></span>|<span data-ttu-id="542eb-1295">Значение</span><span class="sxs-lookup"><span data-stu-id="542eb-1295">Value</span></span>|
|---|---|
|[<span data-ttu-id="542eb-1296">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="542eb-1296">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="542eb-1297">1.6</span><span class="sxs-lookup"><span data-stu-id="542eb-1297">1.6</span></span>|
|[<span data-ttu-id="542eb-1298">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="542eb-1298">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="542eb-1299">ReadItem</span><span class="sxs-lookup"><span data-stu-id="542eb-1299">ReadItem</span></span>|
|[<span data-ttu-id="542eb-1300">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="542eb-1300">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="542eb-1301">Чтение</span><span class="sxs-lookup"><span data-stu-id="542eb-1301">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="542eb-1302">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="542eb-1302">Returns:</span></span>

<span data-ttu-id="542eb-p175">Объект, содержащий массив строк, которые соответствуют регулярным выражениям, определяемым в XML-файле манифеста. Имя каждого массива равно соответствующему значению атрибута `RegExName` подходящего правила `ItemHasRegularExpressionMatch` или атрибута `FilterName` соответствующего правила `ItemHasKnownEntity`.</span><span class="sxs-lookup"><span data-stu-id="542eb-p175">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

##### <a name="example"></a><span data-ttu-id="542eb-1305">Пример</span><span class="sxs-lookup"><span data-stu-id="542eb-1305">Example</span></span>

<span data-ttu-id="542eb-1306">В приведенном ниже примере показано, как получить доступ к массиву совпадений с элементами `fruits` и `veggies` правил активации регулярных выражений, указанными в манифесте.</span><span class="sxs-lookup"><span data-stu-id="542eb-1306">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```javascript
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

#### <a name="getsharedpropertiesasyncoptions-callback"></a><span data-ttu-id="542eb-1307">getSharedPropertiesAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="542eb-1307">getSharedPropertiesAsync([options], callback)</span></span>

<span data-ttu-id="542eb-1308">Получает свойства выбранного встречи или сообщения в общей папке, календаре или почтовом ящике.</span><span class="sxs-lookup"><span data-stu-id="542eb-1308">Gets the properties of the selected appointment or message in a shared folder, calendar, or mailbox.</span></span>

##### <a name="parameters"></a><span data-ttu-id="542eb-1309">Параметры:</span><span class="sxs-lookup"><span data-stu-id="542eb-1309">Parameters:</span></span>

|<span data-ttu-id="542eb-1310">Имя</span><span class="sxs-lookup"><span data-stu-id="542eb-1310">Name</span></span>|<span data-ttu-id="542eb-1311">Тип</span><span class="sxs-lookup"><span data-stu-id="542eb-1311">Type</span></span>|<span data-ttu-id="542eb-1312">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="542eb-1312">Attributes</span></span>|<span data-ttu-id="542eb-1313">Описание</span><span class="sxs-lookup"><span data-stu-id="542eb-1313">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="542eb-1314">Object</span><span class="sxs-lookup"><span data-stu-id="542eb-1314">Object</span></span>|<span data-ttu-id="542eb-1315">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="542eb-1315">&lt;optional&gt;</span></span>|<span data-ttu-id="542eb-1316">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="542eb-1316">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="542eb-1317">Object</span><span class="sxs-lookup"><span data-stu-id="542eb-1317">Object</span></span>|<span data-ttu-id="542eb-1318">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="542eb-1318">&lt;optional&gt;</span></span>|<span data-ttu-id="542eb-1319">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="542eb-1319">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="542eb-1320">функция</span><span class="sxs-lookup"><span data-stu-id="542eb-1320">function</span></span>||<span data-ttu-id="542eb-1321">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="542eb-1321">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="542eb-1322">Общие свойства предоставляются в виде объекта [`SharedProperties`](/javascript/api/outlook/office.sharedproperties) в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="542eb-1322">The custom properties are provided as a [`SharedProperties`](/javascript/api/outlook/office.sharedproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="542eb-1323">Этот объект можно использовать для получения общих свойств элемента.</span><span class="sxs-lookup"><span data-stu-id="542eb-1323">This object can be used to get the item's shared properties.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="542eb-1324">Требования</span><span class="sxs-lookup"><span data-stu-id="542eb-1324">Requirements</span></span>

|<span data-ttu-id="542eb-1325">Requirement</span><span class="sxs-lookup"><span data-stu-id="542eb-1325">Requirement</span></span>|<span data-ttu-id="542eb-1326">Значение</span><span class="sxs-lookup"><span data-stu-id="542eb-1326">Value</span></span>|
|---|---|
|[<span data-ttu-id="542eb-1327">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="542eb-1327">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="542eb-1328">Предварительная версия</span><span class="sxs-lookup"><span data-stu-id="542eb-1328">Preview</span></span>|
|[<span data-ttu-id="542eb-1329">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="542eb-1329">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="542eb-1330">ReadItem</span><span class="sxs-lookup"><span data-stu-id="542eb-1330">ReadItem</span></span>|
|[<span data-ttu-id="542eb-1331">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="542eb-1331">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="542eb-1332">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="542eb-1332">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="542eb-1333">Пример</span><span class="sxs-lookup"><span data-stu-id="542eb-1333">Example</span></span>

```js
Office.context.mailbox.item.getSharedPropertiesAsync(callback);
function callback (asyncResult) {
  var context=asyncResult.context;
  var sharedProperties = asyncResult.value;
}
```

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="542eb-1334">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="542eb-1334">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="542eb-1335">Асинхронно загружает настраиваемые свойства для надстройки для выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="542eb-1335">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="542eb-p177">Настраиваемые свойства сохраняются в виде пар "ключ-значение" для каждого приложения и каждого элемента. Этот метод возвращает объект `CustomProperties` при обратном вызове, который предоставляет методы для доступа к настраиваемым свойствам, характерным для текущего элемента и текущей надстройки. Настраиваемые свойства не шифруются для элемента, поэтому этот способ хранения не является безопасным.</span><span class="sxs-lookup"><span data-stu-id="542eb-p177">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="542eb-1339">Параметры</span><span class="sxs-lookup"><span data-stu-id="542eb-1339">Parameters:</span></span>

|<span data-ttu-id="542eb-1340">Имя</span><span class="sxs-lookup"><span data-stu-id="542eb-1340">Name</span></span>|<span data-ttu-id="542eb-1341">Тип</span><span class="sxs-lookup"><span data-stu-id="542eb-1341">Type</span></span>|<span data-ttu-id="542eb-1342">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="542eb-1342">Attributes</span></span>|<span data-ttu-id="542eb-1343">Описание</span><span class="sxs-lookup"><span data-stu-id="542eb-1343">Description</span></span>|
|---|---|---|---|
|`callback`|<span data-ttu-id="542eb-1344">function</span><span class="sxs-lookup"><span data-stu-id="542eb-1344">function</span></span>||<span data-ttu-id="542eb-1345">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="542eb-1345">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="542eb-1346">Настраиваемые свойства предоставляются в виде объекта [`CustomProperties`](/javascript/api/outlook/office.customproperties) в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="542eb-1346">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="542eb-1347">Этот объект позволяет получить, задать и удалить настраиваемые свойства для элемента, а также сохранить изменения, внесенные в набор настраиваемых свойств, на сервере.</span><span class="sxs-lookup"><span data-stu-id="542eb-1347">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`|<span data-ttu-id="542eb-1348">Объект</span><span class="sxs-lookup"><span data-stu-id="542eb-1348">Object</span></span>|<span data-ttu-id="542eb-1349">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="542eb-1349">&lt;optional&gt;</span></span>|<span data-ttu-id="542eb-1350">Разработчики могут указать любой объект, к которому необходимо получить доступ, в функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="542eb-1350">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="542eb-1351">Доступ к этому объекту можно получить с помощью свойства `asyncResult.asyncContext` в функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="542eb-1351">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="542eb-1352">Требования</span><span class="sxs-lookup"><span data-stu-id="542eb-1352">Requirements</span></span>

|<span data-ttu-id="542eb-1353">Requirement</span><span class="sxs-lookup"><span data-stu-id="542eb-1353">Requirement</span></span>|<span data-ttu-id="542eb-1354">Значение</span><span class="sxs-lookup"><span data-stu-id="542eb-1354">Value</span></span>|
|---|---|
|[<span data-ttu-id="542eb-1355">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="542eb-1355">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="542eb-1356">1.0</span><span class="sxs-lookup"><span data-stu-id="542eb-1356">1.0</span></span>|
|[<span data-ttu-id="542eb-1357">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="542eb-1357">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="542eb-1358">ReadItem</span><span class="sxs-lookup"><span data-stu-id="542eb-1358">ReadItem</span></span>|
|[<span data-ttu-id="542eb-1359">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="542eb-1359">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="542eb-1360">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="542eb-1360">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="542eb-1361">Пример</span><span class="sxs-lookup"><span data-stu-id="542eb-1361">Example</span></span>

<span data-ttu-id="542eb-p180">Приведенный ниже пример кода показывает, как асинхронно загружать настраиваемые свойства, характерные для текущего элемента, с помощью метода `loadCustomPropertiesAsync`. Этот пример также показывает, как сохранять эти свойства на сервере с помощью метода `CustomProperties.saveAsync`. После загрузки настраиваемых свойств в этом примере кода метод `CustomProperties.get` используется для считывания настраиваемого свойства `myProp`, метод `CustomProperties.set` — для записи настраиваемого свойства `otherProp`, а метод `saveAsync` — для сохранения настраиваемых свойств.</span><span class="sxs-lookup"><span data-stu-id="542eb-p180">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

```javascript
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

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="542eb-1365">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="542eb-1365">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="542eb-1366">Удаляет вложение из сообщения или встречи.</span><span class="sxs-lookup"><span data-stu-id="542eb-1366">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="542eb-1367">Метод `removeAttachmentAsync` удаляет из элемента вложение с указанным идентификатором.</span><span class="sxs-lookup"><span data-stu-id="542eb-1367">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="542eb-1368">Идентификатор вложения рекомендуется использовать для удаления вложения, только если оно добавлено тем же почтовым приложением в ходе текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="542eb-1368">As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session.</span></span> <span data-ttu-id="542eb-1369">В Outlook Web App и Outlook Web App для устройств идентификатор вложения действителен только в рамках одного сеанса.</span><span class="sxs-lookup"><span data-stu-id="542eb-1369">In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="542eb-1370">Сеанс завершается, когда пользователь закрывает приложение или начинает создавать элемент во встроенной форме, а затем переходит из формы в отдельное окно.</span><span class="sxs-lookup"><span data-stu-id="542eb-1370">A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="542eb-1371">Параметры:</span><span class="sxs-lookup"><span data-stu-id="542eb-1371">Parameters:</span></span>

|<span data-ttu-id="542eb-1372">Имя</span><span class="sxs-lookup"><span data-stu-id="542eb-1372">Name</span></span>|<span data-ttu-id="542eb-1373">Тип</span><span class="sxs-lookup"><span data-stu-id="542eb-1373">Type</span></span>|<span data-ttu-id="542eb-1374">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="542eb-1374">Attributes</span></span>|<span data-ttu-id="542eb-1375">Описание</span><span class="sxs-lookup"><span data-stu-id="542eb-1375">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="542eb-1376">String</span><span class="sxs-lookup"><span data-stu-id="542eb-1376">String</span></span>||<span data-ttu-id="542eb-p182">Идентификатор удаляемого вложения. Максимальная длина строки — 100 символов.</span><span class="sxs-lookup"><span data-stu-id="542eb-p182">The identifier of the attachment to remove. The maximum length of the string is 100 characters.</span></span>|
|`options`|<span data-ttu-id="542eb-1379">Object</span><span class="sxs-lookup"><span data-stu-id="542eb-1379">Object</span></span>|<span data-ttu-id="542eb-1380">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="542eb-1380">&lt;optional&gt;</span></span>|<span data-ttu-id="542eb-1381">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="542eb-1381">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="542eb-1382">Object</span><span class="sxs-lookup"><span data-stu-id="542eb-1382">Object</span></span>|<span data-ttu-id="542eb-1383">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="542eb-1383">&lt;optional&gt;</span></span>|<span data-ttu-id="542eb-1384">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="542eb-1384">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="542eb-1385">функция</span><span class="sxs-lookup"><span data-stu-id="542eb-1385">function</span></span>|<span data-ttu-id="542eb-1386">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="542eb-1386">&lt;optional&gt;</span></span>|<span data-ttu-id="542eb-1387">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="542eb-1387">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="542eb-1388">Если удалить вложение не удается, свойство `asyncResult.error` содержит код ошибки с указанием ее причины.</span><span class="sxs-lookup"><span data-stu-id="542eb-1388">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="542eb-1389">Ошибки</span><span class="sxs-lookup"><span data-stu-id="542eb-1389">Errors</span></span>

|<span data-ttu-id="542eb-1390">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="542eb-1390">Error code</span></span>|<span data-ttu-id="542eb-1391">Описание</span><span class="sxs-lookup"><span data-stu-id="542eb-1391">Description</span></span>|
|------------|-------------|
|`InvalidAttachmentId`|<span data-ttu-id="542eb-1392">Идентификатор вложения не существует.</span><span class="sxs-lookup"><span data-stu-id="542eb-1392">The attachment identifier does not exist.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="542eb-1393">Требования</span><span class="sxs-lookup"><span data-stu-id="542eb-1393">Requirements</span></span>

|<span data-ttu-id="542eb-1394">Requirement</span><span class="sxs-lookup"><span data-stu-id="542eb-1394">Requirement</span></span>|<span data-ttu-id="542eb-1395">Значение</span><span class="sxs-lookup"><span data-stu-id="542eb-1395">Value</span></span>|
|---|---|
|[<span data-ttu-id="542eb-1396">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="542eb-1396">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="542eb-1397">1.1</span><span class="sxs-lookup"><span data-stu-id="542eb-1397">1.1</span></span>|
|[<span data-ttu-id="542eb-1398">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="542eb-1398">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="542eb-1399">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="542eb-1399">ReadWriteItem</span></span>|
|[<span data-ttu-id="542eb-1400">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="542eb-1400">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="542eb-1401">Создание</span><span class="sxs-lookup"><span data-stu-id="542eb-1401">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="542eb-1402">Пример</span><span class="sxs-lookup"><span data-stu-id="542eb-1402">Example</span></span>

<span data-ttu-id="542eb-1403">Указанный ниже код удаляет вложение с идентификатором "0".</span><span class="sxs-lookup"><span data-stu-id="542eb-1403">The following code removes an attachment with an identifier of '0'.</span></span>

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

####  <a name="removehandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="542eb-1404">removeHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="542eb-1404">removeHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="542eb-1405">Удаляет обработчик для поддерживаемого события.</span><span class="sxs-lookup"><span data-stu-id="542eb-1405">Removes an event handler for a supported event.</span></span>

<span data-ttu-id="542eb-1406">Сейчас поддерживаются следующие типы событий: `Office.EventType.AttachmentsChanged`, `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged` и `Office.EventType.RecurrenceChanged`.</span><span class="sxs-lookup"><span data-stu-id="542eb-1406">Currently the supported event types are `Office.EventType.AttachmentsChanged`, `Office.EventType.AppointmentTimeChanged`, and `Office.EventType.RecipientsChanged`</span></span>

##### <a name="parameters"></a><span data-ttu-id="542eb-1407">Параметры:</span><span class="sxs-lookup"><span data-stu-id="542eb-1407">Parameters:</span></span>

| <span data-ttu-id="542eb-1408">Имя</span><span class="sxs-lookup"><span data-stu-id="542eb-1408">Name</span></span> | <span data-ttu-id="542eb-1409">Тип</span><span class="sxs-lookup"><span data-stu-id="542eb-1409">Type</span></span> | <span data-ttu-id="542eb-1410">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="542eb-1410">Attributes</span></span> | <span data-ttu-id="542eb-1411">Описание</span><span class="sxs-lookup"><span data-stu-id="542eb-1411">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="542eb-1412">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="542eb-1412">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="542eb-1413">Событие, которое должно отменить обработчик.</span><span class="sxs-lookup"><span data-stu-id="542eb-1413">The event that should revoke the handler.</span></span> |
| `handler` | <span data-ttu-id="542eb-1414">Функция</span><span class="sxs-lookup"><span data-stu-id="542eb-1414">Function</span></span> || <span data-ttu-id="542eb-p183">Функция для обработки события. Функция должна принимать один параметр, представляющий собой объектный литерал. Значение свойства `type` параметра совпадет со значением параметра `eventType`, переданного методу `removeHandlerAsync`.</span><span class="sxs-lookup"><span data-stu-id="542eb-p183">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `removeHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="542eb-1418">Объект</span><span class="sxs-lookup"><span data-stu-id="542eb-1418">Object</span></span> | <span data-ttu-id="542eb-1419">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="542eb-1419">&lt;optional&gt;</span></span> | <span data-ttu-id="542eb-1420">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="542eb-1420">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="542eb-1421">Object</span><span class="sxs-lookup"><span data-stu-id="542eb-1421">Object</span></span> | <span data-ttu-id="542eb-1422">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="542eb-1422">&lt;optional&gt;</span></span> | <span data-ttu-id="542eb-1423">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="542eb-1423">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="542eb-1424">функция</span><span class="sxs-lookup"><span data-stu-id="542eb-1424">function</span></span>| <span data-ttu-id="542eb-1425">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="542eb-1425">&lt;optional&gt;</span></span>|<span data-ttu-id="542eb-1426">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="542eb-1426">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="542eb-1427">Требования</span><span class="sxs-lookup"><span data-stu-id="542eb-1427">Requirements</span></span>

|<span data-ttu-id="542eb-1428">Requirement</span><span class="sxs-lookup"><span data-stu-id="542eb-1428">Requirement</span></span>| <span data-ttu-id="542eb-1429">Значение</span><span class="sxs-lookup"><span data-stu-id="542eb-1429">Value</span></span>|
|---|---|
|[<span data-ttu-id="542eb-1430">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="542eb-1430">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="542eb-1431">1.7</span><span class="sxs-lookup"><span data-stu-id="542eb-1431">1.7</span></span> |
|[<span data-ttu-id="542eb-1432">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="542eb-1432">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="542eb-1433">ReadItem</span><span class="sxs-lookup"><span data-stu-id="542eb-1433">ReadItem</span></span> |
|[<span data-ttu-id="542eb-1434">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="542eb-1434">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="542eb-1435">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="542eb-1435">Compose or read</span></span> |

####  <a name="saveasyncoptions-callback"></a><span data-ttu-id="542eb-1436">saveAsync([options], обратный вызов)</span><span class="sxs-lookup"><span data-stu-id="542eb-1436">saveAsync([options], callback)</span></span>

<span data-ttu-id="542eb-1437">Асинхронно сохраняет элемент.</span><span class="sxs-lookup"><span data-stu-id="542eb-1437">Asynchronously saves an item.</span></span>

<span data-ttu-id="542eb-p184">При вызове этот метод сохраняет текущее сообщение в виде черновика и возвращает идентификатор элемента с помощью метода обратного вызова. В Outlook Web App или интерактивном режиме Outlook этот элемент сохраняется на сервере. В Outlook в режиме кэширования этот элемент сохраняется в локальном кэше.</span><span class="sxs-lookup"><span data-stu-id="542eb-p184">When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="542eb-1441">Если в режиме создания надстройка вызывает для элемента метод `saveAsync`, чтобы получить параметр `itemId` для использования с EWS или REST API, необходимо помнить, что синхронизация элемента с сервером может занять много времени, если Outlook работает в режиме кэширования данных.</span><span class="sxs-lookup"><span data-stu-id="542eb-1441">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="542eb-1442">До окончания синхронизации элемента применение параметра `itemId` будет приводить к ошибке.</span><span class="sxs-lookup"><span data-stu-id="542eb-1442">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="542eb-p186">Если метод `saveAsync` вызывается для встречи в режиме создания, она сохраняется как обычная встреча в календаре пользователя, а не как черновик. При сохранении новой встречи приглашения не отправляются. При сохранении существующей встречи уведомления отправляются добавленным или удаленным участникам.</span><span class="sxs-lookup"><span data-stu-id="542eb-p186">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="542eb-1446">Следующие клиенты отличаются другим поведением `saveAsync` в отношении встреч в режиме создания:</span><span class="sxs-lookup"><span data-stu-id="542eb-1446">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="542eb-1447">Outlook для Mac не поддерживает `saveAsync` для собраний в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="542eb-1447">Mac Outlook does not support `saveAsync` on a meeting in compose mode.</span></span> <span data-ttu-id="542eb-1448">При вызове `saveAsync` для собрания в Outlook для Mac возвращается ошибка.</span><span class="sxs-lookup"><span data-stu-id="542eb-1448">Calling `saveAsync` on a meeting in Mac Outlook will return an error.</span></span>
> - <span data-ttu-id="542eb-1449">Outlook в Интернете всегда отправляет приглашение или обновление при вызове `saveAsync` для встречи в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="542eb-1449">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="542eb-1450">Параметры:</span><span class="sxs-lookup"><span data-stu-id="542eb-1450">Parameters:</span></span>

|<span data-ttu-id="542eb-1451">Имя</span><span class="sxs-lookup"><span data-stu-id="542eb-1451">Name</span></span>|<span data-ttu-id="542eb-1452">Тип</span><span class="sxs-lookup"><span data-stu-id="542eb-1452">Type</span></span>|<span data-ttu-id="542eb-1453">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="542eb-1453">Attributes</span></span>|<span data-ttu-id="542eb-1454">Описание</span><span class="sxs-lookup"><span data-stu-id="542eb-1454">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="542eb-1455">Object</span><span class="sxs-lookup"><span data-stu-id="542eb-1455">Object</span></span>|<span data-ttu-id="542eb-1456">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="542eb-1456">&lt;optional&gt;</span></span>|<span data-ttu-id="542eb-1457">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="542eb-1457">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="542eb-1458">Object</span><span class="sxs-lookup"><span data-stu-id="542eb-1458">Object</span></span>|<span data-ttu-id="542eb-1459">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="542eb-1459">&lt;optional&gt;</span></span>|<span data-ttu-id="542eb-1460">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="542eb-1460">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="542eb-1461">функция</span><span class="sxs-lookup"><span data-stu-id="542eb-1461">function</span></span>||<span data-ttu-id="542eb-1462">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="542eb-1462">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="542eb-1463">После успешного выполнения идентификатор элемента будет предоставлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="542eb-1463">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="542eb-1464">Требования</span><span class="sxs-lookup"><span data-stu-id="542eb-1464">Requirements</span></span>

|<span data-ttu-id="542eb-1465">Requirement</span><span class="sxs-lookup"><span data-stu-id="542eb-1465">Requirement</span></span>|<span data-ttu-id="542eb-1466">Значение</span><span class="sxs-lookup"><span data-stu-id="542eb-1466">Value</span></span>|
|---|---|
|[<span data-ttu-id="542eb-1467">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="542eb-1467">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="542eb-1468">1.3</span><span class="sxs-lookup"><span data-stu-id="542eb-1468">1.3</span></span>|
|[<span data-ttu-id="542eb-1469">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="542eb-1469">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="542eb-1470">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="542eb-1470">ReadWriteItem</span></span>|
|[<span data-ttu-id="542eb-1471">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="542eb-1471">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="542eb-1472">Создание</span><span class="sxs-lookup"><span data-stu-id="542eb-1472">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="542eb-1473">Примеры</span><span class="sxs-lookup"><span data-stu-id="542eb-1473">Examples</span></span>

```javascript
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result
  });
```

<span data-ttu-id="542eb-p188">Ниже приведен пример параметра `result`, переданного функции обратного вызова. Свойство `value` содержит идентификатор элемента.</span><span class="sxs-lookup"><span data-stu-id="542eb-p188">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

####  <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="542eb-1476">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="542eb-1476">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="542eb-1477">Асинхронно вставляет данные в текст или тему сообщения.</span><span class="sxs-lookup"><span data-stu-id="542eb-1477">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="542eb-p189">Метод `setSelectedDataAsync` вставляет указанную строку в местоположение курсора в теме или тексте элемента либо, если текст выделен в редакторе, он заменяет выделенный текст. Если курсор находится вне текста или темы элемента, возвращается ошибка. После вставки курсор помещается в конец вставленного содержимого.</span><span class="sxs-lookup"><span data-stu-id="542eb-p189">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="542eb-1481">Параметры:</span><span class="sxs-lookup"><span data-stu-id="542eb-1481">Parameters:</span></span>

|<span data-ttu-id="542eb-1482">Имя</span><span class="sxs-lookup"><span data-stu-id="542eb-1482">Name</span></span>|<span data-ttu-id="542eb-1483">Тип</span><span class="sxs-lookup"><span data-stu-id="542eb-1483">Type</span></span>|<span data-ttu-id="542eb-1484">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="542eb-1484">Attributes</span></span>|<span data-ttu-id="542eb-1485">Описание</span><span class="sxs-lookup"><span data-stu-id="542eb-1485">Description</span></span>|
|---|---|---|---|
|`data`|<span data-ttu-id="542eb-1486">String</span><span class="sxs-lookup"><span data-stu-id="542eb-1486">String</span></span>||<span data-ttu-id="542eb-p190">Вставляемые данные. Объем данных не должен превышать 1 000 000 символов. Если передано больше 1 000 000 символов, возвращается исключение `ArgumentOutOfRange`.</span><span class="sxs-lookup"><span data-stu-id="542eb-p190">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`|<span data-ttu-id="542eb-1490">Object</span><span class="sxs-lookup"><span data-stu-id="542eb-1490">Object</span></span>|<span data-ttu-id="542eb-1491">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="542eb-1491">&lt;optional&gt;</span></span>|<span data-ttu-id="542eb-1492">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="542eb-1492">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="542eb-1493">Object</span><span class="sxs-lookup"><span data-stu-id="542eb-1493">Object</span></span>|<span data-ttu-id="542eb-1494">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="542eb-1494">&lt;optional&gt;</span></span>|<span data-ttu-id="542eb-1495">В методе обратного вызова разработчики могут указать любой объект, к которому необходимо получить доступ.</span><span class="sxs-lookup"><span data-stu-id="542eb-1495">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="542eb-1496">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="542eb-1496">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="542eb-1497">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="542eb-1497">&lt;optional&gt;</span></span>|<span data-ttu-id="542eb-p191">Если задано значение `text`, текущий стиль применяется в Outlook Web App и Outlook. Если поле представляет собой редактор HTML, вставляются только текстовые данные, даже если они имеют формат HTML.</span><span class="sxs-lookup"><span data-stu-id="542eb-p191">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="542eb-p192">Если задано значение `html` и поле (не тема) поддерживает HTML, в Outlook Web App применяется текущий стиль, а в Outlook — стиль по умолчанию. Если поле является текстовым, возвращается ошибка `InvalidDataFormat`.</span><span class="sxs-lookup"><span data-stu-id="542eb-p192">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="542eb-1502">Если свойство `coercionType` не задано, результат зависит от поля: если поле имеет формат HTML, используется текст в формате HTML, а если поле текстовое, применяется обычный текст.</span><span class="sxs-lookup"><span data-stu-id="542eb-1502">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`|<span data-ttu-id="542eb-1503">функция</span><span class="sxs-lookup"><span data-stu-id="542eb-1503">function</span></span>||<span data-ttu-id="542eb-1504">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="542eb-1504">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="542eb-1505">Требования</span><span class="sxs-lookup"><span data-stu-id="542eb-1505">Requirements</span></span>

|<span data-ttu-id="542eb-1506">Requirement</span><span class="sxs-lookup"><span data-stu-id="542eb-1506">Requirement</span></span>|<span data-ttu-id="542eb-1507">Значение</span><span class="sxs-lookup"><span data-stu-id="542eb-1507">Value</span></span>|
|---|---|
|[<span data-ttu-id="542eb-1508">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="542eb-1508">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="542eb-1509">1.2</span><span class="sxs-lookup"><span data-stu-id="542eb-1509">1.2</span></span>|
|[<span data-ttu-id="542eb-1510">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="542eb-1510">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="542eb-1511">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="542eb-1511">ReadWriteItem</span></span>|
|[<span data-ttu-id="542eb-1512">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="542eb-1512">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="542eb-1513">Создание</span><span class="sxs-lookup"><span data-stu-id="542eb-1513">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="542eb-1514">Пример</span><span class="sxs-lookup"><span data-stu-id="542eb-1514">Example</span></span>

```javascript
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```