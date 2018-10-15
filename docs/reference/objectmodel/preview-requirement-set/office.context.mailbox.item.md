
# <a name="item"></a><span data-ttu-id="8b9b5-101">item</span><span class="sxs-lookup"><span data-stu-id="8b9b5-101">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="8b9b5-102">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="8b9b5-102">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="8b9b5-p101">Пространство имен `item` используется для доступа к выбранному в данный момент сообщению, приглашению на собрание или описанию встречи. Вы можете определить тип пространства имен `item` с помощью свойства [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="8b9b5-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="8b9b5-105">Требования</span><span class="sxs-lookup"><span data-stu-id="8b9b5-105">Requirements</span></span>

|<span data-ttu-id="8b9b5-106">Требование</span><span class="sxs-lookup"><span data-stu-id="8b9b5-106">Requirement</span></span>|<span data-ttu-id="8b9b5-107">Значение</span><span class="sxs-lookup"><span data-stu-id="8b9b5-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="8b9b5-108">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="8b9b5-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8b9b5-109">1.0</span><span class="sxs-lookup"><span data-stu-id="8b9b5-109">1.0</span></span>|
|[<span data-ttu-id="8b9b5-110">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="8b9b5-110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8b9b5-111">Ограниченный доступ</span><span class="sxs-lookup"><span data-stu-id="8b9b5-111">Restricted</span></span>|
|[<span data-ttu-id="8b9b5-112">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="8b9b5-112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8b9b5-113">Cоздание или чтение</span><span class="sxs-lookup"><span data-stu-id="8b9b5-113">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="8b9b5-114">Члены и методы</span><span class="sxs-lookup"><span data-stu-id="8b9b5-114">Members and methods</span></span>

| <span data-ttu-id="8b9b5-115">Член</span><span class="sxs-lookup"><span data-stu-id="8b9b5-115">Member</span></span> | <span data-ttu-id="8b9b5-116">Тип</span><span class="sxs-lookup"><span data-stu-id="8b9b5-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="8b9b5-117">attachments</span><span class="sxs-lookup"><span data-stu-id="8b9b5-117">attachments</span></span>](#attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails) | <span data-ttu-id="8b9b5-118">Член</span><span class="sxs-lookup"><span data-stu-id="8b9b5-118">Member</span></span> |
| [<span data-ttu-id="8b9b5-119">bcc</span><span class="sxs-lookup"><span data-stu-id="8b9b5-119">bcc</span></span>](#bcc-recipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="8b9b5-120">Член</span><span class="sxs-lookup"><span data-stu-id="8b9b5-120">Member</span></span> |
| [<span data-ttu-id="8b9b5-121">body</span><span class="sxs-lookup"><span data-stu-id="8b9b5-121">body</span></span>](#body-bodyjavascriptapioutlookofficebody) | <span data-ttu-id="8b9b5-122">Член</span><span class="sxs-lookup"><span data-stu-id="8b9b5-122">Member</span></span> |
| [<span data-ttu-id="8b9b5-123">cc</span><span class="sxs-lookup"><span data-stu-id="8b9b5-123">cc</span></span>](#cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="8b9b5-124">Член</span><span class="sxs-lookup"><span data-stu-id="8b9b5-124">Member</span></span> |
| [<span data-ttu-id="8b9b5-125">conversationId</span><span class="sxs-lookup"><span data-stu-id="8b9b5-125">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="8b9b5-126">Член</span><span class="sxs-lookup"><span data-stu-id="8b9b5-126">Member</span></span> |
| [<span data-ttu-id="8b9b5-127">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="8b9b5-127">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="8b9b5-128">Член</span><span class="sxs-lookup"><span data-stu-id="8b9b5-128">Member</span></span> |
| [<span data-ttu-id="8b9b5-129">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="8b9b5-129">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="8b9b5-130">Член</span><span class="sxs-lookup"><span data-stu-id="8b9b5-130">Member</span></span> |
| [<span data-ttu-id="8b9b5-131">end</span><span class="sxs-lookup"><span data-stu-id="8b9b5-131">end</span></span>](#end-datetimejavascriptapioutlookofficetime) | <span data-ttu-id="8b9b5-132">Член</span><span class="sxs-lookup"><span data-stu-id="8b9b5-132">Member</span></span> |
| [<span data-ttu-id="8b9b5-133">from</span><span class="sxs-lookup"><span data-stu-id="8b9b5-133">from</span></span>](#from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom) | <span data-ttu-id="8b9b5-134">Член</span><span class="sxs-lookup"><span data-stu-id="8b9b5-134">Member</span></span> |
| [<span data-ttu-id="8b9b5-135">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="8b9b5-135">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="8b9b5-136">Член</span><span class="sxs-lookup"><span data-stu-id="8b9b5-136">Member</span></span> |
| [<span data-ttu-id="8b9b5-137">itemClass</span><span class="sxs-lookup"><span data-stu-id="8b9b5-137">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="8b9b5-138">Член</span><span class="sxs-lookup"><span data-stu-id="8b9b5-138">Member</span></span> |
| [<span data-ttu-id="8b9b5-139">itemId</span><span class="sxs-lookup"><span data-stu-id="8b9b5-139">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="8b9b5-140">Член</span><span class="sxs-lookup"><span data-stu-id="8b9b5-140">Member</span></span> |
| [<span data-ttu-id="8b9b5-141">itemType</span><span class="sxs-lookup"><span data-stu-id="8b9b5-141">itemType</span></span>](#itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype) | <span data-ttu-id="8b9b5-142">Член</span><span class="sxs-lookup"><span data-stu-id="8b9b5-142">Member</span></span> |
| [<span data-ttu-id="8b9b5-143">location</span><span class="sxs-lookup"><span data-stu-id="8b9b5-143">location</span></span>](#location-stringlocationjavascriptapioutlookofficelocation) | <span data-ttu-id="8b9b5-144">Член</span><span class="sxs-lookup"><span data-stu-id="8b9b5-144">Member</span></span> |
| [<span data-ttu-id="8b9b5-145">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="8b9b5-145">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="8b9b5-146">Член</span><span class="sxs-lookup"><span data-stu-id="8b9b5-146">Member</span></span> |
| [<span data-ttu-id="8b9b5-147">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="8b9b5-147">notificationMessages</span></span>](#notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessages) | <span data-ttu-id="8b9b5-148">Член</span><span class="sxs-lookup"><span data-stu-id="8b9b5-148">Member</span></span> |
| [<span data-ttu-id="8b9b5-149">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="8b9b5-149">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="8b9b5-150">Член</span><span class="sxs-lookup"><span data-stu-id="8b9b5-150">Member</span></span> |
| [<span data-ttu-id="8b9b5-151">organizer</span><span class="sxs-lookup"><span data-stu-id="8b9b5-151">organizer</span></span>](#organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsorganizerjavascriptapioutlookofficeorganizer) | <span data-ttu-id="8b9b5-152">Член</span><span class="sxs-lookup"><span data-stu-id="8b9b5-152">Member</span></span> |
| [<span data-ttu-id="8b9b5-153">recurrence</span><span class="sxs-lookup"><span data-stu-id="8b9b5-153">recurrence</span></span>](#nullable-recurrence-recurrencejavascriptapioutlookofficerecurrence) | <span data-ttu-id="8b9b5-154">Член</span><span class="sxs-lookup"><span data-stu-id="8b9b5-154">Member</span></span> |
| [<span data-ttu-id="8b9b5-155">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="8b9b5-155">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="8b9b5-156">Член</span><span class="sxs-lookup"><span data-stu-id="8b9b5-156">Member</span></span> |
| [<span data-ttu-id="8b9b5-157">sender</span><span class="sxs-lookup"><span data-stu-id="8b9b5-157">sender</span></span>](#sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails) | <span data-ttu-id="8b9b5-158">Член</span><span class="sxs-lookup"><span data-stu-id="8b9b5-158">Member</span></span> |
| [<span data-ttu-id="8b9b5-159">seriesId</span><span class="sxs-lookup"><span data-stu-id="8b9b5-159">seriesId</span></span>](#nullable-seriesid-string) | <span data-ttu-id="8b9b5-160">Член</span><span class="sxs-lookup"><span data-stu-id="8b9b5-160">Member</span></span> |
| [<span data-ttu-id="8b9b5-161">start</span><span class="sxs-lookup"><span data-stu-id="8b9b5-161">start</span></span>](#start-datetimejavascriptapioutlookofficetime) | <span data-ttu-id="8b9b5-162">Член</span><span class="sxs-lookup"><span data-stu-id="8b9b5-162">Member</span></span> |
| [<span data-ttu-id="8b9b5-163">subject</span><span class="sxs-lookup"><span data-stu-id="8b9b5-163">subject</span></span>](#subject-stringsubjectjavascriptapioutlookofficesubject) | <span data-ttu-id="8b9b5-164">Член</span><span class="sxs-lookup"><span data-stu-id="8b9b5-164">Member</span></span> |
| [<span data-ttu-id="8b9b5-165">to</span><span class="sxs-lookup"><span data-stu-id="8b9b5-165">to</span></span>](#to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="8b9b5-166">Член</span><span class="sxs-lookup"><span data-stu-id="8b9b5-166">Member</span></span> |
| [<span data-ttu-id="8b9b5-167">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="8b9b5-167">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="8b9b5-168">Метод</span><span class="sxs-lookup"><span data-stu-id="8b9b5-168">Method</span></span> |
| [<span data-ttu-id="8b9b5-169">addFileAttachmentFromBase64Async</span><span class="sxs-lookup"><span data-stu-id="8b9b5-169">addFileAttachmentFromBase64Async</span></span>](#addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback) | <span data-ttu-id="8b9b5-170">Метод</span><span class="sxs-lookup"><span data-stu-id="8b9b5-170">Method</span></span> |
| [<span data-ttu-id="8b9b5-171">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="8b9b5-171">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="8b9b5-172">Метод</span><span class="sxs-lookup"><span data-stu-id="8b9b5-172">Method</span></span> |
| [<span data-ttu-id="8b9b5-173">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="8b9b5-173">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="8b9b5-174">Метод</span><span class="sxs-lookup"><span data-stu-id="8b9b5-174">Method</span></span> |
| [<span data-ttu-id="8b9b5-175">close</span><span class="sxs-lookup"><span data-stu-id="8b9b5-175">close</span></span>](#close) | <span data-ttu-id="8b9b5-176">Метод</span><span class="sxs-lookup"><span data-stu-id="8b9b5-176">Method</span></span> |
| [<span data-ttu-id="8b9b5-177">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="8b9b5-177">displayReplyAllForm</span></span>](#displayreplyallformformdata) | <span data-ttu-id="8b9b5-178">Метод</span><span class="sxs-lookup"><span data-stu-id="8b9b5-178">Method</span></span> |
| [<span data-ttu-id="8b9b5-179">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="8b9b5-179">displayReplyForm</span></span>](#displayreplyformformdata) | <span data-ttu-id="8b9b5-180">Метод</span><span class="sxs-lookup"><span data-stu-id="8b9b5-180">Method</span></span> |
| [<span data-ttu-id="8b9b5-181">getEntities</span><span class="sxs-lookup"><span data-stu-id="8b9b5-181">getEntities</span></span>](#getentities--entitiesjavascriptapioutlookofficeentities) | <span data-ttu-id="8b9b5-182">Метод</span><span class="sxs-lookup"><span data-stu-id="8b9b5-182">Method</span></span> |
| [<span data-ttu-id="8b9b5-183">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="8b9b5-183">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion) | <span data-ttu-id="8b9b5-184">Метод</span><span class="sxs-lookup"><span data-stu-id="8b9b5-184">Method</span></span> |
| [<span data-ttu-id="8b9b5-185">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="8b9b5-185">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion) | <span data-ttu-id="8b9b5-186">Метод</span><span class="sxs-lookup"><span data-stu-id="8b9b5-186">Method</span></span> |
| [<span data-ttu-id="8b9b5-187">getInitializationContextAsync</span><span class="sxs-lookup"><span data-stu-id="8b9b5-187">getInitializationContextAsync</span></span>](#getinitializationcontextasyncoptions-callback) | <span data-ttu-id="8b9b5-188">Метод</span><span class="sxs-lookup"><span data-stu-id="8b9b5-188">Method</span></span> |
| [<span data-ttu-id="8b9b5-189">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="8b9b5-189">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="8b9b5-190">Метод</span><span class="sxs-lookup"><span data-stu-id="8b9b5-190">Method</span></span> |
| [<span data-ttu-id="8b9b5-191">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="8b9b5-191">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="8b9b5-192">Метод</span><span class="sxs-lookup"><span data-stu-id="8b9b5-192">Method</span></span> |
| [<span data-ttu-id="8b9b5-193">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="8b9b5-193">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="8b9b5-194">Метод</span><span class="sxs-lookup"><span data-stu-id="8b9b5-194">Method</span></span> |
| [<span data-ttu-id="8b9b5-195">getSelectedEntities</span><span class="sxs-lookup"><span data-stu-id="8b9b5-195">getSelectedEntities</span></span>](#getselectedentities--entitiesjavascriptapioutlookofficeentities) | <span data-ttu-id="8b9b5-196">Метод</span><span class="sxs-lookup"><span data-stu-id="8b9b5-196">Method</span></span> |
| [<span data-ttu-id="8b9b5-197">getSelectedRegExMatches</span><span class="sxs-lookup"><span data-stu-id="8b9b5-197">getSelectedRegExMatches</span></span>](#getselectedregexmatches--object) | <span data-ttu-id="8b9b5-198">Метод</span><span class="sxs-lookup"><span data-stu-id="8b9b5-198">Method</span></span> |
| [<span data-ttu-id="8b9b5-199">getSharedPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="8b9b5-199">getSharedPropertiesAsync</span></span>](#getsharedpropertiesasyncoptions-callback) | <span data-ttu-id="8b9b5-200">Метод</span><span class="sxs-lookup"><span data-stu-id="8b9b5-200">Method</span></span> |
| [<span data-ttu-id="8b9b5-201">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="8b9b5-201">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="8b9b5-202">Метод</span><span class="sxs-lookup"><span data-stu-id="8b9b5-202">Method</span></span> |
| [<span data-ttu-id="8b9b5-203">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="8b9b5-203">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="8b9b5-204">Метод</span><span class="sxs-lookup"><span data-stu-id="8b9b5-204">Method</span></span> |
| [<span data-ttu-id="8b9b5-205">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="8b9b5-205">removeHandlerAsync</span></span>](#removehandlerasynceventtype-handler-options-callback) | <span data-ttu-id="8b9b5-206">Метод</span><span class="sxs-lookup"><span data-stu-id="8b9b5-206">Method</span></span> |
| [<span data-ttu-id="8b9b5-207">saveAsync</span><span class="sxs-lookup"><span data-stu-id="8b9b5-207">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="8b9b5-208">Метод</span><span class="sxs-lookup"><span data-stu-id="8b9b5-208">Method</span></span> |
| [<span data-ttu-id="8b9b5-209">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="8b9b5-209">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="8b9b5-210">Метод</span><span class="sxs-lookup"><span data-stu-id="8b9b5-210">Method</span></span> |

### <a name="example"></a><span data-ttu-id="8b9b5-211">Пример</span><span class="sxs-lookup"><span data-stu-id="8b9b5-211">Example</span></span>

<span data-ttu-id="8b9b5-212">В приведенном ниже примере кода JavaScript показано, как получить доступ к свойству `subject` текущего элемента в Outlook.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-212">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="8b9b5-213">Члены</span><span class="sxs-lookup"><span data-stu-id="8b9b5-213">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails"></a><span data-ttu-id="8b9b5-214">attachments :Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="8b9b5-214">attachments :Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

<span data-ttu-id="8b9b5-p102">Получает массив вложений для элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="8b9b5-217">Некоторые типы файлов блокируются Outlook из-за потенциальных проблем безопасности и поэтому не возвращаются.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-217">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="8b9b5-218">Дополнительные сведения см. в статье [Блокированные вложения в Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="8b9b5-218">For more information see [Payments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="8b9b5-219">Тип:</span><span class="sxs-lookup"><span data-stu-id="8b9b5-219">Type:</span></span>

*   <span data-ttu-id="8b9b5-220">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="8b9b5-220">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="8b9b5-221">Требования</span><span class="sxs-lookup"><span data-stu-id="8b9b5-221">Requirements</span></span>

|<span data-ttu-id="8b9b5-222">Требование</span><span class="sxs-lookup"><span data-stu-id="8b9b5-222">Requirement</span></span>|<span data-ttu-id="8b9b5-223">Значение</span><span class="sxs-lookup"><span data-stu-id="8b9b5-223">Value</span></span>|
|---|---|
|[<span data-ttu-id="8b9b5-224">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="8b9b5-224">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8b9b5-225">1.0</span><span class="sxs-lookup"><span data-stu-id="8b9b5-225">1.0</span></span>|
|[<span data-ttu-id="8b9b5-226">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="8b9b5-226">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8b9b5-227">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8b9b5-227">ReadItem</span></span>|
|[<span data-ttu-id="8b9b5-228">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="8b9b5-228">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8b9b5-229">Чтение</span><span class="sxs-lookup"><span data-stu-id="8b9b5-229">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8b9b5-230">Пример</span><span class="sxs-lookup"><span data-stu-id="8b9b5-230">Example</span></span>

<span data-ttu-id="8b9b5-231">С помощью приведенного ниже кода можно создать HTML-строку с подробными сведениями обо всех вложениях для текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-231">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

####  <a name="bcc-recipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="8b9b5-232">bcc :[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="8b9b5-232">bcc :[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="8b9b5-233">Получает объект, который предоставляет методы для получения или обновления получателей в строке Bcc (скрытой копии) сообщения.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-233">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="8b9b5-234">Только в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-234">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="8b9b5-235">Тип:</span><span class="sxs-lookup"><span data-stu-id="8b9b5-235">Type:</span></span>

*   [<span data-ttu-id="8b9b5-236">Recipients</span><span class="sxs-lookup"><span data-stu-id="8b9b5-236">Recipients</span></span>](/javascript/api/outlook/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="8b9b5-237">Требования</span><span class="sxs-lookup"><span data-stu-id="8b9b5-237">Requirements</span></span>

|<span data-ttu-id="8b9b5-238">Требование</span><span class="sxs-lookup"><span data-stu-id="8b9b5-238">Requirement</span></span>|<span data-ttu-id="8b9b5-239">Значение</span><span class="sxs-lookup"><span data-stu-id="8b9b5-239">Value</span></span>|
|---|---|
|[<span data-ttu-id="8b9b5-240">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="8b9b5-240">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8b9b5-241">1.1</span><span class="sxs-lookup"><span data-stu-id="8b9b5-241">1.1</span></span>|
|[<span data-ttu-id="8b9b5-242">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="8b9b5-242">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8b9b5-243">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8b9b5-243">ReadItem</span></span>|
|[<span data-ttu-id="8b9b5-244">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="8b9b5-244">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8b9b5-245">Создание</span><span class="sxs-lookup"><span data-stu-id="8b9b5-245">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="8b9b5-246">Пример</span><span class="sxs-lookup"><span data-stu-id="8b9b5-246">Example</span></span>

```
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlookofficebody"></a><span data-ttu-id="8b9b5-247">body :[Body](/javascript/api/outlook/office.body)</span><span class="sxs-lookup"><span data-stu-id="8b9b5-247">body :[Body](/javascript/api/outlook/office.body)</span></span>

<span data-ttu-id="8b9b5-248">Получает объект, предоставляющий методы для работы с основным текстом элемента.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-248">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="8b9b5-249">Тип:</span><span class="sxs-lookup"><span data-stu-id="8b9b5-249">Type:</span></span>

*   [<span data-ttu-id="8b9b5-250">Body</span><span class="sxs-lookup"><span data-stu-id="8b9b5-250">Body</span></span>](/javascript/api/outlook/office.body)

##### <a name="requirements"></a><span data-ttu-id="8b9b5-251">Требования</span><span class="sxs-lookup"><span data-stu-id="8b9b5-251">Requirements</span></span>

|<span data-ttu-id="8b9b5-252">Требование</span><span class="sxs-lookup"><span data-stu-id="8b9b5-252">Requirement</span></span>|<span data-ttu-id="8b9b5-253">Значение</span><span class="sxs-lookup"><span data-stu-id="8b9b5-253">Value</span></span>|
|---|---|
|[<span data-ttu-id="8b9b5-254">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="8b9b5-254">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8b9b5-255">1.1</span><span class="sxs-lookup"><span data-stu-id="8b9b5-255">1.1</span></span>|
|[<span data-ttu-id="8b9b5-256">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="8b9b5-256">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8b9b5-257">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8b9b5-257">ReadItem</span></span>|
|[<span data-ttu-id="8b9b5-258">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="8b9b5-258">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8b9b5-259">Cоздание или чтение</span><span class="sxs-lookup"><span data-stu-id="8b9b5-259">Compose or read</span></span>|

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="8b9b5-260">cc :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="8b9b5-260">cc :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="8b9b5-261">Предоставляет доступ к получателям Cc (копии) сообщения.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-261">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="8b9b5-262">Уровень доступа и тип объекта, зависит от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-262">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="8b9b5-263">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="8b9b5-263">Read mode</span></span>

<span data-ttu-id="8b9b5-p106">Свойство `cc` возвращает массив, который содержит объект `EmailAddressDetails`, каждому получателю, указанному в строке **Cc (копия)** сообщения. Коллекция может включать не более 100 членов.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="8b9b5-266">Режим создания</span><span class="sxs-lookup"><span data-stu-id="8b9b5-266">Compose mode</span></span>

<span data-ttu-id="8b9b5-267">Свойство `cc` возвращает объект `Recipients`, предоставляющий методы для получения или обновления получателей, которые указаны в строке **Cc (копия)** сообщения.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-267">The `cc` property returns a `Recipients` object that provides methods for manipulating the recipients on the **Cc** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="8b9b5-268">Тип:</span><span class="sxs-lookup"><span data-stu-id="8b9b5-268">Type:</span></span>

*   <span data-ttu-id="8b9b5-269">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="8b9b5-269">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="8b9b5-270">Требования</span><span class="sxs-lookup"><span data-stu-id="8b9b5-270">Requirements</span></span>

|<span data-ttu-id="8b9b5-271">Требование</span><span class="sxs-lookup"><span data-stu-id="8b9b5-271">Requirement</span></span>|<span data-ttu-id="8b9b5-272">Значение</span><span class="sxs-lookup"><span data-stu-id="8b9b5-272">Value</span></span>|
|---|---|
|[<span data-ttu-id="8b9b5-273">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="8b9b5-273">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8b9b5-274">1.0</span><span class="sxs-lookup"><span data-stu-id="8b9b5-274">1.0</span></span>|
|[<span data-ttu-id="8b9b5-275">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="8b9b5-275">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8b9b5-276">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8b9b5-276">ReadItem</span></span>|
|[<span data-ttu-id="8b9b5-277">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="8b9b5-277">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8b9b5-278">Cоздание или чтение</span><span class="sxs-lookup"><span data-stu-id="8b9b5-278">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="8b9b5-279">Пример</span><span class="sxs-lookup"><span data-stu-id="8b9b5-279">Example</span></span>

```
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="8b9b5-280">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="8b9b5-280">(nullable) conversationId :String</span></span>

<span data-ttu-id="8b9b5-281">Получает идентификатор разговора по электронной почте, содержащего конкретное сообщение.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-281">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="8b9b5-p107">Вы можете получить целочисленное значение этого свойства, если ваше почтовое приложение активируется в формах просмотра или формах создания ответов. Если пользователь в свою очередь изменит тему ответа, после его отправки идентификатор беседы будет изменен, и полученное ранее значение будет недействительным.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="8b9b5-p108">Это свойство имеет значение NULL для нового элемента в форме создания. Свойство `conversationId` вернет значение, если пользователь задаст тему и сохранит элемент.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="8b9b5-286">Тип:</span><span class="sxs-lookup"><span data-stu-id="8b9b5-286">Type:</span></span>

*   <span data-ttu-id="8b9b5-287">String</span><span class="sxs-lookup"><span data-stu-id="8b9b5-287">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="8b9b5-288">Требования</span><span class="sxs-lookup"><span data-stu-id="8b9b5-288">Requirements</span></span>

|<span data-ttu-id="8b9b5-289">Требование</span><span class="sxs-lookup"><span data-stu-id="8b9b5-289">Requirement</span></span>|<span data-ttu-id="8b9b5-290">Значение</span><span class="sxs-lookup"><span data-stu-id="8b9b5-290">Value</span></span>|
|---|---|
|[<span data-ttu-id="8b9b5-291">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="8b9b5-291">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8b9b5-292">1.0</span><span class="sxs-lookup"><span data-stu-id="8b9b5-292">1.0</span></span>|
|[<span data-ttu-id="8b9b5-293">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="8b9b5-293">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8b9b5-294">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8b9b5-294">ReadItem</span></span>|
|[<span data-ttu-id="8b9b5-295">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="8b9b5-295">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8b9b5-296">Cоздание или чтение</span><span class="sxs-lookup"><span data-stu-id="8b9b5-296">Compose or read</span></span>|

#### <a name="datetimecreated-date"></a><span data-ttu-id="8b9b5-297">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="8b9b5-297">dateTimeCreated :Date</span></span>

<span data-ttu-id="8b9b5-p109">Получает дату и время создания элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="8b9b5-300">Тип:</span><span class="sxs-lookup"><span data-stu-id="8b9b5-300">Type:</span></span>

*   <span data-ttu-id="8b9b5-301">Date</span><span class="sxs-lookup"><span data-stu-id="8b9b5-301">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="8b9b5-302">Требования</span><span class="sxs-lookup"><span data-stu-id="8b9b5-302">Requirements</span></span>

|<span data-ttu-id="8b9b5-303">Требование</span><span class="sxs-lookup"><span data-stu-id="8b9b5-303">Requirement</span></span>|<span data-ttu-id="8b9b5-304">Значение</span><span class="sxs-lookup"><span data-stu-id="8b9b5-304">Value</span></span>|
|---|---|
|[<span data-ttu-id="8b9b5-305">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="8b9b5-305">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8b9b5-306">1.0</span><span class="sxs-lookup"><span data-stu-id="8b9b5-306">1.0</span></span>|
|[<span data-ttu-id="8b9b5-307">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="8b9b5-307">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8b9b5-308">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8b9b5-308">ReadItem</span></span>|
|[<span data-ttu-id="8b9b5-309">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="8b9b5-309">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8b9b5-310">Чтение</span><span class="sxs-lookup"><span data-stu-id="8b9b5-310">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8b9b5-311">Пример</span><span class="sxs-lookup"><span data-stu-id="8b9b5-311">Example</span></span>

```
var created = Office.context.mailbox.item.dateTimeCreated;
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="8b9b5-312">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="8b9b5-312">dateTimeModified :Date</span></span>

<span data-ttu-id="8b9b5-p110">Получает дату и время последнего изменения элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="8b9b5-315">Этот элемент не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-315">Note: This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="8b9b5-316">Тип:</span><span class="sxs-lookup"><span data-stu-id="8b9b5-316">Type:</span></span>

*   <span data-ttu-id="8b9b5-317">Date</span><span class="sxs-lookup"><span data-stu-id="8b9b5-317">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="8b9b5-318">Требования</span><span class="sxs-lookup"><span data-stu-id="8b9b5-318">Requirements</span></span>

|<span data-ttu-id="8b9b5-319">Требование</span><span class="sxs-lookup"><span data-stu-id="8b9b5-319">Requirement</span></span>|<span data-ttu-id="8b9b5-320">Значение</span><span class="sxs-lookup"><span data-stu-id="8b9b5-320">Value</span></span>|
|---|---|
|[<span data-ttu-id="8b9b5-321">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="8b9b5-321">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8b9b5-322">1.0</span><span class="sxs-lookup"><span data-stu-id="8b9b5-322">1.0</span></span>|
|[<span data-ttu-id="8b9b5-323">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="8b9b5-323">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8b9b5-324">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8b9b5-324">ReadItem</span></span>|
|[<span data-ttu-id="8b9b5-325">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="8b9b5-325">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8b9b5-326">Чтение</span><span class="sxs-lookup"><span data-stu-id="8b9b5-326">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8b9b5-327">Пример</span><span class="sxs-lookup"><span data-stu-id="8b9b5-327">Example</span></span>

```
var modified = Office.context.mailbox.item.dateTimeModified;
```

####  <a name="end-datetimejavascriptapioutlookofficetime"></a><span data-ttu-id="8b9b5-328">end :Date|[Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="8b9b5-328">end :Date|[Time](/javascript/api/outlook/office.time)</span></span>

<span data-ttu-id="8b9b5-329">Получает или задает дату и время окончания встречи.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-329">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="8b9b5-p111">Свойство `end` представлено в виде значения даты и времени в формате UTC. Преобразовать значение свойства end в местные значения даты и времени клиента можно с помощью метода [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime).</span><span class="sxs-lookup"><span data-stu-id="8b9b5-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="8b9b5-332">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="8b9b5-332">Read mode</span></span>

<span data-ttu-id="8b9b5-333">Свойство `end` возвращает объект `Date`.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-333">The `end` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="8b9b5-334">Режим создания</span><span class="sxs-lookup"><span data-stu-id="8b9b5-334">Compose mode</span></span>

<span data-ttu-id="8b9b5-335">Свойство `end` возвращает объект `Time`.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-335">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="8b9b5-336">Когда вы используете метод [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) для того, чтобы задать время окончания, вы должны использовать метод [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) , чтобы преобразовать местное время на клиенте в формат UTC.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-336">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="8b9b5-337">Тип:</span><span class="sxs-lookup"><span data-stu-id="8b9b5-337">Type:</span></span>

*   <span data-ttu-id="8b9b5-338">Date | [Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="8b9b5-338">Date | [Time](/javascript/api/outlook/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="8b9b5-339">Требования</span><span class="sxs-lookup"><span data-stu-id="8b9b5-339">Requirements</span></span>

|<span data-ttu-id="8b9b5-340">Требование</span><span class="sxs-lookup"><span data-stu-id="8b9b5-340">Requirement</span></span>|<span data-ttu-id="8b9b5-341">Значение</span><span class="sxs-lookup"><span data-stu-id="8b9b5-341">Value</span></span>|
|---|---|
|[<span data-ttu-id="8b9b5-342">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="8b9b5-342">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8b9b5-343">1.0</span><span class="sxs-lookup"><span data-stu-id="8b9b5-343">1.0</span></span>|
|[<span data-ttu-id="8b9b5-344">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="8b9b5-344">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8b9b5-345">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8b9b5-345">ReadItem</span></span>|
|[<span data-ttu-id="8b9b5-346">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="8b9b5-346">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8b9b5-347">Cоздание или чтение</span><span class="sxs-lookup"><span data-stu-id="8b9b5-347">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="8b9b5-348">Пример</span><span class="sxs-lookup"><span data-stu-id="8b9b5-348">Example</span></span>

<span data-ttu-id="8b9b5-349">В примере ниже показано, как с помощью метода [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) объекта `Time` задать время окончания встречи в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-349">The following example sets the end time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom"></a><span data-ttu-id="8b9b5-350">from :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[From](/javascript/api/outlook/office.from)</span><span class="sxs-lookup"><span data-stu-id="8b9b5-350">from :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[From](/javascript/api/outlook/office.from)</span></span>

<span data-ttu-id="8b9b5-351">Получает адрес электронной почты отправителя сообщения.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-351">Gets the email address of the sender of a message.</span></span>

<span data-ttu-id="8b9b5-p112">Свойства `from` и [`sender`](#sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails) представляют одно лицо, пока сообщение не будет отправлено представителем. В противном случае свойство `from` представляет лицо, делегировавшее полномочия, а свойство sender — представителя.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-p112">The `from` and [`sender`](#sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="8b9b5-354">Свойство `recipientType` объекта `EmailAddressDetails` в свойстве `from` — `undefined`.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-354">Note: The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="8b9b5-355">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="8b9b5-355">Read mode</span></span>

<span data-ttu-id="8b9b5-356">Свойство `from` возвращает объект `EmailAddressDetails`.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-356">The `from` property returns a `EmailAddressDetails` object.</span></span>

```
var subject = Office.context.mailbox.item.from;
```

##### <a name="compose-mode"></a><span data-ttu-id="8b9b5-357">Режим создания</span><span class="sxs-lookup"><span data-stu-id="8b9b5-357">Compose mode</span></span>

<span data-ttu-id="8b9b5-358">Свойство `from` возвращает объект `From`, который обеспечивает метод получения объекта из значения.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-358">The `from` property returns a `From` object that provides a method to get the from value.</span></span>

```
Office.context.mailbox.item.from.getAsync(callback);

function callback(asyncResult) {
  var from = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="8b9b5-359">Тип:</span><span class="sxs-lookup"><span data-stu-id="8b9b5-359">Type:</span></span>

*   <span data-ttu-id="8b9b5-360">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [From](/javascript/api/outlook/office.from)</span><span class="sxs-lookup"><span data-stu-id="8b9b5-360">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [From](/javascript/api/outlook/office.from)</span></span>

##### <a name="requirements"></a><span data-ttu-id="8b9b5-361">Требования</span><span class="sxs-lookup"><span data-stu-id="8b9b5-361">Requirements</span></span>

|<span data-ttu-id="8b9b5-362">Требование</span><span class="sxs-lookup"><span data-stu-id="8b9b5-362">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="8b9b5-363">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="8b9b5-363">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8b9b5-364">1.0</span><span class="sxs-lookup"><span data-stu-id="8b9b5-364">1.0</span></span>|<span data-ttu-id="8b9b5-365">1.7</span><span class="sxs-lookup"><span data-stu-id="8b9b5-365">17 </span></span>|
|[<span data-ttu-id="8b9b5-366">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="8b9b5-366">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8b9b5-367">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8b9b5-367">ReadItem</span></span>|<span data-ttu-id="8b9b5-368">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="8b9b5-368">ReadWriteItem</span></span>|
|[<span data-ttu-id="8b9b5-369">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="8b9b5-369">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8b9b5-370">Чтение</span><span class="sxs-lookup"><span data-stu-id="8b9b5-370">Read</span></span>|<span data-ttu-id="8b9b5-371">Создание</span><span class="sxs-lookup"><span data-stu-id="8b9b5-371">Compose</span></span>|

#### <a name="internetmessageid-string"></a><span data-ttu-id="8b9b5-372">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="8b9b5-372">internetMessageId :String</span></span>

<span data-ttu-id="8b9b5-p113">Получает идентификатор интернет-сообщения для электронного сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-p113">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="8b9b5-375">Тип:</span><span class="sxs-lookup"><span data-stu-id="8b9b5-375">Type:</span></span>

*   <span data-ttu-id="8b9b5-376">String</span><span class="sxs-lookup"><span data-stu-id="8b9b5-376">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="8b9b5-377">Требования</span><span class="sxs-lookup"><span data-stu-id="8b9b5-377">Requirements</span></span>

|<span data-ttu-id="8b9b5-378">Требование</span><span class="sxs-lookup"><span data-stu-id="8b9b5-378">Requirement</span></span>|<span data-ttu-id="8b9b5-379">Значение</span><span class="sxs-lookup"><span data-stu-id="8b9b5-379">Value</span></span>|
|---|---|
|[<span data-ttu-id="8b9b5-380">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="8b9b5-380">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8b9b5-381">1.0</span><span class="sxs-lookup"><span data-stu-id="8b9b5-381">1.0</span></span>|
|[<span data-ttu-id="8b9b5-382">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="8b9b5-382">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8b9b5-383">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8b9b5-383">ReadItem</span></span>|
|[<span data-ttu-id="8b9b5-384">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="8b9b5-384">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8b9b5-385">Чтение</span><span class="sxs-lookup"><span data-stu-id="8b9b5-385">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8b9b5-386">Пример</span><span class="sxs-lookup"><span data-stu-id="8b9b5-386">Example</span></span>

```
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="8b9b5-387">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="8b9b5-387">itemClass :String</span></span>

<span data-ttu-id="8b9b5-p114">Получает класс элемента веб-служб Exchange для выбранного элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-p114">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="8b9b5-p115">Свойство `itemClass` указывает класс сообщения выбранного элемента. Ниже приводятся классы сообщения по умолчанию для элемента сообщения или встречи.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-p115">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

|<span data-ttu-id="8b9b5-392">Тип</span><span class="sxs-lookup"><span data-stu-id="8b9b5-392">Type</span></span>|<span data-ttu-id="8b9b5-393">Описание</span><span class="sxs-lookup"><span data-stu-id="8b9b5-393">Description</span></span>|<span data-ttu-id="8b9b5-394">item class</span><span class="sxs-lookup"><span data-stu-id="8b9b5-394">item class</span></span>|
|---|---|---|
|<span data-ttu-id="8b9b5-395">Элементы встречи</span><span class="sxs-lookup"><span data-stu-id="8b9b5-395">Appointment items</span></span>|<span data-ttu-id="8b9b5-396">Это элементы календаря для класса элемента `IPM.Appointment` или `IPM.Appointment.Occurence`.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-396">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurence`.</span></span>|`IPM.Appointment`<br />`IPM.Appointment.Occurence`|
|<span data-ttu-id="8b9b5-397">Элементы сообщения</span><span class="sxs-lookup"><span data-stu-id="8b9b5-397">Message items</span></span>|<span data-ttu-id="8b9b5-398">Сюда входят электронные сообщения, для которых по умолчанию задан класс сообщения `IPM.Note`, а также приглашения на собрания, ответы на них и уведомления об их отмене, использующие `IPM.Schedule.Meeting` в качестве базового класса сообщений.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-398">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span>|`IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled`|

<span data-ttu-id="8b9b5-399">Вы можете создавать настраиваемые классы сообщения, расширяющие классы сообщения по умолчанию, например, настраиваемый класс сообщения о встрече `IPM.Appointment.Contoso`.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-399">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="8b9b5-400">Тип:</span><span class="sxs-lookup"><span data-stu-id="8b9b5-400">Type:</span></span>

*   <span data-ttu-id="8b9b5-401">String</span><span class="sxs-lookup"><span data-stu-id="8b9b5-401">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="8b9b5-402">Требования</span><span class="sxs-lookup"><span data-stu-id="8b9b5-402">Requirements</span></span>

|<span data-ttu-id="8b9b5-403">Требование</span><span class="sxs-lookup"><span data-stu-id="8b9b5-403">Requirement</span></span>|<span data-ttu-id="8b9b5-404">Значение</span><span class="sxs-lookup"><span data-stu-id="8b9b5-404">Value</span></span>|
|---|---|
|[<span data-ttu-id="8b9b5-405">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="8b9b5-405">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8b9b5-406">1.0</span><span class="sxs-lookup"><span data-stu-id="8b9b5-406">1.0</span></span>|
|[<span data-ttu-id="8b9b5-407">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="8b9b5-407">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8b9b5-408">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8b9b5-408">ReadItem</span></span>|
|[<span data-ttu-id="8b9b5-409">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="8b9b5-409">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8b9b5-410">Чтение</span><span class="sxs-lookup"><span data-stu-id="8b9b5-410">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8b9b5-411">Пример</span><span class="sxs-lookup"><span data-stu-id="8b9b5-411">Example</span></span>

```
var itemClass = Office.context.mailbox.item.itemClass;
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="8b9b5-412">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="8b9b5-412">(nullable) itemId :String</span></span>

<span data-ttu-id="8b9b5-p116">Получает идентификатор элемента веб-служб Exchange для текущего элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-p116">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="8b9b5-415">Идентификатор, возвращаемый свойством `itemId`, совпадает с идентификатором элемента веб-служб Exchange.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-415">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="8b9b5-416">Свойство  `itemId` не совпадает с идентификатором записи Outlook или идентификатором, используемым API-Интерфейсом REST Outlook.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-416">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="8b9b5-417">Прежде чем осуществлять вызовы API-Интерфейса REST с помощью этого значения, его следует преобразовать с помощью [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="8b9b5-417">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="8b9b5-418">Дополнительные сведения см. в статье [Использование API REST для Outlook из надстройки Outlook](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="8b9b5-418">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="8b9b5-p118">Свойство `itemId` недоступно в режиме создания. Если требуется идентификатор элемента, с помощью метода [`saveAsync`](#saveasyncoptions-callback) можно сохранить элемент в хранилище. При этом в параметре [`AsyncResult.value`](/javascript/api/office/office.asyncresult) функции обратного вызова возвращается идентификатор элемента.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-p118">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="8b9b5-421">Тип:</span><span class="sxs-lookup"><span data-stu-id="8b9b5-421">Type:</span></span>

*   <span data-ttu-id="8b9b5-422">String</span><span class="sxs-lookup"><span data-stu-id="8b9b5-422">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="8b9b5-423">Требования</span><span class="sxs-lookup"><span data-stu-id="8b9b5-423">Requirements</span></span>

|<span data-ttu-id="8b9b5-424">Требование</span><span class="sxs-lookup"><span data-stu-id="8b9b5-424">Requirement</span></span>|<span data-ttu-id="8b9b5-425">Значение</span><span class="sxs-lookup"><span data-stu-id="8b9b5-425">Value</span></span>|
|---|---|
|[<span data-ttu-id="8b9b5-426">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="8b9b5-426">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8b9b5-427">1.0</span><span class="sxs-lookup"><span data-stu-id="8b9b5-427">1.0</span></span>|
|[<span data-ttu-id="8b9b5-428">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="8b9b5-428">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8b9b5-429">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8b9b5-429">ReadItem</span></span>|
|[<span data-ttu-id="8b9b5-430">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="8b9b5-430">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8b9b5-431">Чтение</span><span class="sxs-lookup"><span data-stu-id="8b9b5-431">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8b9b5-432">Пример</span><span class="sxs-lookup"><span data-stu-id="8b9b5-432">Example</span></span>

<span data-ttu-id="8b9b5-p119">Указанный ниже код проверяет наличие идентификатора элемента. Если свойство `itemId` возвращает значение `null` или `undefined`, элемент будет сохранен в хранилище, а из асинхронного результата будет получен идентификатор элемента.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-p119">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result){
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype"></a><span data-ttu-id="8b9b5-435">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="8b9b5-435">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="8b9b5-436">Получает тип элемента, который представляет экземпляр.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-436">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="8b9b5-437">Свойство `itemType` возвращает одно из значений перечисления `ItemType`, которое указывает, является ли экземпляр объекта `item` сообщением или встречей.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-437">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="8b9b5-438">Тип:</span><span class="sxs-lookup"><span data-stu-id="8b9b5-438">Type:</span></span>

*   [<span data-ttu-id="8b9b5-439">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="8b9b5-439">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="8b9b5-440">Требования</span><span class="sxs-lookup"><span data-stu-id="8b9b5-440">Requirements</span></span>

|<span data-ttu-id="8b9b5-441">Требование</span><span class="sxs-lookup"><span data-stu-id="8b9b5-441">Requirement</span></span>|<span data-ttu-id="8b9b5-442">Значение</span><span class="sxs-lookup"><span data-stu-id="8b9b5-442">Value</span></span>|
|---|---|
|[<span data-ttu-id="8b9b5-443">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="8b9b5-443">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8b9b5-444">1.0</span><span class="sxs-lookup"><span data-stu-id="8b9b5-444">1.0</span></span>|
|[<span data-ttu-id="8b9b5-445">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="8b9b5-445">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8b9b5-446">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8b9b5-446">ReadItem</span></span>|
|[<span data-ttu-id="8b9b5-447">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="8b9b5-447">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8b9b5-448">Cоздание или чтение</span><span class="sxs-lookup"><span data-stu-id="8b9b5-448">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="8b9b5-449">Пример</span><span class="sxs-lookup"><span data-stu-id="8b9b5-449">Example</span></span>

```
if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Message)
  // do something
else
  // do something else
```

####  <a name="location-stringlocationjavascriptapioutlookofficelocation"></a><span data-ttu-id="8b9b5-450">location :String|[Location](/javascript/api/outlook/office.location)</span><span class="sxs-lookup"><span data-stu-id="8b9b5-450">location :String|[Location](/javascript/api/outlook/office.location)</span></span>

<span data-ttu-id="8b9b5-451">Получает или задает место встречи.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-451">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="8b9b5-452">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="8b9b5-452">Read mode</span></span>

<span data-ttu-id="8b9b5-453">Свойство `location` возвращает строку, содержащую сведения о месте встречи.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-453">The `location` property returns a string that contains the location of the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="8b9b5-454">Режим создания</span><span class="sxs-lookup"><span data-stu-id="8b9b5-454">Compose mode</span></span>

<span data-ttu-id="8b9b5-455">Свойство `location` возвращает объект `Location`, предоставляющий методы, которые используются для получения и задания места встречи.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-455">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="8b9b5-456">Тип:</span><span class="sxs-lookup"><span data-stu-id="8b9b5-456">Type:</span></span>

*   <span data-ttu-id="8b9b5-457">String | [Location](/javascript/api/outlook/office.location)</span><span class="sxs-lookup"><span data-stu-id="8b9b5-457">String | [Location](/javascript/api/outlook/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="8b9b5-458">Требования</span><span class="sxs-lookup"><span data-stu-id="8b9b5-458">Requirements</span></span>

|<span data-ttu-id="8b9b5-459">Требование</span><span class="sxs-lookup"><span data-stu-id="8b9b5-459">Requirement</span></span>|<span data-ttu-id="8b9b5-460">Значение</span><span class="sxs-lookup"><span data-stu-id="8b9b5-460">Value</span></span>|
|---|---|
|[<span data-ttu-id="8b9b5-461">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="8b9b5-461">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8b9b5-462">1.0</span><span class="sxs-lookup"><span data-stu-id="8b9b5-462">1.0</span></span>|
|[<span data-ttu-id="8b9b5-463">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="8b9b5-463">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8b9b5-464">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8b9b5-464">ReadItem</span></span>|
|[<span data-ttu-id="8b9b5-465">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="8b9b5-465">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8b9b5-466">Cоздание или чтение</span><span class="sxs-lookup"><span data-stu-id="8b9b5-466">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="8b9b5-467">Пример</span><span class="sxs-lookup"><span data-stu-id="8b9b5-467">Example</span></span>

```
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

#### <a name="normalizedsubject-string"></a><span data-ttu-id="8b9b5-468">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="8b9b5-468">normalizedSubject :String</span></span>

<span data-ttu-id="8b9b5-p120">Получает тему элемента со всеми удаленными префиксами (включая `RE:` и `FWD:`). Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-p120">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="8b9b5-p121">Свойство normalizedSubject получает тему элемента со стандартными префиксами (такими как `RE:` и `FW:`), добавляемыми почтовыми программами. Для получения темы элемента с неизмененными префиксами используйте свойство [`subject`](#subject-stringsubjectjavascriptapioutlookofficesubject).</span><span class="sxs-lookup"><span data-stu-id="8b9b5-p121">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubjectjavascriptapioutlookofficesubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="8b9b5-473">Тип:</span><span class="sxs-lookup"><span data-stu-id="8b9b5-473">Type:</span></span>

*   <span data-ttu-id="8b9b5-474">String</span><span class="sxs-lookup"><span data-stu-id="8b9b5-474">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="8b9b5-475">Требования</span><span class="sxs-lookup"><span data-stu-id="8b9b5-475">Requirements</span></span>

|<span data-ttu-id="8b9b5-476">Требование</span><span class="sxs-lookup"><span data-stu-id="8b9b5-476">Requirement</span></span>|<span data-ttu-id="8b9b5-477">Значение</span><span class="sxs-lookup"><span data-stu-id="8b9b5-477">Value</span></span>|
|---|---|
|[<span data-ttu-id="8b9b5-478">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="8b9b5-478">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8b9b5-479">1.0</span><span class="sxs-lookup"><span data-stu-id="8b9b5-479">1.0</span></span>|
|[<span data-ttu-id="8b9b5-480">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="8b9b5-480">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8b9b5-481">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8b9b5-481">ReadItem</span></span>|
|[<span data-ttu-id="8b9b5-482">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="8b9b5-482">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8b9b5-483">Чтение</span><span class="sxs-lookup"><span data-stu-id="8b9b5-483">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8b9b5-484">Пример</span><span class="sxs-lookup"><span data-stu-id="8b9b5-484">Example</span></span>

```
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
```

####  <a name="notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessages"></a><span data-ttu-id="8b9b5-485">notificationMessages :[NotificationMessages](/javascript/api/outlook/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="8b9b5-485">notificationMessages :[NotificationMessages](/javascript/api/outlook/office.notificationmessages)</span></span>

<span data-ttu-id="8b9b5-486">Получает сообщения уведомления для элемента.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-486">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="8b9b5-487">Тип:</span><span class="sxs-lookup"><span data-stu-id="8b9b5-487">Type:</span></span>

*   [<span data-ttu-id="8b9b5-488">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="8b9b5-488">NotificationMessages</span></span>](/javascript/api/outlook/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="8b9b5-489">Требования</span><span class="sxs-lookup"><span data-stu-id="8b9b5-489">Requirements</span></span>

|<span data-ttu-id="8b9b5-490">Требование</span><span class="sxs-lookup"><span data-stu-id="8b9b5-490">Requirement</span></span>|<span data-ttu-id="8b9b5-491">Значение</span><span class="sxs-lookup"><span data-stu-id="8b9b5-491">Value</span></span>|
|---|---|
|[<span data-ttu-id="8b9b5-492">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="8b9b5-492">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8b9b5-493">1.3</span><span class="sxs-lookup"><span data-stu-id="8b9b5-493">1.3</span></span>|
|[<span data-ttu-id="8b9b5-494">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="8b9b5-494">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8b9b5-495">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8b9b5-495">ReadItem</span></span>|
|[<span data-ttu-id="8b9b5-496">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="8b9b5-496">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8b9b5-497">Cоздание или чтение</span><span class="sxs-lookup"><span data-stu-id="8b9b5-497">Compose or read</span></span>|

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="8b9b5-498">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="8b9b5-498">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="8b9b5-499">Предоставляет доступ к необязательным участникам события.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-499">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="8b9b5-500">Уровень доступа и тип объекта, зависит от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-500">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="8b9b5-501">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="8b9b5-501">Read mode</span></span>

<span data-ttu-id="8b9b5-502">Свойство `optionalAttendees` возвращает массив, содержащий объект `EmailAddressDetails` для каждого необязательного участника собрания.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-502">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="8b9b5-503">Режим создания</span><span class="sxs-lookup"><span data-stu-id="8b9b5-503">Compose mode</span></span>

<span data-ttu-id="8b9b5-504">Свойство `optionalAttendees` возвращает объект `Recipients`, который предоставляет методы для получения и обновления необязательных участников собрания.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-504">The `optionalAttendees` property returns a `Recipients` object that provides methods to get and set the optional attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="8b9b5-505">Тип:</span><span class="sxs-lookup"><span data-stu-id="8b9b5-505">Type:</span></span>

*   <span data-ttu-id="8b9b5-506">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="8b9b5-506">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="8b9b5-507">Требования</span><span class="sxs-lookup"><span data-stu-id="8b9b5-507">Requirements</span></span>

|<span data-ttu-id="8b9b5-508">Требование</span><span class="sxs-lookup"><span data-stu-id="8b9b5-508">Requirement</span></span>|<span data-ttu-id="8b9b5-509">Значение</span><span class="sxs-lookup"><span data-stu-id="8b9b5-509">Value</span></span>|
|---|---|
|[<span data-ttu-id="8b9b5-510">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="8b9b5-510">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8b9b5-511">1.0</span><span class="sxs-lookup"><span data-stu-id="8b9b5-511">1.0</span></span>|
|[<span data-ttu-id="8b9b5-512">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="8b9b5-512">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8b9b5-513">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8b9b5-513">ReadItem</span></span>|
|[<span data-ttu-id="8b9b5-514">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="8b9b5-514">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8b9b5-515">Cоздание или чтение</span><span class="sxs-lookup"><span data-stu-id="8b9b5-515">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="8b9b5-516">Пример</span><span class="sxs-lookup"><span data-stu-id="8b9b5-516">Example</span></span>

```
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsorganizerjavascriptapioutlookofficeorganizer"></a><span data-ttu-id="8b9b5-517">organizer :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[Organizer](/javascript/api/outlook/office.organizer)</span><span class="sxs-lookup"><span data-stu-id="8b9b5-517">organizer :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[Organizer](/javascript/api/outlook/office.organizer)</span></span>

<span data-ttu-id="8b9b5-518">Получает адрес электронной почты организатора указанного собрания.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-518">Gets the email address of the meeting organizer for a specified meeting.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="8b9b5-519">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="8b9b5-519">Read mode</span></span>

<span data-ttu-id="8b9b5-520">Свойство `organizer` возвращает объект [EmailAddressDetails,](/javascript/api/outlook/office.emailaddressdetails) который представляет организатора собрания.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-520">The `organizer` property returns an [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) object that represents the meeting organizer.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="8b9b5-521">Режим создания</span><span class="sxs-lookup"><span data-stu-id="8b9b5-521">Compose mode</span></span>

<span data-ttu-id="8b9b5-522">Свойство `organizer` возвращает объект [Organizer](/javascript/api/outlook/office.organizer), который предоставляет метод для получения значения организатора.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-522">The `organizer` property returns an [Organizer](/javascript/api/outlook/office.organizer) object that provides a method to get the organizer value.</span></span>

##### <a name="type"></a><span data-ttu-id="8b9b5-523">Тип:</span><span class="sxs-lookup"><span data-stu-id="8b9b5-523">Type:</span></span>

*   <span data-ttu-id="8b9b5-524">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [Organizer](/javascript/api/outlook/office.organizer)</span><span class="sxs-lookup"><span data-stu-id="8b9b5-524">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [Organizer](/javascript/api/outlook/office.organizer)</span></span>

##### <a name="requirements"></a><span data-ttu-id="8b9b5-525">Требования</span><span class="sxs-lookup"><span data-stu-id="8b9b5-525">Requirements</span></span>

|<span data-ttu-id="8b9b5-526">Требование</span><span class="sxs-lookup"><span data-stu-id="8b9b5-526">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="8b9b5-527">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="8b9b5-527">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8b9b5-528">1.0</span><span class="sxs-lookup"><span data-stu-id="8b9b5-528">1.0</span></span>|<span data-ttu-id="8b9b5-529">1.7</span><span class="sxs-lookup"><span data-stu-id="8b9b5-529">17 </span></span>|
|[<span data-ttu-id="8b9b5-530">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="8b9b5-530">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8b9b5-531">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8b9b5-531">ReadItem</span></span>|<span data-ttu-id="8b9b5-532">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="8b9b5-532">ReadWriteItem</span></span>|
|[<span data-ttu-id="8b9b5-533">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="8b9b5-533">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8b9b5-534">Чтение</span><span class="sxs-lookup"><span data-stu-id="8b9b5-534">Read</span></span>|<span data-ttu-id="8b9b5-535">Создание</span><span class="sxs-lookup"><span data-stu-id="8b9b5-535">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="8b9b5-536">Пример</span><span class="sxs-lookup"><span data-stu-id="8b9b5-536">Example</span></span>

```
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
```

#### <a name="nullable-recurrence-recurrencejavascriptapioutlookofficerecurrence"></a><span data-ttu-id="8b9b5-537">(nullable) recurrence :[Recurrence](/javascript/api/outlook/office.recurrence)</span><span class="sxs-lookup"><span data-stu-id="8b9b5-537">(nullable) recurrence :[Recurrence](/javascript/api/outlook/office.recurrence)</span></span>

<span data-ttu-id="8b9b5-538">Получает или задает расписание повторения встречи.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-538">Gets or sets the location of an appointment.</span></span> <span data-ttu-id="8b9b5-539">Получает расписание повторения приглашения на собрание.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-539">Gets the recurrence pattern of a meeting request.</span></span> <span data-ttu-id="8b9b5-540">Чтение и создание режимов для элементов встречи.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-540">Read and compose modes for appointment items.</span></span> <span data-ttu-id="8b9b5-541">Режим чтения для элементов запроса на собрание.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-541">Read mode for meeting request items.</span></span>

<span data-ttu-id="8b9b5-542">Свойство `recurrence` возвращает объект [recurrence](/javascript/api/outlook/office.recurrence) для повторения запросов на встречи или собрания, если элемент или экземпляр являются серийными.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-542">The `recurrence` property returns a [recurrence](/javascript/api/outlook/office.recurrence) object for recurring appointments or meetings requests if an item is a series or an instance in a series.</span></span> <span data-ttu-id="8b9b5-543">`null` возвращается для одиночных встреч и запросов на собрания одиночных встреч.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-543">`null` is returned for single appointments and meeting requests of single appointments.</span></span> <span data-ttu-id="8b9b5-544">`undefined` возвращается для сообщений, которые не являются запросами на собрания.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-544">`undefined` is returned for messages that are not meeting requests.</span></span>

> <span data-ttu-id="8b9b5-545">Примечание: запросы на собрание имеют значение IPM.Schedule.Meeting.Request `itemClass`.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-545">Note: Meeting requests have an `itemClass` value of IPM.Schedule.Meeting.Request.</span></span>

> <span data-ttu-id="8b9b5-546">Примечание: если объектом повторения является `null`, это указывает на то, что объект является одиночной встречей или запросом на собрание одиночной встречи и НЕ является частью серии.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-546">Note: If the recurrence object is `null`, this indicates that the object is a single appointment or a meeting request of a single appointment and NOT a part of a series.</span></span>

##### <a name="type"></a><span data-ttu-id="8b9b5-547">Тип:</span><span class="sxs-lookup"><span data-stu-id="8b9b5-547">Type:</span></span>

* [<span data-ttu-id="8b9b5-548">Recurrence</span><span class="sxs-lookup"><span data-stu-id="8b9b5-548">recurrence</span></span>](/javascript/api/outlook/office.recurrence)

|<span data-ttu-id="8b9b5-549">Требование</span><span class="sxs-lookup"><span data-stu-id="8b9b5-549">Requirement</span></span>|<span data-ttu-id="8b9b5-550">Значение</span><span class="sxs-lookup"><span data-stu-id="8b9b5-550">Value</span></span>|
|---|---|
|[<span data-ttu-id="8b9b5-551">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="8b9b5-551">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8b9b5-552">1.7</span><span class="sxs-lookup"><span data-stu-id="8b9b5-552">17 </span></span>|
|[<span data-ttu-id="8b9b5-553">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="8b9b5-553">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8b9b5-554">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8b9b5-554">ReadItem</span></span>|
|[<span data-ttu-id="8b9b5-555">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="8b9b5-555">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8b9b5-556">Cоздание или чтение</span><span class="sxs-lookup"><span data-stu-id="8b9b5-556">Compose or read</span></span>|

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="8b9b5-557">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="8b9b5-557">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="8b9b5-558">Предоставляет доступ к обязательным участникам события.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-558">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="8b9b5-559">Уровень доступа и тип объекта, зависит от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-559">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="8b9b5-560">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="8b9b5-560">Read mode</span></span>

<span data-ttu-id="8b9b5-561">Свойство `requiredAttendees` возвращает массив, содержащий объект `EmailAddressDetails`, каждому обязательному участнику собрания.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-561">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="8b9b5-562">Режим создания</span><span class="sxs-lookup"><span data-stu-id="8b9b5-562">Compose mode</span></span>

<span data-ttu-id="8b9b5-563">Свойство `requiredAttendees` возвращает объект `Recipients`, который предоставляет методы для получения и обновления обязательных участников собрания.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-563">The `requiredAttendees` property returns a `Recipients` object that provides methods to get and set the required attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="8b9b5-564">Тип:</span><span class="sxs-lookup"><span data-stu-id="8b9b5-564">Type:</span></span>

*   <span data-ttu-id="8b9b5-565">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="8b9b5-565">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="8b9b5-566">Требования</span><span class="sxs-lookup"><span data-stu-id="8b9b5-566">Requirements</span></span>

|<span data-ttu-id="8b9b5-567">Требование</span><span class="sxs-lookup"><span data-stu-id="8b9b5-567">Requirement</span></span>|<span data-ttu-id="8b9b5-568">Значение</span><span class="sxs-lookup"><span data-stu-id="8b9b5-568">Value</span></span>|
|---|---|
|[<span data-ttu-id="8b9b5-569">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="8b9b5-569">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8b9b5-570">1.0</span><span class="sxs-lookup"><span data-stu-id="8b9b5-570">1.0</span></span>|
|[<span data-ttu-id="8b9b5-571">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="8b9b5-571">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8b9b5-572">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8b9b5-572">ReadItem</span></span>|
|[<span data-ttu-id="8b9b5-573">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="8b9b5-573">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8b9b5-574">Cоздание или чтение</span><span class="sxs-lookup"><span data-stu-id="8b9b5-574">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="8b9b5-575">Пример</span><span class="sxs-lookup"><span data-stu-id="8b9b5-575">Example</span></span>

```
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
}
```

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails"></a><span data-ttu-id="8b9b5-576">sender :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="8b9b5-576">sender :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)</span></span>

<span data-ttu-id="8b9b5-p126">Получает электронный адрес отправителя электронного сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="8b9b5-p127">Свойства [`from`](#from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom) и `sender` представляют одно и то же лицо, если сообщение не отправлено делегатом. В противном случае свойство `from` представляет лицо, делегировавшее полномочия, а свойство sender — делегата.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-p127">The [`from`](#from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="8b9b5-581">Свойство `recipientType` объекта `EmailAddressDetails` в свойстве `sender` — `undefined`.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-581">Note: The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="8b9b5-582">Тип:</span><span class="sxs-lookup"><span data-stu-id="8b9b5-582">Type:</span></span>

*   [<span data-ttu-id="8b9b5-583">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="8b9b5-583">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="8b9b5-584">Требования</span><span class="sxs-lookup"><span data-stu-id="8b9b5-584">Requirements</span></span>

|<span data-ttu-id="8b9b5-585">Требование</span><span class="sxs-lookup"><span data-stu-id="8b9b5-585">Requirement</span></span>|<span data-ttu-id="8b9b5-586">Значение</span><span class="sxs-lookup"><span data-stu-id="8b9b5-586">Value</span></span>|
|---|---|
|[<span data-ttu-id="8b9b5-587">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="8b9b5-587">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8b9b5-588">1.0</span><span class="sxs-lookup"><span data-stu-id="8b9b5-588">1.0</span></span>|
|[<span data-ttu-id="8b9b5-589">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="8b9b5-589">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8b9b5-590">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8b9b5-590">ReadItem</span></span>|
|[<span data-ttu-id="8b9b5-591">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="8b9b5-591">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8b9b5-592">Чтение</span><span class="sxs-lookup"><span data-stu-id="8b9b5-592">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8b9b5-593">Пример</span><span class="sxs-lookup"><span data-stu-id="8b9b5-593">Example</span></span>

```
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
```

#### <a name="nullable-seriesid-string"></a><span data-ttu-id="8b9b5-594">(nullable) seriesId :String</span><span class="sxs-lookup"><span data-stu-id="8b9b5-594">(nullable) seriesId :String</span></span>

<span data-ttu-id="8b9b5-595">Получает идентификатор серии, к которой принадлежит экземпляр.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-595">Gets the id of the series that an instance belongs to.</span></span>

<span data-ttu-id="8b9b5-596">В OWA и Outlook `seriesId` возвращает идентификатор веб-служб Exchange (EWS) родительского (серийного) элемента, к которому принадлежит этот элемент.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-596">In OWA and Outlook, the `seriesId` returns the Exchange Web Services (EWS) ID of the parent (series) item that this item belongs to.</span></span> <span data-ttu-id="8b9b5-597">Однако в iOS и Android `seriesId` возвращает REST идентификатор родительского элемента.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-597">However, in iOS and Android, the `seriesId` returns the REST ID of the parent item.</span></span>

> [!NOTE]
> <span data-ttu-id="8b9b5-598">Идентификатор, возвращаемый свойством `seriesId`, совпадает с идентификатором элемента веб-служб Exchange.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-598">The identifier returned by the `seriesId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="8b9b5-599">Свойство `seriesId` не идентично идентификаторам Outlook, используемым API-Интерфейсом REST Outlook.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-599">The `seriesId` property is not identical to the Outlook IDs used by the Outlook REST API.</span></span> <span data-ttu-id="8b9b5-600">Прежде чем осуществлять вызовы API-Интерфейса REST с помощью этого значения, его следует преобразовать с помощью [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="8b9b5-600">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="8b9b5-601">Для получения дополнительных сведений см. [Использование API REST Outlook из надстройки Outlook](https://docs.microsoft.com/outlook/add-ins/use-rest-api).</span><span class="sxs-lookup"><span data-stu-id="8b9b5-601">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api).</span></span>

<span data-ttu-id="8b9b5-602">Свойство `seriesId` возвращает `null` для элементов, у которых нет родительских элементов, таких как одиночные встречи, элементы серии или запросы на собрания и возвращает `undefined` для любых других элементов, которые не являются запросами на собрание.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-602">The `seriesId` property returns `null` for items that do not have parent items such as single appointments, series items, or meeting requests and returns `undefined` for any other items that are not meeting requests.</span></span>

##### <a name="type"></a><span data-ttu-id="8b9b5-603">Тип:</span><span class="sxs-lookup"><span data-stu-id="8b9b5-603">Type:</span></span>

* <span data-ttu-id="8b9b5-604">String</span><span class="sxs-lookup"><span data-stu-id="8b9b5-604">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="8b9b5-605">Требования</span><span class="sxs-lookup"><span data-stu-id="8b9b5-605">Requirements</span></span>

|<span data-ttu-id="8b9b5-606">Требование</span><span class="sxs-lookup"><span data-stu-id="8b9b5-606">Requirement</span></span>|<span data-ttu-id="8b9b5-607">Значение</span><span class="sxs-lookup"><span data-stu-id="8b9b5-607">Value</span></span>|
|---|---|
|[<span data-ttu-id="8b9b5-608">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="8b9b5-608">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8b9b5-609">1.7</span><span class="sxs-lookup"><span data-stu-id="8b9b5-609">17 </span></span>|
|[<span data-ttu-id="8b9b5-610">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="8b9b5-610">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8b9b5-611">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8b9b5-611">ReadItem</span></span>|
|[<span data-ttu-id="8b9b5-612">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="8b9b5-612">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8b9b5-613">Cоздание или чтение</span><span class="sxs-lookup"><span data-stu-id="8b9b5-613">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="8b9b5-614">Пример</span><span class="sxs-lookup"><span data-stu-id="8b9b5-614">Example</span></span>

```
var seriesId = Office.context.mailbox.item.seriesId;
var isSeries = (seriesId == null);
```

####  <a name="start-datetimejavascriptapioutlookofficetime"></a><span data-ttu-id="8b9b5-615">start :Date|[Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="8b9b5-615">start :Date|[Time](/javascript/api/outlook/office.time)</span></span>

<span data-ttu-id="8b9b5-616">Получает или задает дату и время начала встречи.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-616">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="8b9b5-p130">Свойство `start` представлено в виде значения даты и времени в формате UTC. Это значение можно преобразовать в местные значения даты и времени клиента с помощью метода [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime).</span><span class="sxs-lookup"><span data-stu-id="8b9b5-p130">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="8b9b5-619">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="8b9b5-619">Read mode</span></span>

<span data-ttu-id="8b9b5-620">Свойство `start` возвращает объект `Date`.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-620">The `start` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="8b9b5-621">Режим создания</span><span class="sxs-lookup"><span data-stu-id="8b9b5-621">Compose mode</span></span>

<span data-ttu-id="8b9b5-622">Свойство `start` возвращает объект `Time`.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-622">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="8b9b5-623">Если вы задаете время начала с помощью метода [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-), необходимо использовать метод [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) для преобразования местного времени на клиенте в формат UTC для сервера.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-623">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="8b9b5-624">Тип:</span><span class="sxs-lookup"><span data-stu-id="8b9b5-624">Type:</span></span>

*   <span data-ttu-id="8b9b5-625">Date | [Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="8b9b5-625">Date | [Time](/javascript/api/outlook/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="8b9b5-626">Требования</span><span class="sxs-lookup"><span data-stu-id="8b9b5-626">Requirements</span></span>

|<span data-ttu-id="8b9b5-627">Требование</span><span class="sxs-lookup"><span data-stu-id="8b9b5-627">Requirement</span></span>|<span data-ttu-id="8b9b5-628">Значение</span><span class="sxs-lookup"><span data-stu-id="8b9b5-628">Value</span></span>|
|---|---|
|[<span data-ttu-id="8b9b5-629">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="8b9b5-629">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8b9b5-630">1.0</span><span class="sxs-lookup"><span data-stu-id="8b9b5-630">1.0</span></span>|
|[<span data-ttu-id="8b9b5-631">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="8b9b5-631">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8b9b5-632">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8b9b5-632">ReadItem</span></span>|
|[<span data-ttu-id="8b9b5-633">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="8b9b5-633">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8b9b5-634">Cоздание или чтение</span><span class="sxs-lookup"><span data-stu-id="8b9b5-634">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="8b9b5-635">Пример</span><span class="sxs-lookup"><span data-stu-id="8b9b5-635">Example</span></span>

<span data-ttu-id="8b9b5-636">В примере ниже с помощью метода [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) объекта `Time` задается время начала встречи в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-636">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

####  <a name="subject-stringsubjectjavascriptapioutlookofficesubject"></a><span data-ttu-id="8b9b5-637">subject :String|[Subject](/javascript/api/outlook/office.subject)</span><span class="sxs-lookup"><span data-stu-id="8b9b5-637">subject :String|[Subject](/javascript/api/outlook/office.subject)</span></span>

<span data-ttu-id="8b9b5-638">Получает или задает описание, которое отображается в поле темы элемента.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-638">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="8b9b5-639">Свойство `subject` получает или задает всю тему элемента для отправки с почтового сервера.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-639">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="8b9b5-640">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="8b9b5-640">Read mode</span></span>

<span data-ttu-id="8b9b5-p131">Свойство `subject` возвращает строку. С помощью свойства [`normalizedSubject`](#normalizedsubject-string) можно получить тему без начальных префиксов, например, `RE:` и `FW:`.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-p131">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```
var subject = Office.context.mailbox.item.subject;
```

##### <a name="compose-mode"></a><span data-ttu-id="8b9b5-643">Режим создания</span><span class="sxs-lookup"><span data-stu-id="8b9b5-643">Compose mode</span></span>

<span data-ttu-id="8b9b5-644">Свойство `subject` возвращает объект `Subject`, который предоставляет методы для получения и задания темы.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-644">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="8b9b5-645">Тип:</span><span class="sxs-lookup"><span data-stu-id="8b9b5-645">Type:</span></span>

*   <span data-ttu-id="8b9b5-646">String | [Subject](/javascript/api/outlook/office.subject)</span><span class="sxs-lookup"><span data-stu-id="8b9b5-646">String | [Subject](/javascript/api/outlook/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="8b9b5-647">Требования</span><span class="sxs-lookup"><span data-stu-id="8b9b5-647">Requirements</span></span>

|<span data-ttu-id="8b9b5-648">Требование</span><span class="sxs-lookup"><span data-stu-id="8b9b5-648">Requirement</span></span>|<span data-ttu-id="8b9b5-649">Значение</span><span class="sxs-lookup"><span data-stu-id="8b9b5-649">Value</span></span>|
|---|---|
|[<span data-ttu-id="8b9b5-650">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="8b9b5-650">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8b9b5-651">1.0</span><span class="sxs-lookup"><span data-stu-id="8b9b5-651">1.0</span></span>|
|[<span data-ttu-id="8b9b5-652">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="8b9b5-652">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8b9b5-653">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8b9b5-653">ReadItem</span></span>|
|[<span data-ttu-id="8b9b5-654">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="8b9b5-654">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8b9b5-655">Cоздание или чтение</span><span class="sxs-lookup"><span data-stu-id="8b9b5-655">Compose or read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="8b9b5-656">to :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="8b9b5-656">to :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="8b9b5-657">Предоставляет доступ получателей к строке **To (Кому)** в сообщении.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-657">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="8b9b5-658">Уровень доступа и тип объекта, зависит от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-658">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="8b9b5-659">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="8b9b5-659">Read mode</span></span>

<span data-ttu-id="8b9b5-p133">Свойство `to` возвращает массив, содержащий объект `EmailAddressDetails` для каждого получателя в строке **To (Кому)** сообщения. Коллекция может включать не более 100 элементов.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-p133">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="8b9b5-662">Режим создания</span><span class="sxs-lookup"><span data-stu-id="8b9b5-662">Compose mode</span></span>

<span data-ttu-id="8b9b5-663">Свойство `to` возвращает объект `Recipients`, предоставляющий методы для получения или обновления получателей, которые указаны в строке **To (кому)** сообщения.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-663">The to`to` property returns a Recipients`Recipients` object that provides methods for manipulating the recipients on the **To** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="8b9b5-664">Тип:</span><span class="sxs-lookup"><span data-stu-id="8b9b5-664">Type:</span></span>

*   <span data-ttu-id="8b9b5-665">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="8b9b5-665">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="8b9b5-666">Требования</span><span class="sxs-lookup"><span data-stu-id="8b9b5-666">Requirements</span></span>

|<span data-ttu-id="8b9b5-667">Требование</span><span class="sxs-lookup"><span data-stu-id="8b9b5-667">Requirement</span></span>|<span data-ttu-id="8b9b5-668">Значение</span><span class="sxs-lookup"><span data-stu-id="8b9b5-668">Value</span></span>|
|---|---|
|[<span data-ttu-id="8b9b5-669">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="8b9b5-669">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8b9b5-670">1.0</span><span class="sxs-lookup"><span data-stu-id="8b9b5-670">1.0</span></span>|
|[<span data-ttu-id="8b9b5-671">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="8b9b5-671">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8b9b5-672">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8b9b5-672">ReadItem</span></span>|
|[<span data-ttu-id="8b9b5-673">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="8b9b5-673">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8b9b5-674">Cоздание или чтение</span><span class="sxs-lookup"><span data-stu-id="8b9b5-674">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="8b9b5-675">Пример</span><span class="sxs-lookup"><span data-stu-id="8b9b5-675">Example</span></span>

```
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

### <a name="methods"></a><span data-ttu-id="8b9b5-676">Методы</span><span class="sxs-lookup"><span data-stu-id="8b9b5-676">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="8b9b5-677">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="8b9b5-677">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="8b9b5-678">Добавляет файл в сообщение или встречу в качестве вложения.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-678">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="8b9b5-679">Метод `addFileAttachmentAsync` передает файл по указанному универсальному коду ресурса (URI) и вкладывает его в элемент в форме создания.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-679">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="8b9b5-680">Идентификатор можно последовательно использовать с методом [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback), чтобы удалить вложение, добавленное во время текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-680">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8b9b5-681">Параметры:</span><span class="sxs-lookup"><span data-stu-id="8b9b5-681">Parameters:</span></span>
|<span data-ttu-id="8b9b5-682">Имя</span><span class="sxs-lookup"><span data-stu-id="8b9b5-682">Name</span></span>|<span data-ttu-id="8b9b5-683">Тип</span><span class="sxs-lookup"><span data-stu-id="8b9b5-683">Type</span></span>|<span data-ttu-id="8b9b5-684">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="8b9b5-684">Attributes</span></span>|<span data-ttu-id="8b9b5-685">Описание</span><span class="sxs-lookup"><span data-stu-id="8b9b5-685">Description</span></span>|
|---|---|---|---|
|`uri`|<span data-ttu-id="8b9b5-686">String</span><span class="sxs-lookup"><span data-stu-id="8b9b5-686">String</span></span>||<span data-ttu-id="8b9b5-p134">URI-адрес, представляющий расположение файла, который нужно вложить в сообщение или встречу. Максимальная длина — 2048 символов.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-p134">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="8b9b5-689">String</span><span class="sxs-lookup"><span data-stu-id="8b9b5-689">String</span></span>||<span data-ttu-id="8b9b5-p135">Имя вложения, которое отображается при передаче вложения. Максимальная длина — 255 символов.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-p135">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="8b9b5-692">Объект</span><span class="sxs-lookup"><span data-stu-id="8b9b5-692">Object</span></span>|<span data-ttu-id="8b9b5-693">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="8b9b5-693">&lt;optional&gt;</span></span>|<span data-ttu-id="8b9b5-694">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-694">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="8b9b5-695">Объект</span><span class="sxs-lookup"><span data-stu-id="8b9b5-695">Object</span></span>|<span data-ttu-id="8b9b5-696">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="8b9b5-696">&lt;optional&gt;</span></span>|<span data-ttu-id="8b9b5-697">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-697">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="8b9b5-698">Boolean</span><span class="sxs-lookup"><span data-stu-id="8b9b5-698">Boolean</span></span>|<span data-ttu-id="8b9b5-699">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="8b9b5-699">&lt;optional&gt;</span></span>|<span data-ttu-id="8b9b5-700">Значение `true` указывает на то, что вложение будет встроено в текст сообщения и не должно отображаться в списке вложений.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-700">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="8b9b5-701">function</span><span class="sxs-lookup"><span data-stu-id="8b9b5-701">function</span></span>|<span data-ttu-id="8b9b5-702">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="8b9b5-702">&lt;optional&gt;</span></span>|<span data-ttu-id="8b9b5-703">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="8b9b5-703">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="8b9b5-704">После успешного выполнения идентификатор вложения будет представлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-704">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="8b9b5-705">Если добавить вложение не удастся, объект `asyncResult` будет содержать объект `Error` с описанием ошибки.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-705">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="8b9b5-706">Ошибки</span><span class="sxs-lookup"><span data-stu-id="8b9b5-706">Errors</span></span>

|<span data-ttu-id="8b9b5-707">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="8b9b5-707">Error code</span></span>|<span data-ttu-id="8b9b5-708">Описание</span><span class="sxs-lookup"><span data-stu-id="8b9b5-708">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="8b9b5-709">Вложение превышает максимальный размер.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-709">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="8b9b5-710">Расширение вложения не поддерживается.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-710">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="8b9b5-711">Сообщение или встреча содержат слишком много вложений.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-711">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8b9b5-712">Требования</span><span class="sxs-lookup"><span data-stu-id="8b9b5-712">Requirements</span></span>

|<span data-ttu-id="8b9b5-713">Требование</span><span class="sxs-lookup"><span data-stu-id="8b9b5-713">Requirement</span></span>|<span data-ttu-id="8b9b5-714">Значение</span><span class="sxs-lookup"><span data-stu-id="8b9b5-714">Value</span></span>|
|---|---|
|[<span data-ttu-id="8b9b5-715">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="8b9b5-715">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8b9b5-716">1.1</span><span class="sxs-lookup"><span data-stu-id="8b9b5-716">1.1</span></span>|
|[<span data-ttu-id="8b9b5-717">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="8b9b5-717">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8b9b5-718">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="8b9b5-718">ReadWriteItem</span></span>|
|[<span data-ttu-id="8b9b5-719">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="8b9b5-719">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8b9b5-720">Создание</span><span class="sxs-lookup"><span data-stu-id="8b9b5-720">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="8b9b5-721">Примеры</span><span class="sxs-lookup"><span data-stu-id="8b9b5-721">Examples</span></span>

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

<span data-ttu-id="8b9b5-722">В приведенном ниже примере файл изображения добавляется в качестве встроенного вложения, а ссылка на вложение добавляется в текст сообщения.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-722">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

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

#### <a name="addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback"></a><span data-ttu-id="8b9b5-723">addFileAttachmentFromBase64Async(base64File, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="8b9b5-723">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="8b9b5-724">Добавляет файл из кодирования  base64 в сообщение или встречу в виде вложения.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-724">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="8b9b5-725">Метод  `addFileAttachmentFromBase64Async` загружает файл из кодирования base64 и вкладывает его в элемент в форме создания.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-725">The `addFileAttachmentFromBase64Async` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span> <span data-ttu-id="8b9b5-726">Этот метод возвращает идентификатор вложения в объекте AsyncResult.value.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-726">This method returns the attachment identifier in the AsyncResult.value object.</span></span>

<span data-ttu-id="8b9b5-727">Идентификатор можно последовательно использовать с методом [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback), чтобы удалить вложение, добавленное во время текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-727">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8b9b5-728">Параметры:</span><span class="sxs-lookup"><span data-stu-id="8b9b5-728">Parameters:</span></span>
|<span data-ttu-id="8b9b5-729">Имя</span><span class="sxs-lookup"><span data-stu-id="8b9b5-729">Name</span></span>|<span data-ttu-id="8b9b5-730">Тип</span><span class="sxs-lookup"><span data-stu-id="8b9b5-730">Type</span></span>|<span data-ttu-id="8b9b5-731">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="8b9b5-731">Attributes</span></span>|<span data-ttu-id="8b9b5-732">Описание</span><span class="sxs-lookup"><span data-stu-id="8b9b5-732">Description</span></span>|
|---|---|---|---|
|`base64File`|<span data-ttu-id="8b9b5-733">String</span><span class="sxs-lookup"><span data-stu-id="8b9b5-733">String</span></span>||<span data-ttu-id="8b9b5-734">Содержимое в формате изображения или файла в сообщение или событие добавляется в кодировке base64.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-734">The base64 encoded content of an image or file to be added to an email or event.</span></span>|
|`attachmentName`|<span data-ttu-id="8b9b5-735">String</span><span class="sxs-lookup"><span data-stu-id="8b9b5-735">String</span></span>||<span data-ttu-id="8b9b5-p137">Имя вложения, которое отображается при передаче вложения. Максимальная длина — 255 символов.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-p137">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="8b9b5-738">Объект</span><span class="sxs-lookup"><span data-stu-id="8b9b5-738">Object</span></span>|<span data-ttu-id="8b9b5-739">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="8b9b5-739">&lt;optional&gt;</span></span>|<span data-ttu-id="8b9b5-740">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-740">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="8b9b5-741">Объект</span><span class="sxs-lookup"><span data-stu-id="8b9b5-741">Object</span></span>|<span data-ttu-id="8b9b5-742">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="8b9b5-742">&lt;optional&gt;</span></span>|<span data-ttu-id="8b9b5-743">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-743">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="8b9b5-744">Boolean</span><span class="sxs-lookup"><span data-stu-id="8b9b5-744">Boolean</span></span>|<span data-ttu-id="8b9b5-745">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="8b9b5-745">&lt;optional&gt;</span></span>|<span data-ttu-id="8b9b5-746">Значение `true` указывает на то, что вложение будет встроено в текст сообщения и не должно отображаться в списке вложений.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-746">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="8b9b5-747">function</span><span class="sxs-lookup"><span data-stu-id="8b9b5-747">function</span></span>|<span data-ttu-id="8b9b5-748">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="8b9b5-748">&lt;optional&gt;</span></span>|<span data-ttu-id="8b9b5-749">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="8b9b5-749">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="8b9b5-750">После успешного выполнения идентификатор вложения будет представлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-750">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="8b9b5-751">Если добавить вложение не удастся, объект `asyncResult` будет содержать объект `Error` с описанием ошибки.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-751">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="8b9b5-752">Ошибки</span><span class="sxs-lookup"><span data-stu-id="8b9b5-752">Errors</span></span>

|<span data-ttu-id="8b9b5-753">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="8b9b5-753">Error code</span></span>|<span data-ttu-id="8b9b5-754">Описание</span><span class="sxs-lookup"><span data-stu-id="8b9b5-754">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="8b9b5-755">Вложение превышает максимальный размер.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-755">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="8b9b5-756">Расширение вложения не поддерживается.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-756">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="8b9b5-757">Сообщение или встреча содержат слишком много вложений.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-757">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8b9b5-758">Требования</span><span class="sxs-lookup"><span data-stu-id="8b9b5-758">Requirements</span></span>

|<span data-ttu-id="8b9b5-759">Требование</span><span class="sxs-lookup"><span data-stu-id="8b9b5-759">Requirement</span></span>|<span data-ttu-id="8b9b5-760">Значение</span><span class="sxs-lookup"><span data-stu-id="8b9b5-760">Value</span></span>|
|---|---|
|[<span data-ttu-id="8b9b5-761">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="8b9b5-761">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8b9b5-762">Предварительная версия</span><span class="sxs-lookup"><span data-stu-id="8b9b5-762">Preview</span></span>|
|[<span data-ttu-id="8b9b5-763">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="8b9b5-763">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8b9b5-764">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="8b9b5-764">ReadWriteItem</span></span>|
|[<span data-ttu-id="8b9b5-765">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="8b9b5-765">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8b9b5-766">Создание</span><span class="sxs-lookup"><span data-stu-id="8b9b5-766">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="8b9b5-767">Примеры</span><span class="sxs-lookup"><span data-stu-id="8b9b5-767">Examples</span></span>

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

####  <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="8b9b5-768">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="8b9b5-768">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="8b9b5-769">Добавляет обработчик для поддерживаемого события.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-769">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="8b9b5-770">В настоящее время поддерживаемые типы событий — `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, и `Office.EventType.RecurrenceChanged`</span><span class="sxs-lookup"><span data-stu-id="8b9b5-770">Currently the supported event types are `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`</span></span>

##### <a name="parameters"></a><span data-ttu-id="8b9b5-771">Параметры:</span><span class="sxs-lookup"><span data-stu-id="8b9b5-771">Parameters:</span></span>

| <span data-ttu-id="8b9b5-772">Имя</span><span class="sxs-lookup"><span data-stu-id="8b9b5-772">Name</span></span> | <span data-ttu-id="8b9b5-773">Тип</span><span class="sxs-lookup"><span data-stu-id="8b9b5-773">Type</span></span> | <span data-ttu-id="8b9b5-774">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="8b9b5-774">Attributes</span></span> | <span data-ttu-id="8b9b5-775">Описание</span><span class="sxs-lookup"><span data-stu-id="8b9b5-775">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="8b9b5-776">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="8b9b5-776">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="8b9b5-777">Событие, которое должно вызвать обработчик.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-777">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="8b9b5-778">Функция</span><span class="sxs-lookup"><span data-stu-id="8b9b5-778">Function</span></span> || <span data-ttu-id="8b9b5-p138">Функция для обработки события. Функция должна принимать один параметр, представляющий собой объектный литерал. Значение свойства `type` параметра совпадет со значением параметра `eventType`, переданного методу `addHandlerAsync`.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-p138">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="8b9b5-782">Объект</span><span class="sxs-lookup"><span data-stu-id="8b9b5-782">Object</span></span> | <span data-ttu-id="8b9b5-783">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="8b9b5-783">&lt;optional&gt;</span></span> | <span data-ttu-id="8b9b5-784">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-784">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="8b9b5-785">Объект</span><span class="sxs-lookup"><span data-stu-id="8b9b5-785">Object</span></span> | <span data-ttu-id="8b9b5-786">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="8b9b5-786">&lt;optional&gt;</span></span> | <span data-ttu-id="8b9b5-787">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-787">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="8b9b5-788">function</span><span class="sxs-lookup"><span data-stu-id="8b9b5-788">function</span></span>| <span data-ttu-id="8b9b5-789">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="8b9b5-789">&lt;optional&gt;</span></span>|<span data-ttu-id="8b9b5-790">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="8b9b5-790">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8b9b5-791">Требования</span><span class="sxs-lookup"><span data-stu-id="8b9b5-791">Requirements</span></span>

|<span data-ttu-id="8b9b5-792">Требование</span><span class="sxs-lookup"><span data-stu-id="8b9b5-792">Requirement</span></span>| <span data-ttu-id="8b9b5-793">Значение</span><span class="sxs-lookup"><span data-stu-id="8b9b5-793">Value</span></span>|
|---|---|
|[<span data-ttu-id="8b9b5-794">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="8b9b5-794">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8b9b5-795">1.7</span><span class="sxs-lookup"><span data-stu-id="8b9b5-795">17 </span></span> |
|[<span data-ttu-id="8b9b5-796">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="8b9b5-796">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8b9b5-797">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8b9b5-797">ReadItem</span></span> |
|[<span data-ttu-id="8b9b5-798">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="8b9b5-798">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8b9b5-799">Cоздание или чтение</span><span class="sxs-lookup"><span data-stu-id="8b9b5-799">Compose or read</span></span> |

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="8b9b5-800">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="8b9b5-800">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="8b9b5-801">Добавляет к сообщению или встрече элемент Exchange (например, сообщение) в виде вложения.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-801">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="8b9b5-p139">С помощью метода `addItemAttachmentAsync` в элемент формы создания можно вложить элемент с указанным идентификатором Exchange. Если указать метод обратного вызова, то этот метод вызывается с помощью параметра `asyncResult`, который содержит идентификатор вложения или код, указывающий на ошибки, которые произошли при вложении элемента. При необходимости можно использовать параметр `options` для передачи сведений о состоянии в метод обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-p139">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="8b9b5-805">Идентификатор можно последовательно использовать с методом [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback), чтобы удалить вложение, добавленное во время текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-805">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="8b9b5-806">Если ваша надстройка Office выполняется в веб-приложении Outlook, метод `addItemAttachmentAsync` обеспечивает вложение элементов в элементы, отличные от редактируемого. Однако это действие не рекомендуется выполнять, так как оно не поддерживается.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-806">If your Office add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8b9b5-807">Параметры:</span><span class="sxs-lookup"><span data-stu-id="8b9b5-807">Parameters:</span></span>

|<span data-ttu-id="8b9b5-808">Имя</span><span class="sxs-lookup"><span data-stu-id="8b9b5-808">Name</span></span>|<span data-ttu-id="8b9b5-809">Тип</span><span class="sxs-lookup"><span data-stu-id="8b9b5-809">Type</span></span>|<span data-ttu-id="8b9b5-810">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="8b9b5-810">Attributes</span></span>|<span data-ttu-id="8b9b5-811">Описание</span><span class="sxs-lookup"><span data-stu-id="8b9b5-811">Description</span></span>|
|---|---|---|---|
|`itemId`|<span data-ttu-id="8b9b5-812">String</span><span class="sxs-lookup"><span data-stu-id="8b9b5-812">String</span></span>||<span data-ttu-id="8b9b5-p140">Идентификатор Exchange для вкладываемого элемента. Максимальная длина — 100 символов.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-p140">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="8b9b5-815">String</span><span class="sxs-lookup"><span data-stu-id="8b9b5-815">String</span></span>||<span data-ttu-id="8b9b5-p141">Тема вкладываемого элемента. Максимальная длина — 255 символов.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-p141">The sujbect of the item to be attached. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="8b9b5-818">Объект</span><span class="sxs-lookup"><span data-stu-id="8b9b5-818">Object</span></span>|<span data-ttu-id="8b9b5-819">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="8b9b5-819">&lt;optional&gt;</span></span>|<span data-ttu-id="8b9b5-820">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-820">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="8b9b5-821">Объект</span><span class="sxs-lookup"><span data-stu-id="8b9b5-821">Object</span></span>|<span data-ttu-id="8b9b5-822">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="8b9b5-822">&lt;optional&gt;</span></span>|<span data-ttu-id="8b9b5-823">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-823">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="8b9b5-824">function</span><span class="sxs-lookup"><span data-stu-id="8b9b5-824">function</span></span>|<span data-ttu-id="8b9b5-825">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="8b9b5-825">&lt;optional&gt;</span></span>|<span data-ttu-id="8b9b5-826">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="8b9b5-826">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="8b9b5-827">После успешного выполнения идентификатор вложения будет представлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-827">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="8b9b5-828">Если добавить вложение не удастся, объект `asyncResult` будет содержать объект `Error` с описанием ошибки.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-828">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="8b9b5-829">Ошибки</span><span class="sxs-lookup"><span data-stu-id="8b9b5-829">Errors</span></span>

|<span data-ttu-id="8b9b5-830">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="8b9b5-830">Error code</span></span>|<span data-ttu-id="8b9b5-831">Описание</span><span class="sxs-lookup"><span data-stu-id="8b9b5-831">Description</span></span>|
|------------|-------------|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="8b9b5-832">Сообщение или встреча содержат слишком много вложений.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-832">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8b9b5-833">Требования</span><span class="sxs-lookup"><span data-stu-id="8b9b5-833">Requirements</span></span>

|<span data-ttu-id="8b9b5-834">Требование</span><span class="sxs-lookup"><span data-stu-id="8b9b5-834">Requirement</span></span>|<span data-ttu-id="8b9b5-835">Значение</span><span class="sxs-lookup"><span data-stu-id="8b9b5-835">Value</span></span>|
|---|---|
|[<span data-ttu-id="8b9b5-836">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="8b9b5-836">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8b9b5-837">1.1</span><span class="sxs-lookup"><span data-stu-id="8b9b5-837">1.1</span></span>|
|[<span data-ttu-id="8b9b5-838">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="8b9b5-838">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8b9b5-839">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="8b9b5-839">ReadWriteItem</span></span>|
|[<span data-ttu-id="8b9b5-840">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="8b9b5-840">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8b9b5-841">Создание</span><span class="sxs-lookup"><span data-stu-id="8b9b5-841">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="8b9b5-842">Пример</span><span class="sxs-lookup"><span data-stu-id="8b9b5-842">Example</span></span>

<span data-ttu-id="8b9b5-843">В следующем примере существующий элемент Outlook добавляется в виде вложения с именем `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-843">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

####  <a name="close"></a><span data-ttu-id="8b9b5-844">close()</span><span class="sxs-lookup"><span data-stu-id="8b9b5-844">close()</span></span>

<span data-ttu-id="8b9b5-845">Закрывает текущий создаваемый элемент.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-845">Closes the current item that is being composed.</span></span>

<span data-ttu-id="8b9b5-p142">Работа метода `close` зависит от текущего состояния создаваемого элемента. Если элемент содержит несохраненные изменения, клиент предложит пользователю сохранить или отклонить их либо отменить действие закрытия.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-p142">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="8b9b5-848">Если элемент является встречей в Outlook в Интернете, и он был ранее сохранен с помощью `saveAsync`, пользователю предлагается сохранить, отменить или удалить его, даже если не произошло каких-либо изменений, поскольку этот элемент был последним сохраненным.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-848">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="8b9b5-849">Если в клиенте Outlook для настольных ПК сообщение представляет собой ответ в тексте, метод `close` не работает.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-849">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="8b9b5-850">Требования</span><span class="sxs-lookup"><span data-stu-id="8b9b5-850">Requirements</span></span>

|<span data-ttu-id="8b9b5-851">Требование</span><span class="sxs-lookup"><span data-stu-id="8b9b5-851">Requirement</span></span>|<span data-ttu-id="8b9b5-852">Значение</span><span class="sxs-lookup"><span data-stu-id="8b9b5-852">Value</span></span>|
|---|---|
|[<span data-ttu-id="8b9b5-853">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="8b9b5-853">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8b9b5-854">1.3</span><span class="sxs-lookup"><span data-stu-id="8b9b5-854">1.3</span></span>|
|[<span data-ttu-id="8b9b5-855">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="8b9b5-855">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8b9b5-856">Ограниченный доступ</span><span class="sxs-lookup"><span data-stu-id="8b9b5-856">Restricted</span></span>|
|[<span data-ttu-id="8b9b5-857">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="8b9b5-857">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8b9b5-858">Создание</span><span class="sxs-lookup"><span data-stu-id="8b9b5-858">Compose</span></span>|

#### <a name="displayreplyallformformdata"></a><span data-ttu-id="8b9b5-859">displayReplyAllForm(formData)</span><span class="sxs-lookup"><span data-stu-id="8b9b5-859">displayReplyAllForm(formData)</span></span>

<span data-ttu-id="8b9b5-860">Отображает форму ответа, включающую отправителя и всех получателей выбранного сообщения или организатора и всех участников выбранной встречи.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-860">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="8b9b5-861">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-861">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="8b9b5-862">В веб-приложении Outlook форма ответа отображается в виде всплывающей формы в представлении с 3 колонками либо всплывающей формы в представлении с 1 или 2 колонками.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-862">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="8b9b5-863">Если любой строчный параметр превышает указанные для него ограничения, `displayReplyAllForm` возвращает исключение.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-863">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="8b9b5-p143">Если в параметре `formData.attachments` указаны вложения, Outlook и веб-приложение Outlook пытаются скачать их и вложить в форму ответа. Если какие-либо вложения добавить не удается, в форме отображается сообщение об ошибке. Если сообщения об ошибках не предусмотрены, то они не отображаются.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-p143">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8b9b5-867">Параметры:</span><span class="sxs-lookup"><span data-stu-id="8b9b5-867">Parameters:</span></span>

|<span data-ttu-id="8b9b5-868">Имя</span><span class="sxs-lookup"><span data-stu-id="8b9b5-868">Name</span></span>|<span data-ttu-id="8b9b5-869">Тип</span><span class="sxs-lookup"><span data-stu-id="8b9b5-869">Type</span></span>|<span data-ttu-id="8b9b5-870">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="8b9b5-870">Attributes</span></span>|<span data-ttu-id="8b9b5-871">Описание</span><span class="sxs-lookup"><span data-stu-id="8b9b5-871">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="8b9b5-872">String | Object</span><span class="sxs-lookup"><span data-stu-id="8b9b5-872">String &#124; Object</span></span>||<span data-ttu-id="8b9b5-p144">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-p144">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="8b9b5-875">**ИЛИ**</span><span class="sxs-lookup"><span data-stu-id="8b9b5-875">**OR**</span></span><br/><span data-ttu-id="8b9b5-p145">Объект, который содержит текст или данные вложения и функцию обратного вызова. Ниже представлено определение этого объекта.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-p145">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="8b9b5-878">String</span><span class="sxs-lookup"><span data-stu-id="8b9b5-878">String</span></span>|<span data-ttu-id="8b9b5-879">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="8b9b5-879">&lt;optional&gt;</span></span>|<span data-ttu-id="8b9b5-p146">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-p146">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="8b9b5-882">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="8b9b5-882">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="8b9b5-883">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="8b9b5-883">&lt;optional&gt;</span></span>|<span data-ttu-id="8b9b5-884">Массив объектов JSON, представляющих собой вложенные файлы или элементы.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-884">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="8b9b5-885">String</span><span class="sxs-lookup"><span data-stu-id="8b9b5-885">String</span></span>||<span data-ttu-id="8b9b5-p147">Указывает тип вложения. Допустимые значения: `file` для вложенного файла и `item` для вложенного элемента.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-p147">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="8b9b5-888">String</span><span class="sxs-lookup"><span data-stu-id="8b9b5-888">String</span></span>||<span data-ttu-id="8b9b5-889">Строка, содержащая имя вложения, длиной до 255 символов.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-889">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="8b9b5-890">String</span><span class="sxs-lookup"><span data-stu-id="8b9b5-890">String</span></span>||<span data-ttu-id="8b9b5-p148">Используется, только если свойству `type` задано значение `file`. URI расположения файла.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-p148">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="8b9b5-893">Логическое значение</span><span class="sxs-lookup"><span data-stu-id="8b9b5-893">Boolean</span></span>||<span data-ttu-id="8b9b5-p149">Используется только в том случае, если свойству `type` задано значение `file`. Значение `true` указывает на то, что вложение будет встроено в текст сообщения и не должно отображаться в списке вложений.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-p149">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="8b9b5-896">String</span><span class="sxs-lookup"><span data-stu-id="8b9b5-896">String</span></span>||<span data-ttu-id="8b9b5-p150">Используется, только если свойству `type` задано значение `item`. Идентификатор вложения EWS. Это строка длиной до 100 символов.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-p150">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="8b9b5-900">function</span><span class="sxs-lookup"><span data-stu-id="8b9b5-900">function</span></span>|<span data-ttu-id="8b9b5-901">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="8b9b5-901">&lt;optional&gt;</span></span>|<span data-ttu-id="8b9b5-902">По завершении работы метода функция, переданная параметру `callback`, вызывается с помощью одного параметра `asyncResult`, представляющего собой объект [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="8b9b5-902">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8b9b5-903">Требования</span><span class="sxs-lookup"><span data-stu-id="8b9b5-903">Requirements</span></span>

|<span data-ttu-id="8b9b5-904">Требование</span><span class="sxs-lookup"><span data-stu-id="8b9b5-904">Requirement</span></span>|<span data-ttu-id="8b9b5-905">Значение</span><span class="sxs-lookup"><span data-stu-id="8b9b5-905">Value</span></span>|
|---|---|
|[<span data-ttu-id="8b9b5-906">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="8b9b5-906">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8b9b5-907">1.0</span><span class="sxs-lookup"><span data-stu-id="8b9b5-907">1.0</span></span>|
|[<span data-ttu-id="8b9b5-908">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="8b9b5-908">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8b9b5-909">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8b9b5-909">ReadItem</span></span>|
|[<span data-ttu-id="8b9b5-910">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="8b9b5-910">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8b9b5-911">Чтение</span><span class="sxs-lookup"><span data-stu-id="8b9b5-911">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="8b9b5-912">Примеры</span><span class="sxs-lookup"><span data-stu-id="8b9b5-912">Examples</span></span>

<span data-ttu-id="8b9b5-913">Приведенный ниже код передает строку в функцию `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-913">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="8b9b5-914">Ответ с пустым текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-914">Reply with an empty body.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="8b9b5-915">Ответ только с текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-915">Reply with just a body.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="8b9b5-916">Ответ с текстом сообщения и вложенным файлом.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-916">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="8b9b5-917">Ответ с текстом сообщения и вложенным элементом.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-917">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="8b9b5-918">Ответ с текстом сообщения, вложенным файлом, вложенным элементом и обратным вызовом.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-918">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata"></a><span data-ttu-id="8b9b5-919">displayReplyForm(formData)</span><span class="sxs-lookup"><span data-stu-id="8b9b5-919">displayReplyForm(formData)</span></span>

<span data-ttu-id="8b9b5-920">Отображает форму ответа, включающую только отправителя выбранного сообщения или организатора выбранной встречи.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-920">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="8b9b5-921">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-921">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="8b9b5-922">В веб-приложении Outlook форма ответа отображается в виде всплывающей формы в представлении с 3 колонками либо всплывающей формы в представлении с 1 или 2 колонками.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-922">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="8b9b5-923">Если любой строчный параметр превышает указанные для него ограничения, `displayReplyForm` возвращает исключение.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-923">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="8b9b5-p151">Если в параметре `formData.attachments` указаны вложения, Outlook и веб-приложение Outlook пытаются скачать их и вложить в форму ответа. Если какие-либо вложения добавить не удается, в форме отображается сообщение об ошибке. Если сообщения об ошибках не предусмотрены, то они не отображаются.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-p151">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8b9b5-927">Параметры:</span><span class="sxs-lookup"><span data-stu-id="8b9b5-927">Parameters:</span></span>

|<span data-ttu-id="8b9b5-928">Имя</span><span class="sxs-lookup"><span data-stu-id="8b9b5-928">Name</span></span>|<span data-ttu-id="8b9b5-929">Тип</span><span class="sxs-lookup"><span data-stu-id="8b9b5-929">Type</span></span>|<span data-ttu-id="8b9b5-930">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="8b9b5-930">Attributes</span></span>|<span data-ttu-id="8b9b5-931">Описание</span><span class="sxs-lookup"><span data-stu-id="8b9b5-931">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="8b9b5-932">String | Object</span><span class="sxs-lookup"><span data-stu-id="8b9b5-932">String &#124; Object</span></span>||<span data-ttu-id="8b9b5-p152">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-p152">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="8b9b5-935">**ИЛИ**</span><span class="sxs-lookup"><span data-stu-id="8b9b5-935">**OR**</span></span><br/><span data-ttu-id="8b9b5-p153">Объект, который содержит текст или данные вложения и функцию обратного вызова. Ниже представлено определение этого объекта.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-p153">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="8b9b5-938">String</span><span class="sxs-lookup"><span data-stu-id="8b9b5-938">String</span></span>|<span data-ttu-id="8b9b5-939">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="8b9b5-939">&lt;optional&gt;</span></span>|<span data-ttu-id="8b9b5-p154">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-p154">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="8b9b5-942">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="8b9b5-942">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="8b9b5-943">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="8b9b5-943">&lt;optional&gt;</span></span>|<span data-ttu-id="8b9b5-944">Массив объектов JSON, представляющих собой вложенные файлы или элементы.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-944">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="8b9b5-945">String</span><span class="sxs-lookup"><span data-stu-id="8b9b5-945">String</span></span>||<span data-ttu-id="8b9b5-p155">Указывает тип вложения. Допустимые значения: `file` для вложенного файла и `item` для вложенного элемента.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-p155">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="8b9b5-948">String</span><span class="sxs-lookup"><span data-stu-id="8b9b5-948">String</span></span>||<span data-ttu-id="8b9b5-949">Строка, содержащая имя вложения, длиной до 255 символов.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-949">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="8b9b5-950">String</span><span class="sxs-lookup"><span data-stu-id="8b9b5-950">String</span></span>||<span data-ttu-id="8b9b5-p156">Используется, только если свойству `type` задано значение `file`. URI расположения файла.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-p156">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="8b9b5-953">Логическое значение</span><span class="sxs-lookup"><span data-stu-id="8b9b5-953">Boolean</span></span>||<span data-ttu-id="8b9b5-p157">Используется только в том случае, если свойству `type` задано значение `file`. Значение `true` указывает на то, что вложение будет встроено в текст сообщения и не должно отображаться в списке вложений.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-p157">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="8b9b5-956">String</span><span class="sxs-lookup"><span data-stu-id="8b9b5-956">String</span></span>||<span data-ttu-id="8b9b5-p158">Используется, только если свойству `type` задано значение `item`. Идентификатор вложения EWS. Это строка длиной до 100 символов.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-p158">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="8b9b5-960">function</span><span class="sxs-lookup"><span data-stu-id="8b9b5-960">function</span></span>|<span data-ttu-id="8b9b5-961">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="8b9b5-961">&lt;optional&gt;</span></span>|<span data-ttu-id="8b9b5-962">По завершении работы метода функция, переданная параметру `callback`, вызывается с помощью одного параметра `asyncResult`, представляющего собой объект [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="8b9b5-962">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8b9b5-963">Требования</span><span class="sxs-lookup"><span data-stu-id="8b9b5-963">Requirements</span></span>

|<span data-ttu-id="8b9b5-964">Требование</span><span class="sxs-lookup"><span data-stu-id="8b9b5-964">Requirement</span></span>|<span data-ttu-id="8b9b5-965">Значение</span><span class="sxs-lookup"><span data-stu-id="8b9b5-965">Value</span></span>|
|---|---|
|[<span data-ttu-id="8b9b5-966">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="8b9b5-966">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8b9b5-967">1.0</span><span class="sxs-lookup"><span data-stu-id="8b9b5-967">1.0</span></span>|
|[<span data-ttu-id="8b9b5-968">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="8b9b5-968">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8b9b5-969">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8b9b5-969">ReadItem</span></span>|
|[<span data-ttu-id="8b9b5-970">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="8b9b5-970">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8b9b5-971">Чтение</span><span class="sxs-lookup"><span data-stu-id="8b9b5-971">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="8b9b5-972">Примеры</span><span class="sxs-lookup"><span data-stu-id="8b9b5-972">Examples</span></span>

<span data-ttu-id="8b9b5-973">Приведенный ниже код передает строку в функцию `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-973">The following code passes a string to the `displayReplyForm` function.</span></span>

```
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="8b9b5-974">Ответ с пустым текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-974">Reply with an empty body.</span></span>

```
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="8b9b5-975">Ответ только с текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-975">Reply with just a body.</span></span>

```
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="8b9b5-976">Ответ с текстом сообщения и вложенным файлом.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-976">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="8b9b5-977">Ответ с текстом сообщения и вложенным элементом.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-977">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="8b9b5-978">Ответ с текстом сообщения, вложенным файлом, вложенным элементом и обратным вызовом.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-978">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlookofficeentities"></a><span data-ttu-id="8b9b5-979">getEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="8b9b5-979">getEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span></span>

<span data-ttu-id="8b9b5-980">Получает сущности, обнаруженные в выбранном тексте элемента.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-980">Gets the entities found in the selected item.</span></span>

> [!NOTE]
> <span data-ttu-id="8b9b5-981">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-981">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="8b9b5-982">Требования</span><span class="sxs-lookup"><span data-stu-id="8b9b5-982">Requirements</span></span>

|<span data-ttu-id="8b9b5-983">Требование</span><span class="sxs-lookup"><span data-stu-id="8b9b5-983">Requirement</span></span>|<span data-ttu-id="8b9b5-984">Значение</span><span class="sxs-lookup"><span data-stu-id="8b9b5-984">Value</span></span>|
|---|---|
|[<span data-ttu-id="8b9b5-985">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="8b9b5-985">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8b9b5-986">1.0</span><span class="sxs-lookup"><span data-stu-id="8b9b5-986">1.0</span></span>|
|[<span data-ttu-id="8b9b5-987">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="8b9b5-987">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8b9b5-988">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8b9b5-988">ReadItem</span></span>|
|[<span data-ttu-id="8b9b5-989">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="8b9b5-989">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8b9b5-990">Чтение</span><span class="sxs-lookup"><span data-stu-id="8b9b5-990">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="8b9b5-991">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="8b9b5-991">Returns:</span></span>

<span data-ttu-id="8b9b5-992">Тип: [Entities](/javascript/api/outlook/office.entities)</span><span class="sxs-lookup"><span data-stu-id="8b9b5-992">Type: [Entities](/javascript/api/outlook/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="8b9b5-993">Пример</span><span class="sxs-lookup"><span data-stu-id="8b9b5-993">Example</span></span>

<span data-ttu-id="8b9b5-994">Ниже приведен пример получения доступа к сущностям контактов в тексте текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-994">The following example accesses the contacts entities on the current item.</span></span>

```
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion"></a><span data-ttu-id="8b9b5-995">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="8b9b5-995">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span></span>

<span data-ttu-id="8b9b5-996">Получает массив всех сущностей указанного типа, обнаруженных в тексте выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-996">Gets an array of all the entities of the specified entity type found in the selected item.</span></span>

> [!NOTE]
> <span data-ttu-id="8b9b5-997">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-997">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8b9b5-998">Параметры:</span><span class="sxs-lookup"><span data-stu-id="8b9b5-998">Parameters:</span></span>

|<span data-ttu-id="8b9b5-999">Имя</span><span class="sxs-lookup"><span data-stu-id="8b9b5-999">Name</span></span>|<span data-ttu-id="8b9b5-1000">Тип</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1000">Type</span></span>|<span data-ttu-id="8b9b5-1001">Описание</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1001">Description</span></span>|
|---|---|---|
|`entityType`|[<span data-ttu-id="8b9b5-1002">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1002">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook/office.mailboxenums.entitytype)|<span data-ttu-id="8b9b5-1003">Одно из значений перечисления EntityType.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1003">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8b9b5-1004">Требования</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1004">Requirements</span></span>

|<span data-ttu-id="8b9b5-1005">Требование</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1005">Requirement</span></span>|<span data-ttu-id="8b9b5-1006">Значение</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1006">Value</span></span>|
|---|---|
|[<span data-ttu-id="8b9b5-1007">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1007">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8b9b5-1008">1.0</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1008">1.0</span></span>|
|[<span data-ttu-id="8b9b5-1009">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1009">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8b9b5-1010">Ограниченный доступ</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1010">Restricted</span></span>|
|[<span data-ttu-id="8b9b5-1011">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1011">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8b9b5-1012">Чтение</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1012">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="8b9b5-1013">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1013">Returns:</span></span>

<span data-ttu-id="8b9b5-1014">Если значение, переданное в `entityType`, не является допустимым членом перечисления `EntityType`, метод возвращает значение NULL.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1014">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="8b9b5-1015">Если в тексте элемента отсутствуют сущности указанного типа, метод возвращает пустой массив.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1015">If no entities of the specified type are present on the item, the method returns an empty array.</span></span> <span data-ttu-id="8b9b5-1016">В противном случае тип объектов в возвращаемом массиве зависит от типа сущности, запрошенной в параметре `entityType`.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1016">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="8b9b5-1017">Хотя минимальный уровень разрешений для использования этого метода — **Restricted**, для некоторых типов сущностей требуется доступ на уровне **ReadItem**, как указано в приведенной ниже таблице.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1017">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

|<span data-ttu-id="8b9b5-1018">Значение параметра `entityType`</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1018">Value of `entityType`</span></span>|<span data-ttu-id="8b9b5-1019">Тип объектов в возвращаемом массиве</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1019">Type of objects in returned array</span></span>|<span data-ttu-id="8b9b5-1020">Необходимый уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1020">Required Permission Level</span></span>|
|---|---|---|
|`Address`|<span data-ttu-id="8b9b5-1021">String</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1021">String</span></span>|<span data-ttu-id="8b9b5-1022">**С ограничениями**</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1022">**Restricted**</span></span>|
|`Contact`|<span data-ttu-id="8b9b5-1023">Contact</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1023">Contact</span></span>|<span data-ttu-id="8b9b5-1024">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1024">**ReadItem**</span></span>|
|`EmailAddress`|<span data-ttu-id="8b9b5-1025">String</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1025">String</span></span>|<span data-ttu-id="8b9b5-1026">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1026">**ReadItem**</span></span>|
|`MeetingSuggestion`|<span data-ttu-id="8b9b5-1027">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1027">MeetingSuggestion</span></span>|<span data-ttu-id="8b9b5-1028">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1028">**ReadItem**</span></span>|
|`PhoneNumber`|<span data-ttu-id="8b9b5-1029">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1029">PhoneNumber</span></span>|<span data-ttu-id="8b9b5-1030">**С ограничениями**</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1030">**Restricted**</span></span>|
|`TaskSuggestion`|<span data-ttu-id="8b9b5-1031">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1031">TaskSuggestion</span></span>|<span data-ttu-id="8b9b5-1032">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1032">**ReadItem**</span></span>|
|`URL`|<span data-ttu-id="8b9b5-1033">String</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1033">String</span></span>|<span data-ttu-id="8b9b5-1034">**С ограничениями**</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1034">**Restricted**</span></span>|

<span data-ttu-id="8b9b5-1035">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="8b9b5-1035">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="8b9b5-1036">Пример</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1036">Example</span></span>

<span data-ttu-id="8b9b5-1037">В примере ниже показано, как получить доступ к массиву строк, которые представляют собой почтовые адреса в тексте текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1037">The following example shows how to access an array of strings that represent postal addresses in the subject or body of the current item.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion"></a><span data-ttu-id="8b9b5-1038">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1038">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span></span>

<span data-ttu-id="8b9b5-1039">Возвращает известные сущности в выбранном элементе, которые проходят через именованный фильтр, определяемый в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1039">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="8b9b5-1040">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1040">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="8b9b5-1041">Метод `getFilteredEntitiesByName` возвращает сущности, соответствующие регулярному выражению, которое определяется в элементе правила [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) в XML-файле манифеста, с использованием указанного значения элемента `FilterName`.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1041">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8b9b5-1042">Параметры:</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1042">Parameters:</span></span>

|<span data-ttu-id="8b9b5-1043">Имя</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1043">Name</span></span>|<span data-ttu-id="8b9b5-1044">Тип</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1044">Type</span></span>|<span data-ttu-id="8b9b5-1045">Описание</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1045">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="8b9b5-1046">String</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1046">String</span></span>|<span data-ttu-id="8b9b5-1047">Имя элемента правила `ItemHasKnownEntity`, определяющее соответствующий фильтр.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1047">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8b9b5-1048">Требования</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1048">Requirements</span></span>

|<span data-ttu-id="8b9b5-1049">Требование</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1049">Requirement</span></span>|<span data-ttu-id="8b9b5-1050">Значение</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1050">Value</span></span>|
|---|---|
|[<span data-ttu-id="8b9b5-1051">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1051">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8b9b5-1052">1.0</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1052">1.0</span></span>|
|[<span data-ttu-id="8b9b5-1053">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1053">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8b9b5-1054">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1054">ReadItem</span></span>|
|[<span data-ttu-id="8b9b5-1055">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1055">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8b9b5-1056">Чтение</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1056">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="8b9b5-1057">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1057">Returns:</span></span>

<span data-ttu-id="8b9b5-p160">Если в манифесте нет элемента `ItemHasKnownEntity` со значением `FilterName`, соответствующим параметру `name`, метод возвращает `null`. Если параметр `name` соответствует элементу `ItemHasKnownEntity` в манифесте, но при этом в текущем элементе нет соответствующих сущностей, метод возвращает пустой массив.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-p160">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="8b9b5-1060">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="8b9b5-1060">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span></span>

#### <a name="getinitializationcontextasyncoptions-callback"></a><span data-ttu-id="8b9b5-1061">getInitializationContextAsync([options], [callback])</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1061">getInitializationContextAsync([options], [callback])</span></span>

<span data-ttu-id="8b9b5-1062">Получает данные инициализации, передаваемые при [активации надстройки интерактивным сообщением](https://docs.microsoft.com/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1062">Gets initialization data passed when the add-in is [activated by an actionable message](https://docs.microsoft.com/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span></span>

> [!NOTE]
> <span data-ttu-id="8b9b5-1063">Примечание. Этот метод поддерживается только Outlook 2016 для Windows (версии "нажми и работай" с номером больше 16.0.8413.1000) и Outlook в Интернете для Office 365.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1063">Note: This method is only supported by Outlook 2016 for Windows (Click-to-Run versions greater than 16.0.8413.1000) and Outlook on the web for Office 365.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8b9b5-1064">Параметры:</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1064">Parameters:</span></span>
|<span data-ttu-id="8b9b5-1065">Имя</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1065">Name</span></span>|<span data-ttu-id="8b9b5-1066">Тип</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1066">Type</span></span>|<span data-ttu-id="8b9b5-1067">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1067">Attributes</span></span>|<span data-ttu-id="8b9b5-1068">Описание</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1068">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="8b9b5-1069">Объект</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1069">Object</span></span>|<span data-ttu-id="8b9b5-1070">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1070">&lt;optional&gt;</span></span>|<span data-ttu-id="8b9b5-1071">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1071">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="8b9b5-1072">Объект</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1072">Object</span></span>|<span data-ttu-id="8b9b5-1073">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1073">&lt;optional&gt;</span></span>|<span data-ttu-id="8b9b5-1074">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1074">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="8b9b5-1075">function</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1075">function</span></span>|<span data-ttu-id="8b9b5-1076">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1076">&lt;optional&gt;</span></span>|<span data-ttu-id="8b9b5-1077">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1077">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="8b9b5-1078">В случае успешного выполнения инициализации данных предоставляются в свойстве `asyncResult.value` в виде строки.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1078">On success, the intialization data is provided in the `asyncResult.value` property as a string.</span></span><br/><span data-ttu-id="8b9b5-1079">Если контекст инициализации отсутствует, объект `asyncResult` будет содержать объект `Error`, одному свойству которого (`code`) будет присвоено значение `9020`, а другому (`name`) — значение `GenericResponseError`.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1079">If there is no initialization context, the `asyncResult` object will contain an `Error` object with its `code` property set to `9020` and its `name` property set to `GenericResponseError`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8b9b5-1080">Требования</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1080">Requirements</span></span>

|<span data-ttu-id="8b9b5-1081">Требование</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1081">Requirement</span></span>|<span data-ttu-id="8b9b5-1082">Значение</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1082">Value</span></span>|
|---|---|
|[<span data-ttu-id="8b9b5-1083">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1083">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8b9b5-1084">Предварительная версия</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1084">Preview</span></span>|
|[<span data-ttu-id="8b9b5-1085">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1085">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8b9b5-1086">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1086">ReadItem</span></span>|
|[<span data-ttu-id="8b9b5-1087">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1087">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8b9b5-1088">Чтение</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1088">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8b9b5-1089">Пример</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1089">Example</span></span>

```
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

#### <a name="getregexmatches--object"></a><span data-ttu-id="8b9b5-1090">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1090">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="8b9b5-1091">Возвращает строчные значения в выбранном элементе, которые соответствуют регулярным выражениям, определенным в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1091">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="8b9b5-1092">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1092">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="8b9b5-p161">Метод `getRegExMatches` возвращает строки, соответствующие регулярному выражению, которое определяется в каждом элементе правила `ItemHasRegularExpressionMatch` или `ItemHasKnownEntity` в XML-файле манифеста. Для правила `ItemHasRegularExpressionMatch` свойство элемента, указанного этим правилом, должно содержать соответствующую строку. Простой тип `PropertyName` определяет поддерживаемые свойства.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-p161">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="8b9b5-1096">Например, рассмотрим манифест надстройки, который содержит указанный ниже элемент `Rule`.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1096">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="8b9b5-1097">Объект, возвращаемый методом `getRegExMatches`, будет содержать два свойства: `fruits` и `veggies`.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1097">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="8b9b5-p162">Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты. Лучше используйте для этого метод [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-).</span><span class="sxs-lookup"><span data-stu-id="8b9b5-p162">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="8b9b5-1101">Требования</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1101">Requirements</span></span>

|<span data-ttu-id="8b9b5-1102">Требование</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1102">Requirement</span></span>|<span data-ttu-id="8b9b5-1103">Значение</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1103">Value</span></span>|
|---|---|
|[<span data-ttu-id="8b9b5-1104">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1104">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8b9b5-1105">1.0</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1105">1.0</span></span>|
|[<span data-ttu-id="8b9b5-1106">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1106">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8b9b5-1107">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1107">ReadItem</span></span>|
|[<span data-ttu-id="8b9b5-1108">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1108">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8b9b5-1109">Чтение</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1109">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="8b9b5-1110">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1110">Returns:</span></span>

<span data-ttu-id="8b9b5-p163">Объект, содержащий массив строк, которые соответствуют регулярным выражениям, определяемым в XML-файле манифеста. Имя каждого массива равно соответствующему значению атрибута `RegExName` правила сопоставления `ItemHasRegularExpressionMatch` или атрибута `FilterName` правила сопоставления `ItemHasKnownEntity`.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-p163">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="8b9b5-1113">

<dt>Тип</dt>

</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1113">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="8b9b5-1114">Объект</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1114">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="8b9b5-1115">Пример</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1115">Example</span></span>

<span data-ttu-id="8b9b5-1116">В следующем примере показано, как получить доступ к массиву совпадений для элементов правила регулярного выражения `fruits` и `veggies`, которые указаны в манифесте.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1116">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veges = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="8b9b5-1117">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1117">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="8b9b5-1118">Возвращает строчные значения в выбранном элементе, которые соответствуют именованному регулярному выражению, определенному в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1118">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="8b9b5-1119">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1119">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="8b9b5-1120">Метод `getRegExMatchesByName` возвращает строки, соответствующие регулярному выражению, которое определяется в элементе правила `ItemHasRegularExpressionMatch` в XML-файле манифеста, с использованием указанного значения элемента `RegExName`.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1120">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="8b9b5-p164">Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-p164">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8b9b5-1123">Параметры:</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1123">Parameters:</span></span>

|<span data-ttu-id="8b9b5-1124">Имя</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1124">Name</span></span>|<span data-ttu-id="8b9b5-1125">Тип</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1125">Type</span></span>|<span data-ttu-id="8b9b5-1126">Описание</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1126">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="8b9b5-1127">String</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1127">String</span></span>|<span data-ttu-id="8b9b5-1128">Имя элемента правила `ItemHasRegularExpressionMatch`, определяющее соответствующий фильтр.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1128">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8b9b5-1129">Требования</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1129">Requirements</span></span>

|<span data-ttu-id="8b9b5-1130">Требование</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1130">Requirement</span></span>|<span data-ttu-id="8b9b5-1131">Значение</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1131">Value</span></span>|
|---|---|
|[<span data-ttu-id="8b9b5-1132">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1132">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8b9b5-1133">1.0</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1133">1.0</span></span>|
|[<span data-ttu-id="8b9b5-1134">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1134">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8b9b5-1135">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1135">ReadItem</span></span>|
|[<span data-ttu-id="8b9b5-1136">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1136">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8b9b5-1137">Чтение</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1137">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="8b9b5-1138">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1138">Returns:</span></span>

<span data-ttu-id="8b9b5-1139">Массив строк, соответствующих регулярному выражению, определяемому в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1139">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="8b9b5-1140">

<dt>Тип</dt>

</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1140">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="8b9b5-1141">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="8b9b5-1141">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="8b9b5-1142">Пример</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1142">Example</span></span>

```
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="8b9b5-1143">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1143">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="8b9b5-1144">Асинхронно возвращает данные, выбранные в теме или тексте сообщения.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1144">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="8b9b5-p165">Если выделенный фрагмент отсутствует, но курсор находится в тексте или теме, метод возвращает значение NULL для выбранных данных. Если выбраны не текст и не тема, метод возвращает ошибку `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-p165">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8b9b5-1147">Параметры:</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1147">Parameters:</span></span>

|<span data-ttu-id="8b9b5-1148">Имя</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1148">Name</span></span>|<span data-ttu-id="8b9b5-1149">Тип</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1149">Type</span></span>|<span data-ttu-id="8b9b5-1150">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1150">Attributes</span></span>|<span data-ttu-id="8b9b5-1151">Описание</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1151">Description</span></span>|
|---|---|---|---|
|`coercionType`|[<span data-ttu-id="8b9b5-1152">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1152">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="8b9b5-p166">Запрашивает формат данных. Если задано значение Text, метод возвращает обычный текст как строку, удаляя все имеющиеся HTML-теги. Если задано значение HTML, метод возвращает выделенный текст (обычный текст или HTML).</span><span class="sxs-lookup"><span data-stu-id="8b9b5-p166">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`|<span data-ttu-id="8b9b5-1156">Объект</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1156">Object</span></span>|<span data-ttu-id="8b9b5-1157">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1157">&lt;optional&gt;</span></span>|<span data-ttu-id="8b9b5-1158">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1158">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="8b9b5-1159">Объект</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1159">Object</span></span>|<span data-ttu-id="8b9b5-1160">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1160">&lt;optional&gt;</span></span>|<span data-ttu-id="8b9b5-1161">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1161">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="8b9b5-1162">функция</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1162">function</span></span>||<span data-ttu-id="8b9b5-1163">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1163">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="8b9b5-1164">Чтобы получить доступ к выбранным данным из метода обратного вызова, вызовите `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1164">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="8b9b5-1165">Для доступа к исходному свойству, на основе которого созданы выбранные данные, вызовите  `asyncResult.value.sourceProperty`, который может иметь значение `body` или `subject`.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1165">To access the source property that the selection comes from, call , which will be either  or .|</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8b9b5-1166">Требования</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1166">Requirements</span></span>

|<span data-ttu-id="8b9b5-1167">Требование</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1167">Requirement</span></span>|<span data-ttu-id="8b9b5-1168">Значение</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1168">Value</span></span>|
|---|---|
|[<span data-ttu-id="8b9b5-1169">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1169">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8b9b5-1170">1.2</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1170">1.2</span></span>|
|[<span data-ttu-id="8b9b5-1171">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1171">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8b9b5-1172">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1172">ReadWriteItem</span></span>|
|[<span data-ttu-id="8b9b5-1173">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1173">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8b9b5-1174">Создание</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1174">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="8b9b5-1175">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1175">Returns:</span></span>

<span data-ttu-id="8b9b5-1176">Выбранные данные в виде строки с форматом, определенным в параметре `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1176">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="8b9b5-1177">

<dt>Тип</dt>

</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1177">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="8b9b5-1178">String</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1178">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="8b9b5-1179">Пример</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1179">Example</span></span>

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

#### <a name="getselectedentities--entitiesjavascriptapioutlookofficeentities"></a><span data-ttu-id="8b9b5-1180">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1180">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span></span>

<span data-ttu-id="8b9b5-p168">Возвращает сущности, найденные в выделенном совпадении, выбранном пользователем. Выделенные совпадения применяются к [контекстным надстройкам](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="8b9b5-p168">Gets the entities found in a highlighted match a user has selected. Highlighted matches apply to [contextual add-ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="8b9b5-1183">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1183">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="8b9b5-1184">Требования</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1184">Requirements</span></span>

|<span data-ttu-id="8b9b5-1185">Требование</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1185">Requirement</span></span>|<span data-ttu-id="8b9b5-1186">Значение</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1186">Value</span></span>|
|---|---|
|[<span data-ttu-id="8b9b5-1187">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1187">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8b9b5-1188">1.6</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1188">1.6</span></span>|
|[<span data-ttu-id="8b9b5-1189">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1189">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8b9b5-1190">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1190">ReadItem</span></span>|
|[<span data-ttu-id="8b9b5-1191">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1191">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8b9b5-1192">Чтение</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1192">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="8b9b5-1193">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1193">Returns:</span></span>

<span data-ttu-id="8b9b5-1194">Тип: [Entities](/javascript/api/outlook/office.entities)</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1194">Type: [Entities](/javascript/api/outlook/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="8b9b5-1195">Пример</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1195">Example</span></span>

<span data-ttu-id="8b9b5-1196">В приведенном ниже примере показано, как получить доступ к сущностям адресов в выделенном совпадении, выбранном пользователем.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1196">The following example accesses the addresses entities in the highlighted match selected by the user.</span></span>

```
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

#### <a name="getselectedregexmatches--object"></a><span data-ttu-id="8b9b5-1197">getSelectedRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1197">getSelectedRegExMatches() → {Object}</span></span>

<span data-ttu-id="8b9b5-p169">Возвращает строковые значения в выделенном совпадении, которые соответствуют регулярным выражениям, определенным в XML-файле манифеста. Выделенные совпадения применяются к [контекстным надстройкам](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="8b9b5-p169">Returns string values in a highlighted match that match the regular expressions defined in the manifest XML file. Highlighted matches apply to [contextual add-ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="8b9b5-1200">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1200">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="8b9b5-p170">Метод `getSelectedRegExMatches` возвращает строки, соответствующие регулярному выражению, которое определяется в каждом элементе правила `ItemHasRegularExpressionMatch` или `ItemHasKnownEntity` в XML-файле манифеста. Для правила `ItemHasRegularExpressionMatch` свойство элемента, указанного этим правилом, должно содержать соответствующую строку. Простой тип `PropertyName` определяет поддерживаемые свойства.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-p170">The `getSelectedRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="8b9b5-1204">Например, рассмотрим манифест надстройки, который содержит указанный ниже элемент `Rule`.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1204">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="8b9b5-1205">Объект, возвращаемый методом `getRegExMatches`, будет содержать два свойства: `fruits` и `veggies`.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1205">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="8b9b5-p171">Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты. Лучше используйте для этого метод [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-).</span><span class="sxs-lookup"><span data-stu-id="8b9b5-p171">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="8b9b5-1209">Требования</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1209">Requirements</span></span>

|<span data-ttu-id="8b9b5-1210">Требование</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1210">Requirement</span></span>|<span data-ttu-id="8b9b5-1211">Значение</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1211">Value</span></span>|
|---|---|
|[<span data-ttu-id="8b9b5-1212">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1212">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8b9b5-1213">1.6</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1213">1.6</span></span>|
|[<span data-ttu-id="8b9b5-1214">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1214">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8b9b5-1215">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1215">ReadItem</span></span>|
|[<span data-ttu-id="8b9b5-1216">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1216">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8b9b5-1217">Чтение</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1217">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="8b9b5-1218">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1218">Returns:</span></span>

<span data-ttu-id="8b9b5-p172">Объект, содержащий массив строк, которые соответствуют регулярным выражениям, определяемым в XML-файле манифеста. Имя каждого массива равно соответствующему значению атрибута `RegExName` правила сопоставления `ItemHasRegularExpressionMatch` или атрибута `FilterName` правила сопоставления `ItemHasKnownEntity`.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-p172">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

##### <a name="example"></a><span data-ttu-id="8b9b5-1221">Пример</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1221">Example</span></span>

<span data-ttu-id="8b9b5-1222">В следующем примере показано, как получить доступ к массиву совпадений для элементов правила регулярного выражения `fruits` и `veggies`, которые указаны в манифесте.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1222">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

#### <a name="getsharedpropertiesasyncoptions-callback"></a><span data-ttu-id="8b9b5-1223">getSharedPropertiesAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1223">getSharedPropertiesAsync([options], callback)</span></span>

<span data-ttu-id="8b9b5-1224">Получает свойства выбранной встречи или сообщения в общей папке, календаре или почтовом ящике.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1224">Gets the properties of the selected appointment or message in a shared folder, calendar, or mailbox.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8b9b5-1225">Параметры:</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1225">Parameters:</span></span>

|<span data-ttu-id="8b9b5-1226">Имя</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1226">Name</span></span>|<span data-ttu-id="8b9b5-1227">Тип</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1227">Type</span></span>|<span data-ttu-id="8b9b5-1228">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1228">Attributes</span></span>|<span data-ttu-id="8b9b5-1229">Описание</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1229">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="8b9b5-1230">Объект</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1230">Object</span></span>|<span data-ttu-id="8b9b5-1231">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1231">&lt;optional&gt;</span></span>|<span data-ttu-id="8b9b5-1232">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1232">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="8b9b5-1233">Объект</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1233">Object</span></span>|<span data-ttu-id="8b9b5-1234">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1234">&lt;optional&gt;</span></span>|<span data-ttu-id="8b9b5-1235">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1235">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="8b9b5-1236">функция</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1236">function</span></span>||<span data-ttu-id="8b9b5-1237">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1237">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="8b9b5-1238">Настраиваемые свойства предоставляются в виде объекта [`SharedProperties`](/javascript/api/outlook/office.sharedproperties) в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1238">The custom properties are provided as a [`SharedProperties`](/javascript/api/outlook/office.sharedproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="8b9b5-1239">Этот объект можно использовать для получения общих свойств элемента.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1239">This object can be used to get the item's shared properties.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8b9b5-1240">Требования</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1240">Requirements</span></span>

|<span data-ttu-id="8b9b5-1241">Требование</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1241">Requirement</span></span>|<span data-ttu-id="8b9b5-1242">Значение</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1242">Value</span></span>|
|---|---|
|[<span data-ttu-id="8b9b5-1243">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1243">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8b9b5-1244">Предварительная версия</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1244">Preview</span></span>|
|[<span data-ttu-id="8b9b5-1245">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1245">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8b9b5-1246">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1246">ReadItem</span></span>|
|[<span data-ttu-id="8b9b5-1247">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1247">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8b9b5-1248">Cоздание или чтение</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1248">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="8b9b5-1249">Пример</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1249">Example</span></span>

```js
Office.context.mailbox.item.getSharedPropertiesAsync(callback);
function callback (asyncResult) {
  var context=asyncResult.context;
  var sharedProperties = asyncResult.value;
}
```

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="8b9b5-1250">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1250">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="8b9b5-1251">Асинхронно загружает настраиваемые свойства для надстройки выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1251">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="8b9b5-p174">Настраиваемые свойства сохраняются в виде пар "ключ-значение" для каждого приложения и каждого элемента. Этот метод возвращает объект `CustomProperties` при обратном вызове, который предоставляет методы для доступа к настраиваемым свойствам, характерным для текущего элемента и текущей надстройки. Настраиваемые свойства не шифруются для элемента, поэтому этот способ хранения не является безопасным.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-p174">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8b9b5-1255">Параметры:</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1255">Parameters:</span></span>

|<span data-ttu-id="8b9b5-1256">Имя</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1256">Name</span></span>|<span data-ttu-id="8b9b5-1257">Тип</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1257">Type</span></span>|<span data-ttu-id="8b9b5-1258">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1258">Attributes</span></span>|<span data-ttu-id="8b9b5-1259">Описание</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1259">Description</span></span>|
|---|---|---|---|
|`callback`|<span data-ttu-id="8b9b5-1260">function</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1260">function</span></span>||<span data-ttu-id="8b9b5-1261">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1261">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="8b9b5-1262">Настраиваемые свойства предоставляются в виде объекта [`CustomProperties`](/javascript/api/outlook/office.customproperties) в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1262">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="8b9b5-1263">Этот объект можно использовать для получения, задания и удаления настраиваемых свойств из элемента и сохранения изменений настраиваемого свойства на сервере.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1263">The custom properties are provided as a CustomProperties object in the asyncResult.value property. This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`|<span data-ttu-id="8b9b5-1264">Объект</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1264">Object</span></span>|<span data-ttu-id="8b9b5-1265">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1265">&lt;optional&gt;</span></span>|<span data-ttu-id="8b9b5-1266">Разработчики могут указать любой объект, к которому необходимо получить доступ, в функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1266">Developers can provide any object they wish to access in the callback method.</span></span> <span data-ttu-id="8b9b5-1267">Доступ к этому объекту можно получить с помощью свойства `asyncResult.asyncContext` в функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1267">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.|</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8b9b5-1268">Требования</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1268">Requirements</span></span>

|<span data-ttu-id="8b9b5-1269">Требование</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1269">Requirement</span></span>|<span data-ttu-id="8b9b5-1270">Значение</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1270">Value</span></span>|
|---|---|
|[<span data-ttu-id="8b9b5-1271">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1271">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8b9b5-1272">1.0</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1272">1.0</span></span>|
|[<span data-ttu-id="8b9b5-1273">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1273">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8b9b5-1274">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1274">ReadItem</span></span>|
|[<span data-ttu-id="8b9b5-1275">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1275">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8b9b5-1276">Cоздание или чтение</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1276">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="8b9b5-1277">Пример</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1277">Example</span></span>

<span data-ttu-id="8b9b5-p177">В приведенном ниже примере кода показано, как асинхронно загружать настраиваемые свойства, характерные для текущего элемента, с помощью метода `loadCustomPropertiesAsync`. Этот пример также показывает, как сохранять эти свойства на сервере с помощью метода `CustomProperties.saveAsync`. В этом примере кода, после того как выполнена загрузка настраиваемых свойств, метод `CustomProperties.get` используется для считывания настраиваемого свойства `myProp`, метод `CustomProperties.set` — для записи настраиваемого свойства `otherProp`, а метод `saveAsync` — для сохранения настраиваемых свойств.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-p177">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="8b9b5-1281">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1281">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="8b9b5-1282">Удаляет вложение из сообщения или встречи.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1282">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="8b9b5-p178">Метод `removeAttachmentAsync` удаляет из элемента вложение с указанным идентификатором. Идентификатор вложения рекомендуется использовать для удаления вложения, только если оно добавлено тем же почтовым приложением в ходе текущего сеанса. В веб-приложении Outlook и веб-приложении Outlook для устройств идентификатор вложения действителен только в рамках одного сеанса. Сеанс завершается, когда пользователь закрывает приложение или начинает создавать элемент во встроенной форме, а затем переходит из формы в отдельное окно.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-p178">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8b9b5-1287">Параметры:</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1287">Parameters:</span></span>

|<span data-ttu-id="8b9b5-1288">Имя</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1288">Name</span></span>|<span data-ttu-id="8b9b5-1289">Тип</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1289">Type</span></span>|<span data-ttu-id="8b9b5-1290">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1290">Attributes</span></span>|<span data-ttu-id="8b9b5-1291">Описание</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1291">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="8b9b5-1292">String</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1292">String</span></span>||<span data-ttu-id="8b9b5-p179">Идентификатор удаляемого вложения. Максимальная длина строки — 100 символов.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-p179">The identifier of the attachment to remove. The maximum length of the string is 100 characters.</span></span>|
|`options`|<span data-ttu-id="8b9b5-1295">Объект</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1295">Object</span></span>|<span data-ttu-id="8b9b5-1296">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1296">&lt;optional&gt;</span></span>|<span data-ttu-id="8b9b5-1297">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1297">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="8b9b5-1298">Объект</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1298">Object</span></span>|<span data-ttu-id="8b9b5-1299">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1299">&lt;optional&gt;</span></span>|<span data-ttu-id="8b9b5-1300">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1300">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="8b9b5-1301">function</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1301">function</span></span>|<span data-ttu-id="8b9b5-1302">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1302">&lt;optional&gt;</span></span>|<span data-ttu-id="8b9b5-1303">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1303">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="8b9b5-1304">Если удалить вложение не удается, свойство `asyncResult.error` содержит код ошибки с указанием ее причины.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1304">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="8b9b5-1305">Ошибки</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1305">Errors</span></span>

|<span data-ttu-id="8b9b5-1306">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1306">Error code</span></span>|<span data-ttu-id="8b9b5-1307">Описание</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1307">Description</span></span>|
|------------|-------------|
|`InvalidAttachmentId`|<span data-ttu-id="8b9b5-1308">Идентификатор вложения не существует.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1308">The attachment identifier does not exist.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8b9b5-1309">Требования</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1309">Requirements</span></span>

|<span data-ttu-id="8b9b5-1310">Требование</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1310">Requirement</span></span>|<span data-ttu-id="8b9b5-1311">Значение</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1311">Value</span></span>|
|---|---|
|[<span data-ttu-id="8b9b5-1312">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1312">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8b9b5-1313">1.1</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1313">1.1</span></span>|
|[<span data-ttu-id="8b9b5-1314">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1314">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8b9b5-1315">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1315">ReadWriteItem</span></span>|
|[<span data-ttu-id="8b9b5-1316">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1316">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8b9b5-1317">Создание</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1317">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="8b9b5-1318">Пример</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1318">Example</span></span>

<span data-ttu-id="8b9b5-1319">Указанный ниже код удаляет вложение с идентификатором "0".</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1319">The following code removes an attachment with an identifier of '0'.</span></span>

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

####  <a name="removehandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="8b9b5-1320">removeHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1320">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="8b9b5-1321">Удаляет обработчик для поддерживаемого события.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1321">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="8b9b5-1322">В настоящее время поддерживаемые типы событий, `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, и `Office.EventType.RecurrenceChanged`</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1322">Currently the supported event types are `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`</span></span>

##### <a name="parameters"></a><span data-ttu-id="8b9b5-1323">Параметры:</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1323">Parameters:</span></span>

| <span data-ttu-id="8b9b5-1324">Имя</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1324">Name</span></span> | <span data-ttu-id="8b9b5-1325">Тип</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1325">Type</span></span> | <span data-ttu-id="8b9b5-1326">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1326">Attributes</span></span> | <span data-ttu-id="8b9b5-1327">Описание</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1327">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="8b9b5-1328">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1328">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="8b9b5-1329">Событие, которое должно вызвать обработчик.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1329">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="8b9b5-1330">Функция</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1330">Function</span></span> || <span data-ttu-id="8b9b5-p180">Функция для обработки события. Функция должна принимать один параметр, представляющий собой объектный литерал. Значение свойства `type` параметра совпадет со значением параметра `eventType`, переданного методу `removeHandlerAsync`.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-p180">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `removeHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="8b9b5-1334">Объект</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1334">Object</span></span> | <span data-ttu-id="8b9b5-1335">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1335">&lt;optional&gt;</span></span> | <span data-ttu-id="8b9b5-1336">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1336">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="8b9b5-1337">Объект</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1337">Object</span></span> | <span data-ttu-id="8b9b5-1338">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1338">&lt;optional&gt;</span></span> | <span data-ttu-id="8b9b5-1339">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1339">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="8b9b5-1340">function</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1340">function</span></span>| <span data-ttu-id="8b9b5-1341">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1341">&lt;optional&gt;</span></span>|<span data-ttu-id="8b9b5-1342">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1342">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8b9b5-1343">Требования</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1343">Requirements</span></span>

|<span data-ttu-id="8b9b5-1344">Требование</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1344">Requirement</span></span>| <span data-ttu-id="8b9b5-1345">Значение</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1345">Value</span></span>|
|---|---|
|[<span data-ttu-id="8b9b5-1346">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1346">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8b9b5-1347">1.7</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1347">17 </span></span> |
|[<span data-ttu-id="8b9b5-1348">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1348">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8b9b5-1349">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1349">ReadItem</span></span> |
|[<span data-ttu-id="8b9b5-1350">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1350">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8b9b5-1351">Cоздание или чтение</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1351">Compose or read</span></span> |

####  <a name="saveasyncoptions-callback"></a><span data-ttu-id="8b9b5-1352">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1352">saveAsync([options], callback)</span></span>

<span data-ttu-id="8b9b5-1353">Асинхронно сохраняет элемент.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1353">Asynchronously saves an item.</span></span>

<span data-ttu-id="8b9b5-p181">При вызове этот метод сохраняет текущее сообщение в виде черновика и возвращает идентификатор элемента с помощью метода обратного вызова. В веб-приложернии Outlook или интерактивном режиме Outlook этот элемент сохраняется на сервере. В Outlook в режиме кэширования этот элемент сохраняется в локальном кэше.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-p181">When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="8b9b5-1357">Если в режиме создания надстройка вызывает для элемента метод `saveAsync`, чтобы получить параметр `itemId` для использования с EWS или REST API, необходимо помнить, что синхронизация элемента с сервером может занять много времени, если Outlook работает в режиме кэширования данных.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1357">Note: If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server. Until the item is synced, using the  will return an error.</span></span> <span data-ttu-id="8b9b5-1358">До окончания синхронизации применение параметра `itemId` будет приводить к ошибке.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1358">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="8b9b5-p183">Так как для встреч не предусмотрено состояние черновика, если `saveAsync` вызывается для встречи в режиме создания, она сохраняется как обычная встреча в календаре пользователя. При сохранении новой встречи приглашения не отправляются. При сохранении существующей встречи уведомления отправляются добавленным или удаленным участникам.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-p183">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="8b9b5-1362">Следующие клиенты имеют разную реакцию на событие для `saveAsync` для встреч в режиме создания:</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1362">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="8b9b5-1363">Mac Outlook не поддерживает `saveAsync` на собрании в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1363">Note: Mac Outlook does not support `saveAsync` on a meeting in compose mode. Calling  on a meeting in Mac Outlook will return an error.</span></span> <span data-ttu-id="8b9b5-1364">Вызов `saveAsync` на собрании в Mac Outlook возвращает ошибку.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1364">Note: Mac Outlook does not support  on a meeting in compose mode. Calling `saveAsync` on a meeting in Mac Outlook will return an error.</span></span>
> - <span data-ttu-id="8b9b5-1365">Outlook в Интернете всегда отправляет приглашение или обновления при вызове `saveAsync` на встрече в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1365">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8b9b5-1366">Параметры:</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1366">Parameters:</span></span>

|<span data-ttu-id="8b9b5-1367">Имя</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1367">Name</span></span>|<span data-ttu-id="8b9b5-1368">Тип</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1368">Type</span></span>|<span data-ttu-id="8b9b5-1369">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1369">Attributes</span></span>|<span data-ttu-id="8b9b5-1370">Описание</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1370">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="8b9b5-1371">Объект</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1371">Object</span></span>|<span data-ttu-id="8b9b5-1372">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1372">&lt;optional&gt;</span></span>|<span data-ttu-id="8b9b5-1373">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1373">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="8b9b5-1374">Объект</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1374">Object</span></span>|<span data-ttu-id="8b9b5-1375">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1375">&lt;optional&gt;</span></span>|<span data-ttu-id="8b9b5-1376">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1376">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="8b9b5-1377">функция</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1377">function</span></span>||<span data-ttu-id="8b9b5-1378">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1378">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="8b9b5-1379">После успешного выполнения идентификатор элемента будет предоставлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1379">On success, the item identifier is provided in the `asyncResult.value` property.|</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8b9b5-1380">Требования</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1380">Requirements</span></span>

|<span data-ttu-id="8b9b5-1381">Требование</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1381">Requirement</span></span>|<span data-ttu-id="8b9b5-1382">Значение</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1382">Value</span></span>|
|---|---|
|[<span data-ttu-id="8b9b5-1383">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1383">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8b9b5-1384">1.3</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1384">1.3</span></span>|
|[<span data-ttu-id="8b9b5-1385">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1385">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8b9b5-1386">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1386">ReadWriteItem</span></span>|
|[<span data-ttu-id="8b9b5-1387">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1387">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8b9b5-1388">Создание</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1388">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="8b9b5-1389">Примеры</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1389">Examples</span></span>

```
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result
  });
```

<span data-ttu-id="8b9b5-p185">Ниже приведен пример параметра `result`, переданного функции обратного вызова. Свойство `value` содержит идентификатор элемента.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-p185">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

####  <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="8b9b5-1392">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1392">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="8b9b5-1393">Асинхронно вставляет данные в текст или тему сообщения.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1393">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="8b9b5-p186">Метод `setSelectedDataAsync` вставляет указанную строку в местоположение курсора в теме или тексте элемента либо, если текст выделен в редакторе, он заменяет выделенный текст. Если курсор находится вне текста или темы элемента, возвращается ошибка. После вставки курсор помещается в конец вставленного содержимого.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-p186">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8b9b5-1397">Параметры:</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1397">Parameters:</span></span>

|<span data-ttu-id="8b9b5-1398">Имя</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1398">Name</span></span>|<span data-ttu-id="8b9b5-1399">Тип</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1399">Type</span></span>|<span data-ttu-id="8b9b5-1400">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1400">Attributes</span></span>|<span data-ttu-id="8b9b5-1401">Описание</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1401">Description</span></span>|
|---|---|---|---|
|`data`|<span data-ttu-id="8b9b5-1402">String</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1402">String</span></span>||<span data-ttu-id="8b9b5-p187">Вставляемые данные. Объем данных не должен превышать 1 000 000 символов. Если передано больше 1 000 000 символов, возвращается исключение `ArgumentOutOfRange`.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-p187">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`|<span data-ttu-id="8b9b5-1406">Объект</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1406">Object</span></span>|<span data-ttu-id="8b9b5-1407">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1407">&lt;optional&gt;</span></span>|<span data-ttu-id="8b9b5-1408">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1408">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="8b9b5-1409">Объект</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1409">Object</span></span>|<span data-ttu-id="8b9b5-1410">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1410">&lt;optional&gt;</span></span>|<span data-ttu-id="8b9b5-1411">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1411">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="8b9b5-1412">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1412">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="8b9b5-1413">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1413">&lt;optional&gt;</span></span>|<span data-ttu-id="8b9b5-p188">Если задано значение `text`, текущий стиль применяется в Outlook и веб-приложении Outlook. Если поле представляет собой редактор HTML, вставляются только текстовые данные, даже если они имеют формат HTML.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-p188">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="8b9b5-p189">Если `html` и поле поддерживают HTML (а тема не поддерживает), в веб-приложении Outlook применяется текущий стиль, а в Outlook — стиль по умолчанию. Если поле является текстовым, возвращается ошибка `InvalidDataFormat`.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-p189">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="8b9b5-1418">Если тип `coercionType` не установлен, результат зависит от поля: если поле имеет формат HTML, то используется HTML; если поле является текстовым, то используется обычный текст.</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1418">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`|<span data-ttu-id="8b9b5-1419">function</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1419">function</span></span>||<span data-ttu-id="8b9b5-1420">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1420">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8b9b5-1421">Требования</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1421">Requirements</span></span>

|<span data-ttu-id="8b9b5-1422">Требование</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1422">Requirement</span></span>|<span data-ttu-id="8b9b5-1423">Значение</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1423">Value</span></span>|
|---|---|
|[<span data-ttu-id="8b9b5-1424">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1424">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8b9b5-1425">1.2</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1425">1.2</span></span>|
|[<span data-ttu-id="8b9b5-1426">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1426">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8b9b5-1427">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1427">ReadWriteItem</span></span>|
|[<span data-ttu-id="8b9b5-1428">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1428">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8b9b5-1429">Создание</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1429">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="8b9b5-1430">Пример</span><span class="sxs-lookup"><span data-stu-id="8b9b5-1430">Example</span></span>

```
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```