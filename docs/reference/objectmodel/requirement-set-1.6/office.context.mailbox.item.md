
# <a name="item"></a><span data-ttu-id="5f0b1-101">item</span><span class="sxs-lookup"><span data-stu-id="5f0b1-101">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="5f0b1-102">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="5f0b1-102">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="5f0b1-p101">Пространство имен `item` используется для доступа к выбранному в данный момент сообщению, приглашению на собрание или описанию встречи. Вы можете определить тип пространства имен `item` с помощью свойства [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook16officemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="5f0b1-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook16officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="5f0b1-105">Требования</span><span class="sxs-lookup"><span data-stu-id="5f0b1-105">Requirements</span></span>

|<span data-ttu-id="5f0b1-106">Требование</span><span class="sxs-lookup"><span data-stu-id="5f0b1-106">Requirement</span></span>| <span data-ttu-id="5f0b1-107">Значение</span><span class="sxs-lookup"><span data-stu-id="5f0b1-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="5f0b1-108">Версия минимального набора требований для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="5f0b1-108">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5f0b1-109">1.0</span><span class="sxs-lookup"><span data-stu-id="5f0b1-109">1.0</span></span>|
|[<span data-ttu-id="5f0b1-110">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5f0b1-110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5f0b1-111">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="5f0b1-111">Restricted</span></span>|
|[<span data-ttu-id="5f0b1-112">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5f0b1-112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5f0b1-113">Cоздание или чтение</span><span class="sxs-lookup"><span data-stu-id="5f0b1-113">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="5f0b1-114">Члены и методы</span><span class="sxs-lookup"><span data-stu-id="5f0b1-114">Members and methods</span></span>

| <span data-ttu-id="5f0b1-115">Член</span><span class="sxs-lookup"><span data-stu-id="5f0b1-115">Member</span></span> | <span data-ttu-id="5f0b1-116">Тип</span><span class="sxs-lookup"><span data-stu-id="5f0b1-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="5f0b1-117">attachments</span><span class="sxs-lookup"><span data-stu-id="5f0b1-117">attachments</span></span>](#attachments-arrayattachmentdetailsjavascriptapioutlook16officeattachmentdetails) | <span data-ttu-id="5f0b1-118">Член</span><span class="sxs-lookup"><span data-stu-id="5f0b1-118">Member</span></span> |
| [<span data-ttu-id="5f0b1-119">bcc</span><span class="sxs-lookup"><span data-stu-id="5f0b1-119">bcc</span></span>](#bcc-recipientsjavascriptapioutlook16officerecipients) | <span data-ttu-id="5f0b1-120">Член</span><span class="sxs-lookup"><span data-stu-id="5f0b1-120">Member</span></span> |
| [<span data-ttu-id="5f0b1-121">body</span><span class="sxs-lookup"><span data-stu-id="5f0b1-121">body</span></span>](#body-bodyjavascriptapioutlook16officebody) | <span data-ttu-id="5f0b1-122">Член</span><span class="sxs-lookup"><span data-stu-id="5f0b1-122">Member</span></span> |
| [<span data-ttu-id="5f0b1-123">cc</span><span class="sxs-lookup"><span data-stu-id="5f0b1-123">cc</span></span>](#cc-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients) | <span data-ttu-id="5f0b1-124">Член</span><span class="sxs-lookup"><span data-stu-id="5f0b1-124">Member</span></span> |
| [<span data-ttu-id="5f0b1-125">conversationId</span><span class="sxs-lookup"><span data-stu-id="5f0b1-125">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="5f0b1-126">Член</span><span class="sxs-lookup"><span data-stu-id="5f0b1-126">Member</span></span> |
| [<span data-ttu-id="5f0b1-127">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="5f0b1-127">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="5f0b1-128">Член</span><span class="sxs-lookup"><span data-stu-id="5f0b1-128">Member</span></span> |
| [<span data-ttu-id="5f0b1-129">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="5f0b1-129">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="5f0b1-130">Член</span><span class="sxs-lookup"><span data-stu-id="5f0b1-130">Member</span></span> |
| [<span data-ttu-id="5f0b1-131">end</span><span class="sxs-lookup"><span data-stu-id="5f0b1-131">end</span></span>](#end-datetimejavascriptapioutlook16officetime) | <span data-ttu-id="5f0b1-132">Член</span><span class="sxs-lookup"><span data-stu-id="5f0b1-132">Member</span></span> |
| [<span data-ttu-id="5f0b1-133">from</span><span class="sxs-lookup"><span data-stu-id="5f0b1-133">from</span></span>](#from-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails) | <span data-ttu-id="5f0b1-134">Член</span><span class="sxs-lookup"><span data-stu-id="5f0b1-134">Member</span></span> |
| [<span data-ttu-id="5f0b1-135">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="5f0b1-135">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="5f0b1-136">Член</span><span class="sxs-lookup"><span data-stu-id="5f0b1-136">Member</span></span> |
| [<span data-ttu-id="5f0b1-137">itemClass</span><span class="sxs-lookup"><span data-stu-id="5f0b1-137">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="5f0b1-138">Член</span><span class="sxs-lookup"><span data-stu-id="5f0b1-138">Member</span></span> |
| [<span data-ttu-id="5f0b1-139">itemId</span><span class="sxs-lookup"><span data-stu-id="5f0b1-139">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="5f0b1-140">Член</span><span class="sxs-lookup"><span data-stu-id="5f0b1-140">Member</span></span> |
| [<span data-ttu-id="5f0b1-141">itemType</span><span class="sxs-lookup"><span data-stu-id="5f0b1-141">itemType</span></span>](#itemtype-officemailboxenumsitemtypejavascriptapioutlook16officemailboxenumsitemtype) | <span data-ttu-id="5f0b1-142">Член</span><span class="sxs-lookup"><span data-stu-id="5f0b1-142">Member</span></span> |
| [<span data-ttu-id="5f0b1-143">location</span><span class="sxs-lookup"><span data-stu-id="5f0b1-143">location</span></span>](#location-stringlocationjavascriptapioutlook16officelocation) | <span data-ttu-id="5f0b1-144">Член</span><span class="sxs-lookup"><span data-stu-id="5f0b1-144">Member</span></span> |
| [<span data-ttu-id="5f0b1-145">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="5f0b1-145">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="5f0b1-146">Член</span><span class="sxs-lookup"><span data-stu-id="5f0b1-146">Member</span></span> |
| [<span data-ttu-id="5f0b1-147">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="5f0b1-147">notificationMessages</span></span>](#notificationmessages-notificationmessagesjavascriptapioutlook16officenotificationmessages) | <span data-ttu-id="5f0b1-148">Член</span><span class="sxs-lookup"><span data-stu-id="5f0b1-148">Member</span></span> |
| [<span data-ttu-id="5f0b1-149">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="5f0b1-149">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients) | <span data-ttu-id="5f0b1-150">Член</span><span class="sxs-lookup"><span data-stu-id="5f0b1-150">Member</span></span> |
| [<span data-ttu-id="5f0b1-151">organizer</span><span class="sxs-lookup"><span data-stu-id="5f0b1-151">organizer</span></span>](#organizer-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails) | <span data-ttu-id="5f0b1-152">Член</span><span class="sxs-lookup"><span data-stu-id="5f0b1-152">Member</span></span> |
| [<span data-ttu-id="5f0b1-153">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="5f0b1-153">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients) | <span data-ttu-id="5f0b1-154">Член</span><span class="sxs-lookup"><span data-stu-id="5f0b1-154">Member</span></span> |
| [<span data-ttu-id="5f0b1-155">sender</span><span class="sxs-lookup"><span data-stu-id="5f0b1-155">sender</span></span>](#sender-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails) | <span data-ttu-id="5f0b1-156">Член</span><span class="sxs-lookup"><span data-stu-id="5f0b1-156">Member</span></span> |
| [<span data-ttu-id="5f0b1-157">start</span><span class="sxs-lookup"><span data-stu-id="5f0b1-157">start</span></span>](#start-datetimejavascriptapioutlook16officetime) | <span data-ttu-id="5f0b1-158">Член</span><span class="sxs-lookup"><span data-stu-id="5f0b1-158">Member</span></span> |
| [<span data-ttu-id="5f0b1-159">subject</span><span class="sxs-lookup"><span data-stu-id="5f0b1-159">subject</span></span>](#subject-stringsubjectjavascriptapioutlook16officesubject) | <span data-ttu-id="5f0b1-160">Член</span><span class="sxs-lookup"><span data-stu-id="5f0b1-160">Member</span></span> |
| [<span data-ttu-id="5f0b1-161">to</span><span class="sxs-lookup"><span data-stu-id="5f0b1-161">to</span></span>](#to-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients) | <span data-ttu-id="5f0b1-162">Член</span><span class="sxs-lookup"><span data-stu-id="5f0b1-162">Member</span></span> |
| [<span data-ttu-id="5f0b1-163">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="5f0b1-163">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="5f0b1-164">Метод</span><span class="sxs-lookup"><span data-stu-id="5f0b1-164">Method</span></span> |
| [<span data-ttu-id="5f0b1-165">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="5f0b1-165">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="5f0b1-166">Метод</span><span class="sxs-lookup"><span data-stu-id="5f0b1-166">Method</span></span> |
| [<span data-ttu-id="5f0b1-167">close</span><span class="sxs-lookup"><span data-stu-id="5f0b1-167">close</span></span>](#close) | <span data-ttu-id="5f0b1-168">Метод</span><span class="sxs-lookup"><span data-stu-id="5f0b1-168">Method</span></span> |
| [<span data-ttu-id="5f0b1-169">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="5f0b1-169">displayReplyAllForm</span></span>](#displayreplyallformformdata) | <span data-ttu-id="5f0b1-170">Метод</span><span class="sxs-lookup"><span data-stu-id="5f0b1-170">Method</span></span> |
| [<span data-ttu-id="5f0b1-171">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="5f0b1-171">displayReplyForm</span></span>](#displayreplyformformdata) | <span data-ttu-id="5f0b1-172">Метод</span><span class="sxs-lookup"><span data-stu-id="5f0b1-172">Method</span></span> |
| [<span data-ttu-id="5f0b1-173">getEntities</span><span class="sxs-lookup"><span data-stu-id="5f0b1-173">getEntities</span></span>](#getentities--entitiesjavascriptapioutlook16officeentities) | <span data-ttu-id="5f0b1-174">Метод</span><span class="sxs-lookup"><span data-stu-id="5f0b1-174">Method</span></span> |
| [<span data-ttu-id="5f0b1-175">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="5f0b1-175">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook16officecontactmeetingsuggestionjavascriptapioutlook16officemeetingsuggestionphonenumberjavascriptapioutlook16officephonenumbertasksuggestionjavascriptapioutlook16officetasksuggestion) | <span data-ttu-id="5f0b1-176">Метод</span><span class="sxs-lookup"><span data-stu-id="5f0b1-176">Method</span></span> |
| [<span data-ttu-id="5f0b1-177">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="5f0b1-177">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook16officecontactmeetingsuggestionjavascriptapioutlook16officemeetingsuggestionphonenumberjavascriptapioutlook16officephonenumbertasksuggestionjavascriptapioutlook16officetasksuggestion) | <span data-ttu-id="5f0b1-178">Метод</span><span class="sxs-lookup"><span data-stu-id="5f0b1-178">Method</span></span> |
| [<span data-ttu-id="5f0b1-179">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="5f0b1-179">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="5f0b1-180">Метод</span><span class="sxs-lookup"><span data-stu-id="5f0b1-180">Method</span></span> |
| [<span data-ttu-id="5f0b1-181">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="5f0b1-181">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="5f0b1-182">Метод</span><span class="sxs-lookup"><span data-stu-id="5f0b1-182">Method</span></span> |
| [<span data-ttu-id="5f0b1-183">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="5f0b1-183">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="5f0b1-184">Метод</span><span class="sxs-lookup"><span data-stu-id="5f0b1-184">Method</span></span> |
| [<span data-ttu-id="5f0b1-185">getSelectedEntities</span><span class="sxs-lookup"><span data-stu-id="5f0b1-185">getSelectedEntities</span></span>](#getselectedentities--entitiesjavascriptapioutlook16officeentities) | <span data-ttu-id="5f0b1-186">Метод</span><span class="sxs-lookup"><span data-stu-id="5f0b1-186">Method</span></span> |
| [<span data-ttu-id="5f0b1-187">getSelectedRegExMatches</span><span class="sxs-lookup"><span data-stu-id="5f0b1-187">getSelectedRegExMatches</span></span>](#getselectedregexmatches--object) | <span data-ttu-id="5f0b1-188">Метод</span><span class="sxs-lookup"><span data-stu-id="5f0b1-188">Method</span></span> |
| [<span data-ttu-id="5f0b1-189">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="5f0b1-189">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="5f0b1-190">Метод</span><span class="sxs-lookup"><span data-stu-id="5f0b1-190">Method</span></span> |
| [<span data-ttu-id="5f0b1-191">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="5f0b1-191">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="5f0b1-192">Метод</span><span class="sxs-lookup"><span data-stu-id="5f0b1-192">Method</span></span> |
| [<span data-ttu-id="5f0b1-193">saveAsync</span><span class="sxs-lookup"><span data-stu-id="5f0b1-193">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="5f0b1-194">Метод</span><span class="sxs-lookup"><span data-stu-id="5f0b1-194">Method</span></span> |
| [<span data-ttu-id="5f0b1-195">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="5f0b1-195">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="5f0b1-196">Метод</span><span class="sxs-lookup"><span data-stu-id="5f0b1-196">Method</span></span> |

### <a name="example"></a><span data-ttu-id="5f0b1-197">Пример</span><span class="sxs-lookup"><span data-stu-id="5f0b1-197">Example</span></span>

<span data-ttu-id="5f0b1-198">В примере кода JavaScript, приведенном ниже, показано, как получить доступ к свойству `subject` текущего элемента в Outlook.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-198">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="5f0b1-199">Члены</span><span class="sxs-lookup"><span data-stu-id="5f0b1-199">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook16officeattachmentdetails"></a><span data-ttu-id="5f0b1-200">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_6/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="5f0b1-200">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_6/office.attachmentdetails)></span></span>

<span data-ttu-id="5f0b1-p102">Получает массив вложений для элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="5f0b1-203">Некоторые типы файлов блокируются Outlook из-за потенциальных проблем безопасности и поэтому не возвращаются.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-203">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="5f0b1-204">Дополнительные сведения см. в статье [Блокированные вложения в Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="5f0b1-204">For more information see [Payments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="5f0b1-205">Тип:</span><span class="sxs-lookup"><span data-stu-id="5f0b1-205">Type:</span></span>

*   <span data-ttu-id="5f0b1-206">Array.<[AttachmentDetails](/javascript/api/outlook_1_6/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="5f0b1-206">Array.<[AttachmentDetails](/javascript/api/outlook_1_6/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="5f0b1-207">Требования</span><span class="sxs-lookup"><span data-stu-id="5f0b1-207">Requirements</span></span>

|<span data-ttu-id="5f0b1-208">Требование</span><span class="sxs-lookup"><span data-stu-id="5f0b1-208">Requirement</span></span>| <span data-ttu-id="5f0b1-209">Значение</span><span class="sxs-lookup"><span data-stu-id="5f0b1-209">Value</span></span>|
|---|---|
|[<span data-ttu-id="5f0b1-210">Версия минимального набора требований для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="5f0b1-210">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5f0b1-211">1.0</span><span class="sxs-lookup"><span data-stu-id="5f0b1-211">1.0</span></span>|
|[<span data-ttu-id="5f0b1-212">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5f0b1-212">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5f0b1-213">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5f0b1-213">ReadItem</span></span>|
|[<span data-ttu-id="5f0b1-214">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5f0b1-214">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5f0b1-215">Чтение</span><span class="sxs-lookup"><span data-stu-id="5f0b1-215">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="5f0b1-216">Пример</span><span class="sxs-lookup"><span data-stu-id="5f0b1-216">Example</span></span>

<span data-ttu-id="5f0b1-217">С помощью приведенного ниже кода можно создать HTML-строку с подробными сведениями обо всех вложениях для текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-217">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

####  <a name="bcc-recipientsjavascriptapioutlook16officerecipients"></a><span data-ttu-id="5f0b1-218">bcc :[Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="5f0b1-218">bcc :[Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

<span data-ttu-id="5f0b1-219">Получает объект, который предоставляет методы для получения или обновления получателей в строке Bcc (скрытой копии) сообщения.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-219">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="5f0b1-220">Только в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-220">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="5f0b1-221">Тип:</span><span class="sxs-lookup"><span data-stu-id="5f0b1-221">Type:</span></span>

*   [<span data-ttu-id="5f0b1-222">Recipients</span><span class="sxs-lookup"><span data-stu-id="5f0b1-222">Recipients</span></span>](/javascript/api/outlook_1_6/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="5f0b1-223">Требования</span><span class="sxs-lookup"><span data-stu-id="5f0b1-223">Requirements</span></span>

|<span data-ttu-id="5f0b1-224">Требование</span><span class="sxs-lookup"><span data-stu-id="5f0b1-224">Requirement</span></span>| <span data-ttu-id="5f0b1-225">Значение</span><span class="sxs-lookup"><span data-stu-id="5f0b1-225">Value</span></span>|
|---|---|
|[<span data-ttu-id="5f0b1-226">Версия минимального набора требований для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="5f0b1-226">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5f0b1-227">1.1</span><span class="sxs-lookup"><span data-stu-id="5f0b1-227">1.1</span></span>|
|[<span data-ttu-id="5f0b1-228">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5f0b1-228">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5f0b1-229">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5f0b1-229">ReadItem</span></span>|
|[<span data-ttu-id="5f0b1-230">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5f0b1-230">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5f0b1-231">Создание</span><span class="sxs-lookup"><span data-stu-id="5f0b1-231">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="5f0b1-232">Пример</span><span class="sxs-lookup"><span data-stu-id="5f0b1-232">Example</span></span>

```
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlook16officebody"></a><span data-ttu-id="5f0b1-233">body :[Body](/javascript/api/outlook_1_6/office.body)</span><span class="sxs-lookup"><span data-stu-id="5f0b1-233">body :[Body](/javascript/api/outlook_1_6/office.body)</span></span>

<span data-ttu-id="5f0b1-234">Получает объект, предоставляющий методы для работы с основным текстом элемента.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-234">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="5f0b1-235">Тип:</span><span class="sxs-lookup"><span data-stu-id="5f0b1-235">Type:</span></span>

*   [<span data-ttu-id="5f0b1-236">Body</span><span class="sxs-lookup"><span data-stu-id="5f0b1-236">Body</span></span>](/javascript/api/outlook_1_6/office.body)

##### <a name="requirements"></a><span data-ttu-id="5f0b1-237">Требования</span><span class="sxs-lookup"><span data-stu-id="5f0b1-237">Requirements</span></span>

|<span data-ttu-id="5f0b1-238">Требование</span><span class="sxs-lookup"><span data-stu-id="5f0b1-238">Requirement</span></span>| <span data-ttu-id="5f0b1-239">Значение</span><span class="sxs-lookup"><span data-stu-id="5f0b1-239">Value</span></span>|
|---|---|
|[<span data-ttu-id="5f0b1-240">Версия минимального набора требований для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="5f0b1-240">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5f0b1-241">1.1</span><span class="sxs-lookup"><span data-stu-id="5f0b1-241">1.1</span></span>|
|[<span data-ttu-id="5f0b1-242">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5f0b1-242">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5f0b1-243">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5f0b1-243">ReadItem</span></span>|
|[<span data-ttu-id="5f0b1-244">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5f0b1-244">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5f0b1-245">Cоздание или чтение</span><span class="sxs-lookup"><span data-stu-id="5f0b1-245">Compose or read</span></span>|

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients"></a><span data-ttu-id="5f0b1-246">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="5f0b1-246">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

<span data-ttu-id="5f0b1-247">Предоставляет доступ к получателям Cc (копии) сообщения.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-247">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="5f0b1-248">Уровень доступа и тип объекта, зависит от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-248">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="5f0b1-249">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="5f0b1-249">Read mode</span></span>

<span data-ttu-id="5f0b1-p106">Свойство `cc` возвращает массив, который содержит объект `EmailAddressDetails` для каждого получателя, указанного в строке **Копия** сообщения. Коллекция может включать не более 100 членов.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="5f0b1-252">Режим создания</span><span class="sxs-lookup"><span data-stu-id="5f0b1-252">Compose mode</span></span>

<span data-ttu-id="5f0b1-253">Свойство `cc` возвращает объект `Recipients`, предоставляющий методы для получения или обновления получателей, которые указаны в строке **Cc (копия)** сообщения.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-253">The `cc` property returns a `Recipients` object that provides methods for manipulating the recipients on the **Cc** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="5f0b1-254">Тип:</span><span class="sxs-lookup"><span data-stu-id="5f0b1-254">Type:</span></span>

*   <span data-ttu-id="5f0b1-255">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="5f0b1-255">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="5f0b1-256">Требования</span><span class="sxs-lookup"><span data-stu-id="5f0b1-256">Requirements</span></span>

|<span data-ttu-id="5f0b1-257">Требование</span><span class="sxs-lookup"><span data-stu-id="5f0b1-257">Requirement</span></span>| <span data-ttu-id="5f0b1-258">Значение</span><span class="sxs-lookup"><span data-stu-id="5f0b1-258">Value</span></span>|
|---|---|
|[<span data-ttu-id="5f0b1-259">Версия минимального набора требований для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="5f0b1-259">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5f0b1-260">1.0</span><span class="sxs-lookup"><span data-stu-id="5f0b1-260">1.0</span></span>|
|[<span data-ttu-id="5f0b1-261">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5f0b1-261">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5f0b1-262">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5f0b1-262">ReadItem</span></span>|
|[<span data-ttu-id="5f0b1-263">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5f0b1-263">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5f0b1-264">Cоздание или чтение</span><span class="sxs-lookup"><span data-stu-id="5f0b1-264">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="5f0b1-265">Пример</span><span class="sxs-lookup"><span data-stu-id="5f0b1-265">Example</span></span>

```
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="5f0b1-266">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="5f0b1-266">(nullable) conversationId :String</span></span>

<span data-ttu-id="5f0b1-267">Получает идентификатор разговора по электронной почте, содержащего конкретное сообщение.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-267">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="5f0b1-p107">Вы можете получить целочисленное значение этого свойства, если ваше почтовое приложение активируется в формах просмотра или формах создания ответов. Если пользователь изменит тему ответа, после его отправки идентификатор беседы будет изменен, и полученное ранее значение будет недействительным.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="5f0b1-p108">Это свойство имеет значение NULL для нового элемента в форме создания. Свойство `conversationId` вернет значение, если пользователь задаст тему и сохранит элемент.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="5f0b1-272">Тип:</span><span class="sxs-lookup"><span data-stu-id="5f0b1-272">Type:</span></span>

*   <span data-ttu-id="5f0b1-273">String</span><span class="sxs-lookup"><span data-stu-id="5f0b1-273">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="5f0b1-274">Требования</span><span class="sxs-lookup"><span data-stu-id="5f0b1-274">Requirements</span></span>

|<span data-ttu-id="5f0b1-275">Требование</span><span class="sxs-lookup"><span data-stu-id="5f0b1-275">Requirement</span></span>| <span data-ttu-id="5f0b1-276">Значение</span><span class="sxs-lookup"><span data-stu-id="5f0b1-276">Value</span></span>|
|---|---|
|[<span data-ttu-id="5f0b1-277">Версия минимального набора требований для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="5f0b1-277">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5f0b1-278">1.0</span><span class="sxs-lookup"><span data-stu-id="5f0b1-278">1.0</span></span>|
|[<span data-ttu-id="5f0b1-279">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5f0b1-279">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5f0b1-280">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5f0b1-280">ReadItem</span></span>|
|[<span data-ttu-id="5f0b1-281">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5f0b1-281">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5f0b1-282">Cоздание или чтение</span><span class="sxs-lookup"><span data-stu-id="5f0b1-282">Compose or read</span></span>|

#### <a name="datetimecreated-date"></a><span data-ttu-id="5f0b1-283">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="5f0b1-283">dateTimeCreated :Date</span></span>

<span data-ttu-id="5f0b1-p109">Получает дату и время создания элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="5f0b1-286">Тип:</span><span class="sxs-lookup"><span data-stu-id="5f0b1-286">Type:</span></span>

*   <span data-ttu-id="5f0b1-287">Date</span><span class="sxs-lookup"><span data-stu-id="5f0b1-287">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="5f0b1-288">Требования</span><span class="sxs-lookup"><span data-stu-id="5f0b1-288">Requirements</span></span>

|<span data-ttu-id="5f0b1-289">Требование</span><span class="sxs-lookup"><span data-stu-id="5f0b1-289">Requirement</span></span>| <span data-ttu-id="5f0b1-290">Значение</span><span class="sxs-lookup"><span data-stu-id="5f0b1-290">Value</span></span>|
|---|---|
|[<span data-ttu-id="5f0b1-291">Версия минимального набора требований для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="5f0b1-291">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5f0b1-292">1.0</span><span class="sxs-lookup"><span data-stu-id="5f0b1-292">1.0</span></span>|
|[<span data-ttu-id="5f0b1-293">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5f0b1-293">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5f0b1-294">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5f0b1-294">ReadItem</span></span>|
|[<span data-ttu-id="5f0b1-295">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5f0b1-295">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5f0b1-296">Чтение</span><span class="sxs-lookup"><span data-stu-id="5f0b1-296">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="5f0b1-297">Пример</span><span class="sxs-lookup"><span data-stu-id="5f0b1-297">Example</span></span>

```
var created = Office.context.mailbox.item.dateTimeCreated;
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="5f0b1-298">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="5f0b1-298">dateTimeModified :Date</span></span>

<span data-ttu-id="5f0b1-p110">Получает дату и время последнего изменения элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="5f0b1-301">Этот элемент не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-301">Note: This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="5f0b1-302">Тип:</span><span class="sxs-lookup"><span data-stu-id="5f0b1-302">Type:</span></span>

*   <span data-ttu-id="5f0b1-303">Date</span><span class="sxs-lookup"><span data-stu-id="5f0b1-303">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="5f0b1-304">Требования</span><span class="sxs-lookup"><span data-stu-id="5f0b1-304">Requirements</span></span>

|<span data-ttu-id="5f0b1-305">Требование</span><span class="sxs-lookup"><span data-stu-id="5f0b1-305">Requirement</span></span>| <span data-ttu-id="5f0b1-306">Значение</span><span class="sxs-lookup"><span data-stu-id="5f0b1-306">Value</span></span>|
|---|---|
|[<span data-ttu-id="5f0b1-307">Версия минимального набора требований для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="5f0b1-307">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5f0b1-308">1.0</span><span class="sxs-lookup"><span data-stu-id="5f0b1-308">1.0</span></span>|
|[<span data-ttu-id="5f0b1-309">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5f0b1-309">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5f0b1-310">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5f0b1-310">ReadItem</span></span>|
|[<span data-ttu-id="5f0b1-311">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5f0b1-311">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5f0b1-312">Чтение</span><span class="sxs-lookup"><span data-stu-id="5f0b1-312">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="5f0b1-313">Пример</span><span class="sxs-lookup"><span data-stu-id="5f0b1-313">Example</span></span>

```
var modified = Office.context.mailbox.item.dateTimeModified;
```

####  <a name="end-datetimejavascriptapioutlook16officetime"></a><span data-ttu-id="5f0b1-314">end :Date|[Time](/javascript/api/outlook_1_6/office.time)</span><span class="sxs-lookup"><span data-stu-id="5f0b1-314">end :Date|[Time](/javascript/api/outlook_1_6/office.time)</span></span>

<span data-ttu-id="5f0b1-315">Получает или задает дату и время окончания встречи.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-315">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="5f0b1-p111">Свойство `end` представлено в виде значения даты и времени в формате UTC. Преобразовать значение свойства end в местные значения даты и времени клиента можно с помощью метода [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook16officelocalclienttime).</span><span class="sxs-lookup"><span data-stu-id="5f0b1-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook16officelocalclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="5f0b1-318">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="5f0b1-318">Read mode</span></span>

<span data-ttu-id="5f0b1-319">Свойство `end` возвращает объект `Date`.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-319">The `end` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="5f0b1-320">Режим создания</span><span class="sxs-lookup"><span data-stu-id="5f0b1-320">Compose mode</span></span>

<span data-ttu-id="5f0b1-321">Свойство `end` возвращает объект `Time`.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-321">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="5f0b1-322">Если вы задаете время окончания с помощью метода [`Time.setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-), необходимо использовать метод [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) для преобразования местного времени на клиенте в формат UTC для сервера.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-322">When you use the [`Time.setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="5f0b1-323">Тип:</span><span class="sxs-lookup"><span data-stu-id="5f0b1-323">Type:</span></span>

*   <span data-ttu-id="5f0b1-324">Date | [Time](/javascript/api/outlook_1_6/office.time)</span><span class="sxs-lookup"><span data-stu-id="5f0b1-324">Date | [Time](/javascript/api/outlook_1_6/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="5f0b1-325">Требования</span><span class="sxs-lookup"><span data-stu-id="5f0b1-325">Requirements</span></span>

|<span data-ttu-id="5f0b1-326">Требование</span><span class="sxs-lookup"><span data-stu-id="5f0b1-326">Requirement</span></span>| <span data-ttu-id="5f0b1-327">Значение</span><span class="sxs-lookup"><span data-stu-id="5f0b1-327">Value</span></span>|
|---|---|
|[<span data-ttu-id="5f0b1-328">Версия минимального набора требований для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="5f0b1-328">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5f0b1-329">1.0</span><span class="sxs-lookup"><span data-stu-id="5f0b1-329">1.0</span></span>|
|[<span data-ttu-id="5f0b1-330">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5f0b1-330">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5f0b1-331">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5f0b1-331">ReadItem</span></span>|
|[<span data-ttu-id="5f0b1-332">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5f0b1-332">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5f0b1-333">Cоздание или чтение</span><span class="sxs-lookup"><span data-stu-id="5f0b1-333">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="5f0b1-334">Пример</span><span class="sxs-lookup"><span data-stu-id="5f0b1-334">Example</span></span>

<span data-ttu-id="5f0b1-335">В примере ниже показано, как с помощью метода [`setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) объекта `Time` задать время окончания встречи в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-335">The following example sets the end time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

#### <a name="from-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails"></a><span data-ttu-id="5f0b1-336">from :[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="5f0b1-336">from :[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)</span></span>

<span data-ttu-id="5f0b1-p112">Получает электронный адрес отправителя сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-p112">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="5f0b1-p113">Свойства `from` и [`sender`](#sender-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails) представляют одно лицо, если сообщение не отправлено представителем. В противном случае свойство `from` представляет лицо, делегировавшее полномочия, а свойство sender — представителя.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-p113">The `from` and [`sender`](#sender-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="5f0b1-341">Свойство `recipientType` объекта `EmailAddressDetails` в свойстве `from` — `undefined`.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-341">Note: The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="5f0b1-342">Тип:</span><span class="sxs-lookup"><span data-stu-id="5f0b1-342">Type:</span></span>

*   [<span data-ttu-id="5f0b1-343">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="5f0b1-343">EmailAddressDetails</span></span>](/javascript/api/outlook_1_6/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="5f0b1-344">Требования</span><span class="sxs-lookup"><span data-stu-id="5f0b1-344">Requirements</span></span>

|<span data-ttu-id="5f0b1-345">Требование</span><span class="sxs-lookup"><span data-stu-id="5f0b1-345">Requirement</span></span>| <span data-ttu-id="5f0b1-346">Значение</span><span class="sxs-lookup"><span data-stu-id="5f0b1-346">Value</span></span>|
|---|---|
|[<span data-ttu-id="5f0b1-347">Версия минимального набора требований для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="5f0b1-347">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5f0b1-348">1.0</span><span class="sxs-lookup"><span data-stu-id="5f0b1-348">1.0</span></span>|
|[<span data-ttu-id="5f0b1-349">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5f0b1-349">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5f0b1-350">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5f0b1-350">ReadItem</span></span>|
|[<span data-ttu-id="5f0b1-351">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5f0b1-351">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5f0b1-352">Чтение</span><span class="sxs-lookup"><span data-stu-id="5f0b1-352">Read</span></span>|

#### <a name="internetmessageid-string"></a><span data-ttu-id="5f0b1-353">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="5f0b1-353">internetMessageId :String</span></span>

<span data-ttu-id="5f0b1-p114">Получает идентификатор интернет-сообщения для электронного сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-p114">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="5f0b1-356">Тип:</span><span class="sxs-lookup"><span data-stu-id="5f0b1-356">Type:</span></span>

*   <span data-ttu-id="5f0b1-357">String</span><span class="sxs-lookup"><span data-stu-id="5f0b1-357">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="5f0b1-358">Требования</span><span class="sxs-lookup"><span data-stu-id="5f0b1-358">Requirements</span></span>

|<span data-ttu-id="5f0b1-359">Требование</span><span class="sxs-lookup"><span data-stu-id="5f0b1-359">Requirement</span></span>| <span data-ttu-id="5f0b1-360">Значение</span><span class="sxs-lookup"><span data-stu-id="5f0b1-360">Value</span></span>|
|---|---|
|[<span data-ttu-id="5f0b1-361">Версия минимального набора требований для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="5f0b1-361">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5f0b1-362">1.0</span><span class="sxs-lookup"><span data-stu-id="5f0b1-362">1.0</span></span>|
|[<span data-ttu-id="5f0b1-363">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5f0b1-363">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5f0b1-364">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5f0b1-364">ReadItem</span></span>|
|[<span data-ttu-id="5f0b1-365">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5f0b1-365">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5f0b1-366">Чтение</span><span class="sxs-lookup"><span data-stu-id="5f0b1-366">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="5f0b1-367">Пример</span><span class="sxs-lookup"><span data-stu-id="5f0b1-367">Example</span></span>

```
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="5f0b1-368">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="5f0b1-368">itemClass :String</span></span>

<span data-ttu-id="5f0b1-p115">Получает класс элемента веб-служб Exchange для выбранного элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-p115">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="5f0b1-p116">Свойство `itemClass` указывает класс сообщения выбранного элемента. Ниже приводятся классы сообщения по умолчанию для элемента сообщения или встречи.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-p116">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="5f0b1-373">Тип</span><span class="sxs-lookup"><span data-stu-id="5f0b1-373">Type</span></span> | <span data-ttu-id="5f0b1-374">Описание</span><span class="sxs-lookup"><span data-stu-id="5f0b1-374">Description</span></span> | <span data-ttu-id="5f0b1-375">item class</span><span class="sxs-lookup"><span data-stu-id="5f0b1-375">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="5f0b1-376">Элементы встречи</span><span class="sxs-lookup"><span data-stu-id="5f0b1-376">Appointment items</span></span> | <span data-ttu-id="5f0b1-377">Это элементы календаря для класса элемента `IPM.Appointment` или `IPM.Appointment.Occurence`.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-377">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurence` |
| <span data-ttu-id="5f0b1-378">Элементы сообщения</span><span class="sxs-lookup"><span data-stu-id="5f0b1-378">Message items</span></span> | <span data-ttu-id="5f0b1-379">Сюда входят электронные сообщения, для которых по умолчанию задан класс сообщения `IPM.Note`, а также приглашения на собрания, ответы на них и уведомления об их отмене, использующие `IPM.Schedule.Meeting` в качестве базового класса сообщения.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-379">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="5f0b1-380">Можно создавать настраиваемые классы сообщения, расширяющие классы сообщения по умолчанию, например, настраиваемый класс сообщения о встрече `IPM.Appointment.Contoso`.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-380">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="5f0b1-381">Тип:</span><span class="sxs-lookup"><span data-stu-id="5f0b1-381">Type:</span></span>

*   <span data-ttu-id="5f0b1-382">String</span><span class="sxs-lookup"><span data-stu-id="5f0b1-382">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="5f0b1-383">Требования</span><span class="sxs-lookup"><span data-stu-id="5f0b1-383">Requirements</span></span>

|<span data-ttu-id="5f0b1-384">Требование</span><span class="sxs-lookup"><span data-stu-id="5f0b1-384">Requirement</span></span>| <span data-ttu-id="5f0b1-385">Значение</span><span class="sxs-lookup"><span data-stu-id="5f0b1-385">Value</span></span>|
|---|---|
|[<span data-ttu-id="5f0b1-386">Версия минимального набора требований для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="5f0b1-386">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5f0b1-387">1.0</span><span class="sxs-lookup"><span data-stu-id="5f0b1-387">1.0</span></span>|
|[<span data-ttu-id="5f0b1-388">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5f0b1-388">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5f0b1-389">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5f0b1-389">ReadItem</span></span>|
|[<span data-ttu-id="5f0b1-390">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5f0b1-390">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5f0b1-391">Чтение</span><span class="sxs-lookup"><span data-stu-id="5f0b1-391">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="5f0b1-392">Пример</span><span class="sxs-lookup"><span data-stu-id="5f0b1-392">Example</span></span>

```
var itemClass = Office.context.mailbox.item.itemClass;
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="5f0b1-393">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="5f0b1-393">(nullable) itemId :String</span></span>

<span data-ttu-id="5f0b1-p117">Получает идентификатор элемента веб-служб Exchange для текущего элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-p117">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="5f0b1-396">Идентификатор, возвращаемый свойством `itemId`, совпадает с идентификатором элемента веб-служб Exchange.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-396">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="5f0b1-397">Свойство  `itemId` не совпадает с идентификатором записи Outlook или идентификатором, используемым API-Интерфейсом REST Outlook.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-397">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="5f0b1-398">Прежде чем осуществлять вызовы API-Интерфейса REST с помощью этого значения, его следует преобразовать с помощью [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="5f0b1-398">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="5f0b1-399">Дополнительные сведения см. в статье [Использование API REST Outlook из надстройки Outlook](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="5f0b1-399">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="5f0b1-p119">Свойство `itemId` недоступно в режиме создания. Если требуется идентификатор элемента, с помощью метода [`saveAsync`](#saveasyncoptions-callback) можно сохранить элемент в хранилище. При этом в параметре [`AsyncResult.value`](/javascript/api/office/office.asyncresult) функции обратного вызова возвращается идентификатор элемента.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-p119">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="5f0b1-402">Тип:</span><span class="sxs-lookup"><span data-stu-id="5f0b1-402">Type:</span></span>

*   <span data-ttu-id="5f0b1-403">String</span><span class="sxs-lookup"><span data-stu-id="5f0b1-403">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="5f0b1-404">Требования</span><span class="sxs-lookup"><span data-stu-id="5f0b1-404">Requirements</span></span>

|<span data-ttu-id="5f0b1-405">Требование</span><span class="sxs-lookup"><span data-stu-id="5f0b1-405">Requirement</span></span>| <span data-ttu-id="5f0b1-406">Значение</span><span class="sxs-lookup"><span data-stu-id="5f0b1-406">Value</span></span>|
|---|---|
|[<span data-ttu-id="5f0b1-407">Версия минимального набора требований для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="5f0b1-407">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5f0b1-408">1.0</span><span class="sxs-lookup"><span data-stu-id="5f0b1-408">1.0</span></span>|
|[<span data-ttu-id="5f0b1-409">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5f0b1-409">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5f0b1-410">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5f0b1-410">ReadItem</span></span>|
|[<span data-ttu-id="5f0b1-411">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5f0b1-411">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5f0b1-412">Чтение</span><span class="sxs-lookup"><span data-stu-id="5f0b1-412">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="5f0b1-413">Пример</span><span class="sxs-lookup"><span data-stu-id="5f0b1-413">Example</span></span>

<span data-ttu-id="5f0b1-p120">Указанный ниже код проверяет наличие идентификатора элемента. Если свойство `itemId` возвращает значение `null` или `undefined`, элемент будет сохранен в хранилище, а из асинхронного результата будет получен идентификатор элемента.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result){
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook16officemailboxenumsitemtype"></a><span data-ttu-id="5f0b1-416">itemType:[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_6/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="5f0b1-416">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_6/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="5f0b1-417">Получает тип элемента, который представляет экземпляр.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-417">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="5f0b1-418">Свойство `itemType` возвращает одно из значений перечисления `ItemType`, которое указывает, является ли экземпляр объекта `item` сообщением или встречей.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-418">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="5f0b1-419">Тип:</span><span class="sxs-lookup"><span data-stu-id="5f0b1-419">Type:</span></span>

*   [<span data-ttu-id="5f0b1-420">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="5f0b1-420">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook_1_6/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="5f0b1-421">Требования</span><span class="sxs-lookup"><span data-stu-id="5f0b1-421">Requirements</span></span>

|<span data-ttu-id="5f0b1-422">Требование</span><span class="sxs-lookup"><span data-stu-id="5f0b1-422">Requirement</span></span>| <span data-ttu-id="5f0b1-423">Значение</span><span class="sxs-lookup"><span data-stu-id="5f0b1-423">Value</span></span>|
|---|---|
|[<span data-ttu-id="5f0b1-424">Версия минимального набора требований для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="5f0b1-424">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5f0b1-425">1.0</span><span class="sxs-lookup"><span data-stu-id="5f0b1-425">1.0</span></span>|
|[<span data-ttu-id="5f0b1-426">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5f0b1-426">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5f0b1-427">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5f0b1-427">ReadItem</span></span>|
|[<span data-ttu-id="5f0b1-428">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5f0b1-428">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5f0b1-429">Cоздание или чтение</span><span class="sxs-lookup"><span data-stu-id="5f0b1-429">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="5f0b1-430">Пример</span><span class="sxs-lookup"><span data-stu-id="5f0b1-430">Example</span></span>

```
if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Message)
  // do something
else
  // do something else
```

####  <a name="location-stringlocationjavascriptapioutlook16officelocation"></a><span data-ttu-id="5f0b1-431">location :String|[Location](/javascript/api/outlook_1_6/office.location)</span><span class="sxs-lookup"><span data-stu-id="5f0b1-431">location :String|[Location](/javascript/api/outlook_1_6/office.location)</span></span>

<span data-ttu-id="5f0b1-432">Получает или задает место встречи.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-432">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="5f0b1-433">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="5f0b1-433">Read mode</span></span>

<span data-ttu-id="5f0b1-434">Свойство `location` возвращает строку, содержащую сведения о месте встречи.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-434">The `location` property returns a string that contains the location of the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="5f0b1-435">Режим создания</span><span class="sxs-lookup"><span data-stu-id="5f0b1-435">Compose mode</span></span>

<span data-ttu-id="5f0b1-436">Свойство `location` возвращает объект `Location`, предоставляющий методы, которые используются для получения и задания места встречи.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-436">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="5f0b1-437">Тип:</span><span class="sxs-lookup"><span data-stu-id="5f0b1-437">Type:</span></span>

*   <span data-ttu-id="5f0b1-438">String | [Location](/javascript/api/outlook_1_6/office.location)</span><span class="sxs-lookup"><span data-stu-id="5f0b1-438">String | [Location](/javascript/api/outlook_1_6/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="5f0b1-439">Требования</span><span class="sxs-lookup"><span data-stu-id="5f0b1-439">Requirements</span></span>

|<span data-ttu-id="5f0b1-440">Требование</span><span class="sxs-lookup"><span data-stu-id="5f0b1-440">Requirement</span></span>| <span data-ttu-id="5f0b1-441">Значение</span><span class="sxs-lookup"><span data-stu-id="5f0b1-441">Value</span></span>|
|---|---|
|[<span data-ttu-id="5f0b1-442">Версия минимального набора требований для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="5f0b1-442">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5f0b1-443">1.0</span><span class="sxs-lookup"><span data-stu-id="5f0b1-443">1.0</span></span>|
|[<span data-ttu-id="5f0b1-444">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5f0b1-444">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5f0b1-445">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5f0b1-445">ReadItem</span></span>|
|[<span data-ttu-id="5f0b1-446">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5f0b1-446">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5f0b1-447">Cоздание или чтение</span><span class="sxs-lookup"><span data-stu-id="5f0b1-447">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="5f0b1-448">Пример</span><span class="sxs-lookup"><span data-stu-id="5f0b1-448">Example</span></span>

```
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

#### <a name="normalizedsubject-string"></a><span data-ttu-id="5f0b1-449">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="5f0b1-449">normalizedSubject :String</span></span>

<span data-ttu-id="5f0b1-p121">Получает тему элемента со всеми удаленными префиксами (включая `RE:` и `FWD:`). Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="5f0b1-p122">Свойство normalizedSubject получает тему элемента со стандартными префиксами (такими как `RE:` и `FW:`), добавляемыми почтовыми программами. Для получения темы элемента с неизмененными префиксами используйте свойство [`subject`](#subject-stringsubjectjavascriptapioutlook16officesubject).</span><span class="sxs-lookup"><span data-stu-id="5f0b1-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubjectjavascriptapioutlook16officesubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="5f0b1-454">Тип:</span><span class="sxs-lookup"><span data-stu-id="5f0b1-454">Type:</span></span>

*   <span data-ttu-id="5f0b1-455">String</span><span class="sxs-lookup"><span data-stu-id="5f0b1-455">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="5f0b1-456">Требования</span><span class="sxs-lookup"><span data-stu-id="5f0b1-456">Requirements</span></span>

|<span data-ttu-id="5f0b1-457">Требование</span><span class="sxs-lookup"><span data-stu-id="5f0b1-457">Requirement</span></span>| <span data-ttu-id="5f0b1-458">Значение</span><span class="sxs-lookup"><span data-stu-id="5f0b1-458">Value</span></span>|
|---|---|
|[<span data-ttu-id="5f0b1-459">Версия минимального набора требований для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="5f0b1-459">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5f0b1-460">1.0</span><span class="sxs-lookup"><span data-stu-id="5f0b1-460">1.0</span></span>|
|[<span data-ttu-id="5f0b1-461">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5f0b1-461">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5f0b1-462">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5f0b1-462">ReadItem</span></span>|
|[<span data-ttu-id="5f0b1-463">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5f0b1-463">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5f0b1-464">Чтение</span><span class="sxs-lookup"><span data-stu-id="5f0b1-464">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="5f0b1-465">Пример</span><span class="sxs-lookup"><span data-stu-id="5f0b1-465">Example</span></span>

```
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
```

####  <a name="notificationmessages-notificationmessagesjavascriptapioutlook16officenotificationmessages"></a><span data-ttu-id="5f0b1-466">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_6/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="5f0b1-466">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_6/office.notificationmessages)</span></span>

<span data-ttu-id="5f0b1-467">Получает сообщения уведомления для элемента.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-467">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="5f0b1-468">Тип:</span><span class="sxs-lookup"><span data-stu-id="5f0b1-468">Type:</span></span>

*   [<span data-ttu-id="5f0b1-469">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="5f0b1-469">NotificationMessages</span></span>](/javascript/api/outlook_1_6/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="5f0b1-470">Требования</span><span class="sxs-lookup"><span data-stu-id="5f0b1-470">Requirements</span></span>

|<span data-ttu-id="5f0b1-471">Требование</span><span class="sxs-lookup"><span data-stu-id="5f0b1-471">Requirement</span></span>| <span data-ttu-id="5f0b1-472">Значение</span><span class="sxs-lookup"><span data-stu-id="5f0b1-472">Value</span></span>|
|---|---|
|[<span data-ttu-id="5f0b1-473">Версия минимального набора требований для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="5f0b1-473">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5f0b1-474">1.3</span><span class="sxs-lookup"><span data-stu-id="5f0b1-474">1.3</span></span>|
|[<span data-ttu-id="5f0b1-475">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5f0b1-475">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5f0b1-476">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5f0b1-476">ReadItem</span></span>|
|[<span data-ttu-id="5f0b1-477">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5f0b1-477">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5f0b1-478">Cоздание или чтение</span><span class="sxs-lookup"><span data-stu-id="5f0b1-478">Compose or read</span></span>|

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients"></a><span data-ttu-id="5f0b1-479">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="5f0b1-479">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

<span data-ttu-id="5f0b1-480">Предоставляет доступ к необязательным участникам события.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-480">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="5f0b1-481">Уровень доступа и тип объекта, зависит от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-481">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="5f0b1-482">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="5f0b1-482">Read mode</span></span>

<span data-ttu-id="5f0b1-483">Свойство `optionalAttendees` возвращает массив, содержащий объект `EmailAddressDetails` для каждого необязательного участника собрания.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-483">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="5f0b1-484">Режим создания</span><span class="sxs-lookup"><span data-stu-id="5f0b1-484">Compose mode</span></span>

<span data-ttu-id="5f0b1-485">Свойство `optionalAttendees` возвращает объект `Recipients`, который предоставляет методы для получения и обновления необязательных участников собрания.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-485">The `optionalAttendees` property returns a `Recipients` object that provides methods to get and set the optional attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="5f0b1-486">Тип:</span><span class="sxs-lookup"><span data-stu-id="5f0b1-486">Type:</span></span>

*   <span data-ttu-id="5f0b1-487">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="5f0b1-487">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="5f0b1-488">Требования</span><span class="sxs-lookup"><span data-stu-id="5f0b1-488">Requirements</span></span>

|<span data-ttu-id="5f0b1-489">Требование</span><span class="sxs-lookup"><span data-stu-id="5f0b1-489">Requirement</span></span>| <span data-ttu-id="5f0b1-490">Значение</span><span class="sxs-lookup"><span data-stu-id="5f0b1-490">Value</span></span>|
|---|---|
|[<span data-ttu-id="5f0b1-491">Версия минимального набора требований для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="5f0b1-491">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5f0b1-492">1.0</span><span class="sxs-lookup"><span data-stu-id="5f0b1-492">1.0</span></span>|
|[<span data-ttu-id="5f0b1-493">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5f0b1-493">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5f0b1-494">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5f0b1-494">ReadItem</span></span>|
|[<span data-ttu-id="5f0b1-495">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5f0b1-495">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5f0b1-496">Cоздание или чтение</span><span class="sxs-lookup"><span data-stu-id="5f0b1-496">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="5f0b1-497">Пример</span><span class="sxs-lookup"><span data-stu-id="5f0b1-497">Example</span></span>

```
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails"></a><span data-ttu-id="5f0b1-498">organizer :[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="5f0b1-498">organizer :[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)</span></span>

<span data-ttu-id="5f0b1-p124">Получает электронный адрес организатора указанного собрания. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="5f0b1-501">Тип:</span><span class="sxs-lookup"><span data-stu-id="5f0b1-501">Type:</span></span>

*   [<span data-ttu-id="5f0b1-502">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="5f0b1-502">EmailAddressDetails</span></span>](/javascript/api/outlook_1_6/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="5f0b1-503">Требования</span><span class="sxs-lookup"><span data-stu-id="5f0b1-503">Requirements</span></span>

|<span data-ttu-id="5f0b1-504">Требование</span><span class="sxs-lookup"><span data-stu-id="5f0b1-504">Requirement</span></span>| <span data-ttu-id="5f0b1-505">Значение</span><span class="sxs-lookup"><span data-stu-id="5f0b1-505">Value</span></span>|
|---|---|
|[<span data-ttu-id="5f0b1-506">Версия минимального набора требований для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="5f0b1-506">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5f0b1-507">1.0</span><span class="sxs-lookup"><span data-stu-id="5f0b1-507">1.0</span></span>|
|[<span data-ttu-id="5f0b1-508">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5f0b1-508">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5f0b1-509">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5f0b1-509">ReadItem</span></span>|
|[<span data-ttu-id="5f0b1-510">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5f0b1-510">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5f0b1-511">Чтение</span><span class="sxs-lookup"><span data-stu-id="5f0b1-511">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="5f0b1-512">Пример</span><span class="sxs-lookup"><span data-stu-id="5f0b1-512">Example</span></span>

```
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
```

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients"></a><span data-ttu-id="5f0b1-513">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="5f0b1-513">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

<span data-ttu-id="5f0b1-514">Предоставляет доступ к обязательным участникам события.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-514">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="5f0b1-515">Уровень доступа и тип объекта, зависит от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-515">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="5f0b1-516">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="5f0b1-516">Read mode</span></span>

<span data-ttu-id="5f0b1-517">Свойство `requiredAttendees` возвращает массив, содержащий объект `EmailAddressDetails` для каждого обязательного участника собрания.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-517">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="5f0b1-518">Режим создания</span><span class="sxs-lookup"><span data-stu-id="5f0b1-518">Compose mode</span></span>

<span data-ttu-id="5f0b1-519">Свойство `requiredAttendees` возвращает объект `Recipients`, который предоставляет методы для получения и обновления обязательных участников собрания.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-519">The `requiredAttendees` property returns a `Recipients` object that provides methods to get and set the required attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="5f0b1-520">Тип:</span><span class="sxs-lookup"><span data-stu-id="5f0b1-520">Type:</span></span>

*   <span data-ttu-id="5f0b1-521">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="5f0b1-521">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="5f0b1-522">Требования</span><span class="sxs-lookup"><span data-stu-id="5f0b1-522">Requirements</span></span>

|<span data-ttu-id="5f0b1-523">Требование</span><span class="sxs-lookup"><span data-stu-id="5f0b1-523">Requirement</span></span>| <span data-ttu-id="5f0b1-524">Значение</span><span class="sxs-lookup"><span data-stu-id="5f0b1-524">Value</span></span>|
|---|---|
|[<span data-ttu-id="5f0b1-525">Версия минимального набора требований для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="5f0b1-525">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5f0b1-526">1.0</span><span class="sxs-lookup"><span data-stu-id="5f0b1-526">1.0</span></span>|
|[<span data-ttu-id="5f0b1-527">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5f0b1-527">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5f0b1-528">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5f0b1-528">ReadItem</span></span>|
|[<span data-ttu-id="5f0b1-529">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5f0b1-529">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5f0b1-530">Cоздание или чтение</span><span class="sxs-lookup"><span data-stu-id="5f0b1-530">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="5f0b1-531">Пример</span><span class="sxs-lookup"><span data-stu-id="5f0b1-531">Example</span></span>

```
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
}
```

#### <a name="sender-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails"></a><span data-ttu-id="5f0b1-532">sender :[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="5f0b1-532">sender :[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)</span></span>

<span data-ttu-id="5f0b1-p126">Получает электронный адрес отправителя электронного сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="5f0b1-p127">Свойства [`from`](#from-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails) и `sender` представляют одно лицо, если сообщение не отправлено представителем. В противном случае свойство `from` представляет лицо, делегировавшее полномочия, а свойство sender — представителя.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-p127">The [`from`](#from-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="5f0b1-537">Свойство `recipientType` объекта `EmailAddressDetails` в свойстве `sender` — `undefined`.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-537">Note: The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="5f0b1-538">Тип:</span><span class="sxs-lookup"><span data-stu-id="5f0b1-538">Type:</span></span>

*   [<span data-ttu-id="5f0b1-539">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="5f0b1-539">EmailAddressDetails</span></span>](/javascript/api/outlook_1_6/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="5f0b1-540">Требования</span><span class="sxs-lookup"><span data-stu-id="5f0b1-540">Requirements</span></span>

|<span data-ttu-id="5f0b1-541">Требование</span><span class="sxs-lookup"><span data-stu-id="5f0b1-541">Requirement</span></span>| <span data-ttu-id="5f0b1-542">Значение</span><span class="sxs-lookup"><span data-stu-id="5f0b1-542">Value</span></span>|
|---|---|
|[<span data-ttu-id="5f0b1-543">Версия минимального набора требований для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="5f0b1-543">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5f0b1-544">1.0</span><span class="sxs-lookup"><span data-stu-id="5f0b1-544">1.0</span></span>|
|[<span data-ttu-id="5f0b1-545">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5f0b1-545">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5f0b1-546">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5f0b1-546">ReadItem</span></span>|
|[<span data-ttu-id="5f0b1-547">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5f0b1-547">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5f0b1-548">Чтение</span><span class="sxs-lookup"><span data-stu-id="5f0b1-548">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="5f0b1-549">Пример</span><span class="sxs-lookup"><span data-stu-id="5f0b1-549">Example</span></span>

```
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
```

####  <a name="start-datetimejavascriptapioutlook16officetime"></a><span data-ttu-id="5f0b1-550">start :Date|[Time](/javascript/api/outlook_1_6/office.time)</span><span class="sxs-lookup"><span data-stu-id="5f0b1-550">start :Date|[Time](/javascript/api/outlook_1_6/office.time)</span></span>

<span data-ttu-id="5f0b1-551">Получает или задает дату и время начала встречи.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-551">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="5f0b1-p128">Свойство `start` представлено в виде значения даты и времени в формате UTC. Это значение можно преобразовать в местные значения даты и времени клиента с помощью метода [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook16officelocalclienttime).</span><span class="sxs-lookup"><span data-stu-id="5f0b1-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook16officelocalclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="5f0b1-554">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="5f0b1-554">Read mode</span></span>

<span data-ttu-id="5f0b1-555">Свойство `start` возвращает объект `Date`.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-555">The `start` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="5f0b1-556">Режим создания</span><span class="sxs-lookup"><span data-stu-id="5f0b1-556">Compose mode</span></span>

<span data-ttu-id="5f0b1-557">Свойство `start` возвращает объект `Time`.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-557">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="5f0b1-558">Если вы задаете время начала с помощью метода [`Time.setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-), необходимо использовать метод [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) для преобразования местного времени на клиенте в формат UTC для сервера.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-558">When you use the [`Time.setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="5f0b1-559">Тип:</span><span class="sxs-lookup"><span data-stu-id="5f0b1-559">Type:</span></span>

*   <span data-ttu-id="5f0b1-560">Date | [Time](/javascript/api/outlook_1_6/office.time)</span><span class="sxs-lookup"><span data-stu-id="5f0b1-560">Date | [Time](/javascript/api/outlook_1_6/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="5f0b1-561">Требования</span><span class="sxs-lookup"><span data-stu-id="5f0b1-561">Requirements</span></span>

|<span data-ttu-id="5f0b1-562">Требование</span><span class="sxs-lookup"><span data-stu-id="5f0b1-562">Requirement</span></span>| <span data-ttu-id="5f0b1-563">Значение</span><span class="sxs-lookup"><span data-stu-id="5f0b1-563">Value</span></span>|
|---|---|
|[<span data-ttu-id="5f0b1-564">Версия минимального набора требований для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="5f0b1-564">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5f0b1-565">1.0</span><span class="sxs-lookup"><span data-stu-id="5f0b1-565">1.0</span></span>|
|[<span data-ttu-id="5f0b1-566">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5f0b1-566">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5f0b1-567">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5f0b1-567">ReadItem</span></span>|
|[<span data-ttu-id="5f0b1-568">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5f0b1-568">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5f0b1-569">Cоздание или чтение</span><span class="sxs-lookup"><span data-stu-id="5f0b1-569">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="5f0b1-570">Пример</span><span class="sxs-lookup"><span data-stu-id="5f0b1-570">Example</span></span>

<span data-ttu-id="5f0b1-571">В примере ниже с помощью метода [`setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) объекта `Time` задается время начала встречи в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-571">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

####  <a name="subject-stringsubjectjavascriptapioutlook16officesubject"></a><span data-ttu-id="5f0b1-572">subject :String|[Subject](/javascript/api/outlook_1_6/office.subject)</span><span class="sxs-lookup"><span data-stu-id="5f0b1-572">subject :String|[Subject](/javascript/api/outlook_1_6/office.subject)</span></span>

<span data-ttu-id="5f0b1-573">Получает или задает описание, которое отображается в поле темы элемента.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-573">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="5f0b1-574">Свойство `subject` получает или задает всю тему элемента для отправки с почтового сервера.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-574">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="5f0b1-575">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="5f0b1-575">Read mode</span></span>

<span data-ttu-id="5f0b1-p129">Свойство `subject` возвращает строку. С помощью свойства [`normalizedSubject`](#normalizedsubject-string) можно получить тему без начальных префиксов, таких как `RE:` и `FW:`.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-p129">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```
var subject = Office.context.mailbox.item.subject;
```

##### <a name="compose-mode"></a><span data-ttu-id="5f0b1-578">Режим создания</span><span class="sxs-lookup"><span data-stu-id="5f0b1-578">Compose mode</span></span>

<span data-ttu-id="5f0b1-579">Свойство `subject` возвращает объект `Subject`, который предоставляет методы для получения и задания темы.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-579">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="5f0b1-580">Тип:</span><span class="sxs-lookup"><span data-stu-id="5f0b1-580">Type:</span></span>

*   <span data-ttu-id="5f0b1-581">String | [Subject](/javascript/api/outlook_1_6/office.subject)</span><span class="sxs-lookup"><span data-stu-id="5f0b1-581">String | [Subject](/javascript/api/outlook_1_6/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="5f0b1-582">Требования</span><span class="sxs-lookup"><span data-stu-id="5f0b1-582">Requirements</span></span>

|<span data-ttu-id="5f0b1-583">Требование</span><span class="sxs-lookup"><span data-stu-id="5f0b1-583">Requirement</span></span>| <span data-ttu-id="5f0b1-584">Значение</span><span class="sxs-lookup"><span data-stu-id="5f0b1-584">Value</span></span>|
|---|---|
|[<span data-ttu-id="5f0b1-585">Версия минимального набора требований для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="5f0b1-585">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5f0b1-586">1.0</span><span class="sxs-lookup"><span data-stu-id="5f0b1-586">1.0</span></span>|
|[<span data-ttu-id="5f0b1-587">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5f0b1-587">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5f0b1-588">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5f0b1-588">ReadItem</span></span>|
|[<span data-ttu-id="5f0b1-589">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5f0b1-589">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5f0b1-590">Cоздание или чтение</span><span class="sxs-lookup"><span data-stu-id="5f0b1-590">Compose or read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients"></a><span data-ttu-id="5f0b1-591">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="5f0b1-591">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

<span data-ttu-id="5f0b1-592">Предоставляет доступ получателей к строке **To (Кому)** в сообщении.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-592">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="5f0b1-593">Уровень доступа и тип объекта, зависит от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-593">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="5f0b1-594">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="5f0b1-594">Read mode</span></span>

<span data-ttu-id="5f0b1-p131">Свойство `to` возвращает массив, содержащий объект `EmailAddressDetails` для каждого получателя в строке **To (Кому)** сообщения. Коллекция может включать не более 100 элементов.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-p131">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="5f0b1-597">Режим создания</span><span class="sxs-lookup"><span data-stu-id="5f0b1-597">Compose mode</span></span>

<span data-ttu-id="5f0b1-598">Свойство `to` возвращает объект `Recipients`, предоставляющий методы для получения или обновления получателей, которые указаны в строке **To**(Кому) сообщения.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-598">The to`to` property returns a Recipients`Recipients` object that provides methods for manipulating the recipients on the **To** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="5f0b1-599">Тип:</span><span class="sxs-lookup"><span data-stu-id="5f0b1-599">Type:</span></span>

*   <span data-ttu-id="5f0b1-600">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="5f0b1-600">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="5f0b1-601">Требования</span><span class="sxs-lookup"><span data-stu-id="5f0b1-601">Requirements</span></span>

|<span data-ttu-id="5f0b1-602">Требование</span><span class="sxs-lookup"><span data-stu-id="5f0b1-602">Requirement</span></span>| <span data-ttu-id="5f0b1-603">Значение</span><span class="sxs-lookup"><span data-stu-id="5f0b1-603">Value</span></span>|
|---|---|
|[<span data-ttu-id="5f0b1-604">Версия минимального набора требований для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="5f0b1-604">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5f0b1-605">1.0</span><span class="sxs-lookup"><span data-stu-id="5f0b1-605">1.0</span></span>|
|[<span data-ttu-id="5f0b1-606">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5f0b1-606">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5f0b1-607">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5f0b1-607">ReadItem</span></span>|
|[<span data-ttu-id="5f0b1-608">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5f0b1-608">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5f0b1-609">Cоздание или чтение</span><span class="sxs-lookup"><span data-stu-id="5f0b1-609">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="5f0b1-610">Пример</span><span class="sxs-lookup"><span data-stu-id="5f0b1-610">Example</span></span>

```
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

### <a name="methods"></a><span data-ttu-id="5f0b1-611">Методы</span><span class="sxs-lookup"><span data-stu-id="5f0b1-611">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="5f0b1-612">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="5f0b1-612">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="5f0b1-613">Добавляет файл в сообщение или встречу в качестве вложения.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-613">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="5f0b1-614">Метод `addFileAttachmentAsync` передает файл по указанному универсальному коду ресурса (URI) и вкладывает его в элемент в форме создания.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-614">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="5f0b1-615">Идентификатор можно последовательно использовать с методом [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback), чтобы удалить вложение, добавленное во время текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-615">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="5f0b1-616">Параметры:</span><span class="sxs-lookup"><span data-stu-id="5f0b1-616">Parameters:</span></span>

|<span data-ttu-id="5f0b1-617">Имя</span><span class="sxs-lookup"><span data-stu-id="5f0b1-617">Name</span></span>| <span data-ttu-id="5f0b1-618">Тип</span><span class="sxs-lookup"><span data-stu-id="5f0b1-618">Type</span></span>| <span data-ttu-id="5f0b1-619">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="5f0b1-619">Attributes</span></span>| <span data-ttu-id="5f0b1-620">Описание</span><span class="sxs-lookup"><span data-stu-id="5f0b1-620">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="5f0b1-621">String</span><span class="sxs-lookup"><span data-stu-id="5f0b1-621">String</span></span>||<span data-ttu-id="5f0b1-p132">Универсальный код ресурса (URI), представляющий расположение файла, который нужно вложить в сообщение или встречу. Максимальная длина — 2048 символов.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-p132">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="5f0b1-624">String</span><span class="sxs-lookup"><span data-stu-id="5f0b1-624">String</span></span>||<span data-ttu-id="5f0b1-p133">Имя вложения, которое отображается при передаче вложения. Максимальная длина — 255 символов.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-p133">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="5f0b1-627">Объект</span><span class="sxs-lookup"><span data-stu-id="5f0b1-627">Object</span></span>| <span data-ttu-id="5f0b1-628">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="5f0b1-628">&lt;optional&gt;</span></span>|<span data-ttu-id="5f0b1-629">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-629">An object literal that contains one or more of the following properties.</span></span>|
| `options.asyncContext` | <span data-ttu-id="5f0b1-630">Объект</span><span class="sxs-lookup"><span data-stu-id="5f0b1-630">Object</span></span> | <span data-ttu-id="5f0b1-631">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="5f0b1-631">&lt;optional&gt;</span></span> | <span data-ttu-id="5f0b1-632">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-632">Developers can provide any object they wish to access in the callback method.</span></span> |
| `options.isInline` | <span data-ttu-id="5f0b1-633">Логический</span><span class="sxs-lookup"><span data-stu-id="5f0b1-633">Boolean</span></span> | <span data-ttu-id="5f0b1-634">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="5f0b1-634">&lt;optional&gt;</span></span> | <span data-ttu-id="5f0b1-635">Значение `true` указывает на то, что вложение будет встроено в текст сообщения и не должно отображаться в списке вложений.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-635">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
|`callback`| <span data-ttu-id="5f0b1-636">function</span><span class="sxs-lookup"><span data-stu-id="5f0b1-636">function</span></span>| <span data-ttu-id="5f0b1-637">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="5f0b1-637">&lt;optional&gt;</span></span>|<span data-ttu-id="5f0b1-638">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="5f0b1-638">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="5f0b1-639">После успешного выполнения идентификатор вложения будет представлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-639">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="5f0b1-640">Если добавить вложение не удастся, объект `asyncResult` будет содержать объект `Error` с описанием ошибки.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-640">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="5f0b1-641">Ошибки</span><span class="sxs-lookup"><span data-stu-id="5f0b1-641">Errors</span></span>

| <span data-ttu-id="5f0b1-642">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="5f0b1-642">Error code</span></span> | <span data-ttu-id="5f0b1-643">Описание</span><span class="sxs-lookup"><span data-stu-id="5f0b1-643">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="5f0b1-644">Вложение превышает максимальный размер.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-644">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="5f0b1-645">Расширение вложения не поддерживается.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-645">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="5f0b1-646">Сообщение или встреча содержат слишком много вложений.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-646">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="5f0b1-647">Требования</span><span class="sxs-lookup"><span data-stu-id="5f0b1-647">Requirements</span></span>

|<span data-ttu-id="5f0b1-648">Требование</span><span class="sxs-lookup"><span data-stu-id="5f0b1-648">Requirement</span></span>| <span data-ttu-id="5f0b1-649">Значение</span><span class="sxs-lookup"><span data-stu-id="5f0b1-649">Value</span></span>|
|---|---|
|[<span data-ttu-id="5f0b1-650">Версия минимального набора требований для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="5f0b1-650">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5f0b1-651">1.1</span><span class="sxs-lookup"><span data-stu-id="5f0b1-651">1.1</span></span>|
|[<span data-ttu-id="5f0b1-652">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5f0b1-652">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5f0b1-653">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="5f0b1-653">ReadWriteItem</span></span>|
|[<span data-ttu-id="5f0b1-654">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5f0b1-654">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5f0b1-655">Создание</span><span class="sxs-lookup"><span data-stu-id="5f0b1-655">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="5f0b1-656">Примеры</span><span class="sxs-lookup"><span data-stu-id="5f0b1-656">Examples</span></span>

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

<span data-ttu-id="5f0b1-657">В приведенном ниже примере файл изображения добавляется в качестве встроенного вложения, а ссылка на вложение добавляется в текст сообщения.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-657">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

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

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="5f0b1-658">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="5f0b1-658">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="5f0b1-659">Добавляет к сообщению элемент Exchange, например, сообщение, в виде вложения.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-659">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="5f0b1-p134">С помощью метода `addItemAttachmentAsync` в элемент формы создания можно вложить элемент с указанным идентификатором Exchange. Если указать метод обратного вызова, то этот метод вызывается с помощью параметра `asyncResult`, который содержит идентификатор вложения или код, указывающий на ошибки, которые произошли при вложении элемента. При необходимости можно использовать параметр `options` для передачи сведений о состоянии в метод обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-p134">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="5f0b1-663">Идентификатор можно последовательно использовать с методом [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback), чтобы удалить вложение, добавленное во время текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-663">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="5f0b1-664">Если ваша надстройка Office выполняется в веб-приложении Outlook, метод `addItemAttachmentAsync` обеспечивает вложение элементов в элементы, отличные от редактируемого. Однако это действие не рекомендуется выполнять, так как оно не поддерживается.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-664">If your Office add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="5f0b1-665">Параметры:</span><span class="sxs-lookup"><span data-stu-id="5f0b1-665">Parameters:</span></span>

|<span data-ttu-id="5f0b1-666">Имя</span><span class="sxs-lookup"><span data-stu-id="5f0b1-666">Name</span></span>| <span data-ttu-id="5f0b1-667">Тип</span><span class="sxs-lookup"><span data-stu-id="5f0b1-667">Type</span></span>| <span data-ttu-id="5f0b1-668">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="5f0b1-668">Attributes</span></span>| <span data-ttu-id="5f0b1-669">Описание</span><span class="sxs-lookup"><span data-stu-id="5f0b1-669">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="5f0b1-670">String</span><span class="sxs-lookup"><span data-stu-id="5f0b1-670">String</span></span>||<span data-ttu-id="5f0b1-p135">Идентификатор Exchange для вкладываемого элемента. Максимальная длина — 100 символов.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-p135">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="5f0b1-673">String</span><span class="sxs-lookup"><span data-stu-id="5f0b1-673">String</span></span>||<span data-ttu-id="5f0b1-p136">Тема вкладываемого элемента. Максимальная длина — 255 символов.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-p136">The sujbect of the item to be attached. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="5f0b1-676">Объект</span><span class="sxs-lookup"><span data-stu-id="5f0b1-676">Object</span></span>| <span data-ttu-id="5f0b1-677">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="5f0b1-677">&lt;optional&gt;</span></span>|<span data-ttu-id="5f0b1-678">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-678">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="5f0b1-679">Объект</span><span class="sxs-lookup"><span data-stu-id="5f0b1-679">Object</span></span>| <span data-ttu-id="5f0b1-680">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="5f0b1-680">&lt;optional&gt;</span></span>|<span data-ttu-id="5f0b1-681">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-681">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="5f0b1-682">function</span><span class="sxs-lookup"><span data-stu-id="5f0b1-682">function</span></span>| <span data-ttu-id="5f0b1-683">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="5f0b1-683">&lt;optional&gt;</span></span>|<span data-ttu-id="5f0b1-684">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="5f0b1-684">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="5f0b1-685">После успешного выполнения идентификатор вложения будет представлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-685">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="5f0b1-686">Если добавить вложение не удастся, объект `asyncResult` будет содержать объект `Error` с описанием ошибки.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-686">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="5f0b1-687">Ошибки</span><span class="sxs-lookup"><span data-stu-id="5f0b1-687">Errors</span></span>

| <span data-ttu-id="5f0b1-688">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="5f0b1-688">Error code</span></span> | <span data-ttu-id="5f0b1-689">Описание</span><span class="sxs-lookup"><span data-stu-id="5f0b1-689">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="5f0b1-690">Сообщение или встреча содержат слишком много вложений.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-690">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="5f0b1-691">Требования</span><span class="sxs-lookup"><span data-stu-id="5f0b1-691">Requirements</span></span>

|<span data-ttu-id="5f0b1-692">Требование</span><span class="sxs-lookup"><span data-stu-id="5f0b1-692">Requirement</span></span>| <span data-ttu-id="5f0b1-693">Значение</span><span class="sxs-lookup"><span data-stu-id="5f0b1-693">Value</span></span>|
|---|---|
|[<span data-ttu-id="5f0b1-694">Версия минимального набора требований для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="5f0b1-694">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5f0b1-695">1.1</span><span class="sxs-lookup"><span data-stu-id="5f0b1-695">1.1</span></span>|
|[<span data-ttu-id="5f0b1-696">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5f0b1-696">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5f0b1-697">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="5f0b1-697">ReadWriteItem</span></span>|
|[<span data-ttu-id="5f0b1-698">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5f0b1-698">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5f0b1-699">Создание</span><span class="sxs-lookup"><span data-stu-id="5f0b1-699">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="5f0b1-700">Пример</span><span class="sxs-lookup"><span data-stu-id="5f0b1-700">Example</span></span>

<span data-ttu-id="5f0b1-701">В следующем примере существующий элемент Outlook добавляется в виде вложения с именем `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-701">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

####  <a name="close"></a><span data-ttu-id="5f0b1-702">close()</span><span class="sxs-lookup"><span data-stu-id="5f0b1-702">close()</span></span>

<span data-ttu-id="5f0b1-703">Закрывает текущий создаваемый элемент.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-703">Closes the current item that is being composed.</span></span>

<span data-ttu-id="5f0b1-p137">Работа метода `close` зависит от текущего состояния создаваемого элемента. Если элемент содержит несохраненные изменения, клиент предложит пользователю сохранить или отклонить их либо отменить действие закрытия.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-p137">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="5f0b1-706">В Outlook в Интернете, если элемент является встречей, и он был ранее сохранен с помощью `saveAsync`, пользователю предлагается сохранить, отменить или удалить, даже если не произошло каких-либо изменений, поскольку элемент был последним сохраненным.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-706">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="5f0b1-707">Если в клиенте Outlook для настольных ПК сообщение представляет собой ответ в тексте, метод `close` не работает.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-707">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="5f0b1-708">Требования</span><span class="sxs-lookup"><span data-stu-id="5f0b1-708">Requirements</span></span>

|<span data-ttu-id="5f0b1-709">Требование</span><span class="sxs-lookup"><span data-stu-id="5f0b1-709">Requirement</span></span>| <span data-ttu-id="5f0b1-710">Значение</span><span class="sxs-lookup"><span data-stu-id="5f0b1-710">Value</span></span>|
|---|---|
|[<span data-ttu-id="5f0b1-711">Версия минимального набора требований для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="5f0b1-711">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5f0b1-712">1.3</span><span class="sxs-lookup"><span data-stu-id="5f0b1-712">1.3</span></span>|
|[<span data-ttu-id="5f0b1-713">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5f0b1-713">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5f0b1-714">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="5f0b1-714">Restricted</span></span>|
|[<span data-ttu-id="5f0b1-715">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5f0b1-715">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5f0b1-716">Создание</span><span class="sxs-lookup"><span data-stu-id="5f0b1-716">Compose</span></span>|

#### <a name="displayreplyallformformdata"></a><span data-ttu-id="5f0b1-717">displayReplyAllForm(formData)</span><span class="sxs-lookup"><span data-stu-id="5f0b1-717">displayReplyAllForm(formData)</span></span>

<span data-ttu-id="5f0b1-718">Отображает форму ответа, включающую отправителя и всех получателей выбранного сообщения или организатора и всех участников выбранной встречи.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-718">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="5f0b1-719">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-719">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="5f0b1-720">В веб-приложении Outlook форма ответа отображается в виде всплывающей формы в представлении с 3 либо 1 или 2 колонками.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-720">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="5f0b1-721">Если любой строчный параметр превышает указанные для него ограничения, `displayReplyAllForm` возвращает исключение.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-721">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="5f0b1-p138">Если в параметре `formData.attachments` указаны вложения, Outlook и веб-приложение Outlook пытаются скачать их и вложить в форму ответа. Если какие-либо вложения добавить не удается, в форме отображается сообщение об ошибке. Если сообщения об ошибках не предусмотрены, то они не отображаются.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-p138">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="5f0b1-725">Параметры:</span><span class="sxs-lookup"><span data-stu-id="5f0b1-725">Parameters:</span></span>

| <span data-ttu-id="5f0b1-726">Имя</span><span class="sxs-lookup"><span data-stu-id="5f0b1-726">Name</span></span> | <span data-ttu-id="5f0b1-727">Тип</span><span class="sxs-lookup"><span data-stu-id="5f0b1-727">Type</span></span> | <span data-ttu-id="5f0b1-728">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="5f0b1-728">Attributes</span></span> | <span data-ttu-id="5f0b1-729">Описание</span><span class="sxs-lookup"><span data-stu-id="5f0b1-729">Description</span></span> |
|---|---|---|---|
|`formData`| <span data-ttu-id="5f0b1-730">String | Object</span><span class="sxs-lookup"><span data-stu-id="5f0b1-730">String &#124; Object</span></span>| |<span data-ttu-id="5f0b1-p139">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-p139">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="5f0b1-733">**ИЛИ**</span><span class="sxs-lookup"><span data-stu-id="5f0b1-733">**OR**</span></span><br/><span data-ttu-id="5f0b1-p140">Объект, который содержит текст или данные вложения и функцию обратного вызова. Ниже представлено определение этого объекта.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-p140">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="5f0b1-736">String</span><span class="sxs-lookup"><span data-stu-id="5f0b1-736">String</span></span> | <span data-ttu-id="5f0b1-737">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="5f0b1-737">&lt;optional&gt;</span></span> | <span data-ttu-id="5f0b1-p141">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-p141">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="5f0b1-740">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="5f0b1-740">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="5f0b1-741">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="5f0b1-741">&lt;optional&gt;</span></span> | <span data-ttu-id="5f0b1-742">Массив объектов JSON, представляющих собой вложенные файлы или элементы.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-742">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="5f0b1-743">String</span><span class="sxs-lookup"><span data-stu-id="5f0b1-743">String</span></span> | | <span data-ttu-id="5f0b1-p142">Указывает тип вложения. Допустимые значения: `file` для вложенного файла и `item` для вложенного элемента.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-p142">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="5f0b1-746">String</span><span class="sxs-lookup"><span data-stu-id="5f0b1-746">String</span></span> | | <span data-ttu-id="5f0b1-747">Строка, содержащая имя вложения, длиной до 255 символов.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-747">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="5f0b1-748">String</span><span class="sxs-lookup"><span data-stu-id="5f0b1-748">String</span></span> | | <span data-ttu-id="5f0b1-p143">Используется, только если свойству `type` задано значение `file`. URI расположения файла.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-p143">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.isInline` | <span data-ttu-id="5f0b1-751">Логический</span><span class="sxs-lookup"><span data-stu-id="5f0b1-751">Boolean</span></span> | | <span data-ttu-id="5f0b1-p144">Используется только в том случае, если свойству `type` задано значение `file`. Значение `true` указывает на то, что вложение будет встроено в текст сообщения и не должно отображаться в списке вложений.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-p144">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="5f0b1-754">String</span><span class="sxs-lookup"><span data-stu-id="5f0b1-754">String</span></span> | | <span data-ttu-id="5f0b1-p145">Используется, только если свойству `type` задано значение `item`. Идентификатор вложения EWS. Это строка длиной до 100 символов.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-p145">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="5f0b1-758">function</span><span class="sxs-lookup"><span data-stu-id="5f0b1-758">function</span></span> | <span data-ttu-id="5f0b1-759">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="5f0b1-759">&lt;optional&gt;</span></span> | <span data-ttu-id="5f0b1-760">По завершении работы метода функция, переданная параметру `callback`, вызывается с помощью одного параметра `asyncResult`, представляющего собой объект [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="5f0b1-760">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="5f0b1-761">Требования</span><span class="sxs-lookup"><span data-stu-id="5f0b1-761">Requirements</span></span>

|<span data-ttu-id="5f0b1-762">Требование</span><span class="sxs-lookup"><span data-stu-id="5f0b1-762">Requirement</span></span>| <span data-ttu-id="5f0b1-763">Значение</span><span class="sxs-lookup"><span data-stu-id="5f0b1-763">Value</span></span>|
|---|---|
|[<span data-ttu-id="5f0b1-764">Версия минимального набора требований для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="5f0b1-764">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5f0b1-765">1.0</span><span class="sxs-lookup"><span data-stu-id="5f0b1-765">1.0</span></span>|
|[<span data-ttu-id="5f0b1-766">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5f0b1-766">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5f0b1-767">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5f0b1-767">ReadItem</span></span>|
|[<span data-ttu-id="5f0b1-768">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5f0b1-768">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5f0b1-769">Чтение</span><span class="sxs-lookup"><span data-stu-id="5f0b1-769">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="5f0b1-770">Примеры</span><span class="sxs-lookup"><span data-stu-id="5f0b1-770">Examples</span></span>

<span data-ttu-id="5f0b1-771">Приведенный ниже код передает строку в функцию `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-771">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="5f0b1-772">Ответ с пустым текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-772">Reply with an empty body.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="5f0b1-773">Ответ только с текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-773">Reply with just a body.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="5f0b1-774">Ответ с текстом сообщения и вложенным файлом.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-774">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="5f0b1-775">Ответ с текстом сообщения и вложенным элементом.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-775">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="5f0b1-776">Ответ с текстом сообщения, вложенным файлом, вложенным элементом и обратным вызовом.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-776">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata"></a><span data-ttu-id="5f0b1-777">displayReplyForm(formData)</span><span class="sxs-lookup"><span data-stu-id="5f0b1-777">displayReplyForm(formData)</span></span>

<span data-ttu-id="5f0b1-778">Отображает форму ответа, включающую только отправителя выбранного сообщения или организатора выбранной встречи.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-778">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="5f0b1-779">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-779">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="5f0b1-780">В веб-приложении Outlook форма ответа отображается в виде всплывающей формы в представлении с 3 либо 1 или 2 колонками.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-780">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="5f0b1-781">Если любой строчный параметр превышает указанные для него ограничения, `displayReplyForm` возвращает исключение.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-781">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="5f0b1-p146">Если в параметре `formData.attachments` указаны вложения, Outlook и веб-приложение Outlook пытаются скачать их и вложить в форму ответа. Если какие-либо вложения добавить не удается, в форме отображается сообщение об ошибке. Если сообщения об ошибках не предусмотрены, то они не отображаются.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-p146">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="5f0b1-785">Параметры:</span><span class="sxs-lookup"><span data-stu-id="5f0b1-785">Parameters:</span></span>

| <span data-ttu-id="5f0b1-786">Имя</span><span class="sxs-lookup"><span data-stu-id="5f0b1-786">Name</span></span> | <span data-ttu-id="5f0b1-787">Тип</span><span class="sxs-lookup"><span data-stu-id="5f0b1-787">Type</span></span> | <span data-ttu-id="5f0b1-788">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="5f0b1-788">Attributes</span></span> | <span data-ttu-id="5f0b1-789">Описание</span><span class="sxs-lookup"><span data-stu-id="5f0b1-789">Description</span></span> |
|---|---|---|---|
|`formData`| <span data-ttu-id="5f0b1-790">String | Object</span><span class="sxs-lookup"><span data-stu-id="5f0b1-790">String &#124; Object</span></span>| | <span data-ttu-id="5f0b1-p147">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-p147">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="5f0b1-793">**ИЛИ**</span><span class="sxs-lookup"><span data-stu-id="5f0b1-793">**OR**</span></span><br/><span data-ttu-id="5f0b1-p148">Объект, который содержит текст или данные вложения и функцию обратного вызова. Ниже представлено определение этого объекта.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-p148">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="5f0b1-796">String</span><span class="sxs-lookup"><span data-stu-id="5f0b1-796">String</span></span> | <span data-ttu-id="5f0b1-797">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="5f0b1-797">&lt;optional&gt;</span></span> | <span data-ttu-id="5f0b1-p149">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-p149">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="5f0b1-800">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="5f0b1-800">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="5f0b1-801">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="5f0b1-801">&lt;optional&gt;</span></span> | <span data-ttu-id="5f0b1-802">Массив объектов JSON, представляющих собой вложенные файлы или элементы.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-802">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="5f0b1-803">String</span><span class="sxs-lookup"><span data-stu-id="5f0b1-803">String</span></span> | | <span data-ttu-id="5f0b1-p150">Указывает тип вложения. Допустимые значения: `file` для вложенного файла и `item` для вложенного элемента.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-p150">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="5f0b1-806">String</span><span class="sxs-lookup"><span data-stu-id="5f0b1-806">String</span></span> | | <span data-ttu-id="5f0b1-807">Строка, содержащая имя вложения, длиной до 255 символов.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-807">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="5f0b1-808">String</span><span class="sxs-lookup"><span data-stu-id="5f0b1-808">String</span></span> | | <span data-ttu-id="5f0b1-p151">Используется, только если свойству `type` задано значение `file`. URI расположения файла.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-p151">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.isInline` | <span data-ttu-id="5f0b1-811">Логический</span><span class="sxs-lookup"><span data-stu-id="5f0b1-811">Boolean</span></span> | | <span data-ttu-id="5f0b1-p152">Используется только в том случае, если свойству `type` задано значение `file`. Значение `true` указывает на то, что вложение будет встроено в текст сообщения и не должно отображаться в списке вложений.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-p152">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="5f0b1-814">String</span><span class="sxs-lookup"><span data-stu-id="5f0b1-814">String</span></span> | | <span data-ttu-id="5f0b1-p153">Используется, только если свойству `type` задано значение `item`. Идентификатор вложения EWS. Это строка длиной до 100 символов.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-p153">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="5f0b1-818">function</span><span class="sxs-lookup"><span data-stu-id="5f0b1-818">function</span></span> | <span data-ttu-id="5f0b1-819">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="5f0b1-819">&lt;optional&gt;</span></span> | <span data-ttu-id="5f0b1-820">По завершении работы метода функция, переданная параметру `callback`, вызывается с помощью одного параметра `asyncResult`, представляющего собой объект [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="5f0b1-820">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="5f0b1-821">Требования</span><span class="sxs-lookup"><span data-stu-id="5f0b1-821">Requirements</span></span>

|<span data-ttu-id="5f0b1-822">Требование</span><span class="sxs-lookup"><span data-stu-id="5f0b1-822">Requirement</span></span>| <span data-ttu-id="5f0b1-823">Значение</span><span class="sxs-lookup"><span data-stu-id="5f0b1-823">Value</span></span>|
|---|---|
|[<span data-ttu-id="5f0b1-824">Версия минимального набора требований для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="5f0b1-824">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5f0b1-825">1.0</span><span class="sxs-lookup"><span data-stu-id="5f0b1-825">1.0</span></span>|
|[<span data-ttu-id="5f0b1-826">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5f0b1-826">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5f0b1-827">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5f0b1-827">ReadItem</span></span>|
|[<span data-ttu-id="5f0b1-828">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5f0b1-828">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5f0b1-829">Чтение</span><span class="sxs-lookup"><span data-stu-id="5f0b1-829">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="5f0b1-830">Примеры</span><span class="sxs-lookup"><span data-stu-id="5f0b1-830">Examples</span></span>

<span data-ttu-id="5f0b1-831">Приведенный ниже код передает строку в функцию `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-831">The following code passes a string to the `displayReplyForm` function.</span></span>

```
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="5f0b1-832">Ответ с пустым текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-832">Reply with an empty body.</span></span>

```
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="5f0b1-833">Ответ только с текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-833">Reply with just a body.</span></span>

```
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="5f0b1-834">Ответ с текстом сообщения и вложенным файлом.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-834">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="5f0b1-835">Ответ с текстом сообщения и вложенным элементом.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-835">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="5f0b1-836">Ответ с текстом сообщения, вложенным файлом, вложенным элементом и обратным вызовом.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-836">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlook16officeentities"></a><span data-ttu-id="5f0b1-837">getEntities() → {[Entities](/javascript/api/outlook_1_6/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="5f0b1-837">getEntities() → {[Entities](/javascript/api/outlook_1_6/office.entities)}</span></span>

<span data-ttu-id="5f0b1-838">Получает сущности, обнаруженные в выбранном элементе.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-838">Gets the entities found in the selected item.</span></span>

> [!NOTE]
> <span data-ttu-id="5f0b1-839">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-839">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="5f0b1-840">Требования</span><span class="sxs-lookup"><span data-stu-id="5f0b1-840">Requirements</span></span>

|<span data-ttu-id="5f0b1-841">Требование</span><span class="sxs-lookup"><span data-stu-id="5f0b1-841">Requirement</span></span>| <span data-ttu-id="5f0b1-842">Значение</span><span class="sxs-lookup"><span data-stu-id="5f0b1-842">Value</span></span>|
|---|---|
|[<span data-ttu-id="5f0b1-843">Версия минимального набора требований для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="5f0b1-843">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5f0b1-844">1.0</span><span class="sxs-lookup"><span data-stu-id="5f0b1-844">1.0</span></span>|
|[<span data-ttu-id="5f0b1-845">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5f0b1-845">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5f0b1-846">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5f0b1-846">ReadItem</span></span>|
|[<span data-ttu-id="5f0b1-847">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5f0b1-847">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5f0b1-848">Чтение</span><span class="sxs-lookup"><span data-stu-id="5f0b1-848">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="5f0b1-849">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="5f0b1-849">Returns:</span></span>

<span data-ttu-id="5f0b1-850">Тип: [Entities](/javascript/api/outlook_1_6/office.entities)</span><span class="sxs-lookup"><span data-stu-id="5f0b1-850">Type: [Entities](/javascript/api/outlook_1_6/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="5f0b1-851">Пример</span><span class="sxs-lookup"><span data-stu-id="5f0b1-851">Example</span></span>

<span data-ttu-id="5f0b1-852">Ниже приведен пример получения доступа к сущностям контактов в тексте текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-852">The following example accesses the contacts entities on the current item.</span></span>

```
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook16officecontactmeetingsuggestionjavascriptapioutlook16officemeetingsuggestionphonenumberjavascriptapioutlook16officephonenumbertasksuggestionjavascriptapioutlook16officetasksuggestion"></a><span data-ttu-id="5f0b1-853">getEntitiesByType(entityType) → (допускающий значение NULL) {Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="5f0b1-853">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))>}</span></span>

<span data-ttu-id="5f0b1-854">Получает массив всех сущностей указанного типа, обнаруженных в тексте выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-854">Gets an array of all the entities of the specified entity type found in the selected item.</span></span>

> [!NOTE]
> <span data-ttu-id="5f0b1-855">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-855">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="5f0b1-856">Параметры:</span><span class="sxs-lookup"><span data-stu-id="5f0b1-856">Parameters:</span></span>

|<span data-ttu-id="5f0b1-857">Имя</span><span class="sxs-lookup"><span data-stu-id="5f0b1-857">Name</span></span>| <span data-ttu-id="5f0b1-858">Тип</span><span class="sxs-lookup"><span data-stu-id="5f0b1-858">Type</span></span>| <span data-ttu-id="5f0b1-859">Описание</span><span class="sxs-lookup"><span data-stu-id="5f0b1-859">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="5f0b1-860">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="5f0b1-860">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook_1_6/office.mailboxenums.entitytype)|<span data-ttu-id="5f0b1-861">Одно из значений перечисления EntityType.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-861">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="5f0b1-862">Требования</span><span class="sxs-lookup"><span data-stu-id="5f0b1-862">Requirements</span></span>

|<span data-ttu-id="5f0b1-863">Требование</span><span class="sxs-lookup"><span data-stu-id="5f0b1-863">Requirement</span></span>| <span data-ttu-id="5f0b1-864">Значение</span><span class="sxs-lookup"><span data-stu-id="5f0b1-864">Value</span></span>|
|---|---|
|[<span data-ttu-id="5f0b1-865">Версия минимального набора требований для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="5f0b1-865">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5f0b1-866">1.0</span><span class="sxs-lookup"><span data-stu-id="5f0b1-866">1.0</span></span>|
|[<span data-ttu-id="5f0b1-867">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5f0b1-867">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5f0b1-868">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="5f0b1-868">Restricted</span></span>|
|[<span data-ttu-id="5f0b1-869">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5f0b1-869">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5f0b1-870">Чтение</span><span class="sxs-lookup"><span data-stu-id="5f0b1-870">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="5f0b1-871">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="5f0b1-871">Returns:</span></span>

<span data-ttu-id="5f0b1-872">Если значение, переданное в `entityType`, не является допустимым членом перечисления `EntityType`, метод возвращает значение NULL.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-872">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="5f0b1-873">Если в тексте элемента отсутствуют сущности указанного типа, метод возвращает пустой массив.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-873">If no entities of the specified type are present on the item, the method returns an empty array.</span></span> <span data-ttu-id="5f0b1-874">В противном случае тип объектов в возвращаемом массиве зависит от типа сущности, запрошенной в параметре `entityType`.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-874">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="5f0b1-875">Хотя минимальный уровень разрешений для использования этого метода — **Restricted**, для некоторых типов сущностей требуется доступ на уровне **ReadItem**, как указано в приведенной ниже таблице.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-875">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="5f0b1-876">Значение параметра `entityType`</span><span class="sxs-lookup"><span data-stu-id="5f0b1-876">Value of `entityType`</span></span> | <span data-ttu-id="5f0b1-877">Тип объектов в возвращаемом массиве</span><span class="sxs-lookup"><span data-stu-id="5f0b1-877">Type of objects in returned array</span></span> | <span data-ttu-id="5f0b1-878">Необходимый уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5f0b1-878">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="5f0b1-879">String</span><span class="sxs-lookup"><span data-stu-id="5f0b1-879">String</span></span> | <span data-ttu-id="5f0b1-880">**Ограниченный доступ**</span><span class="sxs-lookup"><span data-stu-id="5f0b1-880">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="5f0b1-881">Contact</span><span class="sxs-lookup"><span data-stu-id="5f0b1-881">Contact</span></span> | <span data-ttu-id="5f0b1-882">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="5f0b1-882">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="5f0b1-883">String</span><span class="sxs-lookup"><span data-stu-id="5f0b1-883">String</span></span> | <span data-ttu-id="5f0b1-884">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="5f0b1-884">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="5f0b1-885">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="5f0b1-885">MeetingSuggestion</span></span> | <span data-ttu-id="5f0b1-886">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="5f0b1-886">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="5f0b1-887">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="5f0b1-887">PhoneNumber</span></span> | <span data-ttu-id="5f0b1-888">**Ограниченный доступ**</span><span class="sxs-lookup"><span data-stu-id="5f0b1-888">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="5f0b1-889">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="5f0b1-889">TaskSuggestion</span></span> | <span data-ttu-id="5f0b1-890">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="5f0b1-890">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="5f0b1-891">String</span><span class="sxs-lookup"><span data-stu-id="5f0b1-891">String</span></span> | <span data-ttu-id="5f0b1-892">**Ограниченный доступ**</span><span class="sxs-lookup"><span data-stu-id="5f0b1-892">**Restricted**</span></span> |

<span data-ttu-id="5f0b1-893">Тип: Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="5f0b1-893">Type: Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="5f0b1-894">Пример</span><span class="sxs-lookup"><span data-stu-id="5f0b1-894">Example</span></span>

<span data-ttu-id="5f0b1-895">В примере ниже показано, как получить доступ к массиву строк, которые представляют собой почтовые адреса в тексте текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-895">The following example shows how to access an array of strings that represent postal addresses in the subject or body of the current item.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook16officecontactmeetingsuggestionjavascriptapioutlook16officemeetingsuggestionphonenumberjavascriptapioutlook16officephonenumbertasksuggestionjavascriptapioutlook16officetasksuggestion"></a><span data-ttu-id="5f0b1-896">getFilteredEntitiesByName(name) → (допускающий значение NULL) {Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="5f0b1-896">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))>}</span></span>

<span data-ttu-id="5f0b1-897">Возвращает известные сущности в выбранном элементе, которые проходят через именованный фильтр, определяемый в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-897">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="5f0b1-898">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-898">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="5f0b1-899">Метод `getFilteredEntitiesByName` возвращает сущности, соответствующие регулярному выражению, которое определяется в элементе правила [ItemHasKnownEntity](/javascript/office/manifest/rule#itemhasknownentity-rule) в XML-файле манифеста, с использованием указанного значения элемента `FilterName`.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-899">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/javascript/office/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="5f0b1-900">Параметры:</span><span class="sxs-lookup"><span data-stu-id="5f0b1-900">Parameters:</span></span>

|<span data-ttu-id="5f0b1-901">Имя</span><span class="sxs-lookup"><span data-stu-id="5f0b1-901">Name</span></span>| <span data-ttu-id="5f0b1-902">Тип</span><span class="sxs-lookup"><span data-stu-id="5f0b1-902">Type</span></span>| <span data-ttu-id="5f0b1-903">Описание</span><span class="sxs-lookup"><span data-stu-id="5f0b1-903">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="5f0b1-904">String</span><span class="sxs-lookup"><span data-stu-id="5f0b1-904">String</span></span>|<span data-ttu-id="5f0b1-905">Имя элемента правила `ItemHasKnownEntity`, определяющее соответствующий фильтр.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-905">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="5f0b1-906">Требования</span><span class="sxs-lookup"><span data-stu-id="5f0b1-906">Requirements</span></span>

|<span data-ttu-id="5f0b1-907">Требование</span><span class="sxs-lookup"><span data-stu-id="5f0b1-907">Requirement</span></span>| <span data-ttu-id="5f0b1-908">Значение</span><span class="sxs-lookup"><span data-stu-id="5f0b1-908">Value</span></span>|
|---|---|
|[<span data-ttu-id="5f0b1-909">Версия минимального набора требований для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="5f0b1-909">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5f0b1-910">1.0</span><span class="sxs-lookup"><span data-stu-id="5f0b1-910">1.0</span></span>|
|[<span data-ttu-id="5f0b1-911">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5f0b1-911">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5f0b1-912">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5f0b1-912">ReadItem</span></span>|
|[<span data-ttu-id="5f0b1-913">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5f0b1-913">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5f0b1-914">Чтение</span><span class="sxs-lookup"><span data-stu-id="5f0b1-914">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="5f0b1-915">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="5f0b1-915">Returns:</span></span>

<span data-ttu-id="5f0b1-p155">Если в манифесте нет элемента `ItemHasKnownEntity` со значением `FilterName`, соответствующим параметру `name`, метод возвращает `null`. Если параметр `name` соответствует элементу `ItemHasKnownEntity` в манифесте, но при этом в текущем элементе нет соответствующих сущностей, метод возвращает пустой массив.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-p155">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="5f0b1-918">Тип: Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="5f0b1-918">Type: Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))></span></span>

#### <a name="getregexmatches--object"></a><span data-ttu-id="5f0b1-919">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="5f0b1-919">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="5f0b1-920">Возвращает строчные значения в выбранном элементе, которые соответствуют регулярным выражениям, определенным в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-920">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="5f0b1-921">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-921">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="5f0b1-p156">Метод `getRegExMatches` возвращает строки, соответствующие регулярному выражению, которое определяется в каждом элементе правила `ItemHasRegularExpressionMatch` или `ItemHasKnownEntity` в XML-файле манифеста. Для правила `ItemHasRegularExpressionMatch` соответствующую строку должно содержать свойство элемента, указанного этим правилом. Простой тип `PropertyName` определяет поддерживаемые свойства.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-p156">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="5f0b1-925">Например, рассмотрим манифест надстройки, который содержит указанный ниже элемент `Rule`.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-925">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="5f0b1-926">Объект, возвращаемый методом `getRegExMatches`, будет содержать два свойства: `fruits` и `veggies`.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-926">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="5f0b1-p157">Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты. Лучше используйте метод [`Body.getAsync`](/javascript/api/outlook_1_6/office.body#getasync-coerciontype--options--callback-) для этого.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-p157">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook_1_6/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="5f0b1-930">Требования</span><span class="sxs-lookup"><span data-stu-id="5f0b1-930">Requirements</span></span>

|<span data-ttu-id="5f0b1-931">Требование</span><span class="sxs-lookup"><span data-stu-id="5f0b1-931">Requirement</span></span>| <span data-ttu-id="5f0b1-932">Значение</span><span class="sxs-lookup"><span data-stu-id="5f0b1-932">Value</span></span>|
|---|---|
|[<span data-ttu-id="5f0b1-933">Версия минимального набора требований для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="5f0b1-933">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5f0b1-934">1.0</span><span class="sxs-lookup"><span data-stu-id="5f0b1-934">1.0</span></span>|
|[<span data-ttu-id="5f0b1-935">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5f0b1-935">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5f0b1-936">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5f0b1-936">ReadItem</span></span>|
|[<span data-ttu-id="5f0b1-937">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5f0b1-937">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5f0b1-938">Чтение</span><span class="sxs-lookup"><span data-stu-id="5f0b1-938">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="5f0b1-939">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="5f0b1-939">Returns:</span></span>

<span data-ttu-id="5f0b1-p158">Объект, содержащий массив строк, которые соответствуют регулярным выражениям, определяемым в XML-файле манифеста. Имя каждого массива равно соответствующему значению атрибута `RegExName` правила сопоставления `ItemHasRegularExpressionMatch` или атрибута `FilterName` правила сопоставления `ItemHasKnownEntity`.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-p158">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="5f0b1-942">

<dt>Тип</dt>

</span><span class="sxs-lookup"><span data-stu-id="5f0b1-942">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="5f0b1-943">Объект</span><span class="sxs-lookup"><span data-stu-id="5f0b1-943">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="5f0b1-944">Пример</span><span class="sxs-lookup"><span data-stu-id="5f0b1-944">Example</span></span>

<span data-ttu-id="5f0b1-945">В приведенном ниже примере показано, как получить доступ к массиву совпадений с элементами `fruits` и `veggies` правил активации регулярных выражений, указанными в манифесте.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-945">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veges = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="5f0b1-946">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="5f0b1-946">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="5f0b1-947">Возвращает строчные значения в выбранном элементе, которые соответствуют именованному регулярному выражению, определенному в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-947">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="5f0b1-948">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-948">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="5f0b1-949">Метод `getRegExMatchesByName` возвращает строки, соответствующие регулярному выражению, которое определяется в элементе правила `ItemHasRegularExpressionMatch` в XML-файле манифеста, с использованием указанного значения элемента `RegExName`.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-949">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="5f0b1-p159">Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-p159">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="5f0b1-952">Параметры:</span><span class="sxs-lookup"><span data-stu-id="5f0b1-952">Parameters:</span></span>

|<span data-ttu-id="5f0b1-953">Имя</span><span class="sxs-lookup"><span data-stu-id="5f0b1-953">Name</span></span>| <span data-ttu-id="5f0b1-954">Тип</span><span class="sxs-lookup"><span data-stu-id="5f0b1-954">Type</span></span>| <span data-ttu-id="5f0b1-955">Описание</span><span class="sxs-lookup"><span data-stu-id="5f0b1-955">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="5f0b1-956">String</span><span class="sxs-lookup"><span data-stu-id="5f0b1-956">String</span></span>|<span data-ttu-id="5f0b1-957">Имя элемента правила `ItemHasRegularExpressionMatch`, определяющее соответствующий фильтр.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-957">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="5f0b1-958">Требования</span><span class="sxs-lookup"><span data-stu-id="5f0b1-958">Requirements</span></span>

|<span data-ttu-id="5f0b1-959">Требование</span><span class="sxs-lookup"><span data-stu-id="5f0b1-959">Requirement</span></span>| <span data-ttu-id="5f0b1-960">Значение</span><span class="sxs-lookup"><span data-stu-id="5f0b1-960">Value</span></span>|
|---|---|
|[<span data-ttu-id="5f0b1-961">Версия минимального набора требований для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="5f0b1-961">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5f0b1-962">1.0</span><span class="sxs-lookup"><span data-stu-id="5f0b1-962">1.0</span></span>|
|[<span data-ttu-id="5f0b1-963">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5f0b1-963">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5f0b1-964">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5f0b1-964">ReadItem</span></span>|
|[<span data-ttu-id="5f0b1-965">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5f0b1-965">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5f0b1-966">Чтение</span><span class="sxs-lookup"><span data-stu-id="5f0b1-966">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="5f0b1-967">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="5f0b1-967">Returns:</span></span>

<span data-ttu-id="5f0b1-968">Массив строк, соответствующих регулярному выражению, определяемому в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-968">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="5f0b1-969">

<dt>Тип</dt>

</span><span class="sxs-lookup"><span data-stu-id="5f0b1-969">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="5f0b1-970">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="5f0b1-970">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="5f0b1-971">Пример</span><span class="sxs-lookup"><span data-stu-id="5f0b1-971">Example</span></span>

```
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="5f0b1-972">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="5f0b1-972">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="5f0b1-973">Асинхронно возвращает данные, выбранные в теме или тексте сообщения.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-973">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="5f0b1-p160">Если выделенный фрагмент отсутствует, но курсор находится в тексте или теме, метод возвращает значение NULL для выбранных данных. Если выбраны не текст и не тема, метод возвращает ошибку `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-p160">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="5f0b1-976">Параметры:</span><span class="sxs-lookup"><span data-stu-id="5f0b1-976">Parameters:</span></span>

|<span data-ttu-id="5f0b1-977">Имя</span><span class="sxs-lookup"><span data-stu-id="5f0b1-977">Name</span></span>| <span data-ttu-id="5f0b1-978">Тип</span><span class="sxs-lookup"><span data-stu-id="5f0b1-978">Type</span></span>| <span data-ttu-id="5f0b1-979">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="5f0b1-979">Attributes</span></span>| <span data-ttu-id="5f0b1-980">Описание</span><span class="sxs-lookup"><span data-stu-id="5f0b1-980">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="5f0b1-981">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="5f0b1-981">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="5f0b1-p161">Запрашивает формат данных. Если задано значение Text, метод возвращает обычный текст как строку, удаляя все имеющиеся HTML-теги. Если задано значение HTML, метод возвращает выделенный текст (обычный текст или HTML).</span><span class="sxs-lookup"><span data-stu-id="5f0b1-p161">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="5f0b1-985">Объект</span><span class="sxs-lookup"><span data-stu-id="5f0b1-985">Object</span></span>| <span data-ttu-id="5f0b1-986">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="5f0b1-986">&lt;optional&gt;</span></span>|<span data-ttu-id="5f0b1-987">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-987">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="5f0b1-988">Объект</span><span class="sxs-lookup"><span data-stu-id="5f0b1-988">Object</span></span>| <span data-ttu-id="5f0b1-989">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="5f0b1-989">&lt;optional&gt;</span></span>|<span data-ttu-id="5f0b1-990">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-990">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="5f0b1-991">function</span><span class="sxs-lookup"><span data-stu-id="5f0b1-991">function</span></span>||<span data-ttu-id="5f0b1-992">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="5f0b1-992">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="5f0b1-993">Чтобы получить доступ к выбранным данным из метода обратного вызова, вызовите `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-993">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="5f0b1-994">Для доступа к исходному свойству, на основе которого созданы выбранные данные, вызовите  `asyncResult.value.sourceProperty`, который может иметь значение `body` или `subject`.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-994">To access the source property that the selection comes from, call , which will be either  or .|</span></span>|

##### <a name="requirements"></a><span data-ttu-id="5f0b1-995">Требования</span><span class="sxs-lookup"><span data-stu-id="5f0b1-995">Requirements</span></span>

|<span data-ttu-id="5f0b1-996">Требование</span><span class="sxs-lookup"><span data-stu-id="5f0b1-996">Requirement</span></span>| <span data-ttu-id="5f0b1-997">Значение</span><span class="sxs-lookup"><span data-stu-id="5f0b1-997">Value</span></span>|
|---|---|
|[<span data-ttu-id="5f0b1-998">Версия минимального набора требований для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="5f0b1-998">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5f0b1-999">1.2</span><span class="sxs-lookup"><span data-stu-id="5f0b1-999">1.2</span></span>|
|[<span data-ttu-id="5f0b1-1000">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1000">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5f0b1-1001">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1001">ReadWriteItem</span></span>|
|[<span data-ttu-id="5f0b1-1002">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1002">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5f0b1-1003">Создание</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1003">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="5f0b1-1004">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1004">Returns:</span></span>

<span data-ttu-id="5f0b1-1005">Выбранные данные в виде строки с форматом, определенным в параметре `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1005">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="5f0b1-1006">

<dt>Тип</dt>

</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1006">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="5f0b1-1007">String</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1007">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="5f0b1-1008">Пример</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1008">Example</span></span>

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

#### <a name="getselectedentities--entitiesjavascriptapioutlook16officeentities"></a><span data-ttu-id="5f0b1-1009">getSelectedEntities() → {[Entities](/javascript/api/outlook_1_6/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1009">getSelectedEntities() → {[Entities](/javascript/api/outlook_1_6/office.entities)}</span></span>

<span data-ttu-id="5f0b1-p163">Возвращает сущности, найденные в выделенном совпадении, выбранном пользователем. Выделенные совпадения применяются к [контекстным надстройкам](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="5f0b1-p163">Gets the entities found in a highlighted match a user has selected. Highlighted matches apply to [contextual add-ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="5f0b1-1012">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1012">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="5f0b1-1013">Требования</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1013">Requirements</span></span>

|<span data-ttu-id="5f0b1-1014">Требование</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1014">Requirement</span></span>| <span data-ttu-id="5f0b1-1015">Значение</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1015">Value</span></span>|
|---|---|
|[<span data-ttu-id="5f0b1-1016">Версия минимального набора требований для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1016">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5f0b1-1017">1.6</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1017">1.6</span></span> |
|[<span data-ttu-id="5f0b1-1018">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1018">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5f0b1-1019">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1019">ReadItem</span></span>|
|[<span data-ttu-id="5f0b1-1020">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1020">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5f0b1-1021">Чтение</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1021">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="5f0b1-1022">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1022">Returns:</span></span>

<span data-ttu-id="5f0b1-1023">Тип: [Entities](/javascript/api/outlook_1_6/office.entities)</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1023">Type: [Entities](/javascript/api/outlook_1_6/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="5f0b1-1024">Пример</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1024">Example</span></span>

<span data-ttu-id="5f0b1-1025">В приведенном ниже примере показано, как получить доступ к сущностям адресов в выделенном совпадении, выбранном пользователем.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1025">The following example accesses the addresses entities in the highlighted match selected by the user.</span></span>

```
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

#### <a name="getselectedregexmatches--object"></a><span data-ttu-id="5f0b1-1026">getSelectedRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1026">getSelectedRegExMatches() → {Object}</span></span>

<span data-ttu-id="5f0b1-p164">Возвращает строковые значения в выделенном совпадении, которые соответствуют регулярным выражениям, определенным в XML-файле манифеста. Выделенные совпадения применяются к [контекстным надстройкам](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="5f0b1-p164">Returns string values in a highlighted match that match the regular expressions defined in the manifest XML file. Highlighted matches apply to [contextual add-ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="5f0b1-1029">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1029">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="5f0b1-p165">Метод `getSelectedRegExMatches` возвращает строки, соответствующие регулярному выражению, которое определяется в каждом элементе правила `ItemHasRegularExpressionMatch` или `ItemHasKnownEntity` в XML-файле манифеста. Для правила `ItemHasRegularExpressionMatch` соответствующую строку должно содержать свойство элемента, указанного этим правилом. Простой тип `PropertyName` определяет поддерживаемые свойства.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-p165">The `getSelectedRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="5f0b1-1033">Например, рассмотрим манифест надстройки, который содержит указанный ниже элемент `Rule`.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1033">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="5f0b1-1034">Объект, возвращаемый методом `getRegExMatches`, будет содержать два свойства: `fruits` и `veggies`.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1034">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="5f0b1-p166">Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты. Лучше используйте метод [`Body.getAsync`](/javascript/api/outlook_1_6/office.body#getasync-coerciontype--options--callback-) для этого.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-p166">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook_1_6/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="5f0b1-1038">Требования</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1038">Requirements</span></span>

|<span data-ttu-id="5f0b1-1039">Требование</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1039">Requirement</span></span>| <span data-ttu-id="5f0b1-1040">Значение</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1040">Value</span></span>|
|---|---|
|[<span data-ttu-id="5f0b1-1041">Версия минимального набора требований для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1041">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5f0b1-1042">1.6</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1042">1.6</span></span> |
|[<span data-ttu-id="5f0b1-1043">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1043">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5f0b1-1044">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1044">ReadItem</span></span>|
|[<span data-ttu-id="5f0b1-1045">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1045">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5f0b1-1046">Чтение</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1046">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="5f0b1-1047">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1047">Returns:</span></span>

<span data-ttu-id="5f0b1-p167">Объект, содержащий массив строк, которые соответствуют регулярным выражениям, определяемым в XML-файле манифеста. Имя каждого массива равно соответствующему значению атрибута `RegExName` правила сопоставления `ItemHasRegularExpressionMatch` или атрибута `FilterName` правила сопоставления `ItemHasKnownEntity`.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-p167">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

##### <a name="example"></a><span data-ttu-id="5f0b1-1050">Пример</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1050">Example</span></span>

<span data-ttu-id="5f0b1-1051">В приведенном ниже примере показано, как получить доступ к массиву совпадений с элементами `fruits` и `veggies` правил активации регулярных выражений, указанными в манифесте.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1051">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="5f0b1-1052">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1052">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="5f0b1-1053">Асинхронно загружает настраиваемые свойства для надстройки выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1053">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="5f0b1-p168">Настраиваемые свойства сохраняются в виде пар "ключ-значение" для каждого приложения и каждого элемента. Этот метод возвращает объект `CustomProperties` при обратном вызове, который предоставляет методы для доступа к настраиваемым свойствам, характерным для текущего элемента и текущей надстройки. Настраиваемые свойства не шифруются для элемента, поэтому этот способ хранения не является безопасным.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-p168">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="5f0b1-1057">Параметры:</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1057">Parameters:</span></span>

|<span data-ttu-id="5f0b1-1058">Имя</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1058">Name</span></span>| <span data-ttu-id="5f0b1-1059">Тип</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1059">Type</span></span>| <span data-ttu-id="5f0b1-1060">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1060">Attributes</span></span>| <span data-ttu-id="5f0b1-1061">Описание</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1061">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="5f0b1-1062">function</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1062">function</span></span>||<span data-ttu-id="5f0b1-1063">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1063">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="5f0b1-1064">Настраиваемые свойства предоставляются в виде объекта [`CustomProperties`](/javascript/api/outlook_1_6/office.customproperties) в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1064">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook_1_6/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="5f0b1-1065">Этот объект можно использовать для получения, задания и удаления настраиваемых свойств из элемента и сохранения изменений настраиваемого свойства на сервере.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1065">The custom properties are provided as a CustomProperties object in the asyncResult.value property. This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="5f0b1-1066">Объект</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1066">Object</span></span>| <span data-ttu-id="5f0b1-1067">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1067">&lt;optional&gt;</span></span>|<span data-ttu-id="5f0b1-1068">Разработчики могут указать любой объект, к которому необходимо получить доступ, в функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1068">Developers can provide any object they wish to access in the callback method.</span></span> <span data-ttu-id="5f0b1-1069">Доступ к этому объекту можно получить с помощью свойства `asyncResult.asyncContext` в функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1069">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.|</span></span>|

##### <a name="requirements"></a><span data-ttu-id="5f0b1-1070">Требования</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1070">Requirements</span></span>

|<span data-ttu-id="5f0b1-1071">Требование</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1071">Requirement</span></span>| <span data-ttu-id="5f0b1-1072">Значение</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1072">Value</span></span>|
|---|---|
|[<span data-ttu-id="5f0b1-1073">Версия минимального набора требований для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1073">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5f0b1-1074">1.0</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1074">1.0</span></span>|
|[<span data-ttu-id="5f0b1-1075">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1075">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5f0b1-1076">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1076">ReadItem</span></span>|
|[<span data-ttu-id="5f0b1-1077">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1077">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5f0b1-1078">Cоздание или чтение</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1078">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="5f0b1-1079">Пример</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1079">Example</span></span>

<span data-ttu-id="5f0b1-p171">В приведенном ниже примере кода показано, как асинхронно загружать настраиваемые свойства, характерные для текущего элемента, с помощью метода `loadCustomPropertiesAsync`. Этот пример также показывает, как сохранять эти свойства на сервере с помощью метода `CustomProperties.saveAsync`. В этом примере кода, после того как выполнена загрузка настраиваемых свойств, метод `CustomProperties.get` используется для считывания настраиваемого свойства `myProp`, метод `CustomProperties.set` — для записи настраиваемого свойства `otherProp`, а метод `saveAsync` — для сохранения настраиваемых свойств.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-p171">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="5f0b1-1083">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1083">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="5f0b1-1084">Удаляет вложение из сообщения или встречи.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1084">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="5f0b1-p172">Метод `removeAttachmentAsync` удаляет из элемента вложение с указанным идентификатором. Идентификатор вложения рекомендуется использовать для удаления вложения, только если оно добавлено тем же почтовым приложением в ходе текущего сеанса. В веб-приложении Outlook и веб-приложении Outlook для устройств идентификатор вложения действителен только в рамках одного сеанса. Сеанс завершается, когда пользователь закрывает приложение или начинает создавать элемент во встроенной форме, а затем переходит из формы в отдельное окно.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-p172">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="5f0b1-1089">Параметры:</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1089">Parameters:</span></span>

|<span data-ttu-id="5f0b1-1090">Имя</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1090">Name</span></span>| <span data-ttu-id="5f0b1-1091">Тип</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1091">Type</span></span>| <span data-ttu-id="5f0b1-1092">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1092">Attributes</span></span>| <span data-ttu-id="5f0b1-1093">Описание</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1093">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="5f0b1-1094">String</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1094">String</span></span>||<span data-ttu-id="5f0b1-p173">Идентификатор удаляемого вложения. Максимальная длина строки — 100 символов.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-p173">The identifier of the attachment to remove. The maximum length of the string is 100 characters.</span></span>|
|`options`| <span data-ttu-id="5f0b1-1097">Объект</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1097">Object</span></span>| <span data-ttu-id="5f0b1-1098">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1098">&lt;optional&gt;</span></span>|<span data-ttu-id="5f0b1-1099">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1099">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="5f0b1-1100">Объект</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1100">Object</span></span>| <span data-ttu-id="5f0b1-1101">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1101">&lt;optional&gt;</span></span>|<span data-ttu-id="5f0b1-1102">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1102">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="5f0b1-1103">function</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1103">function</span></span>| <span data-ttu-id="5f0b1-1104">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1104">&lt;optional&gt;</span></span>|<span data-ttu-id="5f0b1-1105">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1105">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="5f0b1-1106">Если удалить вложение не удается, свойство `asyncResult.error` содержит код ошибки с указанием ее причины.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1106">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="5f0b1-1107">Ошибки</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1107">Errors</span></span>

| <span data-ttu-id="5f0b1-1108">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1108">Error code</span></span> | <span data-ttu-id="5f0b1-1109">Описание</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1109">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="5f0b1-1110">Идентификатор вложения не существует.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1110">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="5f0b1-1111">Требования</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1111">Requirements</span></span>

|<span data-ttu-id="5f0b1-1112">Требование</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1112">Requirement</span></span>| <span data-ttu-id="5f0b1-1113">Значение</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1113">Value</span></span>|
|---|---|
|[<span data-ttu-id="5f0b1-1114">Версия минимального набора требований для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1114">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5f0b1-1115">1.1</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1115">1.1</span></span>|
|[<span data-ttu-id="5f0b1-1116">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1116">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5f0b1-1117">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1117">ReadWriteItem</span></span>|
|[<span data-ttu-id="5f0b1-1118">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1118">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5f0b1-1119">Создание</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1119">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="5f0b1-1120">Пример</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1120">Example</span></span>

<span data-ttu-id="5f0b1-1121">Указанный ниже код удаляет вложение с идентификатором "0".</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1121">The following code removes an attachment with an identifier of '0'.</span></span>

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

####  <a name="saveasyncoptions-callback"></a><span data-ttu-id="5f0b1-1122">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1122">saveAsync([options], callback)</span></span>

<span data-ttu-id="5f0b1-1123">Асинхронно сохраняет элемент.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1123">Asynchronously saves an item.</span></span>

<span data-ttu-id="5f0b1-p174">При вызове этот метод сохраняет текущее сообщение в виде черновика и возвращает идентификатор элемента с помощью метода обратного вызова. В Outlook Web App или интерактивном режиме Outlook этот элемент сохраняется на сервере. В Outlook в режиме кэширования этот элемент сохраняется в локальном кэше.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-p174">When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="5f0b1-1127">Если в режиме создания надстройка вызывает для элемента метод `saveAsync`, чтобы получить параметр `itemId` для использования с EWS или REST API, необходимо помнить, что синхронизация элемента с сервером может занять много времени, если Outlook работает в режиме кэширования данных.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1127">Note: If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server. Until the item is synced, using the  will return an error.</span></span> <span data-ttu-id="5f0b1-1128">До окончания синхронизации применение параметра `itemId`  будет приводить к ошибке.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1128">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="5f0b1-p176">Если метод `saveAsync` вызывается для встречи в режиме создания, она сохраняется как обычная встреча в календаре пользователя, а не как черновик. При сохранении новой встречи приглашения не отправляются. При сохранении существующей встречи уведомления отправляются добавленным или удаленным участникам.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-p176">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="5f0b1-1132">Следующие клиенты имеют разную реакцию на событие для `saveAsync` для встреч в режиме создания:</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1132">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="5f0b1-1133">Mac Outlook не поддерживает `saveAsync` на собрании в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1133">Note: Mac Outlook does not support `saveAsync` on a meeting in compose mode. Calling  on a meeting in Mac Outlook will return an error.</span></span> <span data-ttu-id="5f0b1-1134">Вызов `saveAsync` на собрании в Mac Outlook возвращает ошибку.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1134">Note: Mac Outlook does not support  on a meeting in compose mode. Calling `saveAsync` on a meeting in Mac Outlook will return an error.</span></span>
> - <span data-ttu-id="5f0b1-1135">Outlook в Интернете всегда отправляет приглашение или обновления при вызове `saveAsync` на встрече в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1135">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="5f0b1-1136">Параметры:</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1136">Parameters:</span></span>

|<span data-ttu-id="5f0b1-1137">Имя</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1137">Name</span></span>| <span data-ttu-id="5f0b1-1138">Тип</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1138">Type</span></span>| <span data-ttu-id="5f0b1-1139">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1139">Attributes</span></span>| <span data-ttu-id="5f0b1-1140">Описание</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1140">Description</span></span>|
|---|---|---|---|
|`options`| <span data-ttu-id="5f0b1-1141">Объект</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1141">Object</span></span>| <span data-ttu-id="5f0b1-1142">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1142">&lt;optional&gt;</span></span>|<span data-ttu-id="5f0b1-1143">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1143">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="5f0b1-1144">Объект</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1144">Object</span></span>| <span data-ttu-id="5f0b1-1145">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1145">&lt;optional&gt;</span></span>|<span data-ttu-id="5f0b1-1146">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1146">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="5f0b1-1147">function</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1147">function</span></span>||<span data-ttu-id="5f0b1-1148">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1148">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="5f0b1-1149">После успешного выполнения идентификатор элемента будет предоставлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1149">On success, the item identifier is provided in the `asyncResult.value` property.|</span></span>|

##### <a name="requirements"></a><span data-ttu-id="5f0b1-1150">Требования</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1150">Requirements</span></span>

|<span data-ttu-id="5f0b1-1151">Требование</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1151">Requirement</span></span>| <span data-ttu-id="5f0b1-1152">Значение</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1152">Value</span></span>|
|---|---|
|[<span data-ttu-id="5f0b1-1153">Версия минимального набора требований для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1153">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5f0b1-1154">1.3</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1154">1.3</span></span>|
|[<span data-ttu-id="5f0b1-1155">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1155">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5f0b1-1156">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1156">ReadWriteItem</span></span>|
|[<span data-ttu-id="5f0b1-1157">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1157">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5f0b1-1158">Создание</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1158">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="5f0b1-1159">Примеры</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1159">Examples</span></span>

```
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result
  });
```

<span data-ttu-id="5f0b1-p178">Ниже приведен пример параметра `result`, переданного функции обратного вызова. Свойство `value` содержит идентификатор элемента.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-p178">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

####  <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="5f0b1-1162">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1162">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="5f0b1-1163">Асинхронно вставляет данные в текст или тему сообщения.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1163">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="5f0b1-p179">Метод `setSelectedDataAsync` вставляет указанную строку в местоположение курсора в теме или тексте элемента либо, если текст выделен в редакторе, он заменяет выделенный текст. Если курсор находится вне текста или темы элемента, возвращается ошибка. После вставки курсор помещается в конец вставленного содержимого.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-p179">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="5f0b1-1167">Параметры:</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1167">Parameters:</span></span>

|<span data-ttu-id="5f0b1-1168">Имя</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1168">Name</span></span>| <span data-ttu-id="5f0b1-1169">Тип</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1169">Type</span></span>| <span data-ttu-id="5f0b1-1170">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1170">Attributes</span></span>| <span data-ttu-id="5f0b1-1171">Описание</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1171">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="5f0b1-1172">String</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1172">String</span></span>||<span data-ttu-id="5f0b1-p180">Вставляемые данные. Объем данных не должен превышать 1 000 000 символов. Если передано больше 1 000 000 символов, возвращается исключение `ArgumentOutOfRange`.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-p180">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="5f0b1-1176">Объект</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1176">Object</span></span>| <span data-ttu-id="5f0b1-1177">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1177">&lt;optional&gt;</span></span>|<span data-ttu-id="5f0b1-1178">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1178">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="5f0b1-1179">Объект</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1179">Object</span></span>| <span data-ttu-id="5f0b1-1180">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1180">&lt;optional&gt;</span></span>|<span data-ttu-id="5f0b1-1181">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1181">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`| [<span data-ttu-id="5f0b1-1182">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1182">Office.CoercionType</span></span>](office.md#coerciontype-string)| <span data-ttu-id="5f0b1-1183">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1183">&lt;optional&gt;</span></span>|<span data-ttu-id="5f0b1-p181">Если задано значение `text`, текущий стиль применяется в Outlook и веб-приложении Outlook. Если поле представляет собой редактор HTML, вставляются только текстовые данные, даже если они имеют формат HTML.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-p181">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="5f0b1-p182">Если `html` и поле поддерживают HTML (тема не поддерживает), в веб-приложении Outlook применяется текущий стиль, а в Outlook — стиль по умолчанию. Если поле является текстовым, возвращается ошибка `InvalidDataFormat`.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-p182">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="5f0b1-1188">Если тип `coercionType` не установлен, результат зависит от поля: если поле имеет формат HTML, то используется HTML; если поле является текстовым, то используется обычный текст.</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1188">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="5f0b1-1189">function</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1189">function</span></span>||<span data-ttu-id="5f0b1-1190">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1190">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="5f0b1-1191">Требования</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1191">Requirements</span></span>

|<span data-ttu-id="5f0b1-1192">Требование</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1192">Requirement</span></span>| <span data-ttu-id="5f0b1-1193">Значение</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1193">Value</span></span>|
|---|---|
|[<span data-ttu-id="5f0b1-1194">Версия минимального набора требований для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1194">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5f0b1-1195">1.2</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1195">1.2</span></span>|
|[<span data-ttu-id="5f0b1-1196">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1196">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5f0b1-1197">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1197">ReadWriteItem</span></span>|
|[<span data-ttu-id="5f0b1-1198">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1198">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5f0b1-1199">Создание</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1199">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="5f0b1-1200">Пример</span><span class="sxs-lookup"><span data-stu-id="5f0b1-1200">Example</span></span>

```
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```