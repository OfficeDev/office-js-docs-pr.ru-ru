
# <a name="item"></a><span data-ttu-id="a74d4-101">item</span><span class="sxs-lookup"><span data-stu-id="a74d4-101">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="a74d4-102">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="a74d4-102">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="a74d4-p101">Пространство имен `item` используется для доступа к выбранному в данный момент сообщению, приглашению на собрание или описанию встречи. Вы можете определить тип пространства имен `item` с помощью свойства [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook16officemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="a74d4-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook16officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="a74d4-105">Требования</span><span class="sxs-lookup"><span data-stu-id="a74d4-105">Requirements</span></span>

|<span data-ttu-id="a74d4-106">Требование</span><span class="sxs-lookup"><span data-stu-id="a74d4-106">Requirement</span></span>| <span data-ttu-id="a74d4-107">Значение</span><span class="sxs-lookup"><span data-stu-id="a74d4-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="a74d4-108">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="a74d4-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a74d4-109">1.0</span><span class="sxs-lookup"><span data-stu-id="a74d4-109">1.0</span></span>|
|[<span data-ttu-id="a74d4-110">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="a74d4-110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a74d4-111">Ограниченный доступ</span><span class="sxs-lookup"><span data-stu-id="a74d4-111">Restricted</span></span>|
|[<span data-ttu-id="a74d4-112">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="a74d4-112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a74d4-113">Cоздание или чтение</span><span class="sxs-lookup"><span data-stu-id="a74d4-113">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="a74d4-114">Члены и методы</span><span class="sxs-lookup"><span data-stu-id="a74d4-114">Members and methods</span></span>

| <span data-ttu-id="a74d4-115">Член</span><span class="sxs-lookup"><span data-stu-id="a74d4-115">Member</span></span> | <span data-ttu-id="a74d4-116">Тип</span><span class="sxs-lookup"><span data-stu-id="a74d4-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="a74d4-117">attachments</span><span class="sxs-lookup"><span data-stu-id="a74d4-117">attachments</span></span>](#attachments-arrayattachmentdetailsjavascriptapioutlook16officeattachmentdetails) | <span data-ttu-id="a74d4-118">Член</span><span class="sxs-lookup"><span data-stu-id="a74d4-118">Member</span></span> |
| [<span data-ttu-id="a74d4-119">bcc</span><span class="sxs-lookup"><span data-stu-id="a74d4-119">bcc</span></span>](#bcc-recipientsjavascriptapioutlook16officerecipients) | <span data-ttu-id="a74d4-120">Член</span><span class="sxs-lookup"><span data-stu-id="a74d4-120">Member</span></span> |
| [<span data-ttu-id="a74d4-121">body</span><span class="sxs-lookup"><span data-stu-id="a74d4-121">body</span></span>](#body-bodyjavascriptapioutlook16officebody) | <span data-ttu-id="a74d4-122">Член</span><span class="sxs-lookup"><span data-stu-id="a74d4-122">Member</span></span> |
| [<span data-ttu-id="a74d4-123">cc</span><span class="sxs-lookup"><span data-stu-id="a74d4-123">cc</span></span>](#cc-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients) | <span data-ttu-id="a74d4-124">Член</span><span class="sxs-lookup"><span data-stu-id="a74d4-124">Member</span></span> |
| [<span data-ttu-id="a74d4-125">conversationId</span><span class="sxs-lookup"><span data-stu-id="a74d4-125">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="a74d4-126">Член</span><span class="sxs-lookup"><span data-stu-id="a74d4-126">Member</span></span> |
| [<span data-ttu-id="a74d4-127">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="a74d4-127">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="a74d4-128">Член</span><span class="sxs-lookup"><span data-stu-id="a74d4-128">Member</span></span> |
| [<span data-ttu-id="a74d4-129">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="a74d4-129">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="a74d4-130">Член</span><span class="sxs-lookup"><span data-stu-id="a74d4-130">Member</span></span> |
| [<span data-ttu-id="a74d4-131">end</span><span class="sxs-lookup"><span data-stu-id="a74d4-131">end</span></span>](#end-datetimejavascriptapioutlook16officetime) | <span data-ttu-id="a74d4-132">Член</span><span class="sxs-lookup"><span data-stu-id="a74d4-132">Member</span></span> |
| [<span data-ttu-id="a74d4-133">from</span><span class="sxs-lookup"><span data-stu-id="a74d4-133">from</span></span>](#from-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails) | <span data-ttu-id="a74d4-134">Член</span><span class="sxs-lookup"><span data-stu-id="a74d4-134">Member</span></span> |
| [<span data-ttu-id="a74d4-135">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="a74d4-135">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="a74d4-136">Член</span><span class="sxs-lookup"><span data-stu-id="a74d4-136">Member</span></span> |
| [<span data-ttu-id="a74d4-137">itemClass</span><span class="sxs-lookup"><span data-stu-id="a74d4-137">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="a74d4-138">Член</span><span class="sxs-lookup"><span data-stu-id="a74d4-138">Member</span></span> |
| [<span data-ttu-id="a74d4-139">itemId</span><span class="sxs-lookup"><span data-stu-id="a74d4-139">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="a74d4-140">Член</span><span class="sxs-lookup"><span data-stu-id="a74d4-140">Member</span></span> |
| [<span data-ttu-id="a74d4-141">itemType</span><span class="sxs-lookup"><span data-stu-id="a74d4-141">itemType</span></span>](#itemtype-officemailboxenumsitemtypejavascriptapioutlook16officemailboxenumsitemtype) | <span data-ttu-id="a74d4-142">Член</span><span class="sxs-lookup"><span data-stu-id="a74d4-142">Member</span></span> |
| [<span data-ttu-id="a74d4-143">location</span><span class="sxs-lookup"><span data-stu-id="a74d4-143">location</span></span>](#location-stringlocationjavascriptapioutlook16officelocation) | <span data-ttu-id="a74d4-144">Член</span><span class="sxs-lookup"><span data-stu-id="a74d4-144">Member</span></span> |
| [<span data-ttu-id="a74d4-145">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="a74d4-145">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="a74d4-146">Член</span><span class="sxs-lookup"><span data-stu-id="a74d4-146">Member</span></span> |
| [<span data-ttu-id="a74d4-147">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="a74d4-147">notificationMessages</span></span>](#notificationmessages-notificationmessagesjavascriptapioutlook16officenotificationmessages) | <span data-ttu-id="a74d4-148">Член</span><span class="sxs-lookup"><span data-stu-id="a74d4-148">Member</span></span> |
| [<span data-ttu-id="a74d4-149">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="a74d4-149">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients) | <span data-ttu-id="a74d4-150">Член</span><span class="sxs-lookup"><span data-stu-id="a74d4-150">Member</span></span> |
| [<span data-ttu-id="a74d4-151">organizer</span><span class="sxs-lookup"><span data-stu-id="a74d4-151">organizer</span></span>](#organizer-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails) | <span data-ttu-id="a74d4-152">Член</span><span class="sxs-lookup"><span data-stu-id="a74d4-152">Member</span></span> |
| [<span data-ttu-id="a74d4-153">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="a74d4-153">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients) | <span data-ttu-id="a74d4-154">Член</span><span class="sxs-lookup"><span data-stu-id="a74d4-154">Member</span></span> |
| [<span data-ttu-id="a74d4-155">sender</span><span class="sxs-lookup"><span data-stu-id="a74d4-155">sender</span></span>](#sender-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails) | <span data-ttu-id="a74d4-156">Член</span><span class="sxs-lookup"><span data-stu-id="a74d4-156">Member</span></span> |
| [<span data-ttu-id="a74d4-157">start</span><span class="sxs-lookup"><span data-stu-id="a74d4-157">start</span></span>](#start-datetimejavascriptapioutlook16officetime) | <span data-ttu-id="a74d4-158">Член</span><span class="sxs-lookup"><span data-stu-id="a74d4-158">Member</span></span> |
| [<span data-ttu-id="a74d4-159">subject</span><span class="sxs-lookup"><span data-stu-id="a74d4-159">subject</span></span>](#subject-stringsubjectjavascriptapioutlook16officesubject) | <span data-ttu-id="a74d4-160">Член</span><span class="sxs-lookup"><span data-stu-id="a74d4-160">Member</span></span> |
| [<span data-ttu-id="a74d4-161">to</span><span class="sxs-lookup"><span data-stu-id="a74d4-161">to</span></span>](#to-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients) | <span data-ttu-id="a74d4-162">Член</span><span class="sxs-lookup"><span data-stu-id="a74d4-162">Member</span></span> |
| [<span data-ttu-id="a74d4-163">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="a74d4-163">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="a74d4-164">Метод</span><span class="sxs-lookup"><span data-stu-id="a74d4-164">Method</span></span> |
| [<span data-ttu-id="a74d4-165">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="a74d4-165">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="a74d4-166">Метод</span><span class="sxs-lookup"><span data-stu-id="a74d4-166">Method</span></span> |
| [<span data-ttu-id="a74d4-167">close</span><span class="sxs-lookup"><span data-stu-id="a74d4-167">close</span></span>](#close) | <span data-ttu-id="a74d4-168">Метод</span><span class="sxs-lookup"><span data-stu-id="a74d4-168">Method</span></span> |
| [<span data-ttu-id="a74d4-169">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="a74d4-169">displayReplyAllForm</span></span>](#displayreplyallformformdata) | <span data-ttu-id="a74d4-170">Метод</span><span class="sxs-lookup"><span data-stu-id="a74d4-170">Method</span></span> |
| [<span data-ttu-id="a74d4-171">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="a74d4-171">displayReplyForm</span></span>](#displayreplyformformdata) | <span data-ttu-id="a74d4-172">Метод</span><span class="sxs-lookup"><span data-stu-id="a74d4-172">Method</span></span> |
| [<span data-ttu-id="a74d4-173">getEntities</span><span class="sxs-lookup"><span data-stu-id="a74d4-173">getEntities</span></span>](#getentities--entitiesjavascriptapioutlook16officeentities) | <span data-ttu-id="a74d4-174">Метод</span><span class="sxs-lookup"><span data-stu-id="a74d4-174">Method</span></span> |
| [<span data-ttu-id="a74d4-175">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="a74d4-175">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook16officecontactmeetingsuggestionjavascriptapioutlook16officemeetingsuggestionphonenumberjavascriptapioutlook16officephonenumbertasksuggestionjavascriptapioutlook16officetasksuggestion) | <span data-ttu-id="a74d4-176">Метод</span><span class="sxs-lookup"><span data-stu-id="a74d4-176">Method</span></span> |
| [<span data-ttu-id="a74d4-177">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="a74d4-177">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook16officecontactmeetingsuggestionjavascriptapioutlook16officemeetingsuggestionphonenumberjavascriptapioutlook16officephonenumbertasksuggestionjavascriptapioutlook16officetasksuggestion) | <span data-ttu-id="a74d4-178">Метод</span><span class="sxs-lookup"><span data-stu-id="a74d4-178">Method</span></span> |
| [<span data-ttu-id="a74d4-179">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="a74d4-179">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="a74d4-180">Метод</span><span class="sxs-lookup"><span data-stu-id="a74d4-180">Method</span></span> |
| [<span data-ttu-id="a74d4-181">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="a74d4-181">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="a74d4-182">Метод</span><span class="sxs-lookup"><span data-stu-id="a74d4-182">Method</span></span> |
| [<span data-ttu-id="a74d4-183">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="a74d4-183">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="a74d4-184">Метод</span><span class="sxs-lookup"><span data-stu-id="a74d4-184">Method</span></span> |
| [<span data-ttu-id="a74d4-185">getSelectedEntities</span><span class="sxs-lookup"><span data-stu-id="a74d4-185">getSelectedEntities</span></span>](#getselectedentities--entitiesjavascriptapioutlook16officeentities) | <span data-ttu-id="a74d4-186">Метод</span><span class="sxs-lookup"><span data-stu-id="a74d4-186">Method</span></span> |
| [<span data-ttu-id="a74d4-187">getSelectedRegExMatches</span><span class="sxs-lookup"><span data-stu-id="a74d4-187">getSelectedRegExMatches</span></span>](#getselectedregexmatches--object) | <span data-ttu-id="a74d4-188">Метод</span><span class="sxs-lookup"><span data-stu-id="a74d4-188">Method</span></span> |
| [<span data-ttu-id="a74d4-189">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="a74d4-189">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="a74d4-190">Метод</span><span class="sxs-lookup"><span data-stu-id="a74d4-190">Method</span></span> |
| [<span data-ttu-id="a74d4-191">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="a74d4-191">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="a74d4-192">Метод</span><span class="sxs-lookup"><span data-stu-id="a74d4-192">Method</span></span> |
| [<span data-ttu-id="a74d4-193">saveAsync</span><span class="sxs-lookup"><span data-stu-id="a74d4-193">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="a74d4-194">Метод</span><span class="sxs-lookup"><span data-stu-id="a74d4-194">Method</span></span> |
| [<span data-ttu-id="a74d4-195">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="a74d4-195">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="a74d4-196">Метод</span><span class="sxs-lookup"><span data-stu-id="a74d4-196">Method</span></span> |

### <a name="example"></a><span data-ttu-id="a74d4-197">Пример</span><span class="sxs-lookup"><span data-stu-id="a74d4-197">Example</span></span>

<span data-ttu-id="a74d4-198">В приведенном ниже примере кода JavaScript показано, как получить доступ к свойству `subject` текущего элемента в Outlook.</span><span class="sxs-lookup"><span data-stu-id="a74d4-198">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="a74d4-199">Члены</span><span class="sxs-lookup"><span data-stu-id="a74d4-199">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook16officeattachmentdetails"></a><span data-ttu-id="a74d4-200">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_6/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="a74d4-200">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_6/office.attachmentdetails)></span></span>

<span data-ttu-id="a74d4-p102">Получает массив вложений для элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="a74d4-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="a74d4-203">Некоторые типы файлов блокируются Outlook из-за потенциальных проблем безопасности и поэтому не возвращаются.</span><span class="sxs-lookup"><span data-stu-id="a74d4-203">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="a74d4-204">Дополнительные сведения см. в статье [Блокированные вложения в Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="a74d4-204">For more information see [Payments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="a74d4-205">Тип:</span><span class="sxs-lookup"><span data-stu-id="a74d4-205">Type:</span></span>

*   <span data-ttu-id="a74d4-206">Array.<[AttachmentDetails](/javascript/api/outlook_1_6/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="a74d4-206">Array.<[AttachmentDetails](/javascript/api/outlook_1_6/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="a74d4-207">Требования</span><span class="sxs-lookup"><span data-stu-id="a74d4-207">Requirements</span></span>

|<span data-ttu-id="a74d4-208">Требование</span><span class="sxs-lookup"><span data-stu-id="a74d4-208">Requirement</span></span>| <span data-ttu-id="a74d4-209">Значение</span><span class="sxs-lookup"><span data-stu-id="a74d4-209">Value</span></span>|
|---|---|
|[<span data-ttu-id="a74d4-210">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="a74d4-210">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a74d4-211">1.0</span><span class="sxs-lookup"><span data-stu-id="a74d4-211">1.0</span></span>|
|[<span data-ttu-id="a74d4-212">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="a74d4-212">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a74d4-213">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a74d4-213">ReadItem</span></span>|
|[<span data-ttu-id="a74d4-214">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="a74d4-214">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a74d4-215">Чтение</span><span class="sxs-lookup"><span data-stu-id="a74d4-215">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a74d4-216">Пример</span><span class="sxs-lookup"><span data-stu-id="a74d4-216">Example</span></span>

<span data-ttu-id="a74d4-217">С помощью приведенного ниже кода можно создать HTML-строку с подробными сведениями обо всех вложениях для текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="a74d4-217">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

####  <a name="bcc-recipientsjavascriptapioutlook16officerecipients"></a><span data-ttu-id="a74d4-218">bcc :[Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="a74d4-218">bcc :[Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

<span data-ttu-id="a74d4-219">Получает объект, который предоставляет методы для получения или обновления получателей в строке Bcc (скрытой копии) сообщения.</span><span class="sxs-lookup"><span data-stu-id="a74d4-219">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="a74d4-220">Только в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="a74d4-220">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="a74d4-221">Тип:</span><span class="sxs-lookup"><span data-stu-id="a74d4-221">Type:</span></span>

*   [<span data-ttu-id="a74d4-222">Recipients</span><span class="sxs-lookup"><span data-stu-id="a74d4-222">Recipients</span></span>](/javascript/api/outlook_1_6/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="a74d4-223">Требования</span><span class="sxs-lookup"><span data-stu-id="a74d4-223">Requirements</span></span>

|<span data-ttu-id="a74d4-224">Требование</span><span class="sxs-lookup"><span data-stu-id="a74d4-224">Requirement</span></span>| <span data-ttu-id="a74d4-225">Значение</span><span class="sxs-lookup"><span data-stu-id="a74d4-225">Value</span></span>|
|---|---|
|[<span data-ttu-id="a74d4-226">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="a74d4-226">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a74d4-227">1.1</span><span class="sxs-lookup"><span data-stu-id="a74d4-227">1.1</span></span>|
|[<span data-ttu-id="a74d4-228">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="a74d4-228">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a74d4-229">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a74d4-229">ReadItem</span></span>|
|[<span data-ttu-id="a74d4-230">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="a74d4-230">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a74d4-231">Создание</span><span class="sxs-lookup"><span data-stu-id="a74d4-231">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="a74d4-232">Пример</span><span class="sxs-lookup"><span data-stu-id="a74d4-232">Example</span></span>

```
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlook16officebody"></a><span data-ttu-id="a74d4-233">body :[Body](/javascript/api/outlook_1_6/office.body)</span><span class="sxs-lookup"><span data-stu-id="a74d4-233">body :[Body](/javascript/api/outlook_1_6/office.body)</span></span>

<span data-ttu-id="a74d4-234">Получает объект, предоставляющий методы для работы с основным текстом элемента.</span><span class="sxs-lookup"><span data-stu-id="a74d4-234">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="a74d4-235">Тип:</span><span class="sxs-lookup"><span data-stu-id="a74d4-235">Type:</span></span>

*   [<span data-ttu-id="a74d4-236">Body</span><span class="sxs-lookup"><span data-stu-id="a74d4-236">Body</span></span>](/javascript/api/outlook_1_6/office.body)

##### <a name="requirements"></a><span data-ttu-id="a74d4-237">Требования</span><span class="sxs-lookup"><span data-stu-id="a74d4-237">Requirements</span></span>

|<span data-ttu-id="a74d4-238">Требование</span><span class="sxs-lookup"><span data-stu-id="a74d4-238">Requirement</span></span>| <span data-ttu-id="a74d4-239">Значение</span><span class="sxs-lookup"><span data-stu-id="a74d4-239">Value</span></span>|
|---|---|
|[<span data-ttu-id="a74d4-240">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="a74d4-240">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a74d4-241">1.1</span><span class="sxs-lookup"><span data-stu-id="a74d4-241">1.1</span></span>|
|[<span data-ttu-id="a74d4-242">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="a74d4-242">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a74d4-243">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a74d4-243">ReadItem</span></span>|
|[<span data-ttu-id="a74d4-244">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="a74d4-244">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a74d4-245">Cоздание или чтение</span><span class="sxs-lookup"><span data-stu-id="a74d4-245">Compose or read</span></span>|

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients"></a><span data-ttu-id="a74d4-246">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="a74d4-246">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

<span data-ttu-id="a74d4-247">Предоставляет доступ к получателям Cc (копии) сообщения.</span><span class="sxs-lookup"><span data-stu-id="a74d4-247">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="a74d4-248">Уровень доступа и тип объекта, зависит от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="a74d4-248">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="a74d4-249">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="a74d4-249">Read mode</span></span>

<span data-ttu-id="a74d4-p106">Свойство `cc` возвращает массив, который содержит объект `EmailAddressDetails`, каждому получателю, указанному в строке **Cc (копия)** сообщения. Коллекция может включать не более 100 членов.</span><span class="sxs-lookup"><span data-stu-id="a74d4-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="a74d4-252">Режим создания</span><span class="sxs-lookup"><span data-stu-id="a74d4-252">Compose mode</span></span>

<span data-ttu-id="a74d4-253">Свойство `cc` возвращает объект `Recipients`, предоставляющий методы для получения или обновления получателей, которые указаны в строке **Cc (копия)** сообщения.</span><span class="sxs-lookup"><span data-stu-id="a74d4-253">The `cc` property returns a `Recipients` object that provides methods for manipulating the recipients on the **Cc** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="a74d4-254">Тип:</span><span class="sxs-lookup"><span data-stu-id="a74d4-254">Type:</span></span>

*   <span data-ttu-id="a74d4-255">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="a74d4-255">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="a74d4-256">Требования</span><span class="sxs-lookup"><span data-stu-id="a74d4-256">Requirements</span></span>

|<span data-ttu-id="a74d4-257">Требование</span><span class="sxs-lookup"><span data-stu-id="a74d4-257">Requirement</span></span>| <span data-ttu-id="a74d4-258">Значение</span><span class="sxs-lookup"><span data-stu-id="a74d4-258">Value</span></span>|
|---|---|
|[<span data-ttu-id="a74d4-259">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="a74d4-259">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a74d4-260">1.0</span><span class="sxs-lookup"><span data-stu-id="a74d4-260">1.0</span></span>|
|[<span data-ttu-id="a74d4-261">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="a74d4-261">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a74d4-262">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a74d4-262">ReadItem</span></span>|
|[<span data-ttu-id="a74d4-263">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="a74d4-263">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a74d4-264">Cоздание или чтение</span><span class="sxs-lookup"><span data-stu-id="a74d4-264">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="a74d4-265">Пример</span><span class="sxs-lookup"><span data-stu-id="a74d4-265">Example</span></span>

```
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="a74d4-266">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="a74d4-266">(nullable) conversationId :String</span></span>

<span data-ttu-id="a74d4-267">Получает идентификатор разговора по электронной почте, содержащего конкретное сообщение.</span><span class="sxs-lookup"><span data-stu-id="a74d4-267">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="a74d4-p107">Вы можете получить целочисленное значение этого свойства, если ваше почтовое приложение активируется в формах просмотра или формах создания ответов. Если пользователь в свою очередь изменит тему ответа, после его отправки идентификатор беседы будет изменен, и полученное ранее значение будет недействительным.</span><span class="sxs-lookup"><span data-stu-id="a74d4-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="a74d4-p108">Это свойство имеет значение NULL для нового элемента в форме создания. Свойство `conversationId` вернет значение, если пользователь задаст тему и сохранит элемент.</span><span class="sxs-lookup"><span data-stu-id="a74d4-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="a74d4-272">Тип:</span><span class="sxs-lookup"><span data-stu-id="a74d4-272">Type:</span></span>

*   <span data-ttu-id="a74d4-273">String</span><span class="sxs-lookup"><span data-stu-id="a74d4-273">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="a74d4-274">Требования</span><span class="sxs-lookup"><span data-stu-id="a74d4-274">Requirements</span></span>

|<span data-ttu-id="a74d4-275">Требование</span><span class="sxs-lookup"><span data-stu-id="a74d4-275">Requirement</span></span>| <span data-ttu-id="a74d4-276">Значение</span><span class="sxs-lookup"><span data-stu-id="a74d4-276">Value</span></span>|
|---|---|
|[<span data-ttu-id="a74d4-277">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="a74d4-277">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a74d4-278">1.0</span><span class="sxs-lookup"><span data-stu-id="a74d4-278">1.0</span></span>|
|[<span data-ttu-id="a74d4-279">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="a74d4-279">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a74d4-280">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a74d4-280">ReadItem</span></span>|
|[<span data-ttu-id="a74d4-281">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="a74d4-281">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a74d4-282">Cоздание или чтение</span><span class="sxs-lookup"><span data-stu-id="a74d4-282">Compose or read</span></span>|

#### <a name="datetimecreated-date"></a><span data-ttu-id="a74d4-283">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="a74d4-283">dateTimeCreated :Date</span></span>

<span data-ttu-id="a74d4-p109">Получает дату и время создания элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="a74d4-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="a74d4-286">Тип:</span><span class="sxs-lookup"><span data-stu-id="a74d4-286">Type:</span></span>

*   <span data-ttu-id="a74d4-287">Date</span><span class="sxs-lookup"><span data-stu-id="a74d4-287">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="a74d4-288">Требования</span><span class="sxs-lookup"><span data-stu-id="a74d4-288">Requirements</span></span>

|<span data-ttu-id="a74d4-289">Требование</span><span class="sxs-lookup"><span data-stu-id="a74d4-289">Requirement</span></span>| <span data-ttu-id="a74d4-290">Значение</span><span class="sxs-lookup"><span data-stu-id="a74d4-290">Value</span></span>|
|---|---|
|[<span data-ttu-id="a74d4-291">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="a74d4-291">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a74d4-292">1.0</span><span class="sxs-lookup"><span data-stu-id="a74d4-292">1.0</span></span>|
|[<span data-ttu-id="a74d4-293">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="a74d4-293">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a74d4-294">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a74d4-294">ReadItem</span></span>|
|[<span data-ttu-id="a74d4-295">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="a74d4-295">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a74d4-296">Чтение</span><span class="sxs-lookup"><span data-stu-id="a74d4-296">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a74d4-297">Пример</span><span class="sxs-lookup"><span data-stu-id="a74d4-297">Example</span></span>

```
var created = Office.context.mailbox.item.dateTimeCreated;
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="a74d4-298">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="a74d4-298">dateTimeModified :Date</span></span>

<span data-ttu-id="a74d4-p110">Получает дату и время последнего изменения элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="a74d4-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="a74d4-301">Этот элемент не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="a74d4-301">Note: This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="a74d4-302">Тип:</span><span class="sxs-lookup"><span data-stu-id="a74d4-302">Type:</span></span>

*   <span data-ttu-id="a74d4-303">Date</span><span class="sxs-lookup"><span data-stu-id="a74d4-303">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="a74d4-304">Требования</span><span class="sxs-lookup"><span data-stu-id="a74d4-304">Requirements</span></span>

|<span data-ttu-id="a74d4-305">Требование</span><span class="sxs-lookup"><span data-stu-id="a74d4-305">Requirement</span></span>| <span data-ttu-id="a74d4-306">Значение</span><span class="sxs-lookup"><span data-stu-id="a74d4-306">Value</span></span>|
|---|---|
|[<span data-ttu-id="a74d4-307">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="a74d4-307">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a74d4-308">1.0</span><span class="sxs-lookup"><span data-stu-id="a74d4-308">1.0</span></span>|
|[<span data-ttu-id="a74d4-309">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="a74d4-309">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a74d4-310">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a74d4-310">ReadItem</span></span>|
|[<span data-ttu-id="a74d4-311">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="a74d4-311">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a74d4-312">Чтение</span><span class="sxs-lookup"><span data-stu-id="a74d4-312">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a74d4-313">Пример</span><span class="sxs-lookup"><span data-stu-id="a74d4-313">Example</span></span>

```
var modified = Office.context.mailbox.item.dateTimeModified;
```

####  <a name="end-datetimejavascriptapioutlook16officetime"></a><span data-ttu-id="a74d4-314">end :Date|[Time](/javascript/api/outlook_1_6/office.time)</span><span class="sxs-lookup"><span data-stu-id="a74d4-314">end :Date|[Time](/javascript/api/outlook_1_6/office.time)</span></span>

<span data-ttu-id="a74d4-315">Получает или задает дату и время окончания встречи.</span><span class="sxs-lookup"><span data-stu-id="a74d4-315">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="a74d4-p111">Свойство `end` представлено в виде значения даты и времени в формате UTC. Преобразовать значение свойства end в местные значения даты и времени клиента можно с помощью метода [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook16officelocalclienttime).</span><span class="sxs-lookup"><span data-stu-id="a74d4-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook16officelocalclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="a74d4-318">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="a74d4-318">Read mode</span></span>

<span data-ttu-id="a74d4-319">Свойство `end` возвращает объект `Date`.</span><span class="sxs-lookup"><span data-stu-id="a74d4-319">The `end` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="a74d4-320">Режим создания</span><span class="sxs-lookup"><span data-stu-id="a74d4-320">Compose mode</span></span>

<span data-ttu-id="a74d4-321">Свойство `end` возвращает объект `Time`.</span><span class="sxs-lookup"><span data-stu-id="a74d4-321">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="a74d4-322">Когда вы используете метод [`Time.setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) для того, чтобы задать время окончания, вы должны использовать метод [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) , чтобы преобразовать местное время на клиенте в формат UTC.</span><span class="sxs-lookup"><span data-stu-id="a74d4-322">When you use the [`Time.setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="a74d4-323">Тип:</span><span class="sxs-lookup"><span data-stu-id="a74d4-323">Type:</span></span>

*   <span data-ttu-id="a74d4-324">Date | [Time](/javascript/api/outlook_1_6/office.time)</span><span class="sxs-lookup"><span data-stu-id="a74d4-324">Date | [Time](/javascript/api/outlook_1_6/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="a74d4-325">Требования</span><span class="sxs-lookup"><span data-stu-id="a74d4-325">Requirements</span></span>

|<span data-ttu-id="a74d4-326">Требование</span><span class="sxs-lookup"><span data-stu-id="a74d4-326">Requirement</span></span>| <span data-ttu-id="a74d4-327">Значение</span><span class="sxs-lookup"><span data-stu-id="a74d4-327">Value</span></span>|
|---|---|
|[<span data-ttu-id="a74d4-328">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="a74d4-328">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a74d4-329">1.0</span><span class="sxs-lookup"><span data-stu-id="a74d4-329">1.0</span></span>|
|[<span data-ttu-id="a74d4-330">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="a74d4-330">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a74d4-331">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a74d4-331">ReadItem</span></span>|
|[<span data-ttu-id="a74d4-332">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="a74d4-332">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a74d4-333">Cоздание или чтение</span><span class="sxs-lookup"><span data-stu-id="a74d4-333">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="a74d4-334">Пример</span><span class="sxs-lookup"><span data-stu-id="a74d4-334">Example</span></span>

<span data-ttu-id="a74d4-335">В примере ниже показано, как с помощью метода [`setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) объекта `Time` задать время окончания встречи в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="a74d4-335">The following example sets the end time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

#### <a name="from-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails"></a><span data-ttu-id="a74d4-336">from :[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="a74d4-336">from :[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)</span></span>

<span data-ttu-id="a74d4-p112">Получает электронный адрес отправителя сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="a74d4-p112">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="a74d4-p113">Свойства `from` и [`sender`](#sender-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails) представляют одно лицо, если сообщение не отправлено представителем. В противном случае свойство `from` представляет лицо, делегировавшее полномочия, а свойство sender — представителя.</span><span class="sxs-lookup"><span data-stu-id="a74d4-p113">The `from` and [`sender`](#sender-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="a74d4-341">Свойство `recipientType` объекта `EmailAddressDetails` в свойстве `from` — `undefined`.</span><span class="sxs-lookup"><span data-stu-id="a74d4-341">Note: The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="a74d4-342">Тип:</span><span class="sxs-lookup"><span data-stu-id="a74d4-342">Type:</span></span>

*   [<span data-ttu-id="a74d4-343">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="a74d4-343">EmailAddressDetails</span></span>](/javascript/api/outlook_1_6/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="a74d4-344">Требования</span><span class="sxs-lookup"><span data-stu-id="a74d4-344">Requirements</span></span>

|<span data-ttu-id="a74d4-345">Требование</span><span class="sxs-lookup"><span data-stu-id="a74d4-345">Requirement</span></span>| <span data-ttu-id="a74d4-346">Значение</span><span class="sxs-lookup"><span data-stu-id="a74d4-346">Value</span></span>|
|---|---|
|[<span data-ttu-id="a74d4-347">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="a74d4-347">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a74d4-348">1.0</span><span class="sxs-lookup"><span data-stu-id="a74d4-348">1.0</span></span>|
|[<span data-ttu-id="a74d4-349">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="a74d4-349">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a74d4-350">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a74d4-350">ReadItem</span></span>|
|[<span data-ttu-id="a74d4-351">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="a74d4-351">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a74d4-352">Чтение</span><span class="sxs-lookup"><span data-stu-id="a74d4-352">Read</span></span>|

#### <a name="internetmessageid-string"></a><span data-ttu-id="a74d4-353">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="a74d4-353">internetMessageId :String</span></span>

<span data-ttu-id="a74d4-p114">Получает идентификатор интернет-сообщения для электронного сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="a74d4-p114">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="a74d4-356">Тип:</span><span class="sxs-lookup"><span data-stu-id="a74d4-356">Type:</span></span>

*   <span data-ttu-id="a74d4-357">String</span><span class="sxs-lookup"><span data-stu-id="a74d4-357">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="a74d4-358">Требования</span><span class="sxs-lookup"><span data-stu-id="a74d4-358">Requirements</span></span>

|<span data-ttu-id="a74d4-359">Требование</span><span class="sxs-lookup"><span data-stu-id="a74d4-359">Requirement</span></span>| <span data-ttu-id="a74d4-360">Значение</span><span class="sxs-lookup"><span data-stu-id="a74d4-360">Value</span></span>|
|---|---|
|[<span data-ttu-id="a74d4-361">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="a74d4-361">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a74d4-362">1.0</span><span class="sxs-lookup"><span data-stu-id="a74d4-362">1.0</span></span>|
|[<span data-ttu-id="a74d4-363">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="a74d4-363">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a74d4-364">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a74d4-364">ReadItem</span></span>|
|[<span data-ttu-id="a74d4-365">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="a74d4-365">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a74d4-366">Чтение</span><span class="sxs-lookup"><span data-stu-id="a74d4-366">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a74d4-367">Пример</span><span class="sxs-lookup"><span data-stu-id="a74d4-367">Example</span></span>

```
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="a74d4-368">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="a74d4-368">itemClass :String</span></span>

<span data-ttu-id="a74d4-p115">Получает класс элемента веб-служб Exchange для выбранного элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="a74d4-p115">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="a74d4-p116">Свойство `itemClass` указывает класс сообщения выбранного элемента. Ниже приводятся классы сообщения по умолчанию для элемента сообщения или встречи.</span><span class="sxs-lookup"><span data-stu-id="a74d4-p116">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="a74d4-373">Тип</span><span class="sxs-lookup"><span data-stu-id="a74d4-373">Type</span></span> | <span data-ttu-id="a74d4-374">Описание</span><span class="sxs-lookup"><span data-stu-id="a74d4-374">Description</span></span> | <span data-ttu-id="a74d4-375">item class</span><span class="sxs-lookup"><span data-stu-id="a74d4-375">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="a74d4-376">Элементы встречи</span><span class="sxs-lookup"><span data-stu-id="a74d4-376">Appointment items</span></span> | <span data-ttu-id="a74d4-377">Это элементы календаря для класса элемента `IPM.Appointment` или `IPM.Appointment.Occurence`.</span><span class="sxs-lookup"><span data-stu-id="a74d4-377">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurence` |
| <span data-ttu-id="a74d4-378">Элементы сообщения</span><span class="sxs-lookup"><span data-stu-id="a74d4-378">Message items</span></span> | <span data-ttu-id="a74d4-379">Сюда входят электронные сообщения, для которых по умолчанию задан класс сообщения `IPM.Note`, а также приглашения на собрания, ответы на них и уведомления об их отмене, использующие `IPM.Schedule.Meeting` в качестве базового класса сообщений.</span><span class="sxs-lookup"><span data-stu-id="a74d4-379">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="a74d4-380">Вы можете создавать настраиваемые классы сообщения, расширяющие классы сообщения по умолчанию, например, настраиваемый класс сообщения о встрече `IPM.Appointment.Contoso`.</span><span class="sxs-lookup"><span data-stu-id="a74d4-380">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="a74d4-381">Тип:</span><span class="sxs-lookup"><span data-stu-id="a74d4-381">Type:</span></span>

*   <span data-ttu-id="a74d4-382">String</span><span class="sxs-lookup"><span data-stu-id="a74d4-382">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="a74d4-383">Требования</span><span class="sxs-lookup"><span data-stu-id="a74d4-383">Requirements</span></span>

|<span data-ttu-id="a74d4-384">Требование</span><span class="sxs-lookup"><span data-stu-id="a74d4-384">Requirement</span></span>| <span data-ttu-id="a74d4-385">Значение</span><span class="sxs-lookup"><span data-stu-id="a74d4-385">Value</span></span>|
|---|---|
|[<span data-ttu-id="a74d4-386">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="a74d4-386">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a74d4-387">1.0</span><span class="sxs-lookup"><span data-stu-id="a74d4-387">1.0</span></span>|
|[<span data-ttu-id="a74d4-388">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="a74d4-388">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a74d4-389">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a74d4-389">ReadItem</span></span>|
|[<span data-ttu-id="a74d4-390">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="a74d4-390">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a74d4-391">Чтение</span><span class="sxs-lookup"><span data-stu-id="a74d4-391">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a74d4-392">Пример</span><span class="sxs-lookup"><span data-stu-id="a74d4-392">Example</span></span>

```
var itemClass = Office.context.mailbox.item.itemClass;
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="a74d4-393">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="a74d4-393">(nullable) itemId :String</span></span>

<span data-ttu-id="a74d4-p117">Получает идентификатор элемента веб-служб Exchange для текущего элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="a74d4-p117">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="a74d4-396">Идентификатор, возвращаемый свойством `itemId`, совпадает с идентификатором элемента веб-служб Exchange.</span><span class="sxs-lookup"><span data-stu-id="a74d4-396">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="a74d4-397">Свойство  `itemId` не совпадает с идентификатором записи Outlook или идентификатором, используемым API-Интерфейсом REST Outlook.</span><span class="sxs-lookup"><span data-stu-id="a74d4-397">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="a74d4-398">Прежде чем осуществлять вызовы API-Интерфейса REST с помощью этого значения, его следует преобразовать с помощью [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="a74d4-398">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="a74d4-399">Дополнительные сведения см. в статье [Использование API REST для Outlook из надстройки Outlook](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="a74d4-399">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="a74d4-p119">Свойство `itemId` недоступно в режиме создания. Если требуется идентификатор элемента, с помощью метода [`saveAsync`](#saveasyncoptions-callback) можно сохранить элемент в хранилище. При этом в параметре [`AsyncResult.value`](/javascript/api/office/office.asyncresult) функции обратного вызова возвращается идентификатор элемента.</span><span class="sxs-lookup"><span data-stu-id="a74d4-p119">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="a74d4-402">Тип:</span><span class="sxs-lookup"><span data-stu-id="a74d4-402">Type:</span></span>

*   <span data-ttu-id="a74d4-403">String</span><span class="sxs-lookup"><span data-stu-id="a74d4-403">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="a74d4-404">Требования</span><span class="sxs-lookup"><span data-stu-id="a74d4-404">Requirements</span></span>

|<span data-ttu-id="a74d4-405">Требование</span><span class="sxs-lookup"><span data-stu-id="a74d4-405">Requirement</span></span>| <span data-ttu-id="a74d4-406">Значение</span><span class="sxs-lookup"><span data-stu-id="a74d4-406">Value</span></span>|
|---|---|
|[<span data-ttu-id="a74d4-407">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="a74d4-407">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a74d4-408">1.0</span><span class="sxs-lookup"><span data-stu-id="a74d4-408">1.0</span></span>|
|[<span data-ttu-id="a74d4-409">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="a74d4-409">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a74d4-410">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a74d4-410">ReadItem</span></span>|
|[<span data-ttu-id="a74d4-411">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="a74d4-411">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a74d4-412">Чтение</span><span class="sxs-lookup"><span data-stu-id="a74d4-412">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a74d4-413">Пример</span><span class="sxs-lookup"><span data-stu-id="a74d4-413">Example</span></span>

<span data-ttu-id="a74d4-p120">Указанный ниже код проверяет наличие идентификатора элемента. Если свойство `itemId` возвращает значение `null` или `undefined`, элемент будет сохранен в хранилище, а из асинхронного результата будет получен идентификатор элемента.</span><span class="sxs-lookup"><span data-stu-id="a74d4-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result){
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook16officemailboxenumsitemtype"></a><span data-ttu-id="a74d4-416">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_6/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="a74d4-416">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_6/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="a74d4-417">Получает тип элемента, который представляет экземпляр.</span><span class="sxs-lookup"><span data-stu-id="a74d4-417">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="a74d4-418">Свойство `itemType` возвращает одно из значений перечисления `ItemType`, которое указывает, является ли экземпляр объекта `item` сообщением или встречей.</span><span class="sxs-lookup"><span data-stu-id="a74d4-418">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="a74d4-419">Тип:</span><span class="sxs-lookup"><span data-stu-id="a74d4-419">Type:</span></span>

*   [<span data-ttu-id="a74d4-420">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="a74d4-420">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook_1_6/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="a74d4-421">Требования</span><span class="sxs-lookup"><span data-stu-id="a74d4-421">Requirements</span></span>

|<span data-ttu-id="a74d4-422">Требование</span><span class="sxs-lookup"><span data-stu-id="a74d4-422">Requirement</span></span>| <span data-ttu-id="a74d4-423">Значение</span><span class="sxs-lookup"><span data-stu-id="a74d4-423">Value</span></span>|
|---|---|
|[<span data-ttu-id="a74d4-424">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="a74d4-424">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a74d4-425">1.0</span><span class="sxs-lookup"><span data-stu-id="a74d4-425">1.0</span></span>|
|[<span data-ttu-id="a74d4-426">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="a74d4-426">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a74d4-427">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a74d4-427">ReadItem</span></span>|
|[<span data-ttu-id="a74d4-428">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="a74d4-428">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a74d4-429">Cоздание или чтение</span><span class="sxs-lookup"><span data-stu-id="a74d4-429">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="a74d4-430">Пример</span><span class="sxs-lookup"><span data-stu-id="a74d4-430">Example</span></span>

```
if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Message)
  // do something
else
  // do something else
```

####  <a name="location-stringlocationjavascriptapioutlook16officelocation"></a><span data-ttu-id="a74d4-431">location :String|[Location](/javascript/api/outlook_1_6/office.location)</span><span class="sxs-lookup"><span data-stu-id="a74d4-431">location :String|[Location](/javascript/api/outlook_1_6/office.location)</span></span>

<span data-ttu-id="a74d4-432">Получает или задает место встречи.</span><span class="sxs-lookup"><span data-stu-id="a74d4-432">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="a74d4-433">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="a74d4-433">Read mode</span></span>

<span data-ttu-id="a74d4-434">Свойство `location` возвращает строку, содержащую сведения о месте встречи.</span><span class="sxs-lookup"><span data-stu-id="a74d4-434">The `location` property returns a string that contains the location of the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="a74d4-435">Режим создания</span><span class="sxs-lookup"><span data-stu-id="a74d4-435">Compose mode</span></span>

<span data-ttu-id="a74d4-436">Свойство `location` возвращает объект `Location`, предоставляющий методы, которые используются для получения и задания места встречи.</span><span class="sxs-lookup"><span data-stu-id="a74d4-436">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="a74d4-437">Тип:</span><span class="sxs-lookup"><span data-stu-id="a74d4-437">Type:</span></span>

*   <span data-ttu-id="a74d4-438">String | [Location](/javascript/api/outlook_1_6/office.location)</span><span class="sxs-lookup"><span data-stu-id="a74d4-438">String | [Location](/javascript/api/outlook_1_6/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="a74d4-439">Требования</span><span class="sxs-lookup"><span data-stu-id="a74d4-439">Requirements</span></span>

|<span data-ttu-id="a74d4-440">Требование</span><span class="sxs-lookup"><span data-stu-id="a74d4-440">Requirement</span></span>| <span data-ttu-id="a74d4-441">Значение</span><span class="sxs-lookup"><span data-stu-id="a74d4-441">Value</span></span>|
|---|---|
|[<span data-ttu-id="a74d4-442">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="a74d4-442">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a74d4-443">1.0</span><span class="sxs-lookup"><span data-stu-id="a74d4-443">1.0</span></span>|
|[<span data-ttu-id="a74d4-444">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="a74d4-444">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a74d4-445">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a74d4-445">ReadItem</span></span>|
|[<span data-ttu-id="a74d4-446">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="a74d4-446">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a74d4-447">Cоздание или чтение</span><span class="sxs-lookup"><span data-stu-id="a74d4-447">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="a74d4-448">Пример</span><span class="sxs-lookup"><span data-stu-id="a74d4-448">Example</span></span>

```
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

#### <a name="normalizedsubject-string"></a><span data-ttu-id="a74d4-449">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="a74d4-449">normalizedSubject :String</span></span>

<span data-ttu-id="a74d4-p121">Получает тему элемента со всеми удаленными префиксами (включая `RE:` и `FWD:`). Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="a74d4-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="a74d4-p122">Свойство normalizedSubject получает тему элемента со стандартными префиксами (такими как `RE:` и `FW:`), добавляемыми почтовыми программами. Для получения темы элемента с неизмененными префиксами используйте свойство [`subject`](#subject-stringsubjectjavascriptapioutlook16officesubject).</span><span class="sxs-lookup"><span data-stu-id="a74d4-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubjectjavascriptapioutlook16officesubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="a74d4-454">Тип:</span><span class="sxs-lookup"><span data-stu-id="a74d4-454">Type:</span></span>

*   <span data-ttu-id="a74d4-455">String</span><span class="sxs-lookup"><span data-stu-id="a74d4-455">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="a74d4-456">Требования</span><span class="sxs-lookup"><span data-stu-id="a74d4-456">Requirements</span></span>

|<span data-ttu-id="a74d4-457">Требование</span><span class="sxs-lookup"><span data-stu-id="a74d4-457">Requirement</span></span>| <span data-ttu-id="a74d4-458">Значение</span><span class="sxs-lookup"><span data-stu-id="a74d4-458">Value</span></span>|
|---|---|
|[<span data-ttu-id="a74d4-459">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="a74d4-459">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a74d4-460">1.0</span><span class="sxs-lookup"><span data-stu-id="a74d4-460">1.0</span></span>|
|[<span data-ttu-id="a74d4-461">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="a74d4-461">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a74d4-462">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a74d4-462">ReadItem</span></span>|
|[<span data-ttu-id="a74d4-463">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="a74d4-463">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a74d4-464">Чтение</span><span class="sxs-lookup"><span data-stu-id="a74d4-464">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a74d4-465">Пример</span><span class="sxs-lookup"><span data-stu-id="a74d4-465">Example</span></span>

```
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
```

####  <a name="notificationmessages-notificationmessagesjavascriptapioutlook16officenotificationmessages"></a><span data-ttu-id="a74d4-466">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_6/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="a74d4-466">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_6/office.notificationmessages)</span></span>

<span data-ttu-id="a74d4-467">Получает сообщения уведомления для элемента.</span><span class="sxs-lookup"><span data-stu-id="a74d4-467">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="a74d4-468">Тип:</span><span class="sxs-lookup"><span data-stu-id="a74d4-468">Type:</span></span>

*   [<span data-ttu-id="a74d4-469">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="a74d4-469">NotificationMessages</span></span>](/javascript/api/outlook_1_6/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="a74d4-470">Требования</span><span class="sxs-lookup"><span data-stu-id="a74d4-470">Requirements</span></span>

|<span data-ttu-id="a74d4-471">Требование</span><span class="sxs-lookup"><span data-stu-id="a74d4-471">Requirement</span></span>| <span data-ttu-id="a74d4-472">Значение</span><span class="sxs-lookup"><span data-stu-id="a74d4-472">Value</span></span>|
|---|---|
|[<span data-ttu-id="a74d4-473">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="a74d4-473">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a74d4-474">1.3</span><span class="sxs-lookup"><span data-stu-id="a74d4-474">1.3</span></span>|
|[<span data-ttu-id="a74d4-475">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="a74d4-475">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a74d4-476">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a74d4-476">ReadItem</span></span>|
|[<span data-ttu-id="a74d4-477">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="a74d4-477">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a74d4-478">Cоздание или чтение</span><span class="sxs-lookup"><span data-stu-id="a74d4-478">Compose or read</span></span>|

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients"></a><span data-ttu-id="a74d4-479">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="a74d4-479">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

<span data-ttu-id="a74d4-480">Предоставляет доступ к необязательным участникам события.</span><span class="sxs-lookup"><span data-stu-id="a74d4-480">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="a74d4-481">Уровень доступа и тип объекта, зависит от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="a74d4-481">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="a74d4-482">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="a74d4-482">Read mode</span></span>

<span data-ttu-id="a74d4-483">Свойство `optionalAttendees` возвращает массив, содержащий объект `EmailAddressDetails` для каждого необязательного участника собрания.</span><span class="sxs-lookup"><span data-stu-id="a74d4-483">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="a74d4-484">Режим создания</span><span class="sxs-lookup"><span data-stu-id="a74d4-484">Compose mode</span></span>

<span data-ttu-id="a74d4-485">Свойство `optionalAttendees` возвращает объект `Recipients`, который предоставляет методы для получения и обновления необязательных участников собрания.</span><span class="sxs-lookup"><span data-stu-id="a74d4-485">The `optionalAttendees` property returns a `Recipients` object that provides methods to get and set the optional attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="a74d4-486">Тип:</span><span class="sxs-lookup"><span data-stu-id="a74d4-486">Type:</span></span>

*   <span data-ttu-id="a74d4-487">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="a74d4-487">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="a74d4-488">Требования</span><span class="sxs-lookup"><span data-stu-id="a74d4-488">Requirements</span></span>

|<span data-ttu-id="a74d4-489">Требование</span><span class="sxs-lookup"><span data-stu-id="a74d4-489">Requirement</span></span>| <span data-ttu-id="a74d4-490">Значение</span><span class="sxs-lookup"><span data-stu-id="a74d4-490">Value</span></span>|
|---|---|
|[<span data-ttu-id="a74d4-491">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="a74d4-491">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a74d4-492">1.0</span><span class="sxs-lookup"><span data-stu-id="a74d4-492">1.0</span></span>|
|[<span data-ttu-id="a74d4-493">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="a74d4-493">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a74d4-494">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a74d4-494">ReadItem</span></span>|
|[<span data-ttu-id="a74d4-495">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="a74d4-495">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a74d4-496">Cоздание или чтение</span><span class="sxs-lookup"><span data-stu-id="a74d4-496">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="a74d4-497">Пример</span><span class="sxs-lookup"><span data-stu-id="a74d4-497">Example</span></span>

```
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails"></a><span data-ttu-id="a74d4-498">organizer :[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="a74d4-498">organizer :[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)</span></span>

<span data-ttu-id="a74d4-p124">Получает электронный адрес организатора указанного собрания. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="a74d4-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="a74d4-501">Тип:</span><span class="sxs-lookup"><span data-stu-id="a74d4-501">Type:</span></span>

*   [<span data-ttu-id="a74d4-502">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="a74d4-502">EmailAddressDetails</span></span>](/javascript/api/outlook_1_6/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="a74d4-503">Требования</span><span class="sxs-lookup"><span data-stu-id="a74d4-503">Requirements</span></span>

|<span data-ttu-id="a74d4-504">Требование</span><span class="sxs-lookup"><span data-stu-id="a74d4-504">Requirement</span></span>| <span data-ttu-id="a74d4-505">Значение</span><span class="sxs-lookup"><span data-stu-id="a74d4-505">Value</span></span>|
|---|---|
|[<span data-ttu-id="a74d4-506">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="a74d4-506">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a74d4-507">1.0</span><span class="sxs-lookup"><span data-stu-id="a74d4-507">1.0</span></span>|
|[<span data-ttu-id="a74d4-508">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="a74d4-508">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a74d4-509">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a74d4-509">ReadItem</span></span>|
|[<span data-ttu-id="a74d4-510">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="a74d4-510">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a74d4-511">Чтение</span><span class="sxs-lookup"><span data-stu-id="a74d4-511">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a74d4-512">Пример</span><span class="sxs-lookup"><span data-stu-id="a74d4-512">Example</span></span>

```
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
```

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients"></a><span data-ttu-id="a74d4-513">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="a74d4-513">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

<span data-ttu-id="a74d4-514">Предоставляет доступ к обязательным участникам события.</span><span class="sxs-lookup"><span data-stu-id="a74d4-514">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="a74d4-515">Уровень доступа и тип объекта, зависит от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="a74d4-515">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="a74d4-516">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="a74d4-516">Read mode</span></span>

<span data-ttu-id="a74d4-517">Свойство `requiredAttendees` возвращает массив, содержащий объект `EmailAddressDetails`, каждому обязательному участнику собрания.</span><span class="sxs-lookup"><span data-stu-id="a74d4-517">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="a74d4-518">Режим создания</span><span class="sxs-lookup"><span data-stu-id="a74d4-518">Compose mode</span></span>

<span data-ttu-id="a74d4-519">Свойство `requiredAttendees` возвращает объект `Recipients`, который предоставляет методы для получения и обновления обязательных участников собрания.</span><span class="sxs-lookup"><span data-stu-id="a74d4-519">The `requiredAttendees` property returns a `Recipients` object that provides methods to get and set the required attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="a74d4-520">Тип:</span><span class="sxs-lookup"><span data-stu-id="a74d4-520">Type:</span></span>

*   <span data-ttu-id="a74d4-521">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="a74d4-521">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="a74d4-522">Требования</span><span class="sxs-lookup"><span data-stu-id="a74d4-522">Requirements</span></span>

|<span data-ttu-id="a74d4-523">Требование</span><span class="sxs-lookup"><span data-stu-id="a74d4-523">Requirement</span></span>| <span data-ttu-id="a74d4-524">Значение</span><span class="sxs-lookup"><span data-stu-id="a74d4-524">Value</span></span>|
|---|---|
|[<span data-ttu-id="a74d4-525">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="a74d4-525">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a74d4-526">1.0</span><span class="sxs-lookup"><span data-stu-id="a74d4-526">1.0</span></span>|
|[<span data-ttu-id="a74d4-527">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="a74d4-527">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a74d4-528">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a74d4-528">ReadItem</span></span>|
|[<span data-ttu-id="a74d4-529">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="a74d4-529">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a74d4-530">Cоздание или чтение</span><span class="sxs-lookup"><span data-stu-id="a74d4-530">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="a74d4-531">Пример</span><span class="sxs-lookup"><span data-stu-id="a74d4-531">Example</span></span>

```
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
}
```

#### <a name="sender-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails"></a><span data-ttu-id="a74d4-532">sender :[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="a74d4-532">sender :[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)</span></span>

<span data-ttu-id="a74d4-p126">Получает электронный адрес отправителя электронного сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="a74d4-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="a74d4-p127">Свойства [`from`](#from-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails) и `sender` представляют одно и то же лицо, если сообщение не отправлено делегатом. В противном случае свойство `from` представляет лицо, делегировавшее полномочия, а свойство sender — делегата.</span><span class="sxs-lookup"><span data-stu-id="a74d4-p127">The [`from`](#from-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="a74d4-537">Свойство `recipientType` объекта `EmailAddressDetails` в свойстве `sender` — `undefined`.</span><span class="sxs-lookup"><span data-stu-id="a74d4-537">Note: The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="a74d4-538">Тип:</span><span class="sxs-lookup"><span data-stu-id="a74d4-538">Type:</span></span>

*   [<span data-ttu-id="a74d4-539">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="a74d4-539">EmailAddressDetails</span></span>](/javascript/api/outlook_1_6/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="a74d4-540">Требования</span><span class="sxs-lookup"><span data-stu-id="a74d4-540">Requirements</span></span>

|<span data-ttu-id="a74d4-541">Требование</span><span class="sxs-lookup"><span data-stu-id="a74d4-541">Requirement</span></span>| <span data-ttu-id="a74d4-542">Значение</span><span class="sxs-lookup"><span data-stu-id="a74d4-542">Value</span></span>|
|---|---|
|[<span data-ttu-id="a74d4-543">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="a74d4-543">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a74d4-544">1.0</span><span class="sxs-lookup"><span data-stu-id="a74d4-544">1.0</span></span>|
|[<span data-ttu-id="a74d4-545">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="a74d4-545">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a74d4-546">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a74d4-546">ReadItem</span></span>|
|[<span data-ttu-id="a74d4-547">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="a74d4-547">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a74d4-548">Чтение</span><span class="sxs-lookup"><span data-stu-id="a74d4-548">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a74d4-549">Пример</span><span class="sxs-lookup"><span data-stu-id="a74d4-549">Example</span></span>

```
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
```

####  <a name="start-datetimejavascriptapioutlook16officetime"></a><span data-ttu-id="a74d4-550">start :Date|[Time](/javascript/api/outlook_1_6/office.time)</span><span class="sxs-lookup"><span data-stu-id="a74d4-550">start :Date|[Time](/javascript/api/outlook_1_6/office.time)</span></span>

<span data-ttu-id="a74d4-551">Получает или задает дату и время начала встречи.</span><span class="sxs-lookup"><span data-stu-id="a74d4-551">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="a74d4-p128">Свойство `start` представлено в виде значения даты и времени в формате UTC. Это значение можно преобразовать в местные значения даты и времени клиента с помощью метода [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook16officelocalclienttime).</span><span class="sxs-lookup"><span data-stu-id="a74d4-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook16officelocalclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="a74d4-554">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="a74d4-554">Read mode</span></span>

<span data-ttu-id="a74d4-555">Свойство `start` возвращает объект `Date`.</span><span class="sxs-lookup"><span data-stu-id="a74d4-555">The `start` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="a74d4-556">Режим создания</span><span class="sxs-lookup"><span data-stu-id="a74d4-556">Compose mode</span></span>

<span data-ttu-id="a74d4-557">Свойство `start` возвращает объект `Time`.</span><span class="sxs-lookup"><span data-stu-id="a74d4-557">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="a74d4-558">Если вы задаете время начала с помощью метода [`Time.setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-), необходимо использовать метод [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) для преобразования местного времени на клиенте в формат UTC для сервера.</span><span class="sxs-lookup"><span data-stu-id="a74d4-558">When you use the [`Time.setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="a74d4-559">Тип:</span><span class="sxs-lookup"><span data-stu-id="a74d4-559">Type:</span></span>

*   <span data-ttu-id="a74d4-560">Date | [Time](/javascript/api/outlook_1_6/office.time)</span><span class="sxs-lookup"><span data-stu-id="a74d4-560">Date | [Time](/javascript/api/outlook_1_6/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="a74d4-561">Требования</span><span class="sxs-lookup"><span data-stu-id="a74d4-561">Requirements</span></span>

|<span data-ttu-id="a74d4-562">Требование</span><span class="sxs-lookup"><span data-stu-id="a74d4-562">Requirement</span></span>| <span data-ttu-id="a74d4-563">Значение</span><span class="sxs-lookup"><span data-stu-id="a74d4-563">Value</span></span>|
|---|---|
|[<span data-ttu-id="a74d4-564">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="a74d4-564">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a74d4-565">1.0</span><span class="sxs-lookup"><span data-stu-id="a74d4-565">1.0</span></span>|
|[<span data-ttu-id="a74d4-566">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="a74d4-566">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a74d4-567">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a74d4-567">ReadItem</span></span>|
|[<span data-ttu-id="a74d4-568">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="a74d4-568">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a74d4-569">Cоздание или чтение</span><span class="sxs-lookup"><span data-stu-id="a74d4-569">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="a74d4-570">Пример</span><span class="sxs-lookup"><span data-stu-id="a74d4-570">Example</span></span>

<span data-ttu-id="a74d4-571">В примере ниже с помощью метода [`setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) объекта `Time` задается время начала встречи в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="a74d4-571">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

####  <a name="subject-stringsubjectjavascriptapioutlook16officesubject"></a><span data-ttu-id="a74d4-572">subject :String|[Subject](/javascript/api/outlook_1_6/office.subject)</span><span class="sxs-lookup"><span data-stu-id="a74d4-572">subject :String|[Subject](/javascript/api/outlook_1_6/office.subject)</span></span>

<span data-ttu-id="a74d4-573">Получает или задает описание, которое отображается в поле темы элемента.</span><span class="sxs-lookup"><span data-stu-id="a74d4-573">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="a74d4-574">Свойство `subject` получает или задает всю тему элемента для отправки с почтового сервера.</span><span class="sxs-lookup"><span data-stu-id="a74d4-574">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="a74d4-575">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="a74d4-575">Read mode</span></span>

<span data-ttu-id="a74d4-p129">Свойство `subject` возвращает строку. С помощью свойства [`normalizedSubject`](#normalizedsubject-string) можно получить тему без начальных префиксов, например, `RE:` и `FW:`.</span><span class="sxs-lookup"><span data-stu-id="a74d4-p129">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```
var subject = Office.context.mailbox.item.subject;
```

##### <a name="compose-mode"></a><span data-ttu-id="a74d4-578">Режим создания</span><span class="sxs-lookup"><span data-stu-id="a74d4-578">Compose mode</span></span>

<span data-ttu-id="a74d4-579">Свойство `subject` возвращает объект `Subject`, который предоставляет методы для получения и задания темы.</span><span class="sxs-lookup"><span data-stu-id="a74d4-579">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="a74d4-580">Тип:</span><span class="sxs-lookup"><span data-stu-id="a74d4-580">Type:</span></span>

*   <span data-ttu-id="a74d4-581">String | [Subject](/javascript/api/outlook_1_6/office.subject)</span><span class="sxs-lookup"><span data-stu-id="a74d4-581">String | [Subject](/javascript/api/outlook_1_6/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="a74d4-582">Требования</span><span class="sxs-lookup"><span data-stu-id="a74d4-582">Requirements</span></span>

|<span data-ttu-id="a74d4-583">Требование</span><span class="sxs-lookup"><span data-stu-id="a74d4-583">Requirement</span></span>| <span data-ttu-id="a74d4-584">Значение</span><span class="sxs-lookup"><span data-stu-id="a74d4-584">Value</span></span>|
|---|---|
|[<span data-ttu-id="a74d4-585">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="a74d4-585">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a74d4-586">1.0</span><span class="sxs-lookup"><span data-stu-id="a74d4-586">1.0</span></span>|
|[<span data-ttu-id="a74d4-587">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="a74d4-587">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a74d4-588">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a74d4-588">ReadItem</span></span>|
|[<span data-ttu-id="a74d4-589">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="a74d4-589">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a74d4-590">Cоздание или чтение</span><span class="sxs-lookup"><span data-stu-id="a74d4-590">Compose or read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients"></a><span data-ttu-id="a74d4-591">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="a74d4-591">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

<span data-ttu-id="a74d4-592">Предоставляет доступ получателей к строке **To (Кому)** в сообщении.</span><span class="sxs-lookup"><span data-stu-id="a74d4-592">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="a74d4-593">Уровень доступа и тип объекта, зависит от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="a74d4-593">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="a74d4-594">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="a74d4-594">Read mode</span></span>

<span data-ttu-id="a74d4-p131">Свойство `to` возвращает массив, содержащий объект `EmailAddressDetails` для каждого получателя в строке **To (Кому)** сообщения. Коллекция может включать не более 100 элементов.</span><span class="sxs-lookup"><span data-stu-id="a74d4-p131">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="a74d4-597">Режим создания</span><span class="sxs-lookup"><span data-stu-id="a74d4-597">Compose mode</span></span>

<span data-ttu-id="a74d4-598">Свойство `to` возвращает объект `Recipients`, предоставляющий методы для получения или обновления получателей, которые указаны в строке **To (кому)** сообщения.</span><span class="sxs-lookup"><span data-stu-id="a74d4-598">The to`to` property returns a Recipients`Recipients` object that provides methods for manipulating the recipients on the **To** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="a74d4-599">Тип:</span><span class="sxs-lookup"><span data-stu-id="a74d4-599">Type:</span></span>

*   <span data-ttu-id="a74d4-600">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="a74d4-600">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="a74d4-601">Требования</span><span class="sxs-lookup"><span data-stu-id="a74d4-601">Requirements</span></span>

|<span data-ttu-id="a74d4-602">Требование</span><span class="sxs-lookup"><span data-stu-id="a74d4-602">Requirement</span></span>| <span data-ttu-id="a74d4-603">Значение</span><span class="sxs-lookup"><span data-stu-id="a74d4-603">Value</span></span>|
|---|---|
|[<span data-ttu-id="a74d4-604">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="a74d4-604">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a74d4-605">1.0</span><span class="sxs-lookup"><span data-stu-id="a74d4-605">1.0</span></span>|
|[<span data-ttu-id="a74d4-606">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="a74d4-606">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a74d4-607">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a74d4-607">ReadItem</span></span>|
|[<span data-ttu-id="a74d4-608">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="a74d4-608">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a74d4-609">Cоздание или чтение</span><span class="sxs-lookup"><span data-stu-id="a74d4-609">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="a74d4-610">Пример</span><span class="sxs-lookup"><span data-stu-id="a74d4-610">Example</span></span>

```
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

### <a name="methods"></a><span data-ttu-id="a74d4-611">Методы</span><span class="sxs-lookup"><span data-stu-id="a74d4-611">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="a74d4-612">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="a74d4-612">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="a74d4-613">Добавляет файл в сообщение или встречу в качестве вложения.</span><span class="sxs-lookup"><span data-stu-id="a74d4-613">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="a74d4-614">Метод `addFileAttachmentAsync` передает файл по указанному универсальному коду ресурса (URI) и вкладывает его в элемент в форме создания.</span><span class="sxs-lookup"><span data-stu-id="a74d4-614">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="a74d4-615">Идентификатор можно последовательно использовать с методом [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback), чтобы удалить вложение, добавленное во время текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="a74d4-615">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a74d4-616">Параметры:</span><span class="sxs-lookup"><span data-stu-id="a74d4-616">Parameters:</span></span>

|<span data-ttu-id="a74d4-617">Имя</span><span class="sxs-lookup"><span data-stu-id="a74d4-617">Name</span></span>| <span data-ttu-id="a74d4-618">Тип</span><span class="sxs-lookup"><span data-stu-id="a74d4-618">Type</span></span>| <span data-ttu-id="a74d4-619">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="a74d4-619">Attributes</span></span>| <span data-ttu-id="a74d4-620">Описание</span><span class="sxs-lookup"><span data-stu-id="a74d4-620">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="a74d4-621">String</span><span class="sxs-lookup"><span data-stu-id="a74d4-621">String</span></span>||<span data-ttu-id="a74d4-p132">URI-адрес, представляющий расположение файла, который нужно вложить в сообщение или встречу. Максимальная длина — 2048 символов.</span><span class="sxs-lookup"><span data-stu-id="a74d4-p132">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="a74d4-624">String</span><span class="sxs-lookup"><span data-stu-id="a74d4-624">String</span></span>||<span data-ttu-id="a74d4-p133">Имя вложения, которое отображается при передаче вложения. Максимальная длина — 255 символов.</span><span class="sxs-lookup"><span data-stu-id="a74d4-p133">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="a74d4-627">Объект</span><span class="sxs-lookup"><span data-stu-id="a74d4-627">Object</span></span>| <span data-ttu-id="a74d4-628">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a74d4-628">&lt;optional&gt;</span></span>|<span data-ttu-id="a74d4-629">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="a74d4-629">An object literal that contains one or more of the following properties.</span></span>|
| `options.asyncContext` | <span data-ttu-id="a74d4-630">Объект</span><span class="sxs-lookup"><span data-stu-id="a74d4-630">Object</span></span> | <span data-ttu-id="a74d4-631">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="a74d4-631">&lt;optional&gt;</span></span> | <span data-ttu-id="a74d4-632">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="a74d4-632">Developers can provide any object they wish to access in the callback method.</span></span> |
| `options.isInline` | <span data-ttu-id="a74d4-633">Логическое значение</span><span class="sxs-lookup"><span data-stu-id="a74d4-633">Boolean</span></span> | <span data-ttu-id="a74d4-634">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="a74d4-634">&lt;optional&gt;</span></span> | <span data-ttu-id="a74d4-635">Значение `true` указывает на то, что вложение будет встроено в текст сообщения и не должно отображаться в списке вложений.</span><span class="sxs-lookup"><span data-stu-id="a74d4-635">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
|`callback`| <span data-ttu-id="a74d4-636">function</span><span class="sxs-lookup"><span data-stu-id="a74d4-636">function</span></span>| <span data-ttu-id="a74d4-637">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="a74d4-637">&lt;optional&gt;</span></span>|<span data-ttu-id="a74d4-638">Когда метод завершает выполнение, переданная в параметре `callback` функция вызывается с единственным параметром `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="a74d4-638">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="a74d4-639">После успешного выполнения идентификатор вложения будет представлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="a74d4-639">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="a74d4-640">Если добавить вложение не удастся, объект `asyncResult` будет содержать объект `Error` с описанием ошибки.</span><span class="sxs-lookup"><span data-stu-id="a74d4-640">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="a74d4-641">Ошибки</span><span class="sxs-lookup"><span data-stu-id="a74d4-641">Errors</span></span>

| <span data-ttu-id="a74d4-642">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="a74d4-642">Error code</span></span> | <span data-ttu-id="a74d4-643">Описание</span><span class="sxs-lookup"><span data-stu-id="a74d4-643">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="a74d4-644">Вложение превышает максимальный размер.</span><span class="sxs-lookup"><span data-stu-id="a74d4-644">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="a74d4-645">Расширение вложения не поддерживается.</span><span class="sxs-lookup"><span data-stu-id="a74d4-645">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="a74d4-646">Сообщение или встреча содержат слишком много вложений.</span><span class="sxs-lookup"><span data-stu-id="a74d4-646">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="a74d4-647">Требования</span><span class="sxs-lookup"><span data-stu-id="a74d4-647">Requirements</span></span>

|<span data-ttu-id="a74d4-648">Требование</span><span class="sxs-lookup"><span data-stu-id="a74d4-648">Requirement</span></span>| <span data-ttu-id="a74d4-649">Значение</span><span class="sxs-lookup"><span data-stu-id="a74d4-649">Value</span></span>|
|---|---|
|[<span data-ttu-id="a74d4-650">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="a74d4-650">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a74d4-651">1.1</span><span class="sxs-lookup"><span data-stu-id="a74d4-651">1.1</span></span>|
|[<span data-ttu-id="a74d4-652">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="a74d4-652">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a74d4-653">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="a74d4-653">ReadWriteItem</span></span>|
|[<span data-ttu-id="a74d4-654">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="a74d4-654">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a74d4-655">Создание</span><span class="sxs-lookup"><span data-stu-id="a74d4-655">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="a74d4-656">Примеры</span><span class="sxs-lookup"><span data-stu-id="a74d4-656">Examples</span></span>

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

<span data-ttu-id="a74d4-657">В приведенном ниже примере файл изображения добавляется в качестве встроенного вложения, а ссылка на вложение добавляется в текст сообщения.</span><span class="sxs-lookup"><span data-stu-id="a74d4-657">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

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

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="a74d4-658">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="a74d4-658">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="a74d4-659">Добавляет к сообщению или встрече элемент Exchange (например, сообщение) в виде вложения.</span><span class="sxs-lookup"><span data-stu-id="a74d4-659">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="a74d4-p134">С помощью метода `addItemAttachmentAsync` в элемент формы создания можно вложить элемент с указанным идентификатором Exchange. Если указать метод обратного вызова, то этот метод вызывается с помощью параметра `asyncResult`, который содержит идентификатор вложения или код, указывающий на ошибки, которые произошли при вложении элемента. При необходимости можно использовать параметр `options` для передачи сведений о состоянии в метод обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="a74d4-p134">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="a74d4-663">Идентификатор можно последовательно использовать с методом [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback), чтобы удалить вложение, добавленное во время текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="a74d4-663">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="a74d4-664">Если ваша надстройка Office выполняется в веб-приложении Outlook, метод `addItemAttachmentAsync` обеспечивает вложение элементов в элементы, отличные от редактируемого. Однако это действие не рекомендуется выполнять, так как оно не поддерживается.</span><span class="sxs-lookup"><span data-stu-id="a74d4-664">If your Office add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a74d4-665">Параметры:</span><span class="sxs-lookup"><span data-stu-id="a74d4-665">Parameters:</span></span>

|<span data-ttu-id="a74d4-666">Имя</span><span class="sxs-lookup"><span data-stu-id="a74d4-666">Name</span></span>| <span data-ttu-id="a74d4-667">Тип</span><span class="sxs-lookup"><span data-stu-id="a74d4-667">Type</span></span>| <span data-ttu-id="a74d4-668">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="a74d4-668">Attributes</span></span>| <span data-ttu-id="a74d4-669">Описание</span><span class="sxs-lookup"><span data-stu-id="a74d4-669">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="a74d4-670">String</span><span class="sxs-lookup"><span data-stu-id="a74d4-670">String</span></span>||<span data-ttu-id="a74d4-p135">Идентификатор Exchange для вкладываемого элемента. Максимальная длина — 100 символов.</span><span class="sxs-lookup"><span data-stu-id="a74d4-p135">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="a74d4-673">Строка</span><span class="sxs-lookup"><span data-stu-id="a74d4-673">String</span></span>||<span data-ttu-id="a74d4-p136">Тема вкладываемого элемента. Максимальная длина — 255 символов.</span><span class="sxs-lookup"><span data-stu-id="a74d4-p136">The sujbect of the item to be attached. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="a74d4-676">Объект</span><span class="sxs-lookup"><span data-stu-id="a74d4-676">Object</span></span>| <span data-ttu-id="a74d4-677">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a74d4-677">&lt;optional&gt;</span></span>|<span data-ttu-id="a74d4-678">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="a74d4-678">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="a74d4-679">Объект</span><span class="sxs-lookup"><span data-stu-id="a74d4-679">Object</span></span>| <span data-ttu-id="a74d4-680">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="a74d4-680">&lt;optional&gt;</span></span>|<span data-ttu-id="a74d4-681">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="a74d4-681">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="a74d4-682">function</span><span class="sxs-lookup"><span data-stu-id="a74d4-682">function</span></span>| <span data-ttu-id="a74d4-683">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="a74d4-683">&lt;optional&gt;</span></span>|<span data-ttu-id="a74d4-684">Когда метод завершает выполнение, переданная в параметре `callback` функция вызывается с единственным параметром `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="a74d4-684">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="a74d4-685">После успешного выполнения идентификатор вложения будет представлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="a74d4-685">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="a74d4-686">Если добавить вложение не удастся, объект `asyncResult` будет содержать объект `Error` с описанием ошибки.</span><span class="sxs-lookup"><span data-stu-id="a74d4-686">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="a74d4-687">Ошибки</span><span class="sxs-lookup"><span data-stu-id="a74d4-687">Errors</span></span>

| <span data-ttu-id="a74d4-688">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="a74d4-688">Error code</span></span> | <span data-ttu-id="a74d4-689">Описание</span><span class="sxs-lookup"><span data-stu-id="a74d4-689">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="a74d4-690">Сообщение или встреча содержат слишком много вложений.</span><span class="sxs-lookup"><span data-stu-id="a74d4-690">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="a74d4-691">Требования</span><span class="sxs-lookup"><span data-stu-id="a74d4-691">Requirements</span></span>

|<span data-ttu-id="a74d4-692">Требование</span><span class="sxs-lookup"><span data-stu-id="a74d4-692">Requirement</span></span>| <span data-ttu-id="a74d4-693">Значение</span><span class="sxs-lookup"><span data-stu-id="a74d4-693">Value</span></span>|
|---|---|
|[<span data-ttu-id="a74d4-694">Версия минимального набора требований для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="a74d4-694">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a74d4-695">1.1</span><span class="sxs-lookup"><span data-stu-id="a74d4-695">1.1</span></span>|
|[<span data-ttu-id="a74d4-696">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="a74d4-696">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a74d4-697">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="a74d4-697">ReadWriteItem</span></span>|
|[<span data-ttu-id="a74d4-698">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="a74d4-698">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a74d4-699">Создание</span><span class="sxs-lookup"><span data-stu-id="a74d4-699">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="a74d4-700">Пример</span><span class="sxs-lookup"><span data-stu-id="a74d4-700">Example</span></span>

<span data-ttu-id="a74d4-701">В следующем примере существующий элемент Outlook добавляется в виде вложения с именем `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="a74d4-701">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

####  <a name="close"></a><span data-ttu-id="a74d4-702">close()</span><span class="sxs-lookup"><span data-stu-id="a74d4-702">close()</span></span>

<span data-ttu-id="a74d4-703">Закрывает текущий создаваемый элемент.</span><span class="sxs-lookup"><span data-stu-id="a74d4-703">Closes the current item that is being composed.</span></span>

<span data-ttu-id="a74d4-p137">Работа метода `close` зависит от текущего состояния создаваемого элемента. Если элемент содержит несохраненные изменения, клиент предложит пользователю сохранить или отклонить их либо отменить действие закрытия.</span><span class="sxs-lookup"><span data-stu-id="a74d4-p137">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="a74d4-706">Если элемент является встречей в Outlook в Интернете, и он был ранее сохранен с помощью `saveAsync`, пользователю предлагается сохранить, отменить или удалить его, даже если не произошло каких-либо изменений, поскольку этот элемент был последним сохраненным.</span><span class="sxs-lookup"><span data-stu-id="a74d4-706">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="a74d4-707">Если в клиенте Outlook для настольных ПК сообщение представляет собой ответ в тексте, метод `close` не работает.</span><span class="sxs-lookup"><span data-stu-id="a74d4-707">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="a74d4-708">Требования</span><span class="sxs-lookup"><span data-stu-id="a74d4-708">Requirements</span></span>

|<span data-ttu-id="a74d4-709">Требование</span><span class="sxs-lookup"><span data-stu-id="a74d4-709">Requirement</span></span>| <span data-ttu-id="a74d4-710">Значение</span><span class="sxs-lookup"><span data-stu-id="a74d4-710">Value</span></span>|
|---|---|
|[<span data-ttu-id="a74d4-711">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="a74d4-711">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a74d4-712">1.3</span><span class="sxs-lookup"><span data-stu-id="a74d4-712">1.3</span></span>|
|[<span data-ttu-id="a74d4-713">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="a74d4-713">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a74d4-714">Ограниченный доступ</span><span class="sxs-lookup"><span data-stu-id="a74d4-714">Restricted</span></span>|
|[<span data-ttu-id="a74d4-715">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="a74d4-715">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a74d4-716">Создание</span><span class="sxs-lookup"><span data-stu-id="a74d4-716">Compose</span></span>|

#### <a name="displayreplyallformformdata"></a><span data-ttu-id="a74d4-717">displayReplyAllForm(formData)</span><span class="sxs-lookup"><span data-stu-id="a74d4-717">displayReplyAllForm(formData)</span></span>

<span data-ttu-id="a74d4-718">Отображает форму ответа, включающую отправителя и всех получателей выбранного сообщения или организатора и всех участников выбранной встречи.</span><span class="sxs-lookup"><span data-stu-id="a74d4-718">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="a74d4-719">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="a74d4-719">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="a74d4-720">В веб-приложении Outlook форма ответа отображается в виде всплывающей формы в представлении с 3 колонками либо всплывающей формы в представлении с 1 или 2 колонками.</span><span class="sxs-lookup"><span data-stu-id="a74d4-720">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="a74d4-721">Если любой строчный параметр превышает указанные для него ограничения, `displayReplyAllForm` возвращает исключение.</span><span class="sxs-lookup"><span data-stu-id="a74d4-721">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="a74d4-p138">Если в параметре `formData.attachments` указаны вложения, Outlook и веб-приложение Outlook пытаются скачать их и вложить в форму ответа. Если какие-либо вложения добавить не удается, в форме отображается сообщение об ошибке. Если сообщения об ошибках не предусмотрены, то они не отображаются.</span><span class="sxs-lookup"><span data-stu-id="a74d4-p138">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a74d4-725">Параметры:</span><span class="sxs-lookup"><span data-stu-id="a74d4-725">Parameters:</span></span>

| <span data-ttu-id="a74d4-726">Имя</span><span class="sxs-lookup"><span data-stu-id="a74d4-726">Name</span></span> | <span data-ttu-id="a74d4-727">Тип</span><span class="sxs-lookup"><span data-stu-id="a74d4-727">Type</span></span> | <span data-ttu-id="a74d4-728">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="a74d4-728">Attributes</span></span> | <span data-ttu-id="a74d4-729">Описание</span><span class="sxs-lookup"><span data-stu-id="a74d4-729">Description</span></span> |
|---|---|---|---|
|`formData`| <span data-ttu-id="a74d4-730">String | Object</span><span class="sxs-lookup"><span data-stu-id="a74d4-730">String &#124; Object</span></span>| |<span data-ttu-id="a74d4-p139">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="a74d4-p139">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="a74d4-733">**ИЛИ**</span><span class="sxs-lookup"><span data-stu-id="a74d4-733">**OR**</span></span><br/><span data-ttu-id="a74d4-p140">Объект, который содержит текст или данные вложения и функцию обратного вызова. Ниже представлено определение этого объекта.</span><span class="sxs-lookup"><span data-stu-id="a74d4-p140">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="a74d4-736">String</span><span class="sxs-lookup"><span data-stu-id="a74d4-736">String</span></span> | <span data-ttu-id="a74d4-737">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="a74d4-737">&lt;optional&gt;</span></span> | <span data-ttu-id="a74d4-p141">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="a74d4-p141">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="a74d4-740">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="a74d4-740">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="a74d4-741">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="a74d4-741">&lt;optional&gt;</span></span> | <span data-ttu-id="a74d4-742">Массив объектов JSON, представляющих собой вложенные файлы или элементы.</span><span class="sxs-lookup"><span data-stu-id="a74d4-742">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="a74d4-743">String</span><span class="sxs-lookup"><span data-stu-id="a74d4-743">String</span></span> | | <span data-ttu-id="a74d4-p142">Указывает тип вложения. Допустимые значения: `file` для вложенного файла и `item` для вложенного элемента.</span><span class="sxs-lookup"><span data-stu-id="a74d4-p142">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="a74d4-746">String</span><span class="sxs-lookup"><span data-stu-id="a74d4-746">String</span></span> | | <span data-ttu-id="a74d4-747">Строка, содержащая имя вложения, длиной до 255 символов.</span><span class="sxs-lookup"><span data-stu-id="a74d4-747">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="a74d4-748">String</span><span class="sxs-lookup"><span data-stu-id="a74d4-748">String</span></span> | | <span data-ttu-id="a74d4-p143">Используется, только если свойству `type` задано значение `file`. URI расположения файла.</span><span class="sxs-lookup"><span data-stu-id="a74d4-p143">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.isInline` | <span data-ttu-id="a74d4-751">Логическое значение</span><span class="sxs-lookup"><span data-stu-id="a74d4-751">Boolean</span></span> | | <span data-ttu-id="a74d4-p144">Используется только в том случае, если свойству `type` задано значение `file`. Значение `true` указывает на то, что вложение будет встроено в текст сообщения и не должно отображаться в списке вложений.</span><span class="sxs-lookup"><span data-stu-id="a74d4-p144">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="a74d4-754">String</span><span class="sxs-lookup"><span data-stu-id="a74d4-754">String</span></span> | | <span data-ttu-id="a74d4-p145">Используется, только если свойству `type` задано значение `item`. Идентификатор вложения EWS. Это строка длиной до 100 символов.</span><span class="sxs-lookup"><span data-stu-id="a74d4-p145">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="a74d4-758">function</span><span class="sxs-lookup"><span data-stu-id="a74d4-758">function</span></span> | <span data-ttu-id="a74d4-759">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="a74d4-759">&lt;optional&gt;</span></span> | <span data-ttu-id="a74d4-760">По завершении работы метода функция, переданная параметру `callback`, вызывается с помощью одного параметра `asyncResult`, представляющего собой объект [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="a74d4-760">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="a74d4-761">Требования</span><span class="sxs-lookup"><span data-stu-id="a74d4-761">Requirements</span></span>

|<span data-ttu-id="a74d4-762">Требование</span><span class="sxs-lookup"><span data-stu-id="a74d4-762">Requirement</span></span>| <span data-ttu-id="a74d4-763">Значение</span><span class="sxs-lookup"><span data-stu-id="a74d4-763">Value</span></span>|
|---|---|
|[<span data-ttu-id="a74d4-764">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="a74d4-764">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a74d4-765">1.0</span><span class="sxs-lookup"><span data-stu-id="a74d4-765">1.0</span></span>|
|[<span data-ttu-id="a74d4-766">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="a74d4-766">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a74d4-767">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a74d4-767">ReadItem</span></span>|
|[<span data-ttu-id="a74d4-768">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="a74d4-768">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a74d4-769">Чтение</span><span class="sxs-lookup"><span data-stu-id="a74d4-769">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="a74d4-770">Примеры</span><span class="sxs-lookup"><span data-stu-id="a74d4-770">Examples</span></span>

<span data-ttu-id="a74d4-771">Приведенный ниже код передает строку в функцию `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="a74d4-771">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="a74d4-772">Ответ с пустым текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="a74d4-772">Reply with an empty body.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="a74d4-773">Ответ только с текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="a74d4-773">Reply with just a body.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="a74d4-774">Ответ с текстом сообщения и вложенным файлом.</span><span class="sxs-lookup"><span data-stu-id="a74d4-774">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="a74d4-775">Ответ с текстом сообщения и вложенным элементом.</span><span class="sxs-lookup"><span data-stu-id="a74d4-775">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="a74d4-776">Ответ с текстом сообщения, вложенным файлом, вложенным элементом и обратным вызовом.</span><span class="sxs-lookup"><span data-stu-id="a74d4-776">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata"></a><span data-ttu-id="a74d4-777">displayReplyForm(formData)</span><span class="sxs-lookup"><span data-stu-id="a74d4-777">displayReplyForm(formData)</span></span>

<span data-ttu-id="a74d4-778">Отображает форму ответа, включающую только отправителя выбранного сообщения или организатора выбранной встречи.</span><span class="sxs-lookup"><span data-stu-id="a74d4-778">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="a74d4-779">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="a74d4-779">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="a74d4-780">В веб-приложении Outlook форма ответа отображается в виде всплывающей формы в представлении с 3 колонками либо всплывающей формы в представлении с 1 или 2 колонками.</span><span class="sxs-lookup"><span data-stu-id="a74d4-780">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="a74d4-781">Если любой строчный параметр превышает указанные для него ограничения, `displayReplyForm` возвращает исключение.</span><span class="sxs-lookup"><span data-stu-id="a74d4-781">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="a74d4-p146">Если в параметре `formData.attachments` указаны вложения, Outlook и веб-приложение Outlook пытаются скачать их и вложить в форму ответа. Если какие-либо вложения добавить не удается, в форме отображается сообщение об ошибке. Если сообщения об ошибках не предусмотрены, то они не отображаются.</span><span class="sxs-lookup"><span data-stu-id="a74d4-p146">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a74d4-785">Параметры:</span><span class="sxs-lookup"><span data-stu-id="a74d4-785">Parameters:</span></span>

| <span data-ttu-id="a74d4-786">Имя</span><span class="sxs-lookup"><span data-stu-id="a74d4-786">Name</span></span> | <span data-ttu-id="a74d4-787">Тип</span><span class="sxs-lookup"><span data-stu-id="a74d4-787">Type</span></span> | <span data-ttu-id="a74d4-788">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="a74d4-788">Attributes</span></span> | <span data-ttu-id="a74d4-789">Описание</span><span class="sxs-lookup"><span data-stu-id="a74d4-789">Description</span></span> |
|---|---|---|---|
|`formData`| <span data-ttu-id="a74d4-790">String | Object</span><span class="sxs-lookup"><span data-stu-id="a74d4-790">String &#124; Object</span></span>| | <span data-ttu-id="a74d4-p147">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="a74d4-p147">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="a74d4-793">**ИЛИ**</span><span class="sxs-lookup"><span data-stu-id="a74d4-793">**OR**</span></span><br/><span data-ttu-id="a74d4-p148">Объект, который содержит текст или данные вложения и функцию обратного вызова. Ниже представлено определение этого объекта.</span><span class="sxs-lookup"><span data-stu-id="a74d4-p148">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="a74d4-796">String</span><span class="sxs-lookup"><span data-stu-id="a74d4-796">String</span></span> | <span data-ttu-id="a74d4-797">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="a74d4-797">&lt;optional&gt;</span></span> | <span data-ttu-id="a74d4-p149">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="a74d4-p149">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="a74d4-800">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="a74d4-800">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="a74d4-801">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="a74d4-801">&lt;optional&gt;</span></span> | <span data-ttu-id="a74d4-802">Массив объектов JSON, представляющих собой вложенные файлы или элементы.</span><span class="sxs-lookup"><span data-stu-id="a74d4-802">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="a74d4-803">String</span><span class="sxs-lookup"><span data-stu-id="a74d4-803">String</span></span> | | <span data-ttu-id="a74d4-p150">Указывает тип вложения. Допустимые значения: `file` для вложенного файла и `item` для вложенного элемента.</span><span class="sxs-lookup"><span data-stu-id="a74d4-p150">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="a74d4-806">String</span><span class="sxs-lookup"><span data-stu-id="a74d4-806">String</span></span> | | <span data-ttu-id="a74d4-807">Строка, содержащая имя вложения, длиной до 255 символов.</span><span class="sxs-lookup"><span data-stu-id="a74d4-807">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="a74d4-808">String</span><span class="sxs-lookup"><span data-stu-id="a74d4-808">String</span></span> | | <span data-ttu-id="a74d4-p151">Используется, только если свойству `type` задано значение `file`. URI расположения файла.</span><span class="sxs-lookup"><span data-stu-id="a74d4-p151">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.isInline` | <span data-ttu-id="a74d4-811">Логическое значение</span><span class="sxs-lookup"><span data-stu-id="a74d4-811">Boolean</span></span> | | <span data-ttu-id="a74d4-p152">Используется только в том случае, если свойству `type` задано значение `file`. Значение `true` указывает на то, что вложение будет встроено в текст сообщения и не должно отображаться в списке вложений.</span><span class="sxs-lookup"><span data-stu-id="a74d4-p152">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="a74d4-814">String</span><span class="sxs-lookup"><span data-stu-id="a74d4-814">String</span></span> | | <span data-ttu-id="a74d4-p153">Используется, только если свойству `type` задано значение `item`. Идентификатор вложения EWS. Это строка длиной до 100 символов.</span><span class="sxs-lookup"><span data-stu-id="a74d4-p153">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="a74d4-818">function</span><span class="sxs-lookup"><span data-stu-id="a74d4-818">function</span></span> | <span data-ttu-id="a74d4-819">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="a74d4-819">&lt;optional&gt;</span></span> | <span data-ttu-id="a74d4-820">По завершении работы метода функция, переданная параметру `callback`, вызывается с помощью одного параметра `asyncResult`, представляющего собой объект [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="a74d4-820">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="a74d4-821">Требования</span><span class="sxs-lookup"><span data-stu-id="a74d4-821">Requirements</span></span>

|<span data-ttu-id="a74d4-822">Требование</span><span class="sxs-lookup"><span data-stu-id="a74d4-822">Requirement</span></span>| <span data-ttu-id="a74d4-823">Значение</span><span class="sxs-lookup"><span data-stu-id="a74d4-823">Value</span></span>|
|---|---|
|[<span data-ttu-id="a74d4-824">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="a74d4-824">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a74d4-825">1.0</span><span class="sxs-lookup"><span data-stu-id="a74d4-825">1.0</span></span>|
|[<span data-ttu-id="a74d4-826">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="a74d4-826">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a74d4-827">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a74d4-827">ReadItem</span></span>|
|[<span data-ttu-id="a74d4-828">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="a74d4-828">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a74d4-829">Чтение</span><span class="sxs-lookup"><span data-stu-id="a74d4-829">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="a74d4-830">Примеры</span><span class="sxs-lookup"><span data-stu-id="a74d4-830">Examples</span></span>

<span data-ttu-id="a74d4-831">Приведенный ниже код передает строку в функцию `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="a74d4-831">The following code passes a string to the `displayReplyForm` function.</span></span>

```
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="a74d4-832">Ответ с пустым текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="a74d4-832">Reply with an empty body.</span></span>

```
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="a74d4-833">Ответ только с текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="a74d4-833">Reply with just a body.</span></span>

```
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="a74d4-834">Ответ с текстом сообщения и вложенным файлом.</span><span class="sxs-lookup"><span data-stu-id="a74d4-834">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="a74d4-835">Ответ с текстом сообщения и вложенным элементом.</span><span class="sxs-lookup"><span data-stu-id="a74d4-835">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="a74d4-836">Ответ с текстом сообщения, вложенным файлом, вложенным элементом и обратным вызовом.</span><span class="sxs-lookup"><span data-stu-id="a74d4-836">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlook16officeentities"></a><span data-ttu-id="a74d4-837">getEntities() → {[Entities](/javascript/api/outlook_1_6/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="a74d4-837">getEntities() → {[Entities](/javascript/api/outlook_1_6/office.entities)}</span></span>

<span data-ttu-id="a74d4-838">Получает сущности, обнаруженные в выбранном тексте элемента.</span><span class="sxs-lookup"><span data-stu-id="a74d4-838">Gets the entities found in the selected item.</span></span>

> [!NOTE]
> <span data-ttu-id="a74d4-839">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="a74d4-839">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="a74d4-840">Требования</span><span class="sxs-lookup"><span data-stu-id="a74d4-840">Requirements</span></span>

|<span data-ttu-id="a74d4-841">Требование</span><span class="sxs-lookup"><span data-stu-id="a74d4-841">Requirement</span></span>| <span data-ttu-id="a74d4-842">Значение</span><span class="sxs-lookup"><span data-stu-id="a74d4-842">Value</span></span>|
|---|---|
|[<span data-ttu-id="a74d4-843">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="a74d4-843">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a74d4-844">1.0</span><span class="sxs-lookup"><span data-stu-id="a74d4-844">1.0</span></span>|
|[<span data-ttu-id="a74d4-845">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="a74d4-845">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a74d4-846">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a74d4-846">ReadItem</span></span>|
|[<span data-ttu-id="a74d4-847">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="a74d4-847">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a74d4-848">Чтение</span><span class="sxs-lookup"><span data-stu-id="a74d4-848">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="a74d4-849">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="a74d4-849">Returns:</span></span>

<span data-ttu-id="a74d4-850">Тип: [Entities](/javascript/api/outlook_1_6/office.entities)</span><span class="sxs-lookup"><span data-stu-id="a74d4-850">Type: [Entities](/javascript/api/outlook_1_6/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="a74d4-851">Пример</span><span class="sxs-lookup"><span data-stu-id="a74d4-851">Example</span></span>

<span data-ttu-id="a74d4-852">Ниже приведен пример получения доступа к сущностям контактов в тексте текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="a74d4-852">The following example accesses the contacts entities on the current item.</span></span>

```
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook16officecontactmeetingsuggestionjavascriptapioutlook16officemeetingsuggestionphonenumberjavascriptapioutlook16officephonenumbertasksuggestionjavascriptapioutlook16officetasksuggestion"></a><span data-ttu-id="a74d4-853">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="a74d4-853">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))>}</span></span>

<span data-ttu-id="a74d4-854">Получает массив всех сущностей указанного типа, обнаруженных в тексте выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="a74d4-854">Gets an array of all the entities of the specified entity type found in the selected item.</span></span>

> [!NOTE]
> <span data-ttu-id="a74d4-855">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="a74d4-855">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a74d4-856">Параметры:</span><span class="sxs-lookup"><span data-stu-id="a74d4-856">Parameters:</span></span>

|<span data-ttu-id="a74d4-857">Имя</span><span class="sxs-lookup"><span data-stu-id="a74d4-857">Name</span></span>| <span data-ttu-id="a74d4-858">Тип</span><span class="sxs-lookup"><span data-stu-id="a74d4-858">Type</span></span>| <span data-ttu-id="a74d4-859">Описание</span><span class="sxs-lookup"><span data-stu-id="a74d4-859">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="a74d4-860">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="a74d4-860">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook_1_6/office.mailboxenums.entitytype)|<span data-ttu-id="a74d4-861">Одно из значений перечисления EntityType.</span><span class="sxs-lookup"><span data-stu-id="a74d4-861">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a74d4-862">Требования</span><span class="sxs-lookup"><span data-stu-id="a74d4-862">Requirements</span></span>

|<span data-ttu-id="a74d4-863">Требование</span><span class="sxs-lookup"><span data-stu-id="a74d4-863">Requirement</span></span>| <span data-ttu-id="a74d4-864">Значение</span><span class="sxs-lookup"><span data-stu-id="a74d4-864">Value</span></span>|
|---|---|
|[<span data-ttu-id="a74d4-865">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="a74d4-865">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a74d4-866">1.0</span><span class="sxs-lookup"><span data-stu-id="a74d4-866">1.0</span></span>|
|[<span data-ttu-id="a74d4-867">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="a74d4-867">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a74d4-868">Ограниченный доступ</span><span class="sxs-lookup"><span data-stu-id="a74d4-868">Restricted</span></span>|
|[<span data-ttu-id="a74d4-869">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="a74d4-869">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a74d4-870">Чтение</span><span class="sxs-lookup"><span data-stu-id="a74d4-870">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="a74d4-871">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="a74d4-871">Returns:</span></span>

<span data-ttu-id="a74d4-872">Если значение, переданное в `entityType`, не является допустимым членом перечисления `EntityType`, метод возвращает значение NULL.</span><span class="sxs-lookup"><span data-stu-id="a74d4-872">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="a74d4-873">Если в тексте элемента отсутствуют сущности указанного типа, метод возвращает пустой массив.</span><span class="sxs-lookup"><span data-stu-id="a74d4-873">If no entities of the specified type are present on the item, the method returns an empty array.</span></span> <span data-ttu-id="a74d4-874">В противном случае тип объектов в возвращаемом массиве зависит от типа сущности, запрошенной в параметре `entityType`.</span><span class="sxs-lookup"><span data-stu-id="a74d4-874">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="a74d4-875">Хотя минимальный уровень разрешений для использования этого метода — **Restricted**, для некоторых типов сущностей требуется доступ на уровне **ReadItem**, как указано в приведенной ниже таблице.</span><span class="sxs-lookup"><span data-stu-id="a74d4-875">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="a74d4-876">Значение параметра `entityType`</span><span class="sxs-lookup"><span data-stu-id="a74d4-876">Value of `entityType`</span></span> | <span data-ttu-id="a74d4-877">Тип объектов в возвращаемом массиве</span><span class="sxs-lookup"><span data-stu-id="a74d4-877">Type of objects in returned array</span></span> | <span data-ttu-id="a74d4-878">Необходимый уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="a74d4-878">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="a74d4-879">String</span><span class="sxs-lookup"><span data-stu-id="a74d4-879">String</span></span> | <span data-ttu-id="a74d4-880">**С ограничениями**</span><span class="sxs-lookup"><span data-stu-id="a74d4-880">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="a74d4-881">Contact</span><span class="sxs-lookup"><span data-stu-id="a74d4-881">Contact</span></span> | <span data-ttu-id="a74d4-882">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="a74d4-882">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="a74d4-883">String</span><span class="sxs-lookup"><span data-stu-id="a74d4-883">String</span></span> | <span data-ttu-id="a74d4-884">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="a74d4-884">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="a74d4-885">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="a74d4-885">MeetingSuggestion</span></span> | <span data-ttu-id="a74d4-886">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="a74d4-886">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="a74d4-887">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="a74d4-887">PhoneNumber</span></span> | <span data-ttu-id="a74d4-888">**С ограничениями**</span><span class="sxs-lookup"><span data-stu-id="a74d4-888">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="a74d4-889">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="a74d4-889">TaskSuggestion</span></span> | <span data-ttu-id="a74d4-890">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="a74d4-890">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="a74d4-891">String</span><span class="sxs-lookup"><span data-stu-id="a74d4-891">String</span></span> | <span data-ttu-id="a74d4-892">**С ограничениями**</span><span class="sxs-lookup"><span data-stu-id="a74d4-892">**Restricted**</span></span> |

<span data-ttu-id="a74d4-893">Type: Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="a74d4-893">Type: Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="a74d4-894">Пример</span><span class="sxs-lookup"><span data-stu-id="a74d4-894">Example</span></span>

<span data-ttu-id="a74d4-895">В примере ниже показано, как получить доступ к массиву строк, которые представляют собой почтовые адреса в тексте текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="a74d4-895">The following example shows how to access an array of strings that represent postal addresses in the subject or body of the current item.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook16officecontactmeetingsuggestionjavascriptapioutlook16officemeetingsuggestionphonenumberjavascriptapioutlook16officephonenumbertasksuggestionjavascriptapioutlook16officetasksuggestion"></a><span data-ttu-id="a74d4-896">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="a74d4-896">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))>}</span></span>

<span data-ttu-id="a74d4-897">Возвращает известные сущности в выбранном элементе, которые проходят через именованный фильтр, определяемый в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="a74d4-897">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="a74d4-898">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="a74d4-898">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="a74d4-899">Метод `getFilteredEntitiesByName` возвращает сущности, соответствующие регулярному выражению, которое определяется в элементе правила [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) в XML-файле манифеста, с использованием указанного значения элемента `FilterName`.</span><span class="sxs-lookup"><span data-stu-id="a74d4-899">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a74d4-900">Параметры:</span><span class="sxs-lookup"><span data-stu-id="a74d4-900">Parameters:</span></span>

|<span data-ttu-id="a74d4-901">Имя</span><span class="sxs-lookup"><span data-stu-id="a74d4-901">Name</span></span>| <span data-ttu-id="a74d4-902">Тип</span><span class="sxs-lookup"><span data-stu-id="a74d4-902">Type</span></span>| <span data-ttu-id="a74d4-903">Описание</span><span class="sxs-lookup"><span data-stu-id="a74d4-903">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="a74d4-904">String</span><span class="sxs-lookup"><span data-stu-id="a74d4-904">String</span></span>|<span data-ttu-id="a74d4-905">Имя элемента правила `ItemHasKnownEntity`, определяющее соответствующий фильтр.</span><span class="sxs-lookup"><span data-stu-id="a74d4-905">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a74d4-906">Требования</span><span class="sxs-lookup"><span data-stu-id="a74d4-906">Requirements</span></span>

|<span data-ttu-id="a74d4-907">Требование</span><span class="sxs-lookup"><span data-stu-id="a74d4-907">Requirement</span></span>| <span data-ttu-id="a74d4-908">Значение</span><span class="sxs-lookup"><span data-stu-id="a74d4-908">Value</span></span>|
|---|---|
|[<span data-ttu-id="a74d4-909">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="a74d4-909">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a74d4-910">1.0</span><span class="sxs-lookup"><span data-stu-id="a74d4-910">1.0</span></span>|
|[<span data-ttu-id="a74d4-911">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="a74d4-911">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a74d4-912">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a74d4-912">ReadItem</span></span>|
|[<span data-ttu-id="a74d4-913">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="a74d4-913">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a74d4-914">Чтение</span><span class="sxs-lookup"><span data-stu-id="a74d4-914">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="a74d4-915">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="a74d4-915">Returns:</span></span>

<span data-ttu-id="a74d4-p155">Если в манифесте нет элемента `ItemHasKnownEntity` со значением `FilterName`, соответствующим параметру `name`, метод возвращает `null`. Если параметр `name` соответствует элементу `ItemHasKnownEntity` в манифесте, но при этом в текущем элементе нет соответствующих сущностей, метод возвращает пустой массив.</span><span class="sxs-lookup"><span data-stu-id="a74d4-p155">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="a74d4-918">Тип: Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="a74d4-918">Type: Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))></span></span>

#### <a name="getregexmatches--object"></a><span data-ttu-id="a74d4-919">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="a74d4-919">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="a74d4-920">Возвращает строчные значения в выбранном элементе, которые соответствуют регулярным выражениям, определенным в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="a74d4-920">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="a74d4-921">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="a74d4-921">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="a74d4-p156">Метод `getRegExMatches` возвращает строки, соответствующие регулярному выражению, которое определяется в каждом элементе правила `ItemHasRegularExpressionMatch` или `ItemHasKnownEntity` в XML-файле манифеста. Для правила `ItemHasRegularExpressionMatch` свойство элемента, указанного этим правилом, должно содержать соответствующую строку. Простой тип `PropertyName` определяет поддерживаемые свойства.</span><span class="sxs-lookup"><span data-stu-id="a74d4-p156">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="a74d4-925">Например, рассмотрим манифест надстройки, который содержит указанный ниже элемент `Rule`.</span><span class="sxs-lookup"><span data-stu-id="a74d4-925">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="a74d4-926">Объект, возвращаемый методом `getRegExMatches`, будет содержать два свойства: `fruits` и `veggies`.</span><span class="sxs-lookup"><span data-stu-id="a74d4-926">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="a74d4-p157">Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты. Лучше используйте для этого метод [`Body.getAsync`](/javascript/api/outlook_1_6/office.body#getasync-coerciontype--options--callback-).</span><span class="sxs-lookup"><span data-stu-id="a74d4-p157">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook_1_6/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="a74d4-930">Требования</span><span class="sxs-lookup"><span data-stu-id="a74d4-930">Requirements</span></span>

|<span data-ttu-id="a74d4-931">Требование</span><span class="sxs-lookup"><span data-stu-id="a74d4-931">Requirement</span></span>| <span data-ttu-id="a74d4-932">Значение</span><span class="sxs-lookup"><span data-stu-id="a74d4-932">Value</span></span>|
|---|---|
|[<span data-ttu-id="a74d4-933">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="a74d4-933">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a74d4-934">1.0</span><span class="sxs-lookup"><span data-stu-id="a74d4-934">1.0</span></span>|
|[<span data-ttu-id="a74d4-935">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="a74d4-935">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a74d4-936">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a74d4-936">ReadItem</span></span>|
|[<span data-ttu-id="a74d4-937">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="a74d4-937">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a74d4-938">Чтение</span><span class="sxs-lookup"><span data-stu-id="a74d4-938">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="a74d4-939">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="a74d4-939">Returns:</span></span>

<span data-ttu-id="a74d4-p158">Объект, содержащий массив строк, которые соответствуют регулярным выражениям, определяемым в XML-файле манифеста. Имя каждого массива равно соответствующему значению атрибута `RegExName` правила сопоставления `ItemHasRegularExpressionMatch` или атрибута `FilterName` правила сопоставления `ItemHasKnownEntity`.</span><span class="sxs-lookup"><span data-stu-id="a74d4-p158">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="a74d4-942">

<dt>Тип</dt>

</span><span class="sxs-lookup"><span data-stu-id="a74d4-942">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="a74d4-943">Object</span><span class="sxs-lookup"><span data-stu-id="a74d4-943">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="a74d4-944">Пример</span><span class="sxs-lookup"><span data-stu-id="a74d4-944">Example</span></span>

<span data-ttu-id="a74d4-945">В приведенном ниже примере показано, как получить доступ к массиву совпадений с элементами `fruits` и `veggies` правил активации регулярных выражений, указанными в манифесте.</span><span class="sxs-lookup"><span data-stu-id="a74d4-945">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veges = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="a74d4-946">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="a74d4-946">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="a74d4-947">Возвращает строчные значения в выбранном элементе, которые соответствуют именованному регулярному выражению, определенному в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="a74d4-947">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="a74d4-948">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="a74d4-948">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="a74d4-949">Метод `getRegExMatchesByName` возвращает строки, соответствующие регулярному выражению, которое определяется в элементе правила `ItemHasRegularExpressionMatch` в XML-файле манифеста, с использованием указанного значения элемента `RegExName`.</span><span class="sxs-lookup"><span data-stu-id="a74d4-949">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="a74d4-p159">Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты.</span><span class="sxs-lookup"><span data-stu-id="a74d4-p159">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a74d4-952">Параметры:</span><span class="sxs-lookup"><span data-stu-id="a74d4-952">Parameters:</span></span>

|<span data-ttu-id="a74d4-953">Имя</span><span class="sxs-lookup"><span data-stu-id="a74d4-953">Name</span></span>| <span data-ttu-id="a74d4-954">Тип</span><span class="sxs-lookup"><span data-stu-id="a74d4-954">Type</span></span>| <span data-ttu-id="a74d4-955">Описание</span><span class="sxs-lookup"><span data-stu-id="a74d4-955">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="a74d4-956">String</span><span class="sxs-lookup"><span data-stu-id="a74d4-956">String</span></span>|<span data-ttu-id="a74d4-957">Имя элемента правила `ItemHasRegularExpressionMatch`, определяющее соответствующий фильтр.</span><span class="sxs-lookup"><span data-stu-id="a74d4-957">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a74d4-958">Требования</span><span class="sxs-lookup"><span data-stu-id="a74d4-958">Requirements</span></span>

|<span data-ttu-id="a74d4-959">Требование</span><span class="sxs-lookup"><span data-stu-id="a74d4-959">Requirement</span></span>| <span data-ttu-id="a74d4-960">Значение</span><span class="sxs-lookup"><span data-stu-id="a74d4-960">Value</span></span>|
|---|---|
|[<span data-ttu-id="a74d4-961">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="a74d4-961">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a74d4-962">1.0</span><span class="sxs-lookup"><span data-stu-id="a74d4-962">1.0</span></span>|
|[<span data-ttu-id="a74d4-963">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="a74d4-963">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a74d4-964">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a74d4-964">ReadItem</span></span>|
|[<span data-ttu-id="a74d4-965">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="a74d4-965">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a74d4-966">Чтение</span><span class="sxs-lookup"><span data-stu-id="a74d4-966">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="a74d4-967">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="a74d4-967">Returns:</span></span>

<span data-ttu-id="a74d4-968">Массив строк, соответствующих регулярному выражению, определяемому в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="a74d4-968">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="a74d4-969">

<dt>Тип</dt>

</span><span class="sxs-lookup"><span data-stu-id="a74d4-969">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="a74d4-970">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="a74d4-970">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="a74d4-971">Пример</span><span class="sxs-lookup"><span data-stu-id="a74d4-971">Example</span></span>

```
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="a74d4-972">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="a74d4-972">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="a74d4-973">Асинхронно возвращает данные, выбранные в теме или тексте сообщения.</span><span class="sxs-lookup"><span data-stu-id="a74d4-973">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="a74d4-p160">Если выделенный фрагмент отсутствует, но курсор находится в тексте или теме, метод возвращает значение NULL для выбранных данных. Если выбраны не текст и не тема, метод возвращает ошибку `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="a74d4-p160">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a74d4-976">Параметры:</span><span class="sxs-lookup"><span data-stu-id="a74d4-976">Parameters:</span></span>

|<span data-ttu-id="a74d4-977">Имя</span><span class="sxs-lookup"><span data-stu-id="a74d4-977">Name</span></span>| <span data-ttu-id="a74d4-978">Тип</span><span class="sxs-lookup"><span data-stu-id="a74d4-978">Type</span></span>| <span data-ttu-id="a74d4-979">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="a74d4-979">Attributes</span></span>| <span data-ttu-id="a74d4-980">Описание</span><span class="sxs-lookup"><span data-stu-id="a74d4-980">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="a74d4-981">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="a74d4-981">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="a74d4-p161">Запрашивает формат данных. Если задано значение Text, метод возвращает обычный текст как строку, удаляя все имеющиеся HTML-теги. Если задано значение HTML, метод возвращает выделенный текст (обычный текст или HTML).</span><span class="sxs-lookup"><span data-stu-id="a74d4-p161">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="a74d4-985">Объект</span><span class="sxs-lookup"><span data-stu-id="a74d4-985">Object</span></span>| <span data-ttu-id="a74d4-986">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="a74d4-986">&lt;optional&gt;</span></span>|<span data-ttu-id="a74d4-987">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="a74d4-987">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="a74d4-988">Объект</span><span class="sxs-lookup"><span data-stu-id="a74d4-988">Object</span></span>| <span data-ttu-id="a74d4-989">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="a74d4-989">&lt;optional&gt;</span></span>|<span data-ttu-id="a74d4-990">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="a74d4-990">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="a74d4-991">функция</span><span class="sxs-lookup"><span data-stu-id="a74d4-991">function</span></span>||<span data-ttu-id="a74d4-992">Когда метод завершает выполнение, переданная в параметре `callback` функция вызывается с единственным параметром `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="a74d4-992">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="a74d4-993">Чтобы получить доступ к выбранным данным из метода обратного вызова, вызовите `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="a74d4-993">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="a74d4-994">Для доступа к исходному свойству, на основе которого созданы выбранные данные, вызовите  `asyncResult.value.sourceProperty`, который может иметь значение `body` или `subject`.</span><span class="sxs-lookup"><span data-stu-id="a74d4-994">To access the source property that the selection comes from, call , which will be either  or .|</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a74d4-995">Требования</span><span class="sxs-lookup"><span data-stu-id="a74d4-995">Requirements</span></span>

|<span data-ttu-id="a74d4-996">Требование</span><span class="sxs-lookup"><span data-stu-id="a74d4-996">Requirement</span></span>| <span data-ttu-id="a74d4-997">Значение</span><span class="sxs-lookup"><span data-stu-id="a74d4-997">Value</span></span>|
|---|---|
|[<span data-ttu-id="a74d4-998">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="a74d4-998">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a74d4-999">1.2</span><span class="sxs-lookup"><span data-stu-id="a74d4-999">1.2</span></span>|
|[<span data-ttu-id="a74d4-1000">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="a74d4-1000">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a74d4-1001">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="a74d4-1001">ReadWriteItem</span></span>|
|[<span data-ttu-id="a74d4-1002">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="a74d4-1002">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a74d4-1003">Создание</span><span class="sxs-lookup"><span data-stu-id="a74d4-1003">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="a74d4-1004">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="a74d4-1004">Returns:</span></span>

<span data-ttu-id="a74d4-1005">Выбранные данные в виде строки с форматом, определенным в параметре `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="a74d4-1005">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="a74d4-1006">

<dt>Тип</dt>

</span><span class="sxs-lookup"><span data-stu-id="a74d4-1006">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="a74d4-1007">String</span><span class="sxs-lookup"><span data-stu-id="a74d4-1007">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="a74d4-1008">Пример</span><span class="sxs-lookup"><span data-stu-id="a74d4-1008">Example</span></span>

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

#### <a name="getselectedentities--entitiesjavascriptapioutlook16officeentities"></a><span data-ttu-id="a74d4-1009">getSelectedEntities() → {[Entities](/javascript/api/outlook_1_6/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="a74d4-1009">getSelectedEntities() → {[Entities](/javascript/api/outlook_1_6/office.entities)}</span></span>

<span data-ttu-id="a74d4-p163">Возвращает сущности, найденные в выделенном совпадении, выбранном пользователем. Выделенные совпадения применяются к [контекстным надстройкам](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="a74d4-p163">Gets the entities found in a highlighted match a user has selected. Highlighted matches apply to [contextual add-ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="a74d4-1012">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="a74d4-1012">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="a74d4-1013">Требования</span><span class="sxs-lookup"><span data-stu-id="a74d4-1013">Requirements</span></span>

|<span data-ttu-id="a74d4-1014">Требование</span><span class="sxs-lookup"><span data-stu-id="a74d4-1014">Requirement</span></span>| <span data-ttu-id="a74d4-1015">Значение</span><span class="sxs-lookup"><span data-stu-id="a74d4-1015">Value</span></span>|
|---|---|
|[<span data-ttu-id="a74d4-1016">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="a74d4-1016">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a74d4-1017">1.6</span><span class="sxs-lookup"><span data-stu-id="a74d4-1017">1.6</span></span> |
|[<span data-ttu-id="a74d4-1018">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="a74d4-1018">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a74d4-1019">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a74d4-1019">ReadItem</span></span>|
|[<span data-ttu-id="a74d4-1020">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="a74d4-1020">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a74d4-1021">Чтение</span><span class="sxs-lookup"><span data-stu-id="a74d4-1021">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="a74d4-1022">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="a74d4-1022">Returns:</span></span>

<span data-ttu-id="a74d4-1023">Тип: [Entities](/javascript/api/outlook_1_6/office.entities)</span><span class="sxs-lookup"><span data-stu-id="a74d4-1023">Type: [Entities](/javascript/api/outlook_1_6/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="a74d4-1024">Пример</span><span class="sxs-lookup"><span data-stu-id="a74d4-1024">Example</span></span>

<span data-ttu-id="a74d4-1025">В приведенном ниже примере показано, как получить доступ к сущностям адресов в выделенном совпадении, выбранном пользователем.</span><span class="sxs-lookup"><span data-stu-id="a74d4-1025">The following example accesses the addresses entities in the highlighted match selected by the user.</span></span>

```
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

#### <a name="getselectedregexmatches--object"></a><span data-ttu-id="a74d4-1026">getSelectedRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="a74d4-1026">getSelectedRegExMatches() → {Object}</span></span>

<span data-ttu-id="a74d4-p164">Возвращает строковые значения в выделенном совпадении, которые соответствуют регулярным выражениям, определенным в XML-файле манифеста. Выделенные совпадения применяются к [контекстным надстройкам](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="a74d4-p164">Returns string values in a highlighted match that match the regular expressions defined in the manifest XML file. Highlighted matches apply to [contextual add-ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="a74d4-1029">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="a74d4-1029">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="a74d4-p165">Метод `getSelectedRegExMatches` возвращает строки, соответствующие регулярному выражению, которое определяется в каждом элементе правила `ItemHasRegularExpressionMatch` или `ItemHasKnownEntity` в XML-файле манифеста. Для правила `ItemHasRegularExpressionMatch` свойство элемента, указанного этим правилом, должно содержать соответствующую строку. Простой тип `PropertyName` определяет поддерживаемые свойства.</span><span class="sxs-lookup"><span data-stu-id="a74d4-p165">The `getSelectedRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="a74d4-1033">Например, рассмотрим манифест надстройки, который содержит указанный ниже элемент `Rule`.</span><span class="sxs-lookup"><span data-stu-id="a74d4-1033">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="a74d4-1034">Объект, возвращаемый методом `getRegExMatches`, будет содержать два свойства: `fruits` и `veggies`.</span><span class="sxs-lookup"><span data-stu-id="a74d4-1034">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="a74d4-p166">Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты. Лучше используйте для этого метод [`Body.getAsync`](/javascript/api/outlook_1_6/office.body#getasync-coerciontype--options--callback-).</span><span class="sxs-lookup"><span data-stu-id="a74d4-p166">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook_1_6/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="a74d4-1038">Требования</span><span class="sxs-lookup"><span data-stu-id="a74d4-1038">Requirements</span></span>

|<span data-ttu-id="a74d4-1039">Требование</span><span class="sxs-lookup"><span data-stu-id="a74d4-1039">Requirement</span></span>| <span data-ttu-id="a74d4-1040">Значение</span><span class="sxs-lookup"><span data-stu-id="a74d4-1040">Value</span></span>|
|---|---|
|[<span data-ttu-id="a74d4-1041">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="a74d4-1041">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a74d4-1042">1.6</span><span class="sxs-lookup"><span data-stu-id="a74d4-1042">1.6</span></span> |
|[<span data-ttu-id="a74d4-1043">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="a74d4-1043">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a74d4-1044">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a74d4-1044">ReadItem</span></span>|
|[<span data-ttu-id="a74d4-1045">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="a74d4-1045">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a74d4-1046">Чтение</span><span class="sxs-lookup"><span data-stu-id="a74d4-1046">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="a74d4-1047">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="a74d4-1047">Returns:</span></span>

<span data-ttu-id="a74d4-p167">Объект, содержащий массив строк, которые соответствуют регулярным выражениям, определяемым в XML-файле манифеста. Имя каждого массива равно соответствующему значению атрибута `RegExName` правила сопоставления `ItemHasRegularExpressionMatch` или атрибута `FilterName` правила сопоставления `ItemHasKnownEntity`.</span><span class="sxs-lookup"><span data-stu-id="a74d4-p167">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

##### <a name="example"></a><span data-ttu-id="a74d4-1050">Пример</span><span class="sxs-lookup"><span data-stu-id="a74d4-1050">Example</span></span>

<span data-ttu-id="a74d4-1051">В приведенном ниже примере показано, как получить доступ к массиву совпадений с элементами `fruits` и `veggies` правил активации регулярных выражений, указанными в манифесте.</span><span class="sxs-lookup"><span data-stu-id="a74d4-1051">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="a74d4-1052">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="a74d4-1052">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="a74d4-1053">Асинхронно загружает настраиваемые свойства для надстройки выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="a74d4-1053">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="a74d4-p168">Настраиваемые свойства сохраняются в виде пар "ключ-значение" для каждого приложения и каждого элемента. Этот метод возвращает объект `CustomProperties` при обратном вызове, который предоставляет методы для доступа к настраиваемым свойствам, характерным для текущего элемента и текущей надстройки. Настраиваемые свойства не шифруются для элемента, поэтому этот способ хранения не является безопасным.</span><span class="sxs-lookup"><span data-stu-id="a74d4-p168">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a74d4-1057">Параметры:</span><span class="sxs-lookup"><span data-stu-id="a74d4-1057">Parameters:</span></span>

|<span data-ttu-id="a74d4-1058">Имя</span><span class="sxs-lookup"><span data-stu-id="a74d4-1058">Name</span></span>| <span data-ttu-id="a74d4-1059">Тип</span><span class="sxs-lookup"><span data-stu-id="a74d4-1059">Type</span></span>| <span data-ttu-id="a74d4-1060">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="a74d4-1060">Attributes</span></span>| <span data-ttu-id="a74d4-1061">Описание</span><span class="sxs-lookup"><span data-stu-id="a74d4-1061">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="a74d4-1062">function</span><span class="sxs-lookup"><span data-stu-id="a74d4-1062">function</span></span>||<span data-ttu-id="a74d4-1063">Когда метод завершает выполнение, переданная в параметре `callback` функция вызывается с единственным параметром `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="a74d4-1063">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="a74d4-1064">Настраиваемые свойства предоставляются в виде объекта [`CustomProperties`](/javascript/api/outlook_1_6/office.customproperties) в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="a74d4-1064">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook_1_6/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="a74d4-1065">Этот объект можно использовать для получения, задания и удаления настраиваемых свойств из элемента и сохранения изменений настраиваемого свойства на сервере.</span><span class="sxs-lookup"><span data-stu-id="a74d4-1065">The custom properties are provided as a CustomProperties object in the asyncResult.value property. This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="a74d4-1066">Объект</span><span class="sxs-lookup"><span data-stu-id="a74d4-1066">Object</span></span>| <span data-ttu-id="a74d4-1067">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="a74d4-1067">&lt;optional&gt;</span></span>|<span data-ttu-id="a74d4-1068">Разработчики могут указать любой объект, к которому необходимо получить доступ, в функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="a74d4-1068">Developers can provide any object they wish to access in the callback method.</span></span> <span data-ttu-id="a74d4-1069">Доступ к этому объекту можно получить с помощью свойства `asyncResult.asyncContext` в функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="a74d4-1069">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.|</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a74d4-1070">Требования</span><span class="sxs-lookup"><span data-stu-id="a74d4-1070">Requirements</span></span>

|<span data-ttu-id="a74d4-1071">Требование</span><span class="sxs-lookup"><span data-stu-id="a74d4-1071">Requirement</span></span>| <span data-ttu-id="a74d4-1072">Значение</span><span class="sxs-lookup"><span data-stu-id="a74d4-1072">Value</span></span>|
|---|---|
|[<span data-ttu-id="a74d4-1073">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="a74d4-1073">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a74d4-1074">1.0</span><span class="sxs-lookup"><span data-stu-id="a74d4-1074">1.0</span></span>|
|[<span data-ttu-id="a74d4-1075">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="a74d4-1075">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a74d4-1076">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a74d4-1076">ReadItem</span></span>|
|[<span data-ttu-id="a74d4-1077">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="a74d4-1077">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a74d4-1078">Cоздание или чтение</span><span class="sxs-lookup"><span data-stu-id="a74d4-1078">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="a74d4-1079">Пример</span><span class="sxs-lookup"><span data-stu-id="a74d4-1079">Example</span></span>

<span data-ttu-id="a74d4-p171">В приведенном ниже примере кода показано, как асинхронно загружать настраиваемые свойства, характерные для текущего элемента, с помощью метода `loadCustomPropertiesAsync`. Этот пример также показывает, как сохранять эти свойства на сервере с помощью метода `CustomProperties.saveAsync`. В этом примере кода, после того как выполнена загрузка настраиваемых свойств, метод `CustomProperties.get` используется для считывания настраиваемого свойства `myProp`, метод `CustomProperties.set` — для записи настраиваемого свойства `otherProp`, а метод `saveAsync` — для сохранения настраиваемых свойств.</span><span class="sxs-lookup"><span data-stu-id="a74d4-p171">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="a74d4-1083">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="a74d4-1083">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="a74d4-1084">Удаляет вложение из сообщения или встречи.</span><span class="sxs-lookup"><span data-stu-id="a74d4-1084">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="a74d4-p172">Метод `removeAttachmentAsync` удаляет из элемента вложение с указанным идентификатором. Идентификатор вложения рекомендуется использовать для удаления вложения, только если оно добавлено тем же почтовым приложением в ходе текущего сеанса. В веб-приложении Outlook и веб-приложении Outlook для устройств идентификатор вложения действителен только в рамках одного сеанса. Сеанс завершается, когда пользователь закрывает приложение или начинает создавать элемент во встроенной форме, а затем переходит из формы в отдельное окно.</span><span class="sxs-lookup"><span data-stu-id="a74d4-p172">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a74d4-1089">Параметры:</span><span class="sxs-lookup"><span data-stu-id="a74d4-1089">Parameters:</span></span>

|<span data-ttu-id="a74d4-1090">Имя</span><span class="sxs-lookup"><span data-stu-id="a74d4-1090">Name</span></span>| <span data-ttu-id="a74d4-1091">Тип</span><span class="sxs-lookup"><span data-stu-id="a74d4-1091">Type</span></span>| <span data-ttu-id="a74d4-1092">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="a74d4-1092">Attributes</span></span>| <span data-ttu-id="a74d4-1093">Описание</span><span class="sxs-lookup"><span data-stu-id="a74d4-1093">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="a74d4-1094">String</span><span class="sxs-lookup"><span data-stu-id="a74d4-1094">String</span></span>||<span data-ttu-id="a74d4-p173">Идентификатор удаляемого вложения. Максимальная длина строки — 100 символов.</span><span class="sxs-lookup"><span data-stu-id="a74d4-p173">The identifier of the attachment to remove. The maximum length of the string is 100 characters.</span></span>|
|`options`| <span data-ttu-id="a74d4-1097">Object</span><span class="sxs-lookup"><span data-stu-id="a74d4-1097">Object</span></span>| <span data-ttu-id="a74d4-1098">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="a74d4-1098">&lt;optional&gt;</span></span>|<span data-ttu-id="a74d4-1099">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="a74d4-1099">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="a74d4-1100">Объект</span><span class="sxs-lookup"><span data-stu-id="a74d4-1100">Object</span></span>| <span data-ttu-id="a74d4-1101">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="a74d4-1101">&lt;optional&gt;</span></span>|<span data-ttu-id="a74d4-1102">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="a74d4-1102">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="a74d4-1103">function</span><span class="sxs-lookup"><span data-stu-id="a74d4-1103">function</span></span>| <span data-ttu-id="a74d4-1104">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="a74d4-1104">&lt;optional&gt;</span></span>|<span data-ttu-id="a74d4-1105">Когда метод завершает выполнение, переданная в параметре `callback` функция вызывается с единственным параметром `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="a74d4-1105">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="a74d4-1106">Если удалить вложение не удается, свойство `asyncResult.error` содержит код ошибки с указанием ее причины.</span><span class="sxs-lookup"><span data-stu-id="a74d4-1106">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="a74d4-1107">Ошибки</span><span class="sxs-lookup"><span data-stu-id="a74d4-1107">Errors</span></span>

| <span data-ttu-id="a74d4-1108">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="a74d4-1108">Error code</span></span> | <span data-ttu-id="a74d4-1109">Описание</span><span class="sxs-lookup"><span data-stu-id="a74d4-1109">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="a74d4-1110">Идентификатор вложения не существует.</span><span class="sxs-lookup"><span data-stu-id="a74d4-1110">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="a74d4-1111">Требования</span><span class="sxs-lookup"><span data-stu-id="a74d4-1111">Requirements</span></span>

|<span data-ttu-id="a74d4-1112">Требование</span><span class="sxs-lookup"><span data-stu-id="a74d4-1112">Requirement</span></span>| <span data-ttu-id="a74d4-1113">Значение</span><span class="sxs-lookup"><span data-stu-id="a74d4-1113">Value</span></span>|
|---|---|
|[<span data-ttu-id="a74d4-1114">Версия минимального набора требований для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="a74d4-1114">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a74d4-1115">1.1</span><span class="sxs-lookup"><span data-stu-id="a74d4-1115">1.1</span></span>|
|[<span data-ttu-id="a74d4-1116">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="a74d4-1116">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a74d4-1117">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="a74d4-1117">ReadWriteItem</span></span>|
|[<span data-ttu-id="a74d4-1118">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="a74d4-1118">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a74d4-1119">Создание</span><span class="sxs-lookup"><span data-stu-id="a74d4-1119">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="a74d4-1120">Пример</span><span class="sxs-lookup"><span data-stu-id="a74d4-1120">Example</span></span>

<span data-ttu-id="a74d4-1121">Указанный ниже код удаляет вложение с идентификатором "0".</span><span class="sxs-lookup"><span data-stu-id="a74d4-1121">The following code removes an attachment with an identifier of '0'.</span></span>

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

####  <a name="saveasyncoptions-callback"></a><span data-ttu-id="a74d4-1122">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="a74d4-1122">saveAsync([options], callback)</span></span>

<span data-ttu-id="a74d4-1123">Асинхронно сохраняет элемент.</span><span class="sxs-lookup"><span data-stu-id="a74d4-1123">Asynchronously saves an item.</span></span>

<span data-ttu-id="a74d4-p174">При вызове этот метод сохраняет текущее сообщение в виде черновика и возвращает идентификатор элемента с помощью метода обратного вызова. В веб-приложернии Outlook или интерактивном режиме Outlook этот элемент сохраняется на сервере. В Outlook в режиме кэширования этот элемент сохраняется в локальном кэше.</span><span class="sxs-lookup"><span data-stu-id="a74d4-p174">When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="a74d4-1127">Если в режиме создания надстройка вызывает для элемента метод `saveAsync`, чтобы получить параметр `itemId` для использования с EWS или REST API, необходимо помнить, что синхронизация элемента с сервером может занять много времени, если Outlook работает в режиме кэширования данных.</span><span class="sxs-lookup"><span data-stu-id="a74d4-1127">Note: If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server. Until the item is synced, using the  will return an error.</span></span> <span data-ttu-id="a74d4-1128">До окончания синхронизации применение параметра `itemId` будет приводить к ошибке.</span><span class="sxs-lookup"><span data-stu-id="a74d4-1128">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="a74d4-p176">Так как для встреч не предусмотрено состояние черновика, если `saveAsync` вызывается для встречи в режиме создания, она сохраняется как обычная встреча в календаре пользователя. При сохранении новой встречи приглашения не отправляются. При сохранении существующей встречи уведомления отправляются добавленным или удаленным участникам.</span><span class="sxs-lookup"><span data-stu-id="a74d4-p176">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="a74d4-1132">Следующие клиенты имеют разную реакцию на событие для `saveAsync` для встреч в режиме создания:</span><span class="sxs-lookup"><span data-stu-id="a74d4-1132">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="a74d4-1133">Mac Outlook не поддерживает `saveAsync` на собрании в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="a74d4-1133">Note: Mac Outlook does not support `saveAsync` on a meeting in compose mode. Calling  on a meeting in Mac Outlook will return an error.</span></span> <span data-ttu-id="a74d4-1134">Вызов `saveAsync` на собрании в Mac Outlook возвращает ошибку.</span><span class="sxs-lookup"><span data-stu-id="a74d4-1134">Note: Mac Outlook does not support  on a meeting in compose mode. Calling `saveAsync` on a meeting in Mac Outlook will return an error.</span></span>
> - <span data-ttu-id="a74d4-1135">Outlook в Интернете всегда отправляет приглашение или обновления при вызове `saveAsync` на встрече в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="a74d4-1135">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a74d4-1136">Параметры:</span><span class="sxs-lookup"><span data-stu-id="a74d4-1136">Parameters:</span></span>

|<span data-ttu-id="a74d4-1137">Имя</span><span class="sxs-lookup"><span data-stu-id="a74d4-1137">Name</span></span>| <span data-ttu-id="a74d4-1138">Тип</span><span class="sxs-lookup"><span data-stu-id="a74d4-1138">Type</span></span>| <span data-ttu-id="a74d4-1139">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="a74d4-1139">Attributes</span></span>| <span data-ttu-id="a74d4-1140">Описание</span><span class="sxs-lookup"><span data-stu-id="a74d4-1140">Description</span></span>|
|---|---|---|---|
|`options`| <span data-ttu-id="a74d4-1141">Oбъект</span><span class="sxs-lookup"><span data-stu-id="a74d4-1141">Object</span></span>| <span data-ttu-id="a74d4-1142">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="a74d4-1142">&lt;optional&gt;</span></span>|<span data-ttu-id="a74d4-1143">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="a74d4-1143">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="a74d4-1144">Объект</span><span class="sxs-lookup"><span data-stu-id="a74d4-1144">Object</span></span>| <span data-ttu-id="a74d4-1145">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="a74d4-1145">&lt;optional&gt;</span></span>|<span data-ttu-id="a74d4-1146">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="a74d4-1146">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="a74d4-1147">функция</span><span class="sxs-lookup"><span data-stu-id="a74d4-1147">function</span></span>||<span data-ttu-id="a74d4-1148">Когда метод завершает выполнение, переданная в параметре `callback` функция вызывается с единственным параметром `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="a74d4-1148">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="a74d4-1149">После успешного выполнения идентификатор элемента будет предоставлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="a74d4-1149">On success, the item identifier is provided in the `asyncResult.value` property.|</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a74d4-1150">Требования</span><span class="sxs-lookup"><span data-stu-id="a74d4-1150">Requirements</span></span>

|<span data-ttu-id="a74d4-1151">Требование</span><span class="sxs-lookup"><span data-stu-id="a74d4-1151">Requirement</span></span>| <span data-ttu-id="a74d4-1152">Значение</span><span class="sxs-lookup"><span data-stu-id="a74d4-1152">Value</span></span>|
|---|---|
|[<span data-ttu-id="a74d4-1153">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="a74d4-1153">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a74d4-1154">1.3</span><span class="sxs-lookup"><span data-stu-id="a74d4-1154">1.3</span></span>|
|[<span data-ttu-id="a74d4-1155">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="a74d4-1155">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a74d4-1156">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="a74d4-1156">ReadWriteItem</span></span>|
|[<span data-ttu-id="a74d4-1157">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="a74d4-1157">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a74d4-1158">Создание</span><span class="sxs-lookup"><span data-stu-id="a74d4-1158">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="a74d4-1159">Примеры</span><span class="sxs-lookup"><span data-stu-id="a74d4-1159">Examples</span></span>

```
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result
  });
```

<span data-ttu-id="a74d4-p178">Ниже приведен пример параметра `result`, переданного функции обратного вызова. Свойство `value` содержит идентификатор элемента.</span><span class="sxs-lookup"><span data-stu-id="a74d4-p178">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

####  <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="a74d4-1162">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="a74d4-1162">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="a74d4-1163">Асинхронно вставляет данные в текст или тему сообщения.</span><span class="sxs-lookup"><span data-stu-id="a74d4-1163">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="a74d4-p179">Метод `setSelectedDataAsync` вставляет указанную строку в местоположение курсора в теме или тексте элемента либо, если текст выделен в редакторе, он заменяет выделенный текст. Если курсор находится вне текста или темы элемента, возвращается ошибка. После вставки курсор помещается в конец вставленного содержимого.</span><span class="sxs-lookup"><span data-stu-id="a74d4-p179">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a74d4-1167">Параметры:</span><span class="sxs-lookup"><span data-stu-id="a74d4-1167">Parameters:</span></span>

|<span data-ttu-id="a74d4-1168">Имя</span><span class="sxs-lookup"><span data-stu-id="a74d4-1168">Name</span></span>| <span data-ttu-id="a74d4-1169">Тип</span><span class="sxs-lookup"><span data-stu-id="a74d4-1169">Type</span></span>| <span data-ttu-id="a74d4-1170">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="a74d4-1170">Attributes</span></span>| <span data-ttu-id="a74d4-1171">Описание</span><span class="sxs-lookup"><span data-stu-id="a74d4-1171">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="a74d4-1172">String</span><span class="sxs-lookup"><span data-stu-id="a74d4-1172">String</span></span>||<span data-ttu-id="a74d4-p180">Вставляемые данные. Объем данных не должен превышать 1 000 000 символов. Если передано больше 1 000 000 символов, возвращается исключение `ArgumentOutOfRange`.</span><span class="sxs-lookup"><span data-stu-id="a74d4-p180">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="a74d4-1176">Объект</span><span class="sxs-lookup"><span data-stu-id="a74d4-1176">Object</span></span>| <span data-ttu-id="a74d4-1177">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="a74d4-1177">&lt;optional&gt;</span></span>|<span data-ttu-id="a74d4-1178">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="a74d4-1178">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="a74d4-1179">Объект</span><span class="sxs-lookup"><span data-stu-id="a74d4-1179">Object</span></span>| <span data-ttu-id="a74d4-1180">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="a74d4-1180">&lt;optional&gt;</span></span>|<span data-ttu-id="a74d4-1181">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="a74d4-1181">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`| [<span data-ttu-id="a74d4-1182">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="a74d4-1182">Office.CoercionType</span></span>](office.md#coerciontype-string)| <span data-ttu-id="a74d4-1183">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="a74d4-1183">&lt;optional&gt;</span></span>|<span data-ttu-id="a74d4-p181">Если задано значение `text`, текущий стиль применяется в Outlook и веб-приложении Outlook. Если поле представляет собой редактор HTML, вставляются только текстовые данные, даже если они имеют формат HTML.</span><span class="sxs-lookup"><span data-stu-id="a74d4-p181">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="a74d4-p182">Если `html` и поле поддерживают HTML (а тема не поддерживает), в веб-приложении Outlook применяется текущий стиль, а в Outlook — стиль по умолчанию. Если поле является текстовым, возвращается ошибка `InvalidDataFormat`.</span><span class="sxs-lookup"><span data-stu-id="a74d4-p182">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="a74d4-1188">Если тип `coercionType` не установлен, результат зависит от поля: если поле имеет формат HTML, то используется HTML; если поле является текстовым, то используется обычный текст.</span><span class="sxs-lookup"><span data-stu-id="a74d4-1188">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="a74d4-1189">function</span><span class="sxs-lookup"><span data-stu-id="a74d4-1189">function</span></span>||<span data-ttu-id="a74d4-1190">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `callback`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="a74d4-1190">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="a74d4-1191">Требования</span><span class="sxs-lookup"><span data-stu-id="a74d4-1191">Requirements</span></span>

|<span data-ttu-id="a74d4-1192">Требование</span><span class="sxs-lookup"><span data-stu-id="a74d4-1192">Requirement</span></span>| <span data-ttu-id="a74d4-1193">Значение</span><span class="sxs-lookup"><span data-stu-id="a74d4-1193">Value</span></span>|
|---|---|
|[<span data-ttu-id="a74d4-1194">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="a74d4-1194">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a74d4-1195">1.2</span><span class="sxs-lookup"><span data-stu-id="a74d4-1195">1.2</span></span>|
|[<span data-ttu-id="a74d4-1196">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="a74d4-1196">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a74d4-1197">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="a74d4-1197">ReadWriteItem</span></span>|
|[<span data-ttu-id="a74d4-1198">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="a74d4-1198">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a74d4-1199">Создание</span><span class="sxs-lookup"><span data-stu-id="a74d4-1199">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="a74d4-1200">Пример</span><span class="sxs-lookup"><span data-stu-id="a74d4-1200">Example</span></span>

```
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```