
# <a name="item"></a><span data-ttu-id="0e1d3-101">item</span><span class="sxs-lookup"><span data-stu-id="0e1d3-101">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="0e1d3-102">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="0e1d3-102">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="0e1d3-p101">Пространство имен `item` используется для доступа к выбранному в данный момент сообщению, приглашению на собрание или описанию встречи. Вы можете определить тип пространства имен `item` с помощью свойства [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook15officemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="0e1d3-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook15officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="0e1d3-105">Требования</span><span class="sxs-lookup"><span data-stu-id="0e1d3-105">Requirements</span></span>

|<span data-ttu-id="0e1d3-106">Требование</span><span class="sxs-lookup"><span data-stu-id="0e1d3-106">Requirement</span></span>| <span data-ttu-id="0e1d3-107">Значение</span><span class="sxs-lookup"><span data-stu-id="0e1d3-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="0e1d3-108">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="0e1d3-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0e1d3-109">1.0</span><span class="sxs-lookup"><span data-stu-id="0e1d3-109">1.0</span></span>|
|[<span data-ttu-id="0e1d3-110">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="0e1d3-110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0e1d3-111">Ограниченный доступ</span><span class="sxs-lookup"><span data-stu-id="0e1d3-111">Restricted</span></span>|
|[<span data-ttu-id="0e1d3-112">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="0e1d3-112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="0e1d3-113">Cоздание или чтение</span><span class="sxs-lookup"><span data-stu-id="0e1d3-113">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="0e1d3-114">Члены и методы</span><span class="sxs-lookup"><span data-stu-id="0e1d3-114">Members and methods</span></span>

| <span data-ttu-id="0e1d3-115">Член</span><span class="sxs-lookup"><span data-stu-id="0e1d3-115">Member</span></span> | <span data-ttu-id="0e1d3-116">Тип</span><span class="sxs-lookup"><span data-stu-id="0e1d3-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="0e1d3-117">attachments</span><span class="sxs-lookup"><span data-stu-id="0e1d3-117">attachments</span></span>](#attachments-arrayattachmentdetailsjavascriptapioutlook15officeattachmentdetails) | <span data-ttu-id="0e1d3-118">Член</span><span class="sxs-lookup"><span data-stu-id="0e1d3-118">Member</span></span> |
| [<span data-ttu-id="0e1d3-119">bcc</span><span class="sxs-lookup"><span data-stu-id="0e1d3-119">bcc</span></span>](#bcc-recipientsjavascriptapioutlook15officerecipients) | <span data-ttu-id="0e1d3-120">Член</span><span class="sxs-lookup"><span data-stu-id="0e1d3-120">Member</span></span> |
| [<span data-ttu-id="0e1d3-121">body</span><span class="sxs-lookup"><span data-stu-id="0e1d3-121">body</span></span>](#body-bodyjavascriptapioutlook15officebody) | <span data-ttu-id="0e1d3-122">Член</span><span class="sxs-lookup"><span data-stu-id="0e1d3-122">Member</span></span> |
| [<span data-ttu-id="0e1d3-123">cc</span><span class="sxs-lookup"><span data-stu-id="0e1d3-123">cc</span></span>](#cc-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients) | <span data-ttu-id="0e1d3-124">Член</span><span class="sxs-lookup"><span data-stu-id="0e1d3-124">Member</span></span> |
| [<span data-ttu-id="0e1d3-125">conversationId</span><span class="sxs-lookup"><span data-stu-id="0e1d3-125">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="0e1d3-126">Член</span><span class="sxs-lookup"><span data-stu-id="0e1d3-126">Member</span></span> |
| [<span data-ttu-id="0e1d3-127">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="0e1d3-127">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="0e1d3-128">Член</span><span class="sxs-lookup"><span data-stu-id="0e1d3-128">Member</span></span> |
| [<span data-ttu-id="0e1d3-129">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="0e1d3-129">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="0e1d3-130">Член</span><span class="sxs-lookup"><span data-stu-id="0e1d3-130">Member</span></span> |
| [<span data-ttu-id="0e1d3-131">end</span><span class="sxs-lookup"><span data-stu-id="0e1d3-131">end</span></span>](#end-datetimejavascriptapioutlook15officetime) | <span data-ttu-id="0e1d3-132">Член</span><span class="sxs-lookup"><span data-stu-id="0e1d3-132">Member</span></span> |
| [<span data-ttu-id="0e1d3-133">from</span><span class="sxs-lookup"><span data-stu-id="0e1d3-133">from</span></span>](#from-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails) | <span data-ttu-id="0e1d3-134">Член</span><span class="sxs-lookup"><span data-stu-id="0e1d3-134">Member</span></span> |
| [<span data-ttu-id="0e1d3-135">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="0e1d3-135">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="0e1d3-136">Член</span><span class="sxs-lookup"><span data-stu-id="0e1d3-136">Member</span></span> |
| [<span data-ttu-id="0e1d3-137">itemClass</span><span class="sxs-lookup"><span data-stu-id="0e1d3-137">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="0e1d3-138">Член</span><span class="sxs-lookup"><span data-stu-id="0e1d3-138">Member</span></span> |
| [<span data-ttu-id="0e1d3-139">itemId</span><span class="sxs-lookup"><span data-stu-id="0e1d3-139">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="0e1d3-140">Член</span><span class="sxs-lookup"><span data-stu-id="0e1d3-140">Member</span></span> |
| [<span data-ttu-id="0e1d3-141">itemType</span><span class="sxs-lookup"><span data-stu-id="0e1d3-141">itemType</span></span>](#itemtype-officemailboxenumsitemtypejavascriptapioutlook15officemailboxenumsitemtype) | <span data-ttu-id="0e1d3-142">Член</span><span class="sxs-lookup"><span data-stu-id="0e1d3-142">Member</span></span> |
| [<span data-ttu-id="0e1d3-143">location</span><span class="sxs-lookup"><span data-stu-id="0e1d3-143">location</span></span>](#location-stringlocationjavascriptapioutlook15officelocation) | <span data-ttu-id="0e1d3-144">Член</span><span class="sxs-lookup"><span data-stu-id="0e1d3-144">Member</span></span> |
| [<span data-ttu-id="0e1d3-145">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="0e1d3-145">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="0e1d3-146">Член</span><span class="sxs-lookup"><span data-stu-id="0e1d3-146">Member</span></span> |
| [<span data-ttu-id="0e1d3-147">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="0e1d3-147">notificationMessages</span></span>](#notificationmessages-notificationmessagesjavascriptapioutlook15officenotificationmessages) | <span data-ttu-id="0e1d3-148">Член</span><span class="sxs-lookup"><span data-stu-id="0e1d3-148">Member</span></span> |
| [<span data-ttu-id="0e1d3-149">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="0e1d3-149">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients) | <span data-ttu-id="0e1d3-150">Член</span><span class="sxs-lookup"><span data-stu-id="0e1d3-150">Member</span></span> |
| [<span data-ttu-id="0e1d3-151">organizer</span><span class="sxs-lookup"><span data-stu-id="0e1d3-151">organizer</span></span>](#organizer-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails) | <span data-ttu-id="0e1d3-152">Член</span><span class="sxs-lookup"><span data-stu-id="0e1d3-152">Member</span></span> |
| [<span data-ttu-id="0e1d3-153">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="0e1d3-153">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients) | <span data-ttu-id="0e1d3-154">Член</span><span class="sxs-lookup"><span data-stu-id="0e1d3-154">Member</span></span> |
| [<span data-ttu-id="0e1d3-155">sender</span><span class="sxs-lookup"><span data-stu-id="0e1d3-155">sender</span></span>](#sender-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails) | <span data-ttu-id="0e1d3-156">Член</span><span class="sxs-lookup"><span data-stu-id="0e1d3-156">Member</span></span> |
| [<span data-ttu-id="0e1d3-157">start</span><span class="sxs-lookup"><span data-stu-id="0e1d3-157">start</span></span>](#start-datetimejavascriptapioutlook15officetime) | <span data-ttu-id="0e1d3-158">Член</span><span class="sxs-lookup"><span data-stu-id="0e1d3-158">Member</span></span> |
| [<span data-ttu-id="0e1d3-159">subject</span><span class="sxs-lookup"><span data-stu-id="0e1d3-159">subject</span></span>](#subject-stringsubjectjavascriptapioutlook15officesubject) | <span data-ttu-id="0e1d3-160">Член</span><span class="sxs-lookup"><span data-stu-id="0e1d3-160">Member</span></span> |
| [<span data-ttu-id="0e1d3-161">to</span><span class="sxs-lookup"><span data-stu-id="0e1d3-161">to</span></span>](#to-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients) | <span data-ttu-id="0e1d3-162">Член</span><span class="sxs-lookup"><span data-stu-id="0e1d3-162">Member</span></span> |
| [<span data-ttu-id="0e1d3-163">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="0e1d3-163">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="0e1d3-164">Метод</span><span class="sxs-lookup"><span data-stu-id="0e1d3-164">Method</span></span> |
| [<span data-ttu-id="0e1d3-165">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="0e1d3-165">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="0e1d3-166">Метод</span><span class="sxs-lookup"><span data-stu-id="0e1d3-166">Method</span></span> |
| [<span data-ttu-id="0e1d3-167">close</span><span class="sxs-lookup"><span data-stu-id="0e1d3-167">close</span></span>](#close) | <span data-ttu-id="0e1d3-168">Метод</span><span class="sxs-lookup"><span data-stu-id="0e1d3-168">Method</span></span> |
| [<span data-ttu-id="0e1d3-169">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="0e1d3-169">displayReplyAllForm</span></span>](#displayreplyallformformdata) | <span data-ttu-id="0e1d3-170">Метод</span><span class="sxs-lookup"><span data-stu-id="0e1d3-170">Method</span></span> |
| [<span data-ttu-id="0e1d3-171">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="0e1d3-171">displayReplyForm</span></span>](#displayreplyformformdata) | <span data-ttu-id="0e1d3-172">Метод</span><span class="sxs-lookup"><span data-stu-id="0e1d3-172">Method</span></span> |
| [<span data-ttu-id="0e1d3-173">getEntities</span><span class="sxs-lookup"><span data-stu-id="0e1d3-173">getEntities</span></span>](#getentities--entitiesjavascriptapioutlook15officeentities) | <span data-ttu-id="0e1d3-174">Метод</span><span class="sxs-lookup"><span data-stu-id="0e1d3-174">Method</span></span> |
| [<span data-ttu-id="0e1d3-175">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="0e1d3-175">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook15officecontactmeetingsuggestionjavascriptapioutlook15officemeetingsuggestionphonenumberjavascriptapioutlook15officephonenumbertasksuggestionjavascriptapioutlook15officetasksuggestion) | <span data-ttu-id="0e1d3-176">Метод</span><span class="sxs-lookup"><span data-stu-id="0e1d3-176">Method</span></span> |
| [<span data-ttu-id="0e1d3-177">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="0e1d3-177">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook15officecontactmeetingsuggestionjavascriptapioutlook15officemeetingsuggestionphonenumberjavascriptapioutlook15officephonenumbertasksuggestionjavascriptapioutlook15officetasksuggestion) | <span data-ttu-id="0e1d3-178">Метод</span><span class="sxs-lookup"><span data-stu-id="0e1d3-178">Method</span></span> |
| [<span data-ttu-id="0e1d3-179">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="0e1d3-179">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="0e1d3-180">Метод</span><span class="sxs-lookup"><span data-stu-id="0e1d3-180">Method</span></span> |
| [<span data-ttu-id="0e1d3-181">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="0e1d3-181">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="0e1d3-182">Метод</span><span class="sxs-lookup"><span data-stu-id="0e1d3-182">Method</span></span> |
| [<span data-ttu-id="0e1d3-183">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="0e1d3-183">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="0e1d3-184">Метод</span><span class="sxs-lookup"><span data-stu-id="0e1d3-184">Method</span></span> |
| [<span data-ttu-id="0e1d3-185">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="0e1d3-185">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="0e1d3-186">Метод</span><span class="sxs-lookup"><span data-stu-id="0e1d3-186">Method</span></span> |
| [<span data-ttu-id="0e1d3-187">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="0e1d3-187">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="0e1d3-188">Метод</span><span class="sxs-lookup"><span data-stu-id="0e1d3-188">Method</span></span> |
| [<span data-ttu-id="0e1d3-189">saveAsync</span><span class="sxs-lookup"><span data-stu-id="0e1d3-189">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="0e1d3-190">Метод</span><span class="sxs-lookup"><span data-stu-id="0e1d3-190">Method</span></span> |
| [<span data-ttu-id="0e1d3-191">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="0e1d3-191">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="0e1d3-192">Метод</span><span class="sxs-lookup"><span data-stu-id="0e1d3-192">Method</span></span> |

### <a name="example"></a><span data-ttu-id="0e1d3-193">Пример</span><span class="sxs-lookup"><span data-stu-id="0e1d3-193">Example</span></span>

<span data-ttu-id="0e1d3-194">В приведенном ниже примере кода JavaScript показано, как получить доступ к свойству `subject` текущего элемента в Outlook.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-194">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="0e1d3-195">Члены</span><span class="sxs-lookup"><span data-stu-id="0e1d3-195">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook15officeattachmentdetails"></a><span data-ttu-id="0e1d3-196">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_5/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="0e1d3-196">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_5/office.attachmentdetails)></span></span>

<span data-ttu-id="0e1d3-p102">Получает массив вложений для элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="0e1d3-199">Некоторые типы файлов блокируются Outlook из-за потенциальных проблем безопасности и поэтому не возвращаются.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-199">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="0e1d3-200">Дополнительные сведения см. в статье [Блокированные вложения в Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="0e1d3-200">For more information see [Payments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="0e1d3-201">Тип:</span><span class="sxs-lookup"><span data-stu-id="0e1d3-201">Type:</span></span>

*   <span data-ttu-id="0e1d3-202">Array.<[AttachmentDetails](/javascript/api/outlook_1_5/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="0e1d3-202">Array.<[AttachmentDetails](/javascript/api/outlook_1_5/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="0e1d3-203">Требования</span><span class="sxs-lookup"><span data-stu-id="0e1d3-203">Requirements</span></span>

|<span data-ttu-id="0e1d3-204">Требование</span><span class="sxs-lookup"><span data-stu-id="0e1d3-204">Requirement</span></span>| <span data-ttu-id="0e1d3-205">Значение</span><span class="sxs-lookup"><span data-stu-id="0e1d3-205">Value</span></span>|
|---|---|
|[<span data-ttu-id="0e1d3-206">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="0e1d3-206">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0e1d3-207">1.0</span><span class="sxs-lookup"><span data-stu-id="0e1d3-207">1.0</span></span>|
|[<span data-ttu-id="0e1d3-208">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="0e1d3-208">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0e1d3-209">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0e1d3-209">ReadItem</span></span>|
|[<span data-ttu-id="0e1d3-210">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="0e1d3-210">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="0e1d3-211">Чтение</span><span class="sxs-lookup"><span data-stu-id="0e1d3-211">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0e1d3-212">Пример</span><span class="sxs-lookup"><span data-stu-id="0e1d3-212">Example</span></span>

<span data-ttu-id="0e1d3-213">С помощью приведенного ниже кода можно создать HTML-строку с подробными сведениями обо всех вложениях для текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-213">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

####  <a name="bcc-recipientsjavascriptapioutlook15officerecipients"></a><span data-ttu-id="0e1d3-214">bcc :[Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="0e1d3-214">bcc :[Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

<span data-ttu-id="0e1d3-215">Получает объект, который предоставляет методы для получения или обновления получателей в строке Bcc (скрытой копии) сообщения.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-215">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="0e1d3-216">Только в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-216">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="0e1d3-217">Тип:</span><span class="sxs-lookup"><span data-stu-id="0e1d3-217">Type:</span></span>

*   [<span data-ttu-id="0e1d3-218">Recipients</span><span class="sxs-lookup"><span data-stu-id="0e1d3-218">Recipients</span></span>](/javascript/api/outlook_1_5/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="0e1d3-219">Требования</span><span class="sxs-lookup"><span data-stu-id="0e1d3-219">Requirements</span></span>

|<span data-ttu-id="0e1d3-220">Требование</span><span class="sxs-lookup"><span data-stu-id="0e1d3-220">Requirement</span></span>| <span data-ttu-id="0e1d3-221">Значение</span><span class="sxs-lookup"><span data-stu-id="0e1d3-221">Value</span></span>|
|---|---|
|[<span data-ttu-id="0e1d3-222">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="0e1d3-222">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0e1d3-223">1.1</span><span class="sxs-lookup"><span data-stu-id="0e1d3-223">1.1</span></span>|
|[<span data-ttu-id="0e1d3-224">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="0e1d3-224">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0e1d3-225">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0e1d3-225">ReadItem</span></span>|
|[<span data-ttu-id="0e1d3-226">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="0e1d3-226">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="0e1d3-227">Создание</span><span class="sxs-lookup"><span data-stu-id="0e1d3-227">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="0e1d3-228">Пример</span><span class="sxs-lookup"><span data-stu-id="0e1d3-228">Example</span></span>

```
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlook15officebody"></a><span data-ttu-id="0e1d3-229">body :[Body](/javascript/api/outlook_1_5/office.body)</span><span class="sxs-lookup"><span data-stu-id="0e1d3-229">body :[Body](/javascript/api/outlook_1_5/office.body)</span></span>

<span data-ttu-id="0e1d3-230">Получает объект, предоставляющий методы для работы с основным текстом элемента.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-230">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="0e1d3-231">Тип:</span><span class="sxs-lookup"><span data-stu-id="0e1d3-231">Type:</span></span>

*   [<span data-ttu-id="0e1d3-232">Body</span><span class="sxs-lookup"><span data-stu-id="0e1d3-232">Body</span></span>](/javascript/api/outlook_1_5/office.body)

##### <a name="requirements"></a><span data-ttu-id="0e1d3-233">Требования</span><span class="sxs-lookup"><span data-stu-id="0e1d3-233">Requirements</span></span>

|<span data-ttu-id="0e1d3-234">Требование</span><span class="sxs-lookup"><span data-stu-id="0e1d3-234">Requirement</span></span>| <span data-ttu-id="0e1d3-235">Значение</span><span class="sxs-lookup"><span data-stu-id="0e1d3-235">Value</span></span>|
|---|---|
|[<span data-ttu-id="0e1d3-236">Версия минимального набора требований для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="0e1d3-236">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0e1d3-237">1.1</span><span class="sxs-lookup"><span data-stu-id="0e1d3-237">1.1</span></span>|
|[<span data-ttu-id="0e1d3-238">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="0e1d3-238">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0e1d3-239">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0e1d3-239">ReadItem</span></span>|
|[<span data-ttu-id="0e1d3-240">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="0e1d3-240">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="0e1d3-241">Cоздание или чтение</span><span class="sxs-lookup"><span data-stu-id="0e1d3-241">Compose or read</span></span>|

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients"></a><span data-ttu-id="0e1d3-242">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="0e1d3-242">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

<span data-ttu-id="0e1d3-243">Предоставляет доступ к получателям Cc (копии) сообщения.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-243">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="0e1d3-244">Уровень доступа и тип объекта, зависит от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-244">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="0e1d3-245">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="0e1d3-245">Read mode</span></span>

<span data-ttu-id="0e1d3-p106">Свойство `cc` возвращает массив, который содержит объект `EmailAddressDetails`, каждому получателю, указанному в строке **Cc (копия)** сообщения. Коллекция может включать не более 100 членов.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="0e1d3-248">Режим создания</span><span class="sxs-lookup"><span data-stu-id="0e1d3-248">Compose mode</span></span>

<span data-ttu-id="0e1d3-249">Свойство `cc` возвращает объект `Recipients`, предоставляющий методы для получения или обновления получателей, которые указаны в строке **Cc (копия)** сообщения.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-249">The `cc` property returns a `Recipients` object that provides methods for manipulating the recipients on the **Cc** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="0e1d3-250">Тип:</span><span class="sxs-lookup"><span data-stu-id="0e1d3-250">Type:</span></span>

*   <span data-ttu-id="0e1d3-251">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="0e1d3-251">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="0e1d3-252">Требования</span><span class="sxs-lookup"><span data-stu-id="0e1d3-252">Requirements</span></span>

|<span data-ttu-id="0e1d3-253">Требование</span><span class="sxs-lookup"><span data-stu-id="0e1d3-253">Requirement</span></span>| <span data-ttu-id="0e1d3-254">Значение</span><span class="sxs-lookup"><span data-stu-id="0e1d3-254">Value</span></span>|
|---|---|
|[<span data-ttu-id="0e1d3-255">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="0e1d3-255">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0e1d3-256">1.0</span><span class="sxs-lookup"><span data-stu-id="0e1d3-256">1.0</span></span>|
|[<span data-ttu-id="0e1d3-257">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="0e1d3-257">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0e1d3-258">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0e1d3-258">ReadItem</span></span>|
|[<span data-ttu-id="0e1d3-259">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="0e1d3-259">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="0e1d3-260">Cоздание или чтение</span><span class="sxs-lookup"><span data-stu-id="0e1d3-260">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="0e1d3-261">Пример</span><span class="sxs-lookup"><span data-stu-id="0e1d3-261">Example</span></span>

```
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="0e1d3-262">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="0e1d3-262">(nullable) conversationId :String</span></span>

<span data-ttu-id="0e1d3-263">Получает идентификатор разговора по электронной почте, содержащего конкретное сообщение.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-263">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="0e1d3-p107">Вы можете получить целочисленное значение этого свойства, если ваше почтовое приложение активируется в формах просмотра или формах создания ответов. Если пользователь в свою очередь изменит тему ответа, после его отправки идентификатор беседы будет изменен, и полученное ранее значение будет недействительным.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="0e1d3-p108">Это свойство имеет значение NULL для нового элемента в форме создания. Свойство `conversationId` вернет значение, если пользователь задаст тему и сохранит элемент.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="0e1d3-268">Тип:</span><span class="sxs-lookup"><span data-stu-id="0e1d3-268">Type:</span></span>

*   <span data-ttu-id="0e1d3-269">String</span><span class="sxs-lookup"><span data-stu-id="0e1d3-269">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="0e1d3-270">Требования</span><span class="sxs-lookup"><span data-stu-id="0e1d3-270">Requirements</span></span>

|<span data-ttu-id="0e1d3-271">Требование</span><span class="sxs-lookup"><span data-stu-id="0e1d3-271">Requirement</span></span>| <span data-ttu-id="0e1d3-272">Значение</span><span class="sxs-lookup"><span data-stu-id="0e1d3-272">Value</span></span>|
|---|---|
|[<span data-ttu-id="0e1d3-273">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="0e1d3-273">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0e1d3-274">1.0</span><span class="sxs-lookup"><span data-stu-id="0e1d3-274">1.0</span></span>|
|[<span data-ttu-id="0e1d3-275">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="0e1d3-275">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0e1d3-276">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0e1d3-276">ReadItem</span></span>|
|[<span data-ttu-id="0e1d3-277">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="0e1d3-277">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="0e1d3-278">Cоздание или чтение</span><span class="sxs-lookup"><span data-stu-id="0e1d3-278">Compose or read</span></span>|

#### <a name="datetimecreated-date"></a><span data-ttu-id="0e1d3-279">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="0e1d3-279">dateTimeCreated :Date</span></span>

<span data-ttu-id="0e1d3-p109">Получает дату и время создания элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="0e1d3-282">Тип:</span><span class="sxs-lookup"><span data-stu-id="0e1d3-282">Type:</span></span>

*   <span data-ttu-id="0e1d3-283">Date</span><span class="sxs-lookup"><span data-stu-id="0e1d3-283">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="0e1d3-284">Требования</span><span class="sxs-lookup"><span data-stu-id="0e1d3-284">Requirements</span></span>

|<span data-ttu-id="0e1d3-285">Требование</span><span class="sxs-lookup"><span data-stu-id="0e1d3-285">Requirement</span></span>| <span data-ttu-id="0e1d3-286">Значение</span><span class="sxs-lookup"><span data-stu-id="0e1d3-286">Value</span></span>|
|---|---|
|[<span data-ttu-id="0e1d3-287">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="0e1d3-287">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0e1d3-288">1.0</span><span class="sxs-lookup"><span data-stu-id="0e1d3-288">1.0</span></span>|
|[<span data-ttu-id="0e1d3-289">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="0e1d3-289">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0e1d3-290">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0e1d3-290">ReadItem</span></span>|
|[<span data-ttu-id="0e1d3-291">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="0e1d3-291">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="0e1d3-292">Чтение</span><span class="sxs-lookup"><span data-stu-id="0e1d3-292">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0e1d3-293">Пример</span><span class="sxs-lookup"><span data-stu-id="0e1d3-293">Example</span></span>

```
var created = Office.context.mailbox.item.dateTimeCreated;
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="0e1d3-294">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="0e1d3-294">dateTimeModified :Date</span></span>

<span data-ttu-id="0e1d3-p110">Получает дату и время последнего изменения элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="0e1d3-297">Этот элемент не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-297">Note: This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="0e1d3-298">Тип:</span><span class="sxs-lookup"><span data-stu-id="0e1d3-298">Type:</span></span>

*   <span data-ttu-id="0e1d3-299">Date</span><span class="sxs-lookup"><span data-stu-id="0e1d3-299">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="0e1d3-300">Требования</span><span class="sxs-lookup"><span data-stu-id="0e1d3-300">Requirements</span></span>

|<span data-ttu-id="0e1d3-301">Требование</span><span class="sxs-lookup"><span data-stu-id="0e1d3-301">Requirement</span></span>| <span data-ttu-id="0e1d3-302">Значение</span><span class="sxs-lookup"><span data-stu-id="0e1d3-302">Value</span></span>|
|---|---|
|[<span data-ttu-id="0e1d3-303">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="0e1d3-303">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0e1d3-304">1.0</span><span class="sxs-lookup"><span data-stu-id="0e1d3-304">1.0</span></span>|
|[<span data-ttu-id="0e1d3-305">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="0e1d3-305">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0e1d3-306">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0e1d3-306">ReadItem</span></span>|
|[<span data-ttu-id="0e1d3-307">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="0e1d3-307">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="0e1d3-308">Чтение</span><span class="sxs-lookup"><span data-stu-id="0e1d3-308">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0e1d3-309">Пример</span><span class="sxs-lookup"><span data-stu-id="0e1d3-309">Example</span></span>

```
var modified = Office.context.mailbox.item.dateTimeModified;
```

####  <a name="end-datetimejavascriptapioutlook15officetime"></a><span data-ttu-id="0e1d3-310">end :Date|[Time](/javascript/api/outlook_1_5/office.time)</span><span class="sxs-lookup"><span data-stu-id="0e1d3-310">end :Date|[Time](/javascript/api/outlook_1_5/office.time)</span></span>

<span data-ttu-id="0e1d3-311">Получает или задает дату и время окончания встречи.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-311">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="0e1d3-p111">Свойство `end` представлено в виде значения даты и времени в формате UTC. Преобразовать значение свойства end в местные значения даты и времени клиента можно с помощью метода [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook15officelocalclienttime).</span><span class="sxs-lookup"><span data-stu-id="0e1d3-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook15officelocalclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="0e1d3-314">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="0e1d3-314">Read mode</span></span>

<span data-ttu-id="0e1d3-315">Свойство `end` возвращает объект `Date`.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-315">The `end` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="0e1d3-316">Режим создания</span><span class="sxs-lookup"><span data-stu-id="0e1d3-316">Compose mode</span></span>

<span data-ttu-id="0e1d3-317">Свойство `end` возвращает объект `Time`.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-317">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="0e1d3-318">Когда вы используете метод [`Time.setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) для того, чтобы задать время окончания, вы должны использовать метод [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) , чтобы преобразовать местное время на клиенте в формат UTC.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-318">When you use the [`Time.setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="0e1d3-319">Тип:</span><span class="sxs-lookup"><span data-stu-id="0e1d3-319">Type:</span></span>

*   <span data-ttu-id="0e1d3-320">Date | [Time](/javascript/api/outlook_1_5/office.time)</span><span class="sxs-lookup"><span data-stu-id="0e1d3-320">Date | [Time](/javascript/api/outlook_1_5/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="0e1d3-321">Требования</span><span class="sxs-lookup"><span data-stu-id="0e1d3-321">Requirements</span></span>

|<span data-ttu-id="0e1d3-322">Требование</span><span class="sxs-lookup"><span data-stu-id="0e1d3-322">Requirement</span></span>| <span data-ttu-id="0e1d3-323">Значение</span><span class="sxs-lookup"><span data-stu-id="0e1d3-323">Value</span></span>|
|---|---|
|[<span data-ttu-id="0e1d3-324">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="0e1d3-324">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0e1d3-325">1.0</span><span class="sxs-lookup"><span data-stu-id="0e1d3-325">1.0</span></span>|
|[<span data-ttu-id="0e1d3-326">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="0e1d3-326">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0e1d3-327">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0e1d3-327">ReadItem</span></span>|
|[<span data-ttu-id="0e1d3-328">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="0e1d3-328">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="0e1d3-329">Cоздание или чтение</span><span class="sxs-lookup"><span data-stu-id="0e1d3-329">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="0e1d3-330">Пример</span><span class="sxs-lookup"><span data-stu-id="0e1d3-330">Example</span></span>

<span data-ttu-id="0e1d3-331">В примере ниже показано, как с помощью метода [`setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) объекта `Time` задать время окончания встречи в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-331">The following example sets the end time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

#### <a name="from-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails"></a><span data-ttu-id="0e1d3-332">from :[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="0e1d3-332">from :[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)</span></span>

<span data-ttu-id="0e1d3-p112">Получает электронный адрес отправителя сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-p112">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="0e1d3-p113">Свойства `from` и [`sender`](#sender-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails) представляют одно лицо, если сообщение не отправлено представителем. В противном случае свойство `from` представляет лицо, делегировавшее полномочия, а свойство sender — представителя.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-p113">The `from` and [`sender`](#sender-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="0e1d3-337">Свойство `recipientType` объекта `EmailAddressDetails` в свойстве `from` — `undefined`.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-337">Note: The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="0e1d3-338">Тип:</span><span class="sxs-lookup"><span data-stu-id="0e1d3-338">Type:</span></span>

*   [<span data-ttu-id="0e1d3-339">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="0e1d3-339">EmailAddressDetails</span></span>](/javascript/api/outlook_1_5/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="0e1d3-340">Требования</span><span class="sxs-lookup"><span data-stu-id="0e1d3-340">Requirements</span></span>

|<span data-ttu-id="0e1d3-341">Требование</span><span class="sxs-lookup"><span data-stu-id="0e1d3-341">Requirement</span></span>| <span data-ttu-id="0e1d3-342">Значение</span><span class="sxs-lookup"><span data-stu-id="0e1d3-342">Value</span></span>|
|---|---|
|[<span data-ttu-id="0e1d3-343">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="0e1d3-343">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0e1d3-344">1.0</span><span class="sxs-lookup"><span data-stu-id="0e1d3-344">1.0</span></span>|
|[<span data-ttu-id="0e1d3-345">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="0e1d3-345">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0e1d3-346">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0e1d3-346">ReadItem</span></span>|
|[<span data-ttu-id="0e1d3-347">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="0e1d3-347">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="0e1d3-348">Чтение</span><span class="sxs-lookup"><span data-stu-id="0e1d3-348">Read</span></span>|

#### <a name="internetmessageid-string"></a><span data-ttu-id="0e1d3-349">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="0e1d3-349">internetMessageId :String</span></span>

<span data-ttu-id="0e1d3-p114">Получает идентификатор интернет-сообщения для электронного сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-p114">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="0e1d3-352">Тип:</span><span class="sxs-lookup"><span data-stu-id="0e1d3-352">Type:</span></span>

*   <span data-ttu-id="0e1d3-353">String</span><span class="sxs-lookup"><span data-stu-id="0e1d3-353">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="0e1d3-354">Требования</span><span class="sxs-lookup"><span data-stu-id="0e1d3-354">Requirements</span></span>

|<span data-ttu-id="0e1d3-355">Требование</span><span class="sxs-lookup"><span data-stu-id="0e1d3-355">Requirement</span></span>| <span data-ttu-id="0e1d3-356">Значение</span><span class="sxs-lookup"><span data-stu-id="0e1d3-356">Value</span></span>|
|---|---|
|[<span data-ttu-id="0e1d3-357">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="0e1d3-357">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0e1d3-358">1.0</span><span class="sxs-lookup"><span data-stu-id="0e1d3-358">1.0</span></span>|
|[<span data-ttu-id="0e1d3-359">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="0e1d3-359">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0e1d3-360">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0e1d3-360">ReadItem</span></span>|
|[<span data-ttu-id="0e1d3-361">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="0e1d3-361">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="0e1d3-362">Чтение</span><span class="sxs-lookup"><span data-stu-id="0e1d3-362">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0e1d3-363">Пример</span><span class="sxs-lookup"><span data-stu-id="0e1d3-363">Example</span></span>

```
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="0e1d3-364">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="0e1d3-364">itemClass :String</span></span>

<span data-ttu-id="0e1d3-p115">Получает класс элемента веб-служб Exchange для выбранного элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-p115">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="0e1d3-p116">Свойство `itemClass` указывает класс сообщения выбранного элемента. Ниже приводятся классы сообщения по умолчанию для элемента сообщения или встречи.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-p116">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="0e1d3-369">Тип</span><span class="sxs-lookup"><span data-stu-id="0e1d3-369">Type</span></span> | <span data-ttu-id="0e1d3-370">Описание</span><span class="sxs-lookup"><span data-stu-id="0e1d3-370">Description</span></span> | <span data-ttu-id="0e1d3-371">item class</span><span class="sxs-lookup"><span data-stu-id="0e1d3-371">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="0e1d3-372">Элементы встречи</span><span class="sxs-lookup"><span data-stu-id="0e1d3-372">Appointment items</span></span> | <span data-ttu-id="0e1d3-373">Это элементы календаря для класса элемента `IPM.Appointment` или `IPM.Appointment.Occurence`.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-373">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurence` |
| <span data-ttu-id="0e1d3-374">Элементы сообщения</span><span class="sxs-lookup"><span data-stu-id="0e1d3-374">Message items</span></span> | <span data-ttu-id="0e1d3-375">Сюда входят электронные сообщения, для которых по умолчанию задан класс сообщения `IPM.Note`, а также приглашения на собрания, ответы на них и уведомления об их отмене, использующие `IPM.Schedule.Meeting` в качестве базового класса сообщений.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-375">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="0e1d3-376">Вы можете создавать настраиваемые классы сообщения, расширяющие классы сообщения по умолчанию, например, настраиваемый класс сообщения о встрече `IPM.Appointment.Contoso`.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-376">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="0e1d3-377">Тип:</span><span class="sxs-lookup"><span data-stu-id="0e1d3-377">Type:</span></span>

*   <span data-ttu-id="0e1d3-378">String</span><span class="sxs-lookup"><span data-stu-id="0e1d3-378">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="0e1d3-379">Требования</span><span class="sxs-lookup"><span data-stu-id="0e1d3-379">Requirements</span></span>

|<span data-ttu-id="0e1d3-380">Требование</span><span class="sxs-lookup"><span data-stu-id="0e1d3-380">Requirement</span></span>| <span data-ttu-id="0e1d3-381">Значение</span><span class="sxs-lookup"><span data-stu-id="0e1d3-381">Value</span></span>|
|---|---|
|[<span data-ttu-id="0e1d3-382">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="0e1d3-382">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0e1d3-383">1.0</span><span class="sxs-lookup"><span data-stu-id="0e1d3-383">1.0</span></span>|
|[<span data-ttu-id="0e1d3-384">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="0e1d3-384">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0e1d3-385">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0e1d3-385">ReadItem</span></span>|
|[<span data-ttu-id="0e1d3-386">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="0e1d3-386">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="0e1d3-387">Чтение</span><span class="sxs-lookup"><span data-stu-id="0e1d3-387">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0e1d3-388">Пример</span><span class="sxs-lookup"><span data-stu-id="0e1d3-388">Example</span></span>

```
var itemClass = Office.context.mailbox.item.itemClass;
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="0e1d3-389">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="0e1d3-389">(nullable) itemId :String</span></span>

<span data-ttu-id="0e1d3-p117">Получает идентификатор элемента веб-служб Exchange для текущего элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-p117">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="0e1d3-392">Идентификатор, возвращаемый свойством `itemId`, совпадает с идентификатором элемента веб-служб Exchange.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-392">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="0e1d3-393">Свойство  `itemId` не совпадает с идентификатором записи Outlook или идентификатором, используемым API-Интерфейсом REST Outlook.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-393">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="0e1d3-394">Прежде чем осуществлять вызовы API-Интерфейса REST с помощью этого значения, его следует преобразовать с помощью [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="0e1d3-394">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="0e1d3-395">Дополнительные сведения см. в статье [Использование API REST для Outlook из надстройки Outlook](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="0e1d3-395">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="0e1d3-p119">Свойство `itemId` недоступно в режиме создания. Если требуется идентификатор элемента, с помощью метода [`saveAsync`](#saveasyncoptions-callback) можно сохранить элемент в хранилище. При этом в параметре [`AsyncResult.value`](/javascript/api/office/office.asyncresult) функции обратного вызова возвращается идентификатор элемента.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-p119">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="0e1d3-398">Тип:</span><span class="sxs-lookup"><span data-stu-id="0e1d3-398">Type:</span></span>

*   <span data-ttu-id="0e1d3-399">String</span><span class="sxs-lookup"><span data-stu-id="0e1d3-399">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="0e1d3-400">Требования</span><span class="sxs-lookup"><span data-stu-id="0e1d3-400">Requirements</span></span>

|<span data-ttu-id="0e1d3-401">Требование</span><span class="sxs-lookup"><span data-stu-id="0e1d3-401">Requirement</span></span>| <span data-ttu-id="0e1d3-402">Значение</span><span class="sxs-lookup"><span data-stu-id="0e1d3-402">Value</span></span>|
|---|---|
|[<span data-ttu-id="0e1d3-403">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="0e1d3-403">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0e1d3-404">1.0</span><span class="sxs-lookup"><span data-stu-id="0e1d3-404">1.0</span></span>|
|[<span data-ttu-id="0e1d3-405">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="0e1d3-405">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0e1d3-406">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0e1d3-406">ReadItem</span></span>|
|[<span data-ttu-id="0e1d3-407">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="0e1d3-407">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="0e1d3-408">Чтение</span><span class="sxs-lookup"><span data-stu-id="0e1d3-408">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0e1d3-409">Пример</span><span class="sxs-lookup"><span data-stu-id="0e1d3-409">Example</span></span>

<span data-ttu-id="0e1d3-p120">Указанный ниже код проверяет наличие идентификатора элемента. Если свойство `itemId` возвращает значение `null` или `undefined`, элемент будет сохранен в хранилище, а из асинхронного результата будет получен идентификатор элемента.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result){
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook15officemailboxenumsitemtype"></a><span data-ttu-id="0e1d3-412">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_5/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="0e1d3-412">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_5/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="0e1d3-413">Получает тип элемента, который представляет экземпляр.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-413">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="0e1d3-414">Свойство `itemType` возвращает одно из значений перечисления `ItemType`, которое указывает, является ли экземпляр объекта `item` сообщением или встречей.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-414">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="0e1d3-415">Тип:</span><span class="sxs-lookup"><span data-stu-id="0e1d3-415">Type:</span></span>

*   [<span data-ttu-id="0e1d3-416">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="0e1d3-416">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook_1_5/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="0e1d3-417">Требования</span><span class="sxs-lookup"><span data-stu-id="0e1d3-417">Requirements</span></span>

|<span data-ttu-id="0e1d3-418">Требование</span><span class="sxs-lookup"><span data-stu-id="0e1d3-418">Requirement</span></span>| <span data-ttu-id="0e1d3-419">Значение</span><span class="sxs-lookup"><span data-stu-id="0e1d3-419">Value</span></span>|
|---|---|
|[<span data-ttu-id="0e1d3-420">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="0e1d3-420">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0e1d3-421">1.0</span><span class="sxs-lookup"><span data-stu-id="0e1d3-421">1.0</span></span>|
|[<span data-ttu-id="0e1d3-422">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="0e1d3-422">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0e1d3-423">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0e1d3-423">ReadItem</span></span>|
|[<span data-ttu-id="0e1d3-424">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="0e1d3-424">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="0e1d3-425">Cоздание или чтение</span><span class="sxs-lookup"><span data-stu-id="0e1d3-425">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="0e1d3-426">Пример</span><span class="sxs-lookup"><span data-stu-id="0e1d3-426">Example</span></span>

```
if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Message)
  // do something
else
  // do something else
```

####  <a name="location-stringlocationjavascriptapioutlook15officelocation"></a><span data-ttu-id="0e1d3-427">location :String|[Location](/javascript/api/outlook_1_5/office.location)</span><span class="sxs-lookup"><span data-stu-id="0e1d3-427">location :String|[Location](/javascript/api/outlook_1_5/office.location)</span></span>

<span data-ttu-id="0e1d3-428">Получает или задает место встречи.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-428">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="0e1d3-429">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="0e1d3-429">Read mode</span></span>

<span data-ttu-id="0e1d3-430">Свойство `location` возвращает строку, содержащую сведения о месте встречи.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-430">The `location` property returns a string that contains the location of the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="0e1d3-431">Режим создания</span><span class="sxs-lookup"><span data-stu-id="0e1d3-431">Compose mode</span></span>

<span data-ttu-id="0e1d3-432">Свойство `location` возвращает объект `Location`, предоставляющий методы, которые используются для получения и задания места встречи.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-432">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="0e1d3-433">Тип:</span><span class="sxs-lookup"><span data-stu-id="0e1d3-433">Type:</span></span>

*   <span data-ttu-id="0e1d3-434">String | [Location](/javascript/api/outlook_1_5/office.location)</span><span class="sxs-lookup"><span data-stu-id="0e1d3-434">String | [Location](/javascript/api/outlook_1_5/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="0e1d3-435">Требования</span><span class="sxs-lookup"><span data-stu-id="0e1d3-435">Requirements</span></span>

|<span data-ttu-id="0e1d3-436">Требование</span><span class="sxs-lookup"><span data-stu-id="0e1d3-436">Requirement</span></span>| <span data-ttu-id="0e1d3-437">Значение</span><span class="sxs-lookup"><span data-stu-id="0e1d3-437">Value</span></span>|
|---|---|
|[<span data-ttu-id="0e1d3-438">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="0e1d3-438">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0e1d3-439">1.0</span><span class="sxs-lookup"><span data-stu-id="0e1d3-439">1.0</span></span>|
|[<span data-ttu-id="0e1d3-440">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="0e1d3-440">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0e1d3-441">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0e1d3-441">ReadItem</span></span>|
|[<span data-ttu-id="0e1d3-442">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="0e1d3-442">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="0e1d3-443">Cоздание или чтение</span><span class="sxs-lookup"><span data-stu-id="0e1d3-443">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="0e1d3-444">Пример</span><span class="sxs-lookup"><span data-stu-id="0e1d3-444">Example</span></span>

```
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

#### <a name="normalizedsubject-string"></a><span data-ttu-id="0e1d3-445">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="0e1d3-445">normalizedSubject :String</span></span>

<span data-ttu-id="0e1d3-p121">Получает тему элемента со всеми удаленными префиксами (включая `RE:` и `FWD:`). Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="0e1d3-p122">Свойство normalizedSubject получает тему элемента со стандартными префиксами (такими как `RE:` и `FW:`), добавляемыми почтовыми программами. Для получения темы элемента с неизмененными префиксами используйте свойство [`subject`](#subject-stringsubjectjavascriptapioutlook15officesubject).</span><span class="sxs-lookup"><span data-stu-id="0e1d3-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubjectjavascriptapioutlook15officesubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="0e1d3-450">Тип:</span><span class="sxs-lookup"><span data-stu-id="0e1d3-450">Type:</span></span>

*   <span data-ttu-id="0e1d3-451">String</span><span class="sxs-lookup"><span data-stu-id="0e1d3-451">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="0e1d3-452">Требования</span><span class="sxs-lookup"><span data-stu-id="0e1d3-452">Requirements</span></span>

|<span data-ttu-id="0e1d3-453">Требование</span><span class="sxs-lookup"><span data-stu-id="0e1d3-453">Requirement</span></span>| <span data-ttu-id="0e1d3-454">Значение</span><span class="sxs-lookup"><span data-stu-id="0e1d3-454">Value</span></span>|
|---|---|
|[<span data-ttu-id="0e1d3-455">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="0e1d3-455">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0e1d3-456">1.0</span><span class="sxs-lookup"><span data-stu-id="0e1d3-456">1.0</span></span>|
|[<span data-ttu-id="0e1d3-457">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="0e1d3-457">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0e1d3-458">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0e1d3-458">ReadItem</span></span>|
|[<span data-ttu-id="0e1d3-459">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="0e1d3-459">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="0e1d3-460">Чтение</span><span class="sxs-lookup"><span data-stu-id="0e1d3-460">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0e1d3-461">Пример</span><span class="sxs-lookup"><span data-stu-id="0e1d3-461">Example</span></span>

```
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
```

####  <a name="notificationmessages-notificationmessagesjavascriptapioutlook15officenotificationmessages"></a><span data-ttu-id="0e1d3-462">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_5/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="0e1d3-462">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_5/office.notificationmessages)</span></span>

<span data-ttu-id="0e1d3-463">Получает сообщения уведомления для элемента.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-463">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="0e1d3-464">Тип:</span><span class="sxs-lookup"><span data-stu-id="0e1d3-464">Type:</span></span>

*   [<span data-ttu-id="0e1d3-465">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="0e1d3-465">NotificationMessages</span></span>](/javascript/api/outlook_1_5/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="0e1d3-466">Требования</span><span class="sxs-lookup"><span data-stu-id="0e1d3-466">Requirements</span></span>

|<span data-ttu-id="0e1d3-467">Требование</span><span class="sxs-lookup"><span data-stu-id="0e1d3-467">Requirement</span></span>| <span data-ttu-id="0e1d3-468">Значение</span><span class="sxs-lookup"><span data-stu-id="0e1d3-468">Value</span></span>|
|---|---|
|[<span data-ttu-id="0e1d3-469">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="0e1d3-469">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0e1d3-470">1.3</span><span class="sxs-lookup"><span data-stu-id="0e1d3-470">1.3</span></span>|
|[<span data-ttu-id="0e1d3-471">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="0e1d3-471">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0e1d3-472">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0e1d3-472">ReadItem</span></span>|
|[<span data-ttu-id="0e1d3-473">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="0e1d3-473">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="0e1d3-474">Cоздание или чтение</span><span class="sxs-lookup"><span data-stu-id="0e1d3-474">Compose or read</span></span>|

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients"></a><span data-ttu-id="0e1d3-475">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="0e1d3-475">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

<span data-ttu-id="0e1d3-476">Предоставляет доступ к необязательным участникам события.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-476">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="0e1d3-477">Уровень доступа и тип объекта, зависит от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-477">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="0e1d3-478">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="0e1d3-478">Read mode</span></span>

<span data-ttu-id="0e1d3-479">Свойство `optionalAttendees` возвращает массив, содержащий объект `EmailAddressDetails` для каждого необязательного участника собрания.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-479">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="0e1d3-480">Режим создания</span><span class="sxs-lookup"><span data-stu-id="0e1d3-480">Compose mode</span></span>

<span data-ttu-id="0e1d3-481">Свойство `optionalAttendees` возвращает объект `Recipients`, который предоставляет методы для получения и обновления необязательных участников собрания.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-481">The `optionalAttendees` property returns a `Recipients` object that provides methods to get and set the optional attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="0e1d3-482">Тип:</span><span class="sxs-lookup"><span data-stu-id="0e1d3-482">Type:</span></span>

*   <span data-ttu-id="0e1d3-483">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="0e1d3-483">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="0e1d3-484">Требования</span><span class="sxs-lookup"><span data-stu-id="0e1d3-484">Requirements</span></span>

|<span data-ttu-id="0e1d3-485">Требование</span><span class="sxs-lookup"><span data-stu-id="0e1d3-485">Requirement</span></span>| <span data-ttu-id="0e1d3-486">Значение</span><span class="sxs-lookup"><span data-stu-id="0e1d3-486">Value</span></span>|
|---|---|
|[<span data-ttu-id="0e1d3-487">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="0e1d3-487">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0e1d3-488">1.0</span><span class="sxs-lookup"><span data-stu-id="0e1d3-488">1.0</span></span>|
|[<span data-ttu-id="0e1d3-489">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="0e1d3-489">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0e1d3-490">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0e1d3-490">ReadItem</span></span>|
|[<span data-ttu-id="0e1d3-491">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="0e1d3-491">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="0e1d3-492">Cоздание или чтение</span><span class="sxs-lookup"><span data-stu-id="0e1d3-492">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="0e1d3-493">Пример</span><span class="sxs-lookup"><span data-stu-id="0e1d3-493">Example</span></span>

```
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails"></a><span data-ttu-id="0e1d3-494">organizer :[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="0e1d3-494">organizer :[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)</span></span>

<span data-ttu-id="0e1d3-p124">Получает электронный адрес организатора указанного собрания. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="0e1d3-497">Тип:</span><span class="sxs-lookup"><span data-stu-id="0e1d3-497">Type:</span></span>

*   [<span data-ttu-id="0e1d3-498">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="0e1d3-498">EmailAddressDetails</span></span>](/javascript/api/outlook_1_5/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="0e1d3-499">Требования</span><span class="sxs-lookup"><span data-stu-id="0e1d3-499">Requirements</span></span>

|<span data-ttu-id="0e1d3-500">Требование</span><span class="sxs-lookup"><span data-stu-id="0e1d3-500">Requirement</span></span>| <span data-ttu-id="0e1d3-501">Значение</span><span class="sxs-lookup"><span data-stu-id="0e1d3-501">Value</span></span>|
|---|---|
|[<span data-ttu-id="0e1d3-502">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="0e1d3-502">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0e1d3-503">1.0</span><span class="sxs-lookup"><span data-stu-id="0e1d3-503">1.0</span></span>|
|[<span data-ttu-id="0e1d3-504">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="0e1d3-504">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0e1d3-505">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0e1d3-505">ReadItem</span></span>|
|[<span data-ttu-id="0e1d3-506">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="0e1d3-506">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="0e1d3-507">Чтение</span><span class="sxs-lookup"><span data-stu-id="0e1d3-507">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0e1d3-508">Пример</span><span class="sxs-lookup"><span data-stu-id="0e1d3-508">Example</span></span>

```
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
```

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients"></a><span data-ttu-id="0e1d3-509">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="0e1d3-509">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

<span data-ttu-id="0e1d3-510">Предоставляет доступ к обязательным участникам события.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-510">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="0e1d3-511">Уровень доступа и тип объекта, зависит от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-511">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="0e1d3-512">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="0e1d3-512">Read mode</span></span>

<span data-ttu-id="0e1d3-513">Свойство `requiredAttendees` возвращает массив, содержащий объект `EmailAddressDetails`, каждому обязательному участнику собрания.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-513">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="0e1d3-514">Режим создания</span><span class="sxs-lookup"><span data-stu-id="0e1d3-514">Compose mode</span></span>

<span data-ttu-id="0e1d3-515">Свойство `requiredAttendees` возвращает объект `Recipients`, который предоставляет методы для получения и обновления обязательных участников собрания.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-515">The `requiredAttendees` property returns a `Recipients` object that provides methods to get and set the required attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="0e1d3-516">Тип:</span><span class="sxs-lookup"><span data-stu-id="0e1d3-516">Type:</span></span>

*   <span data-ttu-id="0e1d3-517">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="0e1d3-517">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="0e1d3-518">Требования</span><span class="sxs-lookup"><span data-stu-id="0e1d3-518">Requirements</span></span>

|<span data-ttu-id="0e1d3-519">Требование</span><span class="sxs-lookup"><span data-stu-id="0e1d3-519">Requirement</span></span>| <span data-ttu-id="0e1d3-520">Значение</span><span class="sxs-lookup"><span data-stu-id="0e1d3-520">Value</span></span>|
|---|---|
|[<span data-ttu-id="0e1d3-521">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="0e1d3-521">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0e1d3-522">1.0</span><span class="sxs-lookup"><span data-stu-id="0e1d3-522">1.0</span></span>|
|[<span data-ttu-id="0e1d3-523">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="0e1d3-523">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0e1d3-524">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0e1d3-524">ReadItem</span></span>|
|[<span data-ttu-id="0e1d3-525">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="0e1d3-525">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="0e1d3-526">Cоздание или чтение</span><span class="sxs-lookup"><span data-stu-id="0e1d3-526">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="0e1d3-527">Пример</span><span class="sxs-lookup"><span data-stu-id="0e1d3-527">Example</span></span>

```
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
}
```

#### <a name="sender-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails"></a><span data-ttu-id="0e1d3-528">sender :[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="0e1d3-528">sender :[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)</span></span>

<span data-ttu-id="0e1d3-p126">Получает электронный адрес отправителя электронного сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="0e1d3-p127">Свойства [`from`](#from-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails) и `sender` представляют одно и то же лицо, если сообщение не отправлено делегатом. В противном случае свойство `from` представляет лицо, делегировавшее полномочия, а свойство sender — делегата.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-p127">The [`from`](#from-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="0e1d3-533">Свойство `recipientType` объекта `EmailAddressDetails` в свойстве `sender` — `undefined`.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-533">Note: The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="0e1d3-534">Тип:</span><span class="sxs-lookup"><span data-stu-id="0e1d3-534">Type:</span></span>

*   [<span data-ttu-id="0e1d3-535">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="0e1d3-535">EmailAddressDetails</span></span>](/javascript/api/outlook_1_5/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="0e1d3-536">Требования</span><span class="sxs-lookup"><span data-stu-id="0e1d3-536">Requirements</span></span>

|<span data-ttu-id="0e1d3-537">Требование</span><span class="sxs-lookup"><span data-stu-id="0e1d3-537">Requirement</span></span>| <span data-ttu-id="0e1d3-538">Значение</span><span class="sxs-lookup"><span data-stu-id="0e1d3-538">Value</span></span>|
|---|---|
|[<span data-ttu-id="0e1d3-539">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="0e1d3-539">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0e1d3-540">1.0</span><span class="sxs-lookup"><span data-stu-id="0e1d3-540">1.0</span></span>|
|[<span data-ttu-id="0e1d3-541">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="0e1d3-541">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0e1d3-542">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0e1d3-542">ReadItem</span></span>|
|[<span data-ttu-id="0e1d3-543">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="0e1d3-543">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="0e1d3-544">Чтение</span><span class="sxs-lookup"><span data-stu-id="0e1d3-544">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0e1d3-545">Пример</span><span class="sxs-lookup"><span data-stu-id="0e1d3-545">Example</span></span>

```
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
```

####  <a name="start-datetimejavascriptapioutlook15officetime"></a><span data-ttu-id="0e1d3-546">start :Date|[Time](/javascript/api/outlook_1_5/office.time)</span><span class="sxs-lookup"><span data-stu-id="0e1d3-546">start :Date|[Time](/javascript/api/outlook_1_5/office.time)</span></span>

<span data-ttu-id="0e1d3-547">Получает или задает дату и время начала встречи.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-547">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="0e1d3-p128">Свойство `start` представлено в виде значения даты и времени в формате UTC. Это значение можно преобразовать в местные значения даты и времени клиента с помощью метода [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook15officelocalclienttime).</span><span class="sxs-lookup"><span data-stu-id="0e1d3-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook15officelocalclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="0e1d3-550">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="0e1d3-550">Read mode</span></span>

<span data-ttu-id="0e1d3-551">Свойство `start` возвращает объект `Date`.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-551">The `start` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="0e1d3-552">Режим создания</span><span class="sxs-lookup"><span data-stu-id="0e1d3-552">Compose mode</span></span>

<span data-ttu-id="0e1d3-553">Свойство `start` возвращает объект `Time`.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-553">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="0e1d3-554">Если вы задаете время начала с помощью метода [`Time.setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-), необходимо использовать метод [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) для преобразования местного времени на клиенте в формат UTC для сервера.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-554">When you use the [`Time.setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="0e1d3-555">Тип:</span><span class="sxs-lookup"><span data-stu-id="0e1d3-555">Type:</span></span>

*   <span data-ttu-id="0e1d3-556">Date | [Time](/javascript/api/outlook_1_5/office.time)</span><span class="sxs-lookup"><span data-stu-id="0e1d3-556">Date | [Time](/javascript/api/outlook_1_5/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="0e1d3-557">Требования</span><span class="sxs-lookup"><span data-stu-id="0e1d3-557">Requirements</span></span>

|<span data-ttu-id="0e1d3-558">Требование</span><span class="sxs-lookup"><span data-stu-id="0e1d3-558">Requirement</span></span>| <span data-ttu-id="0e1d3-559">Значение</span><span class="sxs-lookup"><span data-stu-id="0e1d3-559">Value</span></span>|
|---|---|
|[<span data-ttu-id="0e1d3-560">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="0e1d3-560">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0e1d3-561">1.0</span><span class="sxs-lookup"><span data-stu-id="0e1d3-561">1.0</span></span>|
|[<span data-ttu-id="0e1d3-562">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="0e1d3-562">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0e1d3-563">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0e1d3-563">ReadItem</span></span>|
|[<span data-ttu-id="0e1d3-564">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="0e1d3-564">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="0e1d3-565">Cоздание или чтение</span><span class="sxs-lookup"><span data-stu-id="0e1d3-565">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="0e1d3-566">Пример</span><span class="sxs-lookup"><span data-stu-id="0e1d3-566">Example</span></span>

<span data-ttu-id="0e1d3-567">В примере ниже с помощью метода [`setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) объекта `Time` задается время начала встречи в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-567">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

####  <a name="subject-stringsubjectjavascriptapioutlook15officesubject"></a><span data-ttu-id="0e1d3-568">subject :String|[Subject](/javascript/api/outlook_1_5/office.subject)</span><span class="sxs-lookup"><span data-stu-id="0e1d3-568">subject :String|[Subject](/javascript/api/outlook_1_5/office.subject)</span></span>

<span data-ttu-id="0e1d3-569">Получает или задает описание, которое отображается в поле темы элемента.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-569">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="0e1d3-570">Свойство `subject` получает или задает всю тему элемента для отправки с почтового сервера.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-570">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="0e1d3-571">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="0e1d3-571">Read mode</span></span>

<span data-ttu-id="0e1d3-p129">Свойство `subject` возвращает строку. С помощью свойства [`normalizedSubject`](#normalizedsubject-string) можно получить тему без начальных префиксов, например, `RE:` и `FW:`.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-p129">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```
var subject = Office.context.mailbox.item.subject;
```

##### <a name="compose-mode"></a><span data-ttu-id="0e1d3-574">Режим создания</span><span class="sxs-lookup"><span data-stu-id="0e1d3-574">Compose mode</span></span>

<span data-ttu-id="0e1d3-575">Свойство `subject` возвращает объект `Subject`, который предоставляет методы для получения и задания темы.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-575">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="0e1d3-576">Тип:</span><span class="sxs-lookup"><span data-stu-id="0e1d3-576">Type:</span></span>

*   <span data-ttu-id="0e1d3-577">String | [Subject](/javascript/api/outlook_1_5/office.subject)</span><span class="sxs-lookup"><span data-stu-id="0e1d3-577">String | [Subject](/javascript/api/outlook_1_5/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="0e1d3-578">Требования</span><span class="sxs-lookup"><span data-stu-id="0e1d3-578">Requirements</span></span>

|<span data-ttu-id="0e1d3-579">Требование</span><span class="sxs-lookup"><span data-stu-id="0e1d3-579">Requirement</span></span>| <span data-ttu-id="0e1d3-580">Значение</span><span class="sxs-lookup"><span data-stu-id="0e1d3-580">Value</span></span>|
|---|---|
|[<span data-ttu-id="0e1d3-581">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="0e1d3-581">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0e1d3-582">1.0</span><span class="sxs-lookup"><span data-stu-id="0e1d3-582">1.0</span></span>|
|[<span data-ttu-id="0e1d3-583">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="0e1d3-583">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0e1d3-584">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0e1d3-584">ReadItem</span></span>|
|[<span data-ttu-id="0e1d3-585">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="0e1d3-585">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="0e1d3-586">Cоздание или чтение</span><span class="sxs-lookup"><span data-stu-id="0e1d3-586">Compose or read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients"></a><span data-ttu-id="0e1d3-587">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="0e1d3-587">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

<span data-ttu-id="0e1d3-588">Предоставляет доступ получателей к строке **To (Кому)** в сообщении.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-588">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="0e1d3-589">Уровень доступа и тип объекта, зависит от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-589">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="0e1d3-590">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="0e1d3-590">Read mode</span></span>

<span data-ttu-id="0e1d3-p131">Свойство `to` возвращает массив, содержащий объект `EmailAddressDetails` для каждого получателя в строке **To (Кому)** сообщения. Коллекция может включать не более 100 элементов.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-p131">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="0e1d3-593">Режим создания</span><span class="sxs-lookup"><span data-stu-id="0e1d3-593">Compose mode</span></span>

<span data-ttu-id="0e1d3-594">Свойство `to` возвращает объект `Recipients`, предоставляющий методы для получения или обновления получателей, которые указаны в строке **To (кому)** сообщения.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-594">The to`to` property returns a Recipients`Recipients` object that provides methods for manipulating the recipients on the **To** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="0e1d3-595">Тип:</span><span class="sxs-lookup"><span data-stu-id="0e1d3-595">Type:</span></span>

*   <span data-ttu-id="0e1d3-596">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="0e1d3-596">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="0e1d3-597">Требования</span><span class="sxs-lookup"><span data-stu-id="0e1d3-597">Requirements</span></span>

|<span data-ttu-id="0e1d3-598">Требование</span><span class="sxs-lookup"><span data-stu-id="0e1d3-598">Requirement</span></span>| <span data-ttu-id="0e1d3-599">Значение</span><span class="sxs-lookup"><span data-stu-id="0e1d3-599">Value</span></span>|
|---|---|
|[<span data-ttu-id="0e1d3-600">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="0e1d3-600">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0e1d3-601">1.0</span><span class="sxs-lookup"><span data-stu-id="0e1d3-601">1.0</span></span>|
|[<span data-ttu-id="0e1d3-602">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="0e1d3-602">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0e1d3-603">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0e1d3-603">ReadItem</span></span>|
|[<span data-ttu-id="0e1d3-604">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="0e1d3-604">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="0e1d3-605">Cоздание или чтение</span><span class="sxs-lookup"><span data-stu-id="0e1d3-605">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="0e1d3-606">Пример</span><span class="sxs-lookup"><span data-stu-id="0e1d3-606">Example</span></span>

```
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

### <a name="methods"></a><span data-ttu-id="0e1d3-607">Методы</span><span class="sxs-lookup"><span data-stu-id="0e1d3-607">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="0e1d3-608">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="0e1d3-608">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="0e1d3-609">Добавляет файл в сообщение или встречу в качестве вложения.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-609">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="0e1d3-610">Метод `addFileAttachmentAsync` передает файл по указанному универсальному коду ресурса (URI) и вкладывает его в элемент в форме создания.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-610">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="0e1d3-611">Идентификатор можно последовательно использовать с методом [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback), чтобы удалить вложение, добавленное во время текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-611">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0e1d3-612">Параметры:</span><span class="sxs-lookup"><span data-stu-id="0e1d3-612">Parameters:</span></span>

|<span data-ttu-id="0e1d3-613">Имя</span><span class="sxs-lookup"><span data-stu-id="0e1d3-613">Name</span></span>| <span data-ttu-id="0e1d3-614">Тип</span><span class="sxs-lookup"><span data-stu-id="0e1d3-614">Type</span></span>| <span data-ttu-id="0e1d3-615">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="0e1d3-615">Attributes</span></span>| <span data-ttu-id="0e1d3-616">Описание</span><span class="sxs-lookup"><span data-stu-id="0e1d3-616">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="0e1d3-617">String</span><span class="sxs-lookup"><span data-stu-id="0e1d3-617">String</span></span>||<span data-ttu-id="0e1d3-p132">URI-адрес, представляющий расположение файла, который нужно вложить в сообщение или встречу. Максимальная длина — 2048 символов.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-p132">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="0e1d3-620">String</span><span class="sxs-lookup"><span data-stu-id="0e1d3-620">String</span></span>||<span data-ttu-id="0e1d3-p133">Имя вложения, которое отображается при передаче вложения. Максимальная длина — 255 символов.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-p133">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="0e1d3-623">Object</span><span class="sxs-lookup"><span data-stu-id="0e1d3-623">Object</span></span>| <span data-ttu-id="0e1d3-624">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="0e1d3-624">&lt;optional&gt;</span></span>|<span data-ttu-id="0e1d3-625">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-625">An object literal that contains one or more of the following properties.</span></span>|
| `options.asyncContext` | <span data-ttu-id="0e1d3-626">Object</span><span class="sxs-lookup"><span data-stu-id="0e1d3-626">Object</span></span> | <span data-ttu-id="0e1d3-627">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="0e1d3-627">&lt;optional&gt;</span></span> | <span data-ttu-id="0e1d3-628">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-628">Developers can provide any object they wish to access in the callback method.</span></span> |
| `options.isInline` | <span data-ttu-id="0e1d3-629">Логическое значение</span><span class="sxs-lookup"><span data-stu-id="0e1d3-629">Boolean</span></span> | <span data-ttu-id="0e1d3-630">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="0e1d3-630">&lt;optional&gt;</span></span> | <span data-ttu-id="0e1d3-631">Значение `true` указывает на то, что вложение будет встроено в текст сообщения и не должно отображаться в списке вложений.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-631">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
|`callback`| <span data-ttu-id="0e1d3-632">function</span><span class="sxs-lookup"><span data-stu-id="0e1d3-632">function</span></span>| <span data-ttu-id="0e1d3-633">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="0e1d3-633">&lt;optional&gt;</span></span>|<span data-ttu-id="0e1d3-634">Когда метод завершает выполнение, переданная в параметре `callback` функция вызывается с единственным параметром `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="0e1d3-634">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="0e1d3-635">После успешного выполнения идентификатор вложения будет представлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-635">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="0e1d3-636">Если добавить вложение не удастся, объект `asyncResult` будет содержать объект `Error` с описанием ошибки.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-636">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="0e1d3-637">Ошибки</span><span class="sxs-lookup"><span data-stu-id="0e1d3-637">Errors</span></span>

| <span data-ttu-id="0e1d3-638">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="0e1d3-638">Error code</span></span> | <span data-ttu-id="0e1d3-639">Описание</span><span class="sxs-lookup"><span data-stu-id="0e1d3-639">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="0e1d3-640">Вложение превышает максимальный размер.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-640">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="0e1d3-641">Расширение вложения не поддерживается.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-641">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="0e1d3-642">Сообщение или встреча содержат слишком много вложений.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-642">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="0e1d3-643">Требования</span><span class="sxs-lookup"><span data-stu-id="0e1d3-643">Requirements</span></span>

|<span data-ttu-id="0e1d3-644">Требование</span><span class="sxs-lookup"><span data-stu-id="0e1d3-644">Requirement</span></span>| <span data-ttu-id="0e1d3-645">Значение</span><span class="sxs-lookup"><span data-stu-id="0e1d3-645">Value</span></span>|
|---|---|
|[<span data-ttu-id="0e1d3-646">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="0e1d3-646">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0e1d3-647">1.1</span><span class="sxs-lookup"><span data-stu-id="0e1d3-647">1.1</span></span>|
|[<span data-ttu-id="0e1d3-648">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="0e1d3-648">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0e1d3-649">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="0e1d3-649">ReadWriteItem</span></span>|
|[<span data-ttu-id="0e1d3-650">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="0e1d3-650">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="0e1d3-651">Создание</span><span class="sxs-lookup"><span data-stu-id="0e1d3-651">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="0e1d3-652">Примеры</span><span class="sxs-lookup"><span data-stu-id="0e1d3-652">Examples</span></span>

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

<span data-ttu-id="0e1d3-653">В приведенном ниже примере файл изображения добавляется в качестве встроенного вложения, а ссылка на вложение добавляется в текст сообщения.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-653">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

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

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="0e1d3-654">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="0e1d3-654">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="0e1d3-655">Добавляет к сообщению или встрече элемент Exchange (например, сообщение) в виде вложения.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-655">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="0e1d3-p134">С помощью метода `addItemAttachmentAsync` в элемент формы создания можно вложить элемент с указанным идентификатором Exchange. Если указать метод обратного вызова, то этот метод вызывается с помощью параметра `asyncResult`, который содержит идентификатор вложения или код, указывающий на ошибки, которые произошли при вложении элемента. При необходимости можно использовать параметр `options` для передачи сведений о состоянии в метод обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-p134">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="0e1d3-659">Идентификатор можно последовательно использовать с методом [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback), чтобы удалить вложение, добавленное во время текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-659">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="0e1d3-660">Если ваша надстройка Office выполняется в веб-приложении Outlook, метод `addItemAttachmentAsync` обеспечивает вложение элементов в элементы, отличные от редактируемого. Однако это действие не рекомендуется выполнять, так как оно не поддерживается.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-660">If your Office add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0e1d3-661">Параметры:</span><span class="sxs-lookup"><span data-stu-id="0e1d3-661">Parameters:</span></span>

|<span data-ttu-id="0e1d3-662">Имя</span><span class="sxs-lookup"><span data-stu-id="0e1d3-662">Name</span></span>| <span data-ttu-id="0e1d3-663">Тип</span><span class="sxs-lookup"><span data-stu-id="0e1d3-663">Type</span></span>| <span data-ttu-id="0e1d3-664">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="0e1d3-664">Attributes</span></span>| <span data-ttu-id="0e1d3-665">Описание</span><span class="sxs-lookup"><span data-stu-id="0e1d3-665">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="0e1d3-666">String</span><span class="sxs-lookup"><span data-stu-id="0e1d3-666">String</span></span>||<span data-ttu-id="0e1d3-p135">Идентификатор Exchange для вкладываемого элемента. Максимальная длина — 100 символов.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-p135">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="0e1d3-669">String</span><span class="sxs-lookup"><span data-stu-id="0e1d3-669">String</span></span>||<span data-ttu-id="0e1d3-p136">Тема вкладываемого элемента. Максимальная длина — 255 символов.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-p136">The sujbect of the item to be attached. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="0e1d3-672">Object</span><span class="sxs-lookup"><span data-stu-id="0e1d3-672">Object</span></span>| <span data-ttu-id="0e1d3-673">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="0e1d3-673">&lt;optional&gt;</span></span>|<span data-ttu-id="0e1d3-674">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-674">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="0e1d3-675">Object</span><span class="sxs-lookup"><span data-stu-id="0e1d3-675">Object</span></span>| <span data-ttu-id="0e1d3-676">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="0e1d3-676">&lt;optional&gt;</span></span>|<span data-ttu-id="0e1d3-677">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-677">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="0e1d3-678">function</span><span class="sxs-lookup"><span data-stu-id="0e1d3-678">function</span></span>| <span data-ttu-id="0e1d3-679">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="0e1d3-679">&lt;optional&gt;</span></span>|<span data-ttu-id="0e1d3-680">Когда метод завершает выполнение, переданная в параметре `callback` функция вызывается с единственным параметром `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="0e1d3-680">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="0e1d3-681">После успешного выполнения идентификатор вложения будет представлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-681">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="0e1d3-682">Если добавить вложение не удастся, объект `asyncResult` будет содержать объект `Error` с описанием ошибки.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-682">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="0e1d3-683">Ошибки</span><span class="sxs-lookup"><span data-stu-id="0e1d3-683">Errors</span></span>

| <span data-ttu-id="0e1d3-684">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="0e1d3-684">Error code</span></span> | <span data-ttu-id="0e1d3-685">Описание</span><span class="sxs-lookup"><span data-stu-id="0e1d3-685">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="0e1d3-686">Сообщение или встреча содержат слишком много вложений.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-686">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="0e1d3-687">Требования</span><span class="sxs-lookup"><span data-stu-id="0e1d3-687">Requirements</span></span>

|<span data-ttu-id="0e1d3-688">Требование</span><span class="sxs-lookup"><span data-stu-id="0e1d3-688">Requirement</span></span>| <span data-ttu-id="0e1d3-689">Значение</span><span class="sxs-lookup"><span data-stu-id="0e1d3-689">Value</span></span>|
|---|---|
|[<span data-ttu-id="0e1d3-690">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="0e1d3-690">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0e1d3-691">1.1</span><span class="sxs-lookup"><span data-stu-id="0e1d3-691">1.1</span></span>|
|[<span data-ttu-id="0e1d3-692">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="0e1d3-692">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0e1d3-693">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="0e1d3-693">ReadWriteItem</span></span>|
|[<span data-ttu-id="0e1d3-694">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="0e1d3-694">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="0e1d3-695">Создание</span><span class="sxs-lookup"><span data-stu-id="0e1d3-695">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="0e1d3-696">Пример</span><span class="sxs-lookup"><span data-stu-id="0e1d3-696">Example</span></span>

<span data-ttu-id="0e1d3-697">В следующем примере существующий элемент Outlook добавляется в виде вложения с именем `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-697">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

####  <a name="close"></a><span data-ttu-id="0e1d3-698">close()</span><span class="sxs-lookup"><span data-stu-id="0e1d3-698">close()</span></span>

<span data-ttu-id="0e1d3-699">Закрывает текущий создаваемый элемент.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-699">Closes the current item that is being composed.</span></span>

<span data-ttu-id="0e1d3-p137">Работа метода `close` зависит от текущего состояния создаваемого элемента. Если элемент содержит несохраненные изменения, клиент предложит пользователю сохранить или отклонить их либо отменить действие закрытия.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-p137">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="0e1d3-702">Если элемент является встречей в Outlook в Интернете, и он был ранее сохранен с помощью `saveAsync`, пользователю предлагается сохранить, отменить или удалить его, даже если не произошло каких-либо изменений, поскольку этот элемент был последним сохраненным.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-702">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="0e1d3-703">Если в клиенте Outlook для настольных ПК сообщение представляет собой ответ в тексте, метод `close` не работает.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-703">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="0e1d3-704">Требования</span><span class="sxs-lookup"><span data-stu-id="0e1d3-704">Requirements</span></span>

|<span data-ttu-id="0e1d3-705">Требование</span><span class="sxs-lookup"><span data-stu-id="0e1d3-705">Requirement</span></span>| <span data-ttu-id="0e1d3-706">Значение</span><span class="sxs-lookup"><span data-stu-id="0e1d3-706">Value</span></span>|
|---|---|
|[<span data-ttu-id="0e1d3-707">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="0e1d3-707">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0e1d3-708">1.3</span><span class="sxs-lookup"><span data-stu-id="0e1d3-708">1.3</span></span>|
|[<span data-ttu-id="0e1d3-709">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="0e1d3-709">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0e1d3-710">Ограниченный доступ</span><span class="sxs-lookup"><span data-stu-id="0e1d3-710">Restricted</span></span>|
|[<span data-ttu-id="0e1d3-711">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="0e1d3-711">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="0e1d3-712">Создание</span><span class="sxs-lookup"><span data-stu-id="0e1d3-712">Compose</span></span>|

#### <a name="displayreplyallformformdata"></a><span data-ttu-id="0e1d3-713">displayReplyAllForm(formData)</span><span class="sxs-lookup"><span data-stu-id="0e1d3-713">displayReplyAllForm(formData)</span></span>

<span data-ttu-id="0e1d3-714">Отображает форму ответа, включающую отправителя и всех получателей выбранного сообщения или организатора и всех участников выбранной встречи.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-714">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="0e1d3-715">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-715">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="0e1d3-716">В веб-приложении Outlook форма ответа отображается в виде всплывающей формы в представлении с 3 колонками либо всплывающей формы в представлении с 1 или 2 колонками.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-716">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="0e1d3-717">Если любой строчный параметр превышает указанные для него ограничения, `displayReplyAllForm` возвращает исключение.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-717">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="0e1d3-p138">Если в параметре `formData.attachments` указаны вложения, Outlook и веб-приложение Outlook пытаются скачать их и вложить в форму ответа. Если какие-либо вложения добавить не удается, в форме отображается сообщение об ошибке. Если сообщения об ошибках не предусмотрены, то они не отображаются.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-p138">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0e1d3-721">Параметры:</span><span class="sxs-lookup"><span data-stu-id="0e1d3-721">Parameters:</span></span>

| <span data-ttu-id="0e1d3-722">Имя</span><span class="sxs-lookup"><span data-stu-id="0e1d3-722">Name</span></span> | <span data-ttu-id="0e1d3-723">Тип</span><span class="sxs-lookup"><span data-stu-id="0e1d3-723">Type</span></span> | <span data-ttu-id="0e1d3-724">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="0e1d3-724">Attributes</span></span> | <span data-ttu-id="0e1d3-725">Описание</span><span class="sxs-lookup"><span data-stu-id="0e1d3-725">Description</span></span> |
|---|---|---|---|
|`formData`| <span data-ttu-id="0e1d3-726">String | Object</span><span class="sxs-lookup"><span data-stu-id="0e1d3-726">String &#124; Object</span></span>| |<span data-ttu-id="0e1d3-p139">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-p139">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="0e1d3-729">**ИЛИ**</span><span class="sxs-lookup"><span data-stu-id="0e1d3-729">**OR**</span></span><br/><span data-ttu-id="0e1d3-p140">Объект, который содержит текст или данные вложения и функцию обратного вызова. Ниже представлено определение этого объекта.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-p140">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="0e1d3-732">String</span><span class="sxs-lookup"><span data-stu-id="0e1d3-732">String</span></span> | <span data-ttu-id="0e1d3-733">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="0e1d3-733">&lt;optional&gt;</span></span> | <span data-ttu-id="0e1d3-p141">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-p141">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="0e1d3-736">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="0e1d3-736">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="0e1d3-737">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="0e1d3-737">&lt;optional&gt;</span></span> | <span data-ttu-id="0e1d3-738">Массив объектов JSON, представляющих собой вложенные файлы или элементы.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-738">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="0e1d3-739">String</span><span class="sxs-lookup"><span data-stu-id="0e1d3-739">String</span></span> | | <span data-ttu-id="0e1d3-p142">Указывает тип вложения. Допустимые значения: `file` для вложенного файла и `item` для вложенного элемента.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-p142">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="0e1d3-742">String</span><span class="sxs-lookup"><span data-stu-id="0e1d3-742">String</span></span> | | <span data-ttu-id="0e1d3-743">Строка, содержащая имя вложения, длиной до 255 символов.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-743">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="0e1d3-744">String</span><span class="sxs-lookup"><span data-stu-id="0e1d3-744">String</span></span> | | <span data-ttu-id="0e1d3-p143">Используется, только если свойству `type` задано значение `file`. URI расположения файла.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-p143">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.isInline` | <span data-ttu-id="0e1d3-747">Логическое значение</span><span class="sxs-lookup"><span data-stu-id="0e1d3-747">Boolean</span></span> | | <span data-ttu-id="0e1d3-p144">Используется только в том случае, если свойству `type` задано значение `file`. Значение `true` указывает на то, что вложение будет встроено в текст сообщения и не должно отображаться в списке вложений.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-p144">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="0e1d3-750">String</span><span class="sxs-lookup"><span data-stu-id="0e1d3-750">String</span></span> | | <span data-ttu-id="0e1d3-p145">Используется, только если свойству `type` задано значение `item`. Идентификатор вложения EWS. Это строка длиной до 100 символов.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-p145">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="0e1d3-754">function</span><span class="sxs-lookup"><span data-stu-id="0e1d3-754">function</span></span> | <span data-ttu-id="0e1d3-755">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="0e1d3-755">&lt;optional&gt;</span></span> | <span data-ttu-id="0e1d3-756">По завершении работы метода функция, переданная параметру `callback`, вызывается с помощью одного параметра `asyncResult`, представляющего собой объект [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="0e1d3-756">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="0e1d3-757">Требования</span><span class="sxs-lookup"><span data-stu-id="0e1d3-757">Requirements</span></span>

|<span data-ttu-id="0e1d3-758">Требование</span><span class="sxs-lookup"><span data-stu-id="0e1d3-758">Requirement</span></span>| <span data-ttu-id="0e1d3-759">Значение</span><span class="sxs-lookup"><span data-stu-id="0e1d3-759">Value</span></span>|
|---|---|
|[<span data-ttu-id="0e1d3-760">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="0e1d3-760">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0e1d3-761">1.0</span><span class="sxs-lookup"><span data-stu-id="0e1d3-761">1.0</span></span>|
|[<span data-ttu-id="0e1d3-762">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="0e1d3-762">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0e1d3-763">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0e1d3-763">ReadItem</span></span>|
|[<span data-ttu-id="0e1d3-764">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="0e1d3-764">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="0e1d3-765">Чтение</span><span class="sxs-lookup"><span data-stu-id="0e1d3-765">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="0e1d3-766">Примеры</span><span class="sxs-lookup"><span data-stu-id="0e1d3-766">Examples</span></span>

<span data-ttu-id="0e1d3-767">Приведенный ниже код передает строку в функцию `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-767">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="0e1d3-768">Ответ с пустым текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-768">Reply with an empty body.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="0e1d3-769">Ответ только с текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-769">Reply with just a body.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="0e1d3-770">Ответ с текстом сообщения и вложенным файлом.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-770">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="0e1d3-771">Ответ с текстом сообщения и вложенным элементом.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-771">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="0e1d3-772">Ответ с текстом сообщения, вложенным файлом, вложенным элементом и обратным вызовом.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-772">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata"></a><span data-ttu-id="0e1d3-773">displayReplyForm(formData)</span><span class="sxs-lookup"><span data-stu-id="0e1d3-773">displayReplyForm(formData)</span></span>

<span data-ttu-id="0e1d3-774">Отображает форму ответа, включающую только отправителя выбранного сообщения или организатора выбранной встречи.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-774">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="0e1d3-775">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-775">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="0e1d3-776">В веб-приложении Outlook форма ответа отображается в виде всплывающей формы в представлении с 3 колонками либо всплывающей формы в представлении с 1 или 2 колонками.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-776">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="0e1d3-777">Если любой строчный параметр превышает указанные для него ограничения, `displayReplyForm` возвращает исключение.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-777">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="0e1d3-p146">Если в параметре `formData.attachments` указаны вложения, Outlook и веб-приложение Outlook пытаются скачать их и вложить в форму ответа. Если какие-либо вложения добавить не удается, в форме отображается сообщение об ошибке. Если сообщения об ошибках не предусмотрены, то они не отображаются.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-p146">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0e1d3-781">Параметры:</span><span class="sxs-lookup"><span data-stu-id="0e1d3-781">Parameters:</span></span>

| <span data-ttu-id="0e1d3-782">Имя</span><span class="sxs-lookup"><span data-stu-id="0e1d3-782">Name</span></span> | <span data-ttu-id="0e1d3-783">Тип</span><span class="sxs-lookup"><span data-stu-id="0e1d3-783">Type</span></span> | <span data-ttu-id="0e1d3-784">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="0e1d3-784">Attributes</span></span> | <span data-ttu-id="0e1d3-785">Описание</span><span class="sxs-lookup"><span data-stu-id="0e1d3-785">Description</span></span> |
|---|---|---|---|
|`formData`| <span data-ttu-id="0e1d3-786">String | Object</span><span class="sxs-lookup"><span data-stu-id="0e1d3-786">String &#124; Object</span></span>| | <span data-ttu-id="0e1d3-p147">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-p147">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="0e1d3-789">**ИЛИ**</span><span class="sxs-lookup"><span data-stu-id="0e1d3-789">**OR**</span></span><br/><span data-ttu-id="0e1d3-p148">Объект, который содержит текст или данные вложения и функцию обратного вызова. Ниже представлено определение этого объекта.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-p148">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="0e1d3-792">String</span><span class="sxs-lookup"><span data-stu-id="0e1d3-792">String</span></span> | <span data-ttu-id="0e1d3-793">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="0e1d3-793">&lt;optional&gt;</span></span> | <span data-ttu-id="0e1d3-p149">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-p149">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="0e1d3-796">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="0e1d3-796">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="0e1d3-797">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="0e1d3-797">&lt;optional&gt;</span></span> | <span data-ttu-id="0e1d3-798">Массив объектов JSON, представляющих собой вложенные файлы или элементы.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-798">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="0e1d3-799">String</span><span class="sxs-lookup"><span data-stu-id="0e1d3-799">String</span></span> | | <span data-ttu-id="0e1d3-p150">Указывает тип вложения. Допустимые значения: `file` для вложенного файла и `item` для вложенного элемента.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-p150">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="0e1d3-802">String</span><span class="sxs-lookup"><span data-stu-id="0e1d3-802">String</span></span> | | <span data-ttu-id="0e1d3-803">Строка, содержащая имя вложения, длиной до 255 символов.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-803">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="0e1d3-804">String</span><span class="sxs-lookup"><span data-stu-id="0e1d3-804">String</span></span> | | <span data-ttu-id="0e1d3-p151">Используется, только если свойству `type` задано значение `file`. URI расположения файла.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-p151">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.isInline` | <span data-ttu-id="0e1d3-807">Логическое значение</span><span class="sxs-lookup"><span data-stu-id="0e1d3-807">Boolean</span></span> | | <span data-ttu-id="0e1d3-p152">Используется только в том случае, если свойству `type` задано значение `file`. Значение `true` указывает на то, что вложение будет встроено в текст сообщения и не должно отображаться в списке вложений.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-p152">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="0e1d3-810">String</span><span class="sxs-lookup"><span data-stu-id="0e1d3-810">String</span></span> | | <span data-ttu-id="0e1d3-p153">Используется, только если свойству `type` задано значение `item`. Идентификатор вложения EWS. Это строка длиной до 100 символов.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-p153">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="0e1d3-814">function</span><span class="sxs-lookup"><span data-stu-id="0e1d3-814">function</span></span> | <span data-ttu-id="0e1d3-815">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="0e1d3-815">&lt;optional&gt;</span></span> | <span data-ttu-id="0e1d3-816">По завершении работы метода функция, переданная параметру `callback`, вызывается с помощью одного параметра `asyncResult`, представляющего собой объект [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="0e1d3-816">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="0e1d3-817">Требования</span><span class="sxs-lookup"><span data-stu-id="0e1d3-817">Requirements</span></span>

|<span data-ttu-id="0e1d3-818">Требование</span><span class="sxs-lookup"><span data-stu-id="0e1d3-818">Requirement</span></span>| <span data-ttu-id="0e1d3-819">Значение</span><span class="sxs-lookup"><span data-stu-id="0e1d3-819">Value</span></span>|
|---|---|
|[<span data-ttu-id="0e1d3-820">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="0e1d3-820">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0e1d3-821">1.0</span><span class="sxs-lookup"><span data-stu-id="0e1d3-821">1.0</span></span>|
|[<span data-ttu-id="0e1d3-822">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="0e1d3-822">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0e1d3-823">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0e1d3-823">ReadItem</span></span>|
|[<span data-ttu-id="0e1d3-824">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="0e1d3-824">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="0e1d3-825">Чтение</span><span class="sxs-lookup"><span data-stu-id="0e1d3-825">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="0e1d3-826">Примеры</span><span class="sxs-lookup"><span data-stu-id="0e1d3-826">Examples</span></span>

<span data-ttu-id="0e1d3-827">Приведенный ниже код передает строку в функцию `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-827">The following code passes a string to the `displayReplyForm` function.</span></span>

```
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="0e1d3-828">Ответ с пустым текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-828">Reply with an empty body.</span></span>

```
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="0e1d3-829">Ответ только с текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-829">Reply with just a body.</span></span>

```
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="0e1d3-830">Ответ с текстом сообщения и вложенным файлом.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-830">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="0e1d3-831">Ответ с текстом сообщения и вложенным элементом.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-831">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="0e1d3-832">Ответ с текстом сообщения, вложенным файлом, вложенным элементом и обратным вызовом.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-832">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlook15officeentities"></a><span data-ttu-id="0e1d3-833">getEntities() → {[Entities](/javascript/api/outlook_1_5/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="0e1d3-833">getEntities() → {[Entities](/javascript/api/outlook_1_5/office.entities)}</span></span>

<span data-ttu-id="0e1d3-834">Получает сущности, обнаруженные в выбранном тексте элемента.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-834">Gets the entities found in the selected item.</span></span>

> [!NOTE]
> <span data-ttu-id="0e1d3-835">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-835">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="0e1d3-836">Требования</span><span class="sxs-lookup"><span data-stu-id="0e1d3-836">Requirements</span></span>

|<span data-ttu-id="0e1d3-837">Требование</span><span class="sxs-lookup"><span data-stu-id="0e1d3-837">Requirement</span></span>| <span data-ttu-id="0e1d3-838">Значение</span><span class="sxs-lookup"><span data-stu-id="0e1d3-838">Value</span></span>|
|---|---|
|[<span data-ttu-id="0e1d3-839">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="0e1d3-839">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0e1d3-840">1.0</span><span class="sxs-lookup"><span data-stu-id="0e1d3-840">1.0</span></span>|
|[<span data-ttu-id="0e1d3-841">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="0e1d3-841">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0e1d3-842">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0e1d3-842">ReadItem</span></span>|
|[<span data-ttu-id="0e1d3-843">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="0e1d3-843">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="0e1d3-844">Чтение</span><span class="sxs-lookup"><span data-stu-id="0e1d3-844">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="0e1d3-845">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="0e1d3-845">Returns:</span></span>

<span data-ttu-id="0e1d3-846">Тип: [Entities](/javascript/api/outlook_1_5/office.entities)</span><span class="sxs-lookup"><span data-stu-id="0e1d3-846">Type: [Entities](/javascript/api/outlook_1_5/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="0e1d3-847">Пример</span><span class="sxs-lookup"><span data-stu-id="0e1d3-847">Example</span></span>

<span data-ttu-id="0e1d3-848">Ниже приведен пример получения доступа к сущностям контактов в тексте текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-848">The following example accesses the contacts entities on the current item.</span></span>

```
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook15officecontactmeetingsuggestionjavascriptapioutlook15officemeetingsuggestionphonenumberjavascriptapioutlook15officephonenumbertasksuggestionjavascriptapioutlook15officetasksuggestion"></a><span data-ttu-id="0e1d3-849">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="0e1d3-849">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))>}</span></span>

<span data-ttu-id="0e1d3-850">Получает массив всех сущностей указанного типа, обнаруженных в тексте выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-850">Gets an array of all the entities of the specified entity type found in the selected item.</span></span>

> [!NOTE]
> <span data-ttu-id="0e1d3-851">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-851">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0e1d3-852">Параметры:</span><span class="sxs-lookup"><span data-stu-id="0e1d3-852">Parameters:</span></span>

|<span data-ttu-id="0e1d3-853">Имя</span><span class="sxs-lookup"><span data-stu-id="0e1d3-853">Name</span></span>| <span data-ttu-id="0e1d3-854">Тип</span><span class="sxs-lookup"><span data-stu-id="0e1d3-854">Type</span></span>| <span data-ttu-id="0e1d3-855">Описание</span><span class="sxs-lookup"><span data-stu-id="0e1d3-855">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="0e1d3-856">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="0e1d3-856">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook_1_5/office.mailboxenums.entitytype)|<span data-ttu-id="0e1d3-857">Одно из значений перечисления EntityType.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-857">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0e1d3-858">Требования</span><span class="sxs-lookup"><span data-stu-id="0e1d3-858">Requirements</span></span>

|<span data-ttu-id="0e1d3-859">Требование</span><span class="sxs-lookup"><span data-stu-id="0e1d3-859">Requirement</span></span>| <span data-ttu-id="0e1d3-860">Значение</span><span class="sxs-lookup"><span data-stu-id="0e1d3-860">Value</span></span>|
|---|---|
|[<span data-ttu-id="0e1d3-861">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="0e1d3-861">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0e1d3-862">1.0</span><span class="sxs-lookup"><span data-stu-id="0e1d3-862">1.0</span></span>|
|[<span data-ttu-id="0e1d3-863">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="0e1d3-863">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0e1d3-864">Ограниченный доступ</span><span class="sxs-lookup"><span data-stu-id="0e1d3-864">Restricted</span></span>|
|[<span data-ttu-id="0e1d3-865">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="0e1d3-865">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="0e1d3-866">Чтение</span><span class="sxs-lookup"><span data-stu-id="0e1d3-866">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="0e1d3-867">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="0e1d3-867">Returns:</span></span>

<span data-ttu-id="0e1d3-868">Если значение, переданное в `entityType`, не является допустимым членом перечисления `EntityType`, метод возвращает значение NULL.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-868">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="0e1d3-869">Если в тексте элемента отсутствуют сущности указанного типа, метод возвращает пустой массив.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-869">If no entities of the specified type are present on the item, the method returns an empty array.</span></span> <span data-ttu-id="0e1d3-870">В противном случае тип объектов в возвращаемом массиве зависит от типа сущности, запрошенной в параметре `entityType`.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-870">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="0e1d3-871">Хотя минимальный уровень разрешений для использования этого метода — **Restricted**, для некоторых типов сущностей требуется доступ на уровне **ReadItem**, как указано в приведенной ниже таблице.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-871">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="0e1d3-872">Значение параметра `entityType`</span><span class="sxs-lookup"><span data-stu-id="0e1d3-872">Value of `entityType`</span></span> | <span data-ttu-id="0e1d3-873">Тип объектов в возвращаемом массиве</span><span class="sxs-lookup"><span data-stu-id="0e1d3-873">Type of objects in returned array</span></span> | <span data-ttu-id="0e1d3-874">Необходимый уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="0e1d3-874">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="0e1d3-875">String</span><span class="sxs-lookup"><span data-stu-id="0e1d3-875">String</span></span> | <span data-ttu-id="0e1d3-876">**С ограничениями**</span><span class="sxs-lookup"><span data-stu-id="0e1d3-876">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="0e1d3-877">Contact</span><span class="sxs-lookup"><span data-stu-id="0e1d3-877">Contact</span></span> | <span data-ttu-id="0e1d3-878">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="0e1d3-878">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="0e1d3-879">String</span><span class="sxs-lookup"><span data-stu-id="0e1d3-879">String</span></span> | <span data-ttu-id="0e1d3-880">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="0e1d3-880">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="0e1d3-881">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="0e1d3-881">MeetingSuggestion</span></span> | <span data-ttu-id="0e1d3-882">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="0e1d3-882">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="0e1d3-883">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="0e1d3-883">PhoneNumber</span></span> | <span data-ttu-id="0e1d3-884">**С ограничениями**</span><span class="sxs-lookup"><span data-stu-id="0e1d3-884">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="0e1d3-885">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="0e1d3-885">TaskSuggestion</span></span> | <span data-ttu-id="0e1d3-886">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="0e1d3-886">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="0e1d3-887">String</span><span class="sxs-lookup"><span data-stu-id="0e1d3-887">String</span></span> | <span data-ttu-id="0e1d3-888">**С ограничениями**</span><span class="sxs-lookup"><span data-stu-id="0e1d3-888">**Restricted**</span></span> |

<span data-ttu-id="0e1d3-889">Type: Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="0e1d3-889">Type: Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="0e1d3-890">Пример</span><span class="sxs-lookup"><span data-stu-id="0e1d3-890">Example</span></span>

<span data-ttu-id="0e1d3-891">В примере ниже показано, как получить доступ к массиву строк, которые представляют собой почтовые адреса в тексте текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-891">The following example shows how to access an array of strings that represent postal addresses in the subject or body of the current item.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook15officecontactmeetingsuggestionjavascriptapioutlook15officemeetingsuggestionphonenumberjavascriptapioutlook15officephonenumbertasksuggestionjavascriptapioutlook15officetasksuggestion"></a><span data-ttu-id="0e1d3-892">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="0e1d3-892">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))>}</span></span>

<span data-ttu-id="0e1d3-893">Возвращает известные сущности в выбранном элементе, которые проходят через именованный фильтр, определяемый в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-893">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="0e1d3-894">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-894">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="0e1d3-895">Метод `getFilteredEntitiesByName` возвращает сущности, соответствующие регулярному выражению, которое определяется в элементе правила [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) в XML-файле манифеста, с использованием указанного значения элемента `FilterName`.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-895">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0e1d3-896">Параметры:</span><span class="sxs-lookup"><span data-stu-id="0e1d3-896">Parameters:</span></span>

|<span data-ttu-id="0e1d3-897">Имя</span><span class="sxs-lookup"><span data-stu-id="0e1d3-897">Name</span></span>| <span data-ttu-id="0e1d3-898">Тип</span><span class="sxs-lookup"><span data-stu-id="0e1d3-898">Type</span></span>| <span data-ttu-id="0e1d3-899">Описание</span><span class="sxs-lookup"><span data-stu-id="0e1d3-899">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="0e1d3-900">String</span><span class="sxs-lookup"><span data-stu-id="0e1d3-900">String</span></span>|<span data-ttu-id="0e1d3-901">Имя элемента правила `ItemHasKnownEntity`, определяющее соответствующий фильтр.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-901">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0e1d3-902">Требования</span><span class="sxs-lookup"><span data-stu-id="0e1d3-902">Requirements</span></span>

|<span data-ttu-id="0e1d3-903">Требование</span><span class="sxs-lookup"><span data-stu-id="0e1d3-903">Requirement</span></span>| <span data-ttu-id="0e1d3-904">Значение</span><span class="sxs-lookup"><span data-stu-id="0e1d3-904">Value</span></span>|
|---|---|
|[<span data-ttu-id="0e1d3-905">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="0e1d3-905">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0e1d3-906">1.0</span><span class="sxs-lookup"><span data-stu-id="0e1d3-906">1.0</span></span>|
|[<span data-ttu-id="0e1d3-907">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="0e1d3-907">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0e1d3-908">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0e1d3-908">ReadItem</span></span>|
|[<span data-ttu-id="0e1d3-909">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="0e1d3-909">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="0e1d3-910">Чтение</span><span class="sxs-lookup"><span data-stu-id="0e1d3-910">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="0e1d3-911">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="0e1d3-911">Returns:</span></span>

<span data-ttu-id="0e1d3-p155">Если в манифесте нет элемента `ItemHasKnownEntity` со значением `FilterName`, соответствующим параметру `name`, метод возвращает `null`. Если параметр `name` соответствует элементу `ItemHasKnownEntity` в манифесте, но при этом в текущем элементе нет соответствующих сущностей, метод возвращает пустой массив.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-p155">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="0e1d3-914">Тип: Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="0e1d3-914">Type: Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))></span></span>

#### <a name="getregexmatches--object"></a><span data-ttu-id="0e1d3-915">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="0e1d3-915">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="0e1d3-916">Возвращает строчные значения в выбранном элементе, которые соответствуют регулярным выражениям, определенным в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-916">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="0e1d3-917">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-917">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="0e1d3-p156">Метод `getRegExMatches` возвращает строки, соответствующие регулярному выражению, которое определяется в каждом элементе правила `ItemHasRegularExpressionMatch` или `ItemHasKnownEntity` в XML-файле манифеста. Для правила `ItemHasRegularExpressionMatch` свойство элемента, указанного этим правилом, должно содержать соответствующую строку. Простой тип `PropertyName` определяет поддерживаемые свойства.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-p156">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="0e1d3-921">Например, рассмотрим манифест надстройки, который содержит указанный ниже элемент `Rule`.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-921">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="0e1d3-922">Объект, возвращаемый методом `getRegExMatches`, будет содержать два свойства: `fruits` и `veggies`.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-922">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="0e1d3-p157">Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты. Лучше используйте для этого метод [`Body.getAsync`](/javascript/api/outlook_1_5/office.body#getasync-coerciontype--options--callback-).</span><span class="sxs-lookup"><span data-stu-id="0e1d3-p157">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook_1_5/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="0e1d3-926">Требования</span><span class="sxs-lookup"><span data-stu-id="0e1d3-926">Requirements</span></span>

|<span data-ttu-id="0e1d3-927">Требование</span><span class="sxs-lookup"><span data-stu-id="0e1d3-927">Requirement</span></span>| <span data-ttu-id="0e1d3-928">Значение</span><span class="sxs-lookup"><span data-stu-id="0e1d3-928">Value</span></span>|
|---|---|
|[<span data-ttu-id="0e1d3-929">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="0e1d3-929">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0e1d3-930">1.0</span><span class="sxs-lookup"><span data-stu-id="0e1d3-930">1.0</span></span>|
|[<span data-ttu-id="0e1d3-931">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="0e1d3-931">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0e1d3-932">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0e1d3-932">ReadItem</span></span>|
|[<span data-ttu-id="0e1d3-933">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="0e1d3-933">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="0e1d3-934">Чтение</span><span class="sxs-lookup"><span data-stu-id="0e1d3-934">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="0e1d3-935">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="0e1d3-935">Returns:</span></span>

<span data-ttu-id="0e1d3-p158">Объект, содержащий массив строк, которые соответствуют регулярным выражениям, определяемым в XML-файле манифеста. Имя каждого массива равно соответствующему значению атрибута `RegExName` правила сопоставления `ItemHasRegularExpressionMatch` или атрибута `FilterName` правила сопоставления `ItemHasKnownEntity`.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-p158">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="0e1d3-938">

<dt>Тип</dt>

</span><span class="sxs-lookup"><span data-stu-id="0e1d3-938">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="0e1d3-939">Object</span><span class="sxs-lookup"><span data-stu-id="0e1d3-939">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="0e1d3-940">Пример</span><span class="sxs-lookup"><span data-stu-id="0e1d3-940">Example</span></span>

<span data-ttu-id="0e1d3-941">В примере ниже показано, как получить доступ к массиву совпадений для элементов `fruits` регулярного выражения<rule> и `veggies`, которые указаны в манифесте.</rule></span><span class="sxs-lookup"><span data-stu-id="0e1d3-941">The following example shows how to access the array of matches for the regular expression <rule>elements `fruits` and `veggies`, which are specified in the manifest.</rule></span></span>

```
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veges = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="0e1d3-942">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="0e1d3-942">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="0e1d3-943">Возвращает строчные значения в выбранном элементе, которые соответствуют именованному регулярному выражению, определенному в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-943">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="0e1d3-944">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-944">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="0e1d3-945">Метод `getRegExMatchesByName` возвращает строки, соответствующие регулярному выражению, которое определяется в элементе правила `ItemHasRegularExpressionMatch` в XML-файле манифеста, с использованием указанного значения элемента `RegExName`.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-945">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="0e1d3-p159">Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-p159">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0e1d3-948">Параметры:</span><span class="sxs-lookup"><span data-stu-id="0e1d3-948">Parameters:</span></span>

|<span data-ttu-id="0e1d3-949">Имя</span><span class="sxs-lookup"><span data-stu-id="0e1d3-949">Name</span></span>| <span data-ttu-id="0e1d3-950">Тип</span><span class="sxs-lookup"><span data-stu-id="0e1d3-950">Type</span></span>| <span data-ttu-id="0e1d3-951">Описание</span><span class="sxs-lookup"><span data-stu-id="0e1d3-951">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="0e1d3-952">String</span><span class="sxs-lookup"><span data-stu-id="0e1d3-952">String</span></span>|<span data-ttu-id="0e1d3-953">Имя элемента правила `ItemHasRegularExpressionMatch`, определяющее соответствующий фильтр.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-953">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0e1d3-954">Требования</span><span class="sxs-lookup"><span data-stu-id="0e1d3-954">Requirements</span></span>

|<span data-ttu-id="0e1d3-955">Требование</span><span class="sxs-lookup"><span data-stu-id="0e1d3-955">Requirement</span></span>| <span data-ttu-id="0e1d3-956">Значение</span><span class="sxs-lookup"><span data-stu-id="0e1d3-956">Value</span></span>|
|---|---|
|[<span data-ttu-id="0e1d3-957">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="0e1d3-957">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0e1d3-958">1.0</span><span class="sxs-lookup"><span data-stu-id="0e1d3-958">1.0</span></span>|
|[<span data-ttu-id="0e1d3-959">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="0e1d3-959">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0e1d3-960">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0e1d3-960">ReadItem</span></span>|
|[<span data-ttu-id="0e1d3-961">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="0e1d3-961">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="0e1d3-962">Чтение</span><span class="sxs-lookup"><span data-stu-id="0e1d3-962">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="0e1d3-963">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="0e1d3-963">Returns:</span></span>

<span data-ttu-id="0e1d3-964">Массив строк, соответствующих регулярному выражению, определяемому в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-964">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="0e1d3-965">

<dt>Тип</dt>

</span><span class="sxs-lookup"><span data-stu-id="0e1d3-965">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="0e1d3-966">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="0e1d3-966">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="0e1d3-967">Пример</span><span class="sxs-lookup"><span data-stu-id="0e1d3-967">Example</span></span>

```
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="0e1d3-968">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="0e1d3-968">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="0e1d3-969">Асинхронно возвращает данные, выбранные в теме или тексте сообщения.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-969">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="0e1d3-p160">Если выделенный фрагмент отсутствует, но курсор находится в тексте или теме, метод возвращает значение NULL для выбранных данных. Если выбраны не текст и не тема, метод возвращает ошибку `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-p160">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0e1d3-972">Параметры:</span><span class="sxs-lookup"><span data-stu-id="0e1d3-972">Parameters:</span></span>

|<span data-ttu-id="0e1d3-973">Имя</span><span class="sxs-lookup"><span data-stu-id="0e1d3-973">Name</span></span>| <span data-ttu-id="0e1d3-974">Тип</span><span class="sxs-lookup"><span data-stu-id="0e1d3-974">Type</span></span>| <span data-ttu-id="0e1d3-975">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="0e1d3-975">Attributes</span></span>| <span data-ttu-id="0e1d3-976">Описание</span><span class="sxs-lookup"><span data-stu-id="0e1d3-976">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="0e1d3-977">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="0e1d3-977">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="0e1d3-p161">Запрашивает формат данных. Если задано значение Text, метод возвращает обычный текст как строку, удаляя все имеющиеся HTML-теги. Если задано значение HTML, метод возвращает выделенный текст (обычный текст или HTML).</span><span class="sxs-lookup"><span data-stu-id="0e1d3-p161">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="0e1d3-981">Object</span><span class="sxs-lookup"><span data-stu-id="0e1d3-981">Object</span></span>| <span data-ttu-id="0e1d3-982">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="0e1d3-982">&lt;optional&gt;</span></span>|<span data-ttu-id="0e1d3-983">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-983">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="0e1d3-984">Object</span><span class="sxs-lookup"><span data-stu-id="0e1d3-984">Object</span></span>| <span data-ttu-id="0e1d3-985">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="0e1d3-985">&lt;optional&gt;</span></span>|<span data-ttu-id="0e1d3-986">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-986">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="0e1d3-987">функция</span><span class="sxs-lookup"><span data-stu-id="0e1d3-987">function</span></span>||<span data-ttu-id="0e1d3-988">Когда метод завершает выполнение, переданная в параметре `callback` функция вызывается с единственным параметром `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="0e1d3-988">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="0e1d3-989">Чтобы получить доступ к выбранным данным из метода обратного вызова, вызовите `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-989">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="0e1d3-990">Для доступа к исходному свойству, на основе которого созданы выбранные данные, вызовите  `asyncResult.value.sourceProperty`, который может иметь значение `body` или `subject`.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-990">To access the source property that the selection comes from, call , which will be either  or .|</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0e1d3-991">Требования</span><span class="sxs-lookup"><span data-stu-id="0e1d3-991">Requirements</span></span>

|<span data-ttu-id="0e1d3-992">Требование</span><span class="sxs-lookup"><span data-stu-id="0e1d3-992">Requirement</span></span>| <span data-ttu-id="0e1d3-993">Значение</span><span class="sxs-lookup"><span data-stu-id="0e1d3-993">Value</span></span>|
|---|---|
|[<span data-ttu-id="0e1d3-994">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="0e1d3-994">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0e1d3-995">1.2</span><span class="sxs-lookup"><span data-stu-id="0e1d3-995">1.2</span></span>|
|[<span data-ttu-id="0e1d3-996">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="0e1d3-996">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0e1d3-997">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="0e1d3-997">ReadWriteItem</span></span>|
|[<span data-ttu-id="0e1d3-998">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="0e1d3-998">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="0e1d3-999">Создание</span><span class="sxs-lookup"><span data-stu-id="0e1d3-999">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="0e1d3-1000">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="0e1d3-1000">Returns:</span></span>

<span data-ttu-id="0e1d3-1001">Выбранные данные в виде строки с форматом, определенным в параметре `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-1001">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="0e1d3-1002">

<dt>Тип</dt>

</span><span class="sxs-lookup"><span data-stu-id="0e1d3-1002">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="0e1d3-1003">String</span><span class="sxs-lookup"><span data-stu-id="0e1d3-1003">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="0e1d3-1004">Пример</span><span class="sxs-lookup"><span data-stu-id="0e1d3-1004">Example</span></span>

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

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="0e1d3-1005">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="0e1d3-1005">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="0e1d3-1006">Асинхронно загружает настраиваемые свойства для надстройки выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-1006">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="0e1d3-p163">Настраиваемые свойства сохраняются в виде пар "ключ-значение" для каждого приложения и каждого элемента. Этот метод возвращает объект `CustomProperties` при обратном вызове, который предоставляет методы для доступа к настраиваемым свойствам, характерным для текущего элемента и текущей надстройки. Настраиваемые свойства не шифруются для элемента, поэтому этот способ хранения не является безопасным.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-p163">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0e1d3-1010">Параметры:</span><span class="sxs-lookup"><span data-stu-id="0e1d3-1010">Parameters:</span></span>

|<span data-ttu-id="0e1d3-1011">Имя</span><span class="sxs-lookup"><span data-stu-id="0e1d3-1011">Name</span></span>| <span data-ttu-id="0e1d3-1012">Тип</span><span class="sxs-lookup"><span data-stu-id="0e1d3-1012">Type</span></span>| <span data-ttu-id="0e1d3-1013">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="0e1d3-1013">Attributes</span></span>| <span data-ttu-id="0e1d3-1014">Описание</span><span class="sxs-lookup"><span data-stu-id="0e1d3-1014">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="0e1d3-1015">function</span><span class="sxs-lookup"><span data-stu-id="0e1d3-1015">function</span></span>||<span data-ttu-id="0e1d3-1016">Когда метод завершает выполнение, переданная в параметре `callback` функция вызывается с единственным параметром `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="0e1d3-1016">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="0e1d3-1017">Настраиваемые свойства предоставляются в виде объекта [`CustomProperties`](/javascript/api/outlook_1_5/office.customproperties) в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-1017">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook_1_5/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="0e1d3-1018">Этот объект можно использовать для получения, задания и удаления настраиваемых свойств из элемента и сохранения изменений настраиваемого свойства на сервере.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-1018">The custom properties are provided as a CustomProperties object in the asyncResult.value property. This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="0e1d3-1019">Object</span><span class="sxs-lookup"><span data-stu-id="0e1d3-1019">Object</span></span>| <span data-ttu-id="0e1d3-1020">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="0e1d3-1020">&lt;optional&gt;</span></span>|<span data-ttu-id="0e1d3-1021">Разработчики могут указать любой объект, к которому необходимо получить доступ, в функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-1021">Developers can provide any object they wish to access in the callback method.</span></span> <span data-ttu-id="0e1d3-1022">Доступ к этому объекту можно получить с помощью свойства `asyncResult.asyncContext` в функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-1022">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.|</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0e1d3-1023">Требования</span><span class="sxs-lookup"><span data-stu-id="0e1d3-1023">Requirements</span></span>

|<span data-ttu-id="0e1d3-1024">Требование</span><span class="sxs-lookup"><span data-stu-id="0e1d3-1024">Requirement</span></span>| <span data-ttu-id="0e1d3-1025">Значение</span><span class="sxs-lookup"><span data-stu-id="0e1d3-1025">Value</span></span>|
|---|---|
|[<span data-ttu-id="0e1d3-1026">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="0e1d3-1026">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0e1d3-1027">1.0</span><span class="sxs-lookup"><span data-stu-id="0e1d3-1027">1.0</span></span>|
|[<span data-ttu-id="0e1d3-1028">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="0e1d3-1028">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0e1d3-1029">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0e1d3-1029">ReadItem</span></span>|
|[<span data-ttu-id="0e1d3-1030">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="0e1d3-1030">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="0e1d3-1031">Cоздание или чтение</span><span class="sxs-lookup"><span data-stu-id="0e1d3-1031">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="0e1d3-1032">Пример</span><span class="sxs-lookup"><span data-stu-id="0e1d3-1032">Example</span></span>

<span data-ttu-id="0e1d3-p166">В приведенном ниже примере кода показано, как асинхронно загружать настраиваемые свойства, характерные для текущего элемента, с помощью метода `loadCustomPropertiesAsync`. Этот пример также показывает, как сохранять эти свойства на сервере с помощью метода `CustomProperties.saveAsync`. В этом примере кода, после того как выполнена загрузка настраиваемых свойств, метод `CustomProperties.get` используется для считывания настраиваемого свойства `myProp`, метод `CustomProperties.set` — для записи настраиваемого свойства `otherProp`, а метод `saveAsync` — для сохранения настраиваемых свойств.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-p166">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="0e1d3-1036">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="0e1d3-1036">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="0e1d3-1037">Удаляет вложение из сообщения или встречи.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-1037">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="0e1d3-p167">Метод `removeAttachmentAsync` удаляет из элемента вложение с указанным идентификатором. Идентификатор вложения рекомендуется использовать для удаления вложения, только если оно добавлено тем же почтовым приложением в ходе текущего сеанса. В веб-приложении Outlook и веб-приложении Outlook для устройств идентификатор вложения действителен только в рамках одного сеанса. Сеанс завершается, когда пользователь закрывает приложение или начинает создавать элемент во встроенной форме, а затем переходит из формы в отдельное окно.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-p167">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0e1d3-1042">Параметры:</span><span class="sxs-lookup"><span data-stu-id="0e1d3-1042">Parameters:</span></span>

|<span data-ttu-id="0e1d3-1043">Имя</span><span class="sxs-lookup"><span data-stu-id="0e1d3-1043">Name</span></span>| <span data-ttu-id="0e1d3-1044">Тип</span><span class="sxs-lookup"><span data-stu-id="0e1d3-1044">Type</span></span>| <span data-ttu-id="0e1d3-1045">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="0e1d3-1045">Attributes</span></span>| <span data-ttu-id="0e1d3-1046">Описание</span><span class="sxs-lookup"><span data-stu-id="0e1d3-1046">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="0e1d3-1047">String</span><span class="sxs-lookup"><span data-stu-id="0e1d3-1047">String</span></span>||<span data-ttu-id="0e1d3-p168">Идентификатор удаляемого вложения. Максимальная длина строки — 100 символов.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-p168">The identifier of the attachment to remove. The maximum length of the string is 100 characters.</span></span>|
|`options`| <span data-ttu-id="0e1d3-1050">Object</span><span class="sxs-lookup"><span data-stu-id="0e1d3-1050">Object</span></span>| <span data-ttu-id="0e1d3-1051">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="0e1d3-1051">&lt;optional&gt;</span></span>|<span data-ttu-id="0e1d3-1052">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-1052">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="0e1d3-1053">Object</span><span class="sxs-lookup"><span data-stu-id="0e1d3-1053">Object</span></span>| <span data-ttu-id="0e1d3-1054">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="0e1d3-1054">&lt;optional&gt;</span></span>|<span data-ttu-id="0e1d3-1055">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-1055">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="0e1d3-1056">function</span><span class="sxs-lookup"><span data-stu-id="0e1d3-1056">function</span></span>| <span data-ttu-id="0e1d3-1057">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="0e1d3-1057">&lt;optional&gt;</span></span>|<span data-ttu-id="0e1d3-1058">Когда метод завершает выполнение, переданная в параметре `callback` функция вызывается с единственным параметром `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="0e1d3-1058">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="0e1d3-1059">Если удалить вложение не удается, свойство `asyncResult.error` содержит код ошибки с указанием ее причины.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-1059">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="0e1d3-1060">Ошибки</span><span class="sxs-lookup"><span data-stu-id="0e1d3-1060">Errors</span></span>

| <span data-ttu-id="0e1d3-1061">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="0e1d3-1061">Error code</span></span> | <span data-ttu-id="0e1d3-1062">Описание</span><span class="sxs-lookup"><span data-stu-id="0e1d3-1062">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="0e1d3-1063">Идентификатор вложения не существует.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-1063">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="0e1d3-1064">Требования</span><span class="sxs-lookup"><span data-stu-id="0e1d3-1064">Requirements</span></span>

|<span data-ttu-id="0e1d3-1065">Требование</span><span class="sxs-lookup"><span data-stu-id="0e1d3-1065">Requirement</span></span>| <span data-ttu-id="0e1d3-1066">Значение</span><span class="sxs-lookup"><span data-stu-id="0e1d3-1066">Value</span></span>|
|---|---|
|[<span data-ttu-id="0e1d3-1067">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="0e1d3-1067">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0e1d3-1068">1.1</span><span class="sxs-lookup"><span data-stu-id="0e1d3-1068">1.1</span></span>|
|[<span data-ttu-id="0e1d3-1069">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="0e1d3-1069">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0e1d3-1070">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="0e1d3-1070">ReadWriteItem</span></span>|
|[<span data-ttu-id="0e1d3-1071">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="0e1d3-1071">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="0e1d3-1072">Создание</span><span class="sxs-lookup"><span data-stu-id="0e1d3-1072">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="0e1d3-1073">Пример</span><span class="sxs-lookup"><span data-stu-id="0e1d3-1073">Example</span></span>

<span data-ttu-id="0e1d3-1074">Указанный ниже код удаляет вложение с идентификатором "0".</span><span class="sxs-lookup"><span data-stu-id="0e1d3-1074">The following code removes an attachment with an identifier of '0'.</span></span>

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

####  <a name="saveasyncoptions-callback"></a><span data-ttu-id="0e1d3-1075">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="0e1d3-1075">saveAsync([options], callback)</span></span>

<span data-ttu-id="0e1d3-1076">Асинхронно сохраняет элемент.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-1076">Asynchronously saves an item.</span></span>

<span data-ttu-id="0e1d3-p169">При вызове этот метод сохраняет текущее сообщение в виде черновика и возвращает идентификатор элемента с помощью метода обратного вызова. В веб-приложернии Outlook или интерактивном режиме Outlook этот элемент сохраняется на сервере. В Outlook в режиме кэширования этот элемент сохраняется в локальном кэше.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-p169">When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="0e1d3-1080">Если в режиме создания надстройка вызывает для элемента метод `saveAsync`, чтобы получить параметр `itemId` для использования с EWS или REST API, необходимо помнить, что синхронизация элемента с сервером может занять много времени, если Outlook работает в режиме кэширования данных.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-1080">Note: If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server. Until the item is synced, using the  will return an error.</span></span> <span data-ttu-id="0e1d3-1081">До окончания синхронизации применение параметра `itemId` будет приводить к ошибке.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-1081">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="0e1d3-p171">Так как для встреч не предусмотрено состояние черновика, если `saveAsync` вызывается для встречи в режиме создания, она сохраняется как обычная встреча в календаре пользователя. При сохранении новой встречи приглашения не отправляются. При сохранении существующей встречи уведомления отправляются добавленным или удаленным участникам.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-p171">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="0e1d3-1085">Следующие клиенты имеют разную реакцию на событие для `saveAsync` для встреч в режиме создания:</span><span class="sxs-lookup"><span data-stu-id="0e1d3-1085">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="0e1d3-1086">Mac Outlook не поддерживает `saveAsync` на собрании в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-1086">Note: Mac Outlook does not support `saveAsync` on a meeting in compose mode. Calling  on a meeting in Mac Outlook will return an error.</span></span> <span data-ttu-id="0e1d3-1087">Вызов `saveAsync` на собрании в Mac Outlook возвращает ошибку.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-1087">Note: Mac Outlook does not support  on a meeting in compose mode. Calling `saveAsync` on a meeting in Mac Outlook will return an error.</span></span>
> - <span data-ttu-id="0e1d3-1088">Outlook в Интернете всегда отправляет приглашение или обновления при вызове `saveAsync` на встрече в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-1088">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0e1d3-1089">Параметры:</span><span class="sxs-lookup"><span data-stu-id="0e1d3-1089">Parameters:</span></span>

|<span data-ttu-id="0e1d3-1090">Имя</span><span class="sxs-lookup"><span data-stu-id="0e1d3-1090">Name</span></span>| <span data-ttu-id="0e1d3-1091">Тип</span><span class="sxs-lookup"><span data-stu-id="0e1d3-1091">Type</span></span>| <span data-ttu-id="0e1d3-1092">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="0e1d3-1092">Attributes</span></span>| <span data-ttu-id="0e1d3-1093">Описание</span><span class="sxs-lookup"><span data-stu-id="0e1d3-1093">Description</span></span>|
|---|---|---|---|
|`options`| <span data-ttu-id="0e1d3-1094">Object</span><span class="sxs-lookup"><span data-stu-id="0e1d3-1094">Object</span></span>| <span data-ttu-id="0e1d3-1095">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="0e1d3-1095">&lt;optional&gt;</span></span>|<span data-ttu-id="0e1d3-1096">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-1096">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="0e1d3-1097">Object</span><span class="sxs-lookup"><span data-stu-id="0e1d3-1097">Object</span></span>| <span data-ttu-id="0e1d3-1098">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="0e1d3-1098">&lt;optional&gt;</span></span>|<span data-ttu-id="0e1d3-1099">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-1099">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="0e1d3-1100">функция</span><span class="sxs-lookup"><span data-stu-id="0e1d3-1100">function</span></span>||<span data-ttu-id="0e1d3-1101">Когда метод завершает выполнение, переданная в параметре `callback` функция вызывается с единственным параметром `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="0e1d3-1101">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="0e1d3-1102">После успешного выполнения идентификатор элемента будет предоставлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-1102">On success, the item identifier is provided in the `asyncResult.value` property.|</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0e1d3-1103">Требования</span><span class="sxs-lookup"><span data-stu-id="0e1d3-1103">Requirements</span></span>

|<span data-ttu-id="0e1d3-1104">Требование</span><span class="sxs-lookup"><span data-stu-id="0e1d3-1104">Requirement</span></span>| <span data-ttu-id="0e1d3-1105">Значение</span><span class="sxs-lookup"><span data-stu-id="0e1d3-1105">Value</span></span>|
|---|---|
|[<span data-ttu-id="0e1d3-1106">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="0e1d3-1106">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0e1d3-1107">1.3</span><span class="sxs-lookup"><span data-stu-id="0e1d3-1107">1.3</span></span>|
|[<span data-ttu-id="0e1d3-1108">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="0e1d3-1108">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0e1d3-1109">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="0e1d3-1109">ReadWriteItem</span></span>|
|[<span data-ttu-id="0e1d3-1110">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="0e1d3-1110">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="0e1d3-1111">Создание</span><span class="sxs-lookup"><span data-stu-id="0e1d3-1111">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="0e1d3-1112">Примеры</span><span class="sxs-lookup"><span data-stu-id="0e1d3-1112">Examples</span></span>

```
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result
  });
```

<span data-ttu-id="0e1d3-p173">Ниже приведен пример параметра `result`, переданного функции обратного вызова. Свойство `value` содержит идентификатор элемента.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-p173">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

####  <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="0e1d3-1115">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="0e1d3-1115">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="0e1d3-1116">Асинхронно вставляет данные в текст или тему сообщения.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-1116">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="0e1d3-p174">Метод `setSelectedDataAsync` вставляет указанную строку в местоположение курсора в теме или тексте элемента либо, если текст выделен в редакторе, он заменяет выделенный текст. Если курсор находится вне текста или темы элемента, возвращается ошибка. После вставки курсор помещается в конец вставленного содержимого.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-p174">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0e1d3-1120">Параметры:</span><span class="sxs-lookup"><span data-stu-id="0e1d3-1120">Parameters:</span></span>

|<span data-ttu-id="0e1d3-1121">Имя</span><span class="sxs-lookup"><span data-stu-id="0e1d3-1121">Name</span></span>| <span data-ttu-id="0e1d3-1122">Тип</span><span class="sxs-lookup"><span data-stu-id="0e1d3-1122">Type</span></span>| <span data-ttu-id="0e1d3-1123">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="0e1d3-1123">Attributes</span></span>| <span data-ttu-id="0e1d3-1124">Описание</span><span class="sxs-lookup"><span data-stu-id="0e1d3-1124">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="0e1d3-1125">String</span><span class="sxs-lookup"><span data-stu-id="0e1d3-1125">String</span></span>||<span data-ttu-id="0e1d3-p175">Вставляемые данные. Объем данных не должен превышать 1 000 000 символов. Если передано больше 1 000 000 символов, возвращается исключение `ArgumentOutOfRange`.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-p175">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="0e1d3-1129">Object</span><span class="sxs-lookup"><span data-stu-id="0e1d3-1129">Object</span></span>| <span data-ttu-id="0e1d3-1130">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="0e1d3-1130">&lt;optional&gt;</span></span>|<span data-ttu-id="0e1d3-1131">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-1131">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="0e1d3-1132">Object</span><span class="sxs-lookup"><span data-stu-id="0e1d3-1132">Object</span></span>| <span data-ttu-id="0e1d3-1133">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="0e1d3-1133">&lt;optional&gt;</span></span>|<span data-ttu-id="0e1d3-1134">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-1134">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`| [<span data-ttu-id="0e1d3-1135">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="0e1d3-1135">Office.CoercionType</span></span>](office.md#coerciontype-string)| <span data-ttu-id="0e1d3-1136">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="0e1d3-1136">&lt;optional&gt;</span></span>|<span data-ttu-id="0e1d3-p176">Если задано значение `text`, текущий стиль применяется в Outlook и веб-приложении Outlook. Если поле представляет собой редактор HTML, вставляются только текстовые данные, даже если они имеют формат HTML.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-p176">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="0e1d3-p177">Если `html` и поле поддерживают HTML (а тема не поддерживает), в веб-приложении Outlook применяется текущий стиль, а в Outlook — стиль по умолчанию. Если поле является текстовым, возвращается ошибка `InvalidDataFormat`.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-p177">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="0e1d3-1141">Если тип `coercionType` не установлен, результат зависит от поля: если поле имеет формат HTML, то используется HTML; если поле является текстовым, то используется обычный текст.</span><span class="sxs-lookup"><span data-stu-id="0e1d3-1141">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="0e1d3-1142">function</span><span class="sxs-lookup"><span data-stu-id="0e1d3-1142">function</span></span>||<span data-ttu-id="0e1d3-1143">Когда метод завершает выполнение, переданная в параметре `callback` функция вызывается с единственным параметром `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="0e1d3-1143">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="0e1d3-1144">Требования</span><span class="sxs-lookup"><span data-stu-id="0e1d3-1144">Requirements</span></span>

|<span data-ttu-id="0e1d3-1145">Требование</span><span class="sxs-lookup"><span data-stu-id="0e1d3-1145">Requirement</span></span>| <span data-ttu-id="0e1d3-1146">Значение</span><span class="sxs-lookup"><span data-stu-id="0e1d3-1146">Value</span></span>|
|---|---|
|[<span data-ttu-id="0e1d3-1147">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="0e1d3-1147">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0e1d3-1148">1.2</span><span class="sxs-lookup"><span data-stu-id="0e1d3-1148">1.2</span></span>|
|[<span data-ttu-id="0e1d3-1149">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="0e1d3-1149">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0e1d3-1150">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="0e1d3-1150">ReadWriteItem</span></span>|
|[<span data-ttu-id="0e1d3-1151">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="0e1d3-1151">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="0e1d3-1152">Создание</span><span class="sxs-lookup"><span data-stu-id="0e1d3-1152">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="0e1d3-1153">Пример</span><span class="sxs-lookup"><span data-stu-id="0e1d3-1153">Example</span></span>

```
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```