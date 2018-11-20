
# <a name="item"></a><span data-ttu-id="ec215-101">item</span><span class="sxs-lookup"><span data-stu-id="ec215-101">item</span></span>

### <span data-ttu-id="ec215-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md). item</span><span class="sxs-lookup"><span data-stu-id="ec215-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md). item</span></span>

<span data-ttu-id="ec215-p102">Пространство имен `item` используется для доступа к выбранному в данный момент сообщению, приглашению на собрание или описанию встречи. Вы можете определить тип пространства имен `item` с помощью свойства [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook11officemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="ec215-p102">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook11officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="ec215-106">Requirements</span><span class="sxs-lookup"><span data-stu-id="ec215-106">Requirements</span></span>

|<span data-ttu-id="ec215-107">Requirement</span><span class="sxs-lookup"><span data-stu-id="ec215-107">Requirement</span></span>| <span data-ttu-id="ec215-108">Значение</span><span class="sxs-lookup"><span data-stu-id="ec215-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="ec215-109">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="ec215-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ec215-110">1.0</span><span class="sxs-lookup"><span data-stu-id="ec215-110">1.0</span></span>|
|[<span data-ttu-id="ec215-111">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ec215-111">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ec215-112">Restricted</span><span class="sxs-lookup"><span data-stu-id="ec215-112">Restricted</span></span>|
|[<span data-ttu-id="ec215-113">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ec215-113">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ec215-114">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="ec215-114">Compose or read</span></span>|

### <a name="example"></a><span data-ttu-id="ec215-115">Пример</span><span class="sxs-lookup"><span data-stu-id="ec215-115">Example</span></span>

<span data-ttu-id="ec215-116">В примере кода JavaScript, приведенном ниже, показано, как получить доступ к свойству `subject` текущего элемента в Outlook.</span><span class="sxs-lookup"><span data-stu-id="ec215-116">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

```JavaScript
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

### <a name="members"></a><span data-ttu-id="ec215-117">Элементы</span><span class="sxs-lookup"><span data-stu-id="ec215-117">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook11officeattachmentdetails"></a><span data-ttu-id="ec215-118">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_1/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="ec215-118">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_1/office.attachmentdetails)></span></span>

<span data-ttu-id="ec215-p103">Получает массив вложений для элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="ec215-p103">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="ec215-121">Outlook блокирует определенные типы файлов из-за потенциальных проблем с безопасностью, поэтому они не возвращаются.</span><span class="sxs-lookup"><span data-stu-id="ec215-121">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="ec215-122">Дополнительные сведения см. в статье [Блокировка вложений в Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="ec215-122">For more information see [Payments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="ec215-123">Тип:</span><span class="sxs-lookup"><span data-stu-id="ec215-123">Type:</span></span>

*   <span data-ttu-id="ec215-124">Array.<[AttachmentDetails](/javascript/api/outlook_1_1/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="ec215-124">Array.<[AttachmentDetails](/javascript/api/outlook_1_1/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="ec215-125">Требования</span><span class="sxs-lookup"><span data-stu-id="ec215-125">Requirements</span></span>

|<span data-ttu-id="ec215-126">Requirement</span><span class="sxs-lookup"><span data-stu-id="ec215-126">Requirement</span></span>| <span data-ttu-id="ec215-127">Значение</span><span class="sxs-lookup"><span data-stu-id="ec215-127">Value</span></span>|
|---|---|
|[<span data-ttu-id="ec215-128">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="ec215-128">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ec215-129">1.0</span><span class="sxs-lookup"><span data-stu-id="ec215-129">1.0</span></span>|
|[<span data-ttu-id="ec215-130">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ec215-130">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ec215-131">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ec215-131">ReadItem</span></span>|
|[<span data-ttu-id="ec215-132">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ec215-132">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ec215-133">Чтение</span><span class="sxs-lookup"><span data-stu-id="ec215-133">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ec215-134">Пример</span><span class="sxs-lookup"><span data-stu-id="ec215-134">Example</span></span>

<span data-ttu-id="ec215-135">С помощью приведенного ниже кода можно создать HTML-строку с подробными сведениями обо всех вложениях для текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="ec215-135">The following code builds an HTML string with details of all attachments on the current item.</span></span>

```JavaScript
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

####  <a name="bcc-recipientsjavascriptapioutlook11officerecipients"></a><span data-ttu-id="ec215-136">bcc :[Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="ec215-136">bcc :[Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

<span data-ttu-id="ec215-137">Получает объект, который предоставляет методы для получения или обновления получателей скрытой копии сообщения.</span><span class="sxs-lookup"><span data-stu-id="ec215-137">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="ec215-138">Только режим создания.</span><span class="sxs-lookup"><span data-stu-id="ec215-138">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="ec215-139">Тип:</span><span class="sxs-lookup"><span data-stu-id="ec215-139">Type:</span></span>

*   [<span data-ttu-id="ec215-140">Recipients</span><span class="sxs-lookup"><span data-stu-id="ec215-140">Recipients</span></span>](/javascript/api/outlook_1_1/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="ec215-141">Требования</span><span class="sxs-lookup"><span data-stu-id="ec215-141">Requirements</span></span>

|<span data-ttu-id="ec215-142">Requirement</span><span class="sxs-lookup"><span data-stu-id="ec215-142">Requirement</span></span>| <span data-ttu-id="ec215-143">Значение</span><span class="sxs-lookup"><span data-stu-id="ec215-143">Value</span></span>|
|---|---|
|[<span data-ttu-id="ec215-144">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="ec215-144">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ec215-145">1.1</span><span class="sxs-lookup"><span data-stu-id="ec215-145">1.1</span></span>|
|[<span data-ttu-id="ec215-146">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ec215-146">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ec215-147">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ec215-147">ReadItem</span></span>|
|[<span data-ttu-id="ec215-148">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ec215-148">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ec215-149">Создание</span><span class="sxs-lookup"><span data-stu-id="ec215-149">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="ec215-150">Пример</span><span class="sxs-lookup"><span data-stu-id="ec215-150">Example</span></span>

```JavaScript
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlook11officebody"></a><span data-ttu-id="ec215-151">body :[Body](/javascript/api/outlook_1_1/office.body)</span><span class="sxs-lookup"><span data-stu-id="ec215-151">body :[Body](/javascript/api/outlook_1_1/office.body)</span></span>

<span data-ttu-id="ec215-152">Получает объект, предоставляющий методы для работы с основным текстом элемента.</span><span class="sxs-lookup"><span data-stu-id="ec215-152">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="ec215-153">Тип:</span><span class="sxs-lookup"><span data-stu-id="ec215-153">Type:</span></span>

*   [<span data-ttu-id="ec215-154">Body</span><span class="sxs-lookup"><span data-stu-id="ec215-154">Body</span></span>](/javascript/api/outlook_1_1/office.body)

##### <a name="requirements"></a><span data-ttu-id="ec215-155">Требования</span><span class="sxs-lookup"><span data-stu-id="ec215-155">Requirements</span></span>

|<span data-ttu-id="ec215-156">Requirement</span><span class="sxs-lookup"><span data-stu-id="ec215-156">Requirement</span></span>| <span data-ttu-id="ec215-157">Значение</span><span class="sxs-lookup"><span data-stu-id="ec215-157">Value</span></span>|
|---|---|
|[<span data-ttu-id="ec215-158">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="ec215-158">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ec215-159">1.1</span><span class="sxs-lookup"><span data-stu-id="ec215-159">1.1</span></span>|
|[<span data-ttu-id="ec215-160">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ec215-160">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ec215-161">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ec215-161">ReadItem</span></span>|
|[<span data-ttu-id="ec215-162">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ec215-162">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ec215-163">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="ec215-163">Compose or read</span></span>|

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlook11officeemailaddressdetailsrecipientsjavascriptapioutlook11officerecipients"></a><span data-ttu-id="ec215-164">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="ec215-164">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

<span data-ttu-id="ec215-165">Предоставляет доступ к получателям копии сообщения.</span><span class="sxs-lookup"><span data-stu-id="ec215-165">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="ec215-166">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="ec215-166">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="ec215-167">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="ec215-167">Read mode</span></span>

<span data-ttu-id="ec215-p107">Свойство `cc` возвращает массив, который содержит объект `EmailAddressDetails` для каждого получателя, указанного в строке **Копия** сообщения. Коллекция может включать не более 100 элементов.</span><span class="sxs-lookup"><span data-stu-id="ec215-p107">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="ec215-170">Режим создания</span><span class="sxs-lookup"><span data-stu-id="ec215-170">Compose mode</span></span>

<span data-ttu-id="ec215-171">Свойство `cc` возвращает объект `Recipients`, предоставляющий методы для получения или обновления получателей, которые указаны в строке **Копия** сообщения.</span><span class="sxs-lookup"><span data-stu-id="ec215-171">The `cc` property returns a `Recipients` object that provides methods for manipulating the recipients on the **Cc** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="ec215-172">Тип:</span><span class="sxs-lookup"><span data-stu-id="ec215-172">Type:</span></span>

*   <span data-ttu-id="ec215-173">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="ec215-173">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="ec215-174">Требования</span><span class="sxs-lookup"><span data-stu-id="ec215-174">Requirements</span></span>

|<span data-ttu-id="ec215-175">Requirement</span><span class="sxs-lookup"><span data-stu-id="ec215-175">Requirement</span></span>| <span data-ttu-id="ec215-176">Значение</span><span class="sxs-lookup"><span data-stu-id="ec215-176">Value</span></span>|
|---|---|
|[<span data-ttu-id="ec215-177">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="ec215-177">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ec215-178">1.0</span><span class="sxs-lookup"><span data-stu-id="ec215-178">1.0</span></span>|
|[<span data-ttu-id="ec215-179">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ec215-179">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ec215-180">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ec215-180">ReadItem</span></span>|
|[<span data-ttu-id="ec215-181">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ec215-181">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ec215-182">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="ec215-182">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="ec215-183">Пример</span><span class="sxs-lookup"><span data-stu-id="ec215-183">Example</span></span>

```JavaScript
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="ec215-184">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="ec215-184">(nullable) conversationId :String</span></span>

<span data-ttu-id="ec215-185">Получает идентификатор разговора по электронной почте, содержащего конкретное сообщение.</span><span class="sxs-lookup"><span data-stu-id="ec215-185">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="ec215-p108">Вы можете получить целочисленное значение этого свойства, если ваше почтовое приложение активируется в формах просмотра или формах создания ответов. Если пользователь изменит тему ответа, после его отправки идентификатор беседы будет изменен, и полученное ранее значение будет недействительным.</span><span class="sxs-lookup"><span data-stu-id="ec215-p108">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="ec215-p109">Это свойство имеет значение NULL для нового элемента в форме создания. Свойство `conversationId` вернет значение, если пользователь задаст тему и сохранит элемент.</span><span class="sxs-lookup"><span data-stu-id="ec215-p109">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="ec215-190">Тип:</span><span class="sxs-lookup"><span data-stu-id="ec215-190">Type:</span></span>

*   <span data-ttu-id="ec215-191">String</span><span class="sxs-lookup"><span data-stu-id="ec215-191">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="ec215-192">Требования</span><span class="sxs-lookup"><span data-stu-id="ec215-192">Requirements</span></span>

|<span data-ttu-id="ec215-193">Requirement</span><span class="sxs-lookup"><span data-stu-id="ec215-193">Requirement</span></span>| <span data-ttu-id="ec215-194">Значение</span><span class="sxs-lookup"><span data-stu-id="ec215-194">Value</span></span>|
|---|---|
|[<span data-ttu-id="ec215-195">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="ec215-195">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ec215-196">1.0</span><span class="sxs-lookup"><span data-stu-id="ec215-196">1.0</span></span>|
|[<span data-ttu-id="ec215-197">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ec215-197">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ec215-198">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ec215-198">ReadItem</span></span>|
|[<span data-ttu-id="ec215-199">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ec215-199">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ec215-200">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="ec215-200">Compose or read</span></span>|

#### <a name="datetimecreated-date"></a><span data-ttu-id="ec215-201">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="ec215-201">dateTimeCreated :Date</span></span>

<span data-ttu-id="ec215-p110">Получает дату и время создания элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="ec215-p110">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="ec215-204">Тип:</span><span class="sxs-lookup"><span data-stu-id="ec215-204">Type:</span></span>

*   <span data-ttu-id="ec215-205">Date</span><span class="sxs-lookup"><span data-stu-id="ec215-205">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="ec215-206">Требования</span><span class="sxs-lookup"><span data-stu-id="ec215-206">Requirements</span></span>

|<span data-ttu-id="ec215-207">Requirement</span><span class="sxs-lookup"><span data-stu-id="ec215-207">Requirement</span></span>| <span data-ttu-id="ec215-208">Значение</span><span class="sxs-lookup"><span data-stu-id="ec215-208">Value</span></span>|
|---|---|
|[<span data-ttu-id="ec215-209">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="ec215-209">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ec215-210">1.0</span><span class="sxs-lookup"><span data-stu-id="ec215-210">1.0</span></span>|
|[<span data-ttu-id="ec215-211">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ec215-211">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ec215-212">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ec215-212">ReadItem</span></span>|
|[<span data-ttu-id="ec215-213">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ec215-213">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ec215-214">Чтение</span><span class="sxs-lookup"><span data-stu-id="ec215-214">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ec215-215">Пример</span><span class="sxs-lookup"><span data-stu-id="ec215-215">Example</span></span>

```JavaScript
var created = Office.context.mailbox.item.dateTimeCreated;
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="ec215-216">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="ec215-216">dateTimeModified :Date</span></span>

<span data-ttu-id="ec215-p111">Получает дату и время последнего изменения элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="ec215-p111">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="ec215-219">Этот элемент не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="ec215-219">Note: This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="ec215-220">Тип:</span><span class="sxs-lookup"><span data-stu-id="ec215-220">Type:</span></span>

*   <span data-ttu-id="ec215-221">Date</span><span class="sxs-lookup"><span data-stu-id="ec215-221">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="ec215-222">Требования</span><span class="sxs-lookup"><span data-stu-id="ec215-222">Requirements</span></span>

|<span data-ttu-id="ec215-223">Requirement</span><span class="sxs-lookup"><span data-stu-id="ec215-223">Requirement</span></span>| <span data-ttu-id="ec215-224">Значение</span><span class="sxs-lookup"><span data-stu-id="ec215-224">Value</span></span>|
|---|---|
|[<span data-ttu-id="ec215-225">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="ec215-225">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ec215-226">1.0</span><span class="sxs-lookup"><span data-stu-id="ec215-226">1.0</span></span>|
|[<span data-ttu-id="ec215-227">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ec215-227">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ec215-228">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ec215-228">ReadItem</span></span>|
|[<span data-ttu-id="ec215-229">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ec215-229">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ec215-230">Чтение</span><span class="sxs-lookup"><span data-stu-id="ec215-230">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ec215-231">Пример</span><span class="sxs-lookup"><span data-stu-id="ec215-231">Example</span></span>

```JavaScript
var modified = Office.context.mailbox.item.dateTimeModified;
```

####  <a name="end-datetimejavascriptapioutlook11officetime"></a><span data-ttu-id="ec215-232">end :Date|[Time](/javascript/api/outlook_1_1/office.time)</span><span class="sxs-lookup"><span data-stu-id="ec215-232">end :Date|[Time](/javascript/api/outlook_1_1/office.time)</span></span>

<span data-ttu-id="ec215-233">Получает или задает дату и время окончания встречи.</span><span class="sxs-lookup"><span data-stu-id="ec215-233">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="ec215-p112">Свойство `end` представлено в виде значения даты и времени в формате UTC. Преобразовать значение свойства end в местные значения даты и времени клиента можно с помощью метода [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook11officelocalclienttime).</span><span class="sxs-lookup"><span data-stu-id="ec215-p112">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook11officelocalclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="ec215-236">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="ec215-236">Read mode</span></span>

<span data-ttu-id="ec215-237">Свойство `end` возвращает объект `Date`.</span><span class="sxs-lookup"><span data-stu-id="ec215-237">The `end` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="ec215-238">Режим создания</span><span class="sxs-lookup"><span data-stu-id="ec215-238">Compose mode</span></span>

<span data-ttu-id="ec215-239">Свойство `end` возвращает объект `Time`.</span><span class="sxs-lookup"><span data-stu-id="ec215-239">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="ec215-240">Если вы задаете время окончания с помощью метода [`Time.setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-), необходимо использовать метод [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) для преобразования местного времени на клиенте в формат UTC для сервера.</span><span class="sxs-lookup"><span data-stu-id="ec215-240">When you use the [`Time.setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="ec215-241">Тип:</span><span class="sxs-lookup"><span data-stu-id="ec215-241">Type:</span></span>

*   <span data-ttu-id="ec215-242">Date | [Time](/javascript/api/outlook_1_1/office.time)</span><span class="sxs-lookup"><span data-stu-id="ec215-242">Date | [Time](/javascript/api/outlook_1_1/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="ec215-243">Требования</span><span class="sxs-lookup"><span data-stu-id="ec215-243">Requirements</span></span>

|<span data-ttu-id="ec215-244">Requirement</span><span class="sxs-lookup"><span data-stu-id="ec215-244">Requirement</span></span>| <span data-ttu-id="ec215-245">Значение</span><span class="sxs-lookup"><span data-stu-id="ec215-245">Value</span></span>|
|---|---|
|[<span data-ttu-id="ec215-246">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="ec215-246">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ec215-247">1.0</span><span class="sxs-lookup"><span data-stu-id="ec215-247">1.0</span></span>|
|[<span data-ttu-id="ec215-248">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ec215-248">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ec215-249">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ec215-249">ReadItem</span></span>|
|[<span data-ttu-id="ec215-250">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ec215-250">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ec215-251">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="ec215-251">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="ec215-252">Пример</span><span class="sxs-lookup"><span data-stu-id="ec215-252">Example</span></span>

<span data-ttu-id="ec215-253">В примере ниже показано, как с помощью метода [`setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) объекта `Time` задать время окончания встречи в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="ec215-253">The following example sets the end time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

```JavaScript
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

#### <a name="from-emailaddressdetailsjavascriptapioutlook11officeemailaddressdetails"></a><span data-ttu-id="ec215-254">from :[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="ec215-254">from :[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)</span></span>

<span data-ttu-id="ec215-p113">Получает электронный адрес отправителя сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="ec215-p113">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="ec215-p114">Свойства `from` и [`sender`](#sender-emailaddressdetailsjavascriptapioutlook11officeemailaddressdetails) представляют одно лицо, если сообщение не отправлено представителем. В противном случае свойство `from` представляет лицо, делегировавшее полномочия, а свойство sender — представителя.</span><span class="sxs-lookup"><span data-stu-id="ec215-p114">The `from` and [`sender`](#sender-emailaddressdetailsjavascriptapioutlook11officeemailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="ec215-259">Свойству `recipientType`, принадлежащему объекту `EmailAddressDetails` в свойстве `from`, задано значение `undefined`.</span><span class="sxs-lookup"><span data-stu-id="ec215-259">Note: The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="ec215-260">Тип:</span><span class="sxs-lookup"><span data-stu-id="ec215-260">Type:</span></span>

*   [<span data-ttu-id="ec215-261">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="ec215-261">EmailAddressDetails</span></span>](/javascript/api/outlook_1_1/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="ec215-262">Требования</span><span class="sxs-lookup"><span data-stu-id="ec215-262">Requirements</span></span>

|<span data-ttu-id="ec215-263">Requirement</span><span class="sxs-lookup"><span data-stu-id="ec215-263">Requirement</span></span>| <span data-ttu-id="ec215-264">Значение</span><span class="sxs-lookup"><span data-stu-id="ec215-264">Value</span></span>|
|---|---|
|[<span data-ttu-id="ec215-265">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="ec215-265">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ec215-266">1.0</span><span class="sxs-lookup"><span data-stu-id="ec215-266">1.0</span></span>|
|[<span data-ttu-id="ec215-267">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ec215-267">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ec215-268">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ec215-268">ReadItem</span></span>|
|[<span data-ttu-id="ec215-269">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ec215-269">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ec215-270">Чтение</span><span class="sxs-lookup"><span data-stu-id="ec215-270">Read</span></span>|

#### <a name="internetmessageid-string"></a><span data-ttu-id="ec215-271">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="ec215-271">internetMessageId :String</span></span>

<span data-ttu-id="ec215-p115">Получает идентификатор интернет-сообщения для электронного сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="ec215-p115">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="ec215-274">Тип:</span><span class="sxs-lookup"><span data-stu-id="ec215-274">Type:</span></span>

*   <span data-ttu-id="ec215-275">String</span><span class="sxs-lookup"><span data-stu-id="ec215-275">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="ec215-276">Требования</span><span class="sxs-lookup"><span data-stu-id="ec215-276">Requirements</span></span>

|<span data-ttu-id="ec215-277">Requirement</span><span class="sxs-lookup"><span data-stu-id="ec215-277">Requirement</span></span>| <span data-ttu-id="ec215-278">Значение</span><span class="sxs-lookup"><span data-stu-id="ec215-278">Value</span></span>|
|---|---|
|[<span data-ttu-id="ec215-279">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="ec215-279">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ec215-280">1.0</span><span class="sxs-lookup"><span data-stu-id="ec215-280">1.0</span></span>|
|[<span data-ttu-id="ec215-281">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ec215-281">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ec215-282">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ec215-282">ReadItem</span></span>|
|[<span data-ttu-id="ec215-283">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ec215-283">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ec215-284">Чтение</span><span class="sxs-lookup"><span data-stu-id="ec215-284">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ec215-285">Пример</span><span class="sxs-lookup"><span data-stu-id="ec215-285">Example</span></span>

```JavaScript
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="ec215-286">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="ec215-286">itemClass :String</span></span>

<span data-ttu-id="ec215-p116">Получает класс элемента веб-служб Exchange для выбранного элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="ec215-p116">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="ec215-p117">Свойство `itemClass` указывает класс сообщения выбранного элемента. Ниже приводятся классы сообщения по умолчанию для элемента сообщения или встречи.</span><span class="sxs-lookup"><span data-stu-id="ec215-p117">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="ec215-291">Тип</span><span class="sxs-lookup"><span data-stu-id="ec215-291">Type</span></span> | <span data-ttu-id="ec215-292">Описание</span><span class="sxs-lookup"><span data-stu-id="ec215-292">Description</span></span> | <span data-ttu-id="ec215-293">Класс элемента</span><span class="sxs-lookup"><span data-stu-id="ec215-293">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="ec215-294">Элементы встречи</span><span class="sxs-lookup"><span data-stu-id="ec215-294">Appointment items</span></span> | <span data-ttu-id="ec215-295">Это элементы календаря для класса элемента `IPM.Appointment` или `IPM.Appointment.Occurence`.</span><span class="sxs-lookup"><span data-stu-id="ec215-295">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurence` |
| <span data-ttu-id="ec215-296">Элементы сообщения</span><span class="sxs-lookup"><span data-stu-id="ec215-296">Message items</span></span> | <span data-ttu-id="ec215-297">Сюда входят электронные сообщения, для которых по умолчанию задан класс сообщения `IPM.Note`, а также приглашения на собрания, ответы на них и уведомления об их отмене, использующие `IPM.Schedule.Meeting` в качестве базового класса сообщения.</span><span class="sxs-lookup"><span data-stu-id="ec215-297">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="ec215-298">Можно создавать настраиваемые классы сообщения, расширяющие классы сообщения по умолчанию, например настраиваемый класс сообщения о встрече `IPM.Appointment.Contoso`.</span><span class="sxs-lookup"><span data-stu-id="ec215-298">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="ec215-299">Тип:</span><span class="sxs-lookup"><span data-stu-id="ec215-299">Type:</span></span>

*   <span data-ttu-id="ec215-300">String</span><span class="sxs-lookup"><span data-stu-id="ec215-300">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="ec215-301">Требования</span><span class="sxs-lookup"><span data-stu-id="ec215-301">Requirements</span></span>

|<span data-ttu-id="ec215-302">Requirement</span><span class="sxs-lookup"><span data-stu-id="ec215-302">Requirement</span></span>| <span data-ttu-id="ec215-303">Значение</span><span class="sxs-lookup"><span data-stu-id="ec215-303">Value</span></span>|
|---|---|
|[<span data-ttu-id="ec215-304">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="ec215-304">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ec215-305">1.0</span><span class="sxs-lookup"><span data-stu-id="ec215-305">1.0</span></span>|
|[<span data-ttu-id="ec215-306">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ec215-306">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ec215-307">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ec215-307">ReadItem</span></span>|
|[<span data-ttu-id="ec215-308">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ec215-308">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ec215-309">Чтение</span><span class="sxs-lookup"><span data-stu-id="ec215-309">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ec215-310">Пример</span><span class="sxs-lookup"><span data-stu-id="ec215-310">Example</span></span>

```JavaScript
var itemClass = Office.context.mailbox.item.itemClass;
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="ec215-311">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="ec215-311">(nullable) itemId :String</span></span>

<span data-ttu-id="ec215-p118">Получает идентификатор элемента веб-служб Exchange для текущего элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="ec215-p118">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="ec215-314">Идентификатор, возвращаемый свойством `itemId`, совпадает с идентификатором элемента веб-служб Exchange.</span><span class="sxs-lookup"><span data-stu-id="ec215-314">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="ec215-315">Свойство `itemId` не совпадает с идентификатором записи Outlook или идентификатором, который используется REST API Outlook.</span><span class="sxs-lookup"><span data-stu-id="ec215-315">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="ec215-316">Прежде чем совершать вызовы REST API, используя это значение, его необходимо преобразовать с помощью метода `Office.context.mailbox.convertToRestId`, который доступен в наборе обязательных элементов, начиная с версии 1.3.</span><span class="sxs-lookup"><span data-stu-id="ec215-316">Before making REST API calls using this value, it should be converted using `Office.context.mailbox.convertToRestId`, which is available starting in requirement set 1.3.</span></span> <span data-ttu-id="ec215-317">Дополнительные сведения см. в статье [Использование REST API Outlook из надстройки Outlook](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="ec215-317">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

##### <a name="type"></a><span data-ttu-id="ec215-318">Тип:</span><span class="sxs-lookup"><span data-stu-id="ec215-318">Type:</span></span>

*   <span data-ttu-id="ec215-319">String</span><span class="sxs-lookup"><span data-stu-id="ec215-319">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="ec215-320">Требования</span><span class="sxs-lookup"><span data-stu-id="ec215-320">Requirements</span></span>

|<span data-ttu-id="ec215-321">Requirement</span><span class="sxs-lookup"><span data-stu-id="ec215-321">Requirement</span></span>| <span data-ttu-id="ec215-322">Значение</span><span class="sxs-lookup"><span data-stu-id="ec215-322">Value</span></span>|
|---|---|
|[<span data-ttu-id="ec215-323">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="ec215-323">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ec215-324">1.0</span><span class="sxs-lookup"><span data-stu-id="ec215-324">1.0</span></span>|
|[<span data-ttu-id="ec215-325">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ec215-325">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ec215-326">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ec215-326">ReadItem</span></span>|
|[<span data-ttu-id="ec215-327">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ec215-327">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ec215-328">Чтение</span><span class="sxs-lookup"><span data-stu-id="ec215-328">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ec215-329">Пример</span><span class="sxs-lookup"><span data-stu-id="ec215-329">Example</span></span>

<span data-ttu-id="ec215-p120">Указанный ниже код проверяет наличие идентификатора элемента. Если свойство `itemId` возвращает значение `null` или `undefined`, элемент будет сохранен в хранилище, а из асинхронного результата будет получен идентификатор элемента.</span><span class="sxs-lookup"><span data-stu-id="ec215-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```JavaScript
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result){
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook11officemailboxenumsitemtype"></a><span data-ttu-id="ec215-332">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_1/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="ec215-332">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_1/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="ec215-333">Получает тип элемента, который представляет экземпляр.</span><span class="sxs-lookup"><span data-stu-id="ec215-333">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="ec215-334">Свойство `itemType` возвращает одно из значений перечисления `ItemType`, которое указывает, является ли экземпляр объекта `item` сообщением или встречей.</span><span class="sxs-lookup"><span data-stu-id="ec215-334">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="ec215-335">Тип:</span><span class="sxs-lookup"><span data-stu-id="ec215-335">Type:</span></span>

*   [<span data-ttu-id="ec215-336">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="ec215-336">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook_1_1/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="ec215-337">Требования</span><span class="sxs-lookup"><span data-stu-id="ec215-337">Requirements</span></span>

|<span data-ttu-id="ec215-338">Requirement</span><span class="sxs-lookup"><span data-stu-id="ec215-338">Requirement</span></span>| <span data-ttu-id="ec215-339">Значение</span><span class="sxs-lookup"><span data-stu-id="ec215-339">Value</span></span>|
|---|---|
|[<span data-ttu-id="ec215-340">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="ec215-340">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ec215-341">1.0</span><span class="sxs-lookup"><span data-stu-id="ec215-341">1.0</span></span>|
|[<span data-ttu-id="ec215-342">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ec215-342">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ec215-343">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ec215-343">ReadItem</span></span>|
|[<span data-ttu-id="ec215-344">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ec215-344">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ec215-345">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="ec215-345">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="ec215-346">Пример</span><span class="sxs-lookup"><span data-stu-id="ec215-346">Example</span></span>

```JavaScript
if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Message)
  // do something
else
  // do something else
```

####  <a name="location-stringlocationjavascriptapioutlook11officelocation"></a><span data-ttu-id="ec215-347">location :String|[Location](/javascript/api/outlook_1_1/office.location)</span><span class="sxs-lookup"><span data-stu-id="ec215-347">location :String|[Location](/javascript/api/outlook_1_1/office.location)</span></span>

<span data-ttu-id="ec215-348">Получает или задает место встречи.</span><span class="sxs-lookup"><span data-stu-id="ec215-348">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="ec215-349">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="ec215-349">Read mode</span></span>

<span data-ttu-id="ec215-350">Свойство `location` возвращает строку, содержащую сведения о месте встречи.</span><span class="sxs-lookup"><span data-stu-id="ec215-350">The `location` property returns a string that contains the location of the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="ec215-351">Режим создания</span><span class="sxs-lookup"><span data-stu-id="ec215-351">Compose mode</span></span>

<span data-ttu-id="ec215-352">Свойство `location` возвращает объект `Location`, предоставляющий методы, которые используются для получения и задания места встречи.</span><span class="sxs-lookup"><span data-stu-id="ec215-352">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="ec215-353">Тип:</span><span class="sxs-lookup"><span data-stu-id="ec215-353">Type:</span></span>

*   <span data-ttu-id="ec215-354">String | [Location](/javascript/api/outlook_1_1/office.location)</span><span class="sxs-lookup"><span data-stu-id="ec215-354">String | [Location](/javascript/api/outlook_1_1/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="ec215-355">Требования</span><span class="sxs-lookup"><span data-stu-id="ec215-355">Requirements</span></span>

|<span data-ttu-id="ec215-356">Requirement</span><span class="sxs-lookup"><span data-stu-id="ec215-356">Requirement</span></span>| <span data-ttu-id="ec215-357">Значение</span><span class="sxs-lookup"><span data-stu-id="ec215-357">Value</span></span>|
|---|---|
|[<span data-ttu-id="ec215-358">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="ec215-358">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ec215-359">1.0</span><span class="sxs-lookup"><span data-stu-id="ec215-359">1.0</span></span>|
|[<span data-ttu-id="ec215-360">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ec215-360">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ec215-361">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ec215-361">ReadItem</span></span>|
|[<span data-ttu-id="ec215-362">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ec215-362">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ec215-363">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="ec215-363">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="ec215-364">Пример</span><span class="sxs-lookup"><span data-stu-id="ec215-364">Example</span></span>

```JavaScript
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

#### <a name="normalizedsubject-string"></a><span data-ttu-id="ec215-365">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="ec215-365">normalizedSubject :String</span></span>

<span data-ttu-id="ec215-p121">Получает тему элемента со всеми удаленными префиксами (включая `RE:` и `FWD:`). Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="ec215-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="ec215-p122">Свойство normalizedSubject получает тему элемента со стандартными префиксами (такими как `RE:` и `FW:`), добавляемыми почтовыми программами. Для получения темы элемента с неизмененными префиксами используйте свойство [`subject`](#subject-stringsubjectjavascriptapioutlook11officesubject).</span><span class="sxs-lookup"><span data-stu-id="ec215-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubjectjavascriptapioutlook11officesubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="ec215-370">Тип:</span><span class="sxs-lookup"><span data-stu-id="ec215-370">Type:</span></span>

*   <span data-ttu-id="ec215-371">String</span><span class="sxs-lookup"><span data-stu-id="ec215-371">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="ec215-372">Требования</span><span class="sxs-lookup"><span data-stu-id="ec215-372">Requirements</span></span>

|<span data-ttu-id="ec215-373">Requirement</span><span class="sxs-lookup"><span data-stu-id="ec215-373">Requirement</span></span>| <span data-ttu-id="ec215-374">Значение</span><span class="sxs-lookup"><span data-stu-id="ec215-374">Value</span></span>|
|---|---|
|[<span data-ttu-id="ec215-375">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="ec215-375">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ec215-376">1.0</span><span class="sxs-lookup"><span data-stu-id="ec215-376">1.0</span></span>|
|[<span data-ttu-id="ec215-377">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ec215-377">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ec215-378">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ec215-378">ReadItem</span></span>|
|[<span data-ttu-id="ec215-379">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ec215-379">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ec215-380">Чтение</span><span class="sxs-lookup"><span data-stu-id="ec215-380">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ec215-381">Пример</span><span class="sxs-lookup"><span data-stu-id="ec215-381">Example</span></span>

```JavaScript
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
```

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook11officeemailaddressdetailsrecipientsjavascriptapioutlook11officerecipients"></a><span data-ttu-id="ec215-382">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="ec215-382">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

<span data-ttu-id="ec215-383">Предоставляет доступ к необязательным участникам события.</span><span class="sxs-lookup"><span data-stu-id="ec215-383">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="ec215-384">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="ec215-384">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="ec215-385">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="ec215-385">Read mode</span></span>

<span data-ttu-id="ec215-386">Свойство `optionalAttendees` возвращает массив, содержащий объект `EmailAddressDetails` для каждого необязательного участника собрания.</span><span class="sxs-lookup"><span data-stu-id="ec215-386">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="ec215-387">Режим создания</span><span class="sxs-lookup"><span data-stu-id="ec215-387">Compose mode</span></span>

<span data-ttu-id="ec215-388">Свойство `optionalAttendees` возвращает объект `Recipients`, который предоставляет методы для получения или обновления необязательных участников собрания.</span><span class="sxs-lookup"><span data-stu-id="ec215-388">The `optionalAttendees` property returns a `Recipients` object that provides methods to get and set the optional attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="ec215-389">Тип:</span><span class="sxs-lookup"><span data-stu-id="ec215-389">Type:</span></span>

*   <span data-ttu-id="ec215-390">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="ec215-390">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="ec215-391">Требования</span><span class="sxs-lookup"><span data-stu-id="ec215-391">Requirements</span></span>

|<span data-ttu-id="ec215-392">Requirement</span><span class="sxs-lookup"><span data-stu-id="ec215-392">Requirement</span></span>| <span data-ttu-id="ec215-393">Значение</span><span class="sxs-lookup"><span data-stu-id="ec215-393">Value</span></span>|
|---|---|
|[<span data-ttu-id="ec215-394">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="ec215-394">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ec215-395">1.0</span><span class="sxs-lookup"><span data-stu-id="ec215-395">1.0</span></span>|
|[<span data-ttu-id="ec215-396">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ec215-396">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ec215-397">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ec215-397">ReadItem</span></span>|
|[<span data-ttu-id="ec215-398">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ec215-398">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ec215-399">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="ec215-399">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="ec215-400">Пример</span><span class="sxs-lookup"><span data-stu-id="ec215-400">Example</span></span>

```JavaScript
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook11officeemailaddressdetails"></a><span data-ttu-id="ec215-401">organizer :[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="ec215-401">organizer :[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)</span></span>

<span data-ttu-id="ec215-p124">Получает электронный адрес организатора указанного собрания. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="ec215-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="ec215-404">Тип:</span><span class="sxs-lookup"><span data-stu-id="ec215-404">Type:</span></span>

*   [<span data-ttu-id="ec215-405">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="ec215-405">EmailAddressDetails</span></span>](/javascript/api/outlook_1_1/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="ec215-406">Требования</span><span class="sxs-lookup"><span data-stu-id="ec215-406">Requirements</span></span>

|<span data-ttu-id="ec215-407">Requirement</span><span class="sxs-lookup"><span data-stu-id="ec215-407">Requirement</span></span>| <span data-ttu-id="ec215-408">Значение</span><span class="sxs-lookup"><span data-stu-id="ec215-408">Value</span></span>|
|---|---|
|[<span data-ttu-id="ec215-409">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="ec215-409">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ec215-410">1.0</span><span class="sxs-lookup"><span data-stu-id="ec215-410">1.0</span></span>|
|[<span data-ttu-id="ec215-411">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ec215-411">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ec215-412">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ec215-412">ReadItem</span></span>|
|[<span data-ttu-id="ec215-413">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ec215-413">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ec215-414">Чтение</span><span class="sxs-lookup"><span data-stu-id="ec215-414">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ec215-415">Пример</span><span class="sxs-lookup"><span data-stu-id="ec215-415">Example</span></span>

```JavaScript
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
```

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook11officeemailaddressdetailsrecipientsjavascriptapioutlook11officerecipients"></a><span data-ttu-id="ec215-416">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="ec215-416">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

<span data-ttu-id="ec215-417">Предоставляет доступ к обязательным участникам события.</span><span class="sxs-lookup"><span data-stu-id="ec215-417">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="ec215-418">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="ec215-418">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="ec215-419">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="ec215-419">Read mode</span></span>

<span data-ttu-id="ec215-420">Свойство `requiredAttendees` возвращает массив, содержащий объект `EmailAddressDetails` для каждого обязательного участника собрания.</span><span class="sxs-lookup"><span data-stu-id="ec215-420">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="ec215-421">Режим создания</span><span class="sxs-lookup"><span data-stu-id="ec215-421">Compose mode</span></span>

<span data-ttu-id="ec215-422">Свойство `requiredAttendees` возвращает объект `Recipients`, предоставляющий методы, с помощью которых можно получить или обновить сведения об обязательных участниках собрания.</span><span class="sxs-lookup"><span data-stu-id="ec215-422">The `requiredAttendees` property returns a `Recipients` object that provides methods to get and set the required attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="ec215-423">Тип:</span><span class="sxs-lookup"><span data-stu-id="ec215-423">Type:</span></span>

*   <span data-ttu-id="ec215-424">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="ec215-424">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="ec215-425">Требования</span><span class="sxs-lookup"><span data-stu-id="ec215-425">Requirements</span></span>

|<span data-ttu-id="ec215-426">Requirement</span><span class="sxs-lookup"><span data-stu-id="ec215-426">Requirement</span></span>| <span data-ttu-id="ec215-427">Значение</span><span class="sxs-lookup"><span data-stu-id="ec215-427">Value</span></span>|
|---|---|
|[<span data-ttu-id="ec215-428">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="ec215-428">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ec215-429">1.0</span><span class="sxs-lookup"><span data-stu-id="ec215-429">1.0</span></span>|
|[<span data-ttu-id="ec215-430">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ec215-430">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ec215-431">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ec215-431">ReadItem</span></span>|
|[<span data-ttu-id="ec215-432">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ec215-432">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ec215-433">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="ec215-433">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="ec215-434">Пример</span><span class="sxs-lookup"><span data-stu-id="ec215-434">Example</span></span>

```JavaScript
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
}
```

#### <a name="sender-emailaddressdetailsjavascriptapioutlook11officeemailaddressdetails"></a><span data-ttu-id="ec215-435">sender :[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="ec215-435">sender :[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)</span></span>

<span data-ttu-id="ec215-p126">Получает электронный адрес отправителя электронного сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="ec215-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="ec215-p127">Свойства [`from`](#from-emailaddressdetailsjavascriptapioutlook11officeemailaddressdetails) и `sender` представляют одно лицо, если сообщение не отправлено представителем. В противном случае свойство `from` представляет лицо, делегировавшее полномочия, а свойство sender — представителя.</span><span class="sxs-lookup"><span data-stu-id="ec215-p127">The [`from`](#from-emailaddressdetailsjavascriptapioutlook11officeemailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="ec215-440">Свойству `recipientType`, принадлежащему объекту `EmailAddressDetails` в свойстве `from`, задано значение `undefined`.</span><span class="sxs-lookup"><span data-stu-id="ec215-440">Note: The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="ec215-441">Тип:</span><span class="sxs-lookup"><span data-stu-id="ec215-441">Type:</span></span>

*   [<span data-ttu-id="ec215-442">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="ec215-442">EmailAddressDetails</span></span>](/javascript/api/outlook_1_1/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="ec215-443">Требования</span><span class="sxs-lookup"><span data-stu-id="ec215-443">Requirements</span></span>

|<span data-ttu-id="ec215-444">Requirement</span><span class="sxs-lookup"><span data-stu-id="ec215-444">Requirement</span></span>| <span data-ttu-id="ec215-445">Значение</span><span class="sxs-lookup"><span data-stu-id="ec215-445">Value</span></span>|
|---|---|
|[<span data-ttu-id="ec215-446">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="ec215-446">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ec215-447">1.0</span><span class="sxs-lookup"><span data-stu-id="ec215-447">1.0</span></span>|
|[<span data-ttu-id="ec215-448">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ec215-448">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ec215-449">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ec215-449">ReadItem</span></span>|
|[<span data-ttu-id="ec215-450">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ec215-450">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ec215-451">Чтение</span><span class="sxs-lookup"><span data-stu-id="ec215-451">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ec215-452">Пример</span><span class="sxs-lookup"><span data-stu-id="ec215-452">Example</span></span>

```JavaScript
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
```

####  <a name="start-datetimejavascriptapioutlook11officetime"></a><span data-ttu-id="ec215-453">start :Date|[Time](/javascript/api/outlook_1_1/office.time)</span><span class="sxs-lookup"><span data-stu-id="ec215-453">start :Date|[Time](/javascript/api/outlook_1_1/office.time)</span></span>

<span data-ttu-id="ec215-454">Получает или задает дату и время начала встречи.</span><span class="sxs-lookup"><span data-stu-id="ec215-454">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="ec215-p128">Свойство `start` представлено в виде значения даты и времени в формате UTC. Это значение можно преобразовать в местные значения даты и времени клиента с помощью метода [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook11officelocalclienttime).</span><span class="sxs-lookup"><span data-stu-id="ec215-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook11officelocalclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="ec215-457">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="ec215-457">Read mode</span></span>

<span data-ttu-id="ec215-458">Свойство `start` возвращает объект `Date`.</span><span class="sxs-lookup"><span data-stu-id="ec215-458">The `start` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="ec215-459">Режим создания</span><span class="sxs-lookup"><span data-stu-id="ec215-459">Compose mode</span></span>

<span data-ttu-id="ec215-460">Свойство `start` возвращает объект `Time`.</span><span class="sxs-lookup"><span data-stu-id="ec215-460">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="ec215-461">Если вы задаете время начала с помощью метода [`Time.setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-), необходимо использовать метод [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) для преобразования местного времени на клиенте в формат UTC для сервера.</span><span class="sxs-lookup"><span data-stu-id="ec215-461">When you use the [`Time.setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="ec215-462">Тип:</span><span class="sxs-lookup"><span data-stu-id="ec215-462">Type:</span></span>

*   <span data-ttu-id="ec215-463">Date | [Time](/javascript/api/outlook_1_1/office.time)</span><span class="sxs-lookup"><span data-stu-id="ec215-463">Date | [Time](/javascript/api/outlook_1_1/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="ec215-464">Требования</span><span class="sxs-lookup"><span data-stu-id="ec215-464">Requirements</span></span>

|<span data-ttu-id="ec215-465">Requirement</span><span class="sxs-lookup"><span data-stu-id="ec215-465">Requirement</span></span>| <span data-ttu-id="ec215-466">Значение</span><span class="sxs-lookup"><span data-stu-id="ec215-466">Value</span></span>|
|---|---|
|[<span data-ttu-id="ec215-467">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="ec215-467">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ec215-468">1.0</span><span class="sxs-lookup"><span data-stu-id="ec215-468">1.0</span></span>|
|[<span data-ttu-id="ec215-469">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ec215-469">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ec215-470">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ec215-470">ReadItem</span></span>|
|[<span data-ttu-id="ec215-471">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ec215-471">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ec215-472">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="ec215-472">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="ec215-473">Пример</span><span class="sxs-lookup"><span data-stu-id="ec215-473">Example</span></span>

<span data-ttu-id="ec215-474">В примере ниже с помощью метода [`setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) объекта `Time` задается время начала встречи в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="ec215-474">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

```JavaScript
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

####  <a name="subject-stringsubjectjavascriptapioutlook11officesubject"></a><span data-ttu-id="ec215-475">subject :String|[Subject](/javascript/api/outlook_1_1/office.subject)</span><span class="sxs-lookup"><span data-stu-id="ec215-475">subject :String|[Subject](/javascript/api/outlook_1_1/office.subject)</span></span>

<span data-ttu-id="ec215-476">Получает или задает описание, которое отображается в поле темы элемента.</span><span class="sxs-lookup"><span data-stu-id="ec215-476">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="ec215-477">Свойство `subject` получает или задает всю тему элемента для отправки с почтового сервера.</span><span class="sxs-lookup"><span data-stu-id="ec215-477">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="ec215-478">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="ec215-478">Read mode</span></span>

<span data-ttu-id="ec215-p129">Свойство `subject` возвращает строку. С помощью свойства [`normalizedSubject`](#normalizedsubject-string) можно получить тему без начальных префиксов, таких как `RE:` и `FW:`.</span><span class="sxs-lookup"><span data-stu-id="ec215-p129">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```js
var subject = Office.context.mailbox.item.subject;
```

##### <a name="compose-mode"></a><span data-ttu-id="ec215-481">Режим создания</span><span class="sxs-lookup"><span data-stu-id="ec215-481">Compose mode</span></span>

<span data-ttu-id="ec215-482">Свойство `subject` возвращает объект `Subject`, который предоставляет методы для получения и задания темы.</span><span class="sxs-lookup"><span data-stu-id="ec215-482">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```JavaScript
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="ec215-483">Тип:</span><span class="sxs-lookup"><span data-stu-id="ec215-483">Type:</span></span>

*   <span data-ttu-id="ec215-484">String | [Subject](/javascript/api/outlook_1_1/office.subject)</span><span class="sxs-lookup"><span data-stu-id="ec215-484">String | [Subject](/javascript/api/outlook_1_1/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="ec215-485">Требования</span><span class="sxs-lookup"><span data-stu-id="ec215-485">Requirements</span></span>

|<span data-ttu-id="ec215-486">Requirement</span><span class="sxs-lookup"><span data-stu-id="ec215-486">Requirement</span></span>| <span data-ttu-id="ec215-487">Значение</span><span class="sxs-lookup"><span data-stu-id="ec215-487">Value</span></span>|
|---|---|
|[<span data-ttu-id="ec215-488">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="ec215-488">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ec215-489">1.0</span><span class="sxs-lookup"><span data-stu-id="ec215-489">1.0</span></span>|
|[<span data-ttu-id="ec215-490">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ec215-490">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ec215-491">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ec215-491">ReadItem</span></span>|
|[<span data-ttu-id="ec215-492">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ec215-492">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ec215-493">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="ec215-493">Compose or read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlook11officeemailaddressdetailsrecipientsjavascriptapioutlook11officerecipients"></a><span data-ttu-id="ec215-494">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="ec215-494">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

<span data-ttu-id="ec215-495">Предоставляет доступ к получателям, указанным в строке **Кому** сообщения.</span><span class="sxs-lookup"><span data-stu-id="ec215-495">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="ec215-496">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="ec215-496">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="ec215-497">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="ec215-497">Read mode</span></span>

<span data-ttu-id="ec215-p131">Свойство `to` возвращает массив, содержащий объект `EmailAddressDetails` для каждого получателя в строке **Кому** сообщения. Коллекция может включать не более 100 элементов.</span><span class="sxs-lookup"><span data-stu-id="ec215-p131">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="ec215-500">Режим создания</span><span class="sxs-lookup"><span data-stu-id="ec215-500">Compose mode</span></span>

<span data-ttu-id="ec215-501">Свойство `to` возвращает объект `Recipients`, предоставляющий методы для получения или обновления получателей, которые указаны в строке **Кому** сообщения.</span><span class="sxs-lookup"><span data-stu-id="ec215-501">The to`to` property returns a Recipients`Recipients` object that provides methods for manipulating the recipients on the **To** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="ec215-502">Тип:</span><span class="sxs-lookup"><span data-stu-id="ec215-502">Type:</span></span>

*   <span data-ttu-id="ec215-503">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="ec215-503">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="ec215-504">Требования</span><span class="sxs-lookup"><span data-stu-id="ec215-504">Requirements</span></span>

|<span data-ttu-id="ec215-505">Requirement</span><span class="sxs-lookup"><span data-stu-id="ec215-505">Requirement</span></span>| <span data-ttu-id="ec215-506">Значение</span><span class="sxs-lookup"><span data-stu-id="ec215-506">Value</span></span>|
|---|---|
|[<span data-ttu-id="ec215-507">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="ec215-507">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ec215-508">1.0</span><span class="sxs-lookup"><span data-stu-id="ec215-508">1.0</span></span>|
|[<span data-ttu-id="ec215-509">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ec215-509">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ec215-510">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ec215-510">ReadItem</span></span>|
|[<span data-ttu-id="ec215-511">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ec215-511">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ec215-512">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="ec215-512">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="ec215-513">Пример</span><span class="sxs-lookup"><span data-stu-id="ec215-513">Example</span></span>

```JavaScript
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

### <a name="methods"></a><span data-ttu-id="ec215-514">Методы</span><span class="sxs-lookup"><span data-stu-id="ec215-514">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="ec215-515">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="ec215-515">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="ec215-516">Добавляет файл в сообщение или встречу в качестве вложения.</span><span class="sxs-lookup"><span data-stu-id="ec215-516">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="ec215-517">Метод `addFileAttachmentAsync` передает файл по указанному универсальному коду ресурса (URI) и вкладывает его в элемент в форме создания.</span><span class="sxs-lookup"><span data-stu-id="ec215-517">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="ec215-518">Идентификатор можно последовательно использовать с методом [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback), чтобы удалить вложение, добавленное во время текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="ec215-518">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ec215-519">Параметры</span><span class="sxs-lookup"><span data-stu-id="ec215-519">Parameters:</span></span>

|<span data-ttu-id="ec215-520">Имя</span><span class="sxs-lookup"><span data-stu-id="ec215-520">Name</span></span>| <span data-ttu-id="ec215-521">Тип</span><span class="sxs-lookup"><span data-stu-id="ec215-521">Type</span></span>| <span data-ttu-id="ec215-522">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="ec215-522">Attributes</span></span>| <span data-ttu-id="ec215-523">Описание</span><span class="sxs-lookup"><span data-stu-id="ec215-523">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="ec215-524">String</span><span class="sxs-lookup"><span data-stu-id="ec215-524">String</span></span>||<span data-ttu-id="ec215-p132">Универсальный код ресурса (URI), представляющий расположение файла, который нужно вложить в сообщение или встречу. Максимальная длина — 2048 символов.</span><span class="sxs-lookup"><span data-stu-id="ec215-p132">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="ec215-527">String</span><span class="sxs-lookup"><span data-stu-id="ec215-527">String</span></span>||<span data-ttu-id="ec215-p133">Имя вложения, которое отображается при передаче вложения. Максимальная длина — 255 символов.</span><span class="sxs-lookup"><span data-stu-id="ec215-p133">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="ec215-530">Object</span><span class="sxs-lookup"><span data-stu-id="ec215-530">Object</span></span>| <span data-ttu-id="ec215-531">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="ec215-531">&lt;optional&gt;</span></span>|<span data-ttu-id="ec215-532">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="ec215-532">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="ec215-533">Object</span><span class="sxs-lookup"><span data-stu-id="ec215-533">Object</span></span>| <span data-ttu-id="ec215-534">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="ec215-534">&lt;optional&gt;</span></span>|<span data-ttu-id="ec215-535">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="ec215-535">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="ec215-536">функция</span><span class="sxs-lookup"><span data-stu-id="ec215-536">function</span></span>| <span data-ttu-id="ec215-537">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="ec215-537">&lt;optional&gt;</span></span>|<span data-ttu-id="ec215-538">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="ec215-538">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="ec215-539">После успешного выполнения идентификатор вложения будет представлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="ec215-539">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="ec215-540">Если добавить вложение не удастся, объект `asyncResult` будет содержать объект `Error` с описанием ошибки.</span><span class="sxs-lookup"><span data-stu-id="ec215-540">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="ec215-541">Ошибки</span><span class="sxs-lookup"><span data-stu-id="ec215-541">Errors</span></span>

| <span data-ttu-id="ec215-542">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="ec215-542">Error code</span></span> | <span data-ttu-id="ec215-543">Описание</span><span class="sxs-lookup"><span data-stu-id="ec215-543">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="ec215-544">Вложение превышает максимальный размер.</span><span class="sxs-lookup"><span data-stu-id="ec215-544">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="ec215-545">Расширение вложения не поддерживается.</span><span class="sxs-lookup"><span data-stu-id="ec215-545">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="ec215-546">Сообщение или встреча содержат слишком много вложений.</span><span class="sxs-lookup"><span data-stu-id="ec215-546">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="ec215-547">Требования</span><span class="sxs-lookup"><span data-stu-id="ec215-547">Requirements</span></span>

|<span data-ttu-id="ec215-548">Requirement</span><span class="sxs-lookup"><span data-stu-id="ec215-548">Requirement</span></span>| <span data-ttu-id="ec215-549">Значение</span><span class="sxs-lookup"><span data-stu-id="ec215-549">Value</span></span>|
|---|---|
|[<span data-ttu-id="ec215-550">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="ec215-550">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ec215-551">1.1</span><span class="sxs-lookup"><span data-stu-id="ec215-551">1.1</span></span>|
|[<span data-ttu-id="ec215-552">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ec215-552">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ec215-553">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="ec215-553">ReadWriteItem</span></span>|
|[<span data-ttu-id="ec215-554">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ec215-554">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ec215-555">Создание</span><span class="sxs-lookup"><span data-stu-id="ec215-555">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="ec215-556">Пример</span><span class="sxs-lookup"><span data-stu-id="ec215-556">Example</span></span>

```JavaScript
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

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="ec215-557">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="ec215-557">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="ec215-558">Добавляет к сообщению элемент Exchange, например сообщение, в виде вложения.</span><span class="sxs-lookup"><span data-stu-id="ec215-558">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="ec215-p134">С помощью метода `addItemAttachmentAsync` можно в элемент формы создания вложить элемент с указанным идентификатором Exchange. Если указать метод обратного вызова, то этот метод вызывается с помощью параметра `asyncResult`, который содержит идентификатор вложения или код, указывающий на ошибки, которые произошли при вложении элемента. При необходимости можно использовать параметр `options` для передачи сведений о состоянии методу обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="ec215-p134">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="ec215-562">Идентификатор можно использовать с методом [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback), чтобы удалить вложение, добавленное во время текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="ec215-562">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="ec215-563">Если ваша надстройка Office выполняется в Outlook Web App, метод `addItemAttachmentAsync` обеспечивает вложение элементов в элементы, отличные от редактируемого. Однако это действие не рекомендуем выполнять, так как оно не поддерживается.</span><span class="sxs-lookup"><span data-stu-id="ec215-563">If your Office add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ec215-564">Параметры:</span><span class="sxs-lookup"><span data-stu-id="ec215-564">Parameters:</span></span>

|<span data-ttu-id="ec215-565">Имя</span><span class="sxs-lookup"><span data-stu-id="ec215-565">Name</span></span>| <span data-ttu-id="ec215-566">Тип</span><span class="sxs-lookup"><span data-stu-id="ec215-566">Type</span></span>| <span data-ttu-id="ec215-567">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="ec215-567">Attributes</span></span>| <span data-ttu-id="ec215-568">Описание</span><span class="sxs-lookup"><span data-stu-id="ec215-568">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="ec215-569">String</span><span class="sxs-lookup"><span data-stu-id="ec215-569">String</span></span>||<span data-ttu-id="ec215-p135">Идентификатор Exchange для вкладываемого элемента. Максимальная длина — 100 символов.</span><span class="sxs-lookup"><span data-stu-id="ec215-p135">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="ec215-572">String</span><span class="sxs-lookup"><span data-stu-id="ec215-572">String</span></span>||<span data-ttu-id="ec215-p136">Тема вкладываемого элемента. Максимальная длина — 255 символов.</span><span class="sxs-lookup"><span data-stu-id="ec215-p136">The sujbect of the item to be attached. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="ec215-575">Object</span><span class="sxs-lookup"><span data-stu-id="ec215-575">Object</span></span>| <span data-ttu-id="ec215-576">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="ec215-576">&lt;optional&gt;</span></span>|<span data-ttu-id="ec215-577">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="ec215-577">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="ec215-578">Object</span><span class="sxs-lookup"><span data-stu-id="ec215-578">Object</span></span>| <span data-ttu-id="ec215-579">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="ec215-579">&lt;optional&gt;</span></span>|<span data-ttu-id="ec215-580">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="ec215-580">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="ec215-581">функция</span><span class="sxs-lookup"><span data-stu-id="ec215-581">function</span></span>| <span data-ttu-id="ec215-582">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="ec215-582">&lt;optional&gt;</span></span>|<span data-ttu-id="ec215-583">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="ec215-583">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="ec215-584">После успешного выполнения идентификатор вложения будет представлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="ec215-584">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="ec215-585">Если добавить вложение не удастся, объект `asyncResult` будет содержать объект `Error` с описанием ошибки.</span><span class="sxs-lookup"><span data-stu-id="ec215-585">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="ec215-586">Ошибки</span><span class="sxs-lookup"><span data-stu-id="ec215-586">Errors</span></span>

| <span data-ttu-id="ec215-587">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="ec215-587">Error code</span></span> | <span data-ttu-id="ec215-588">Описание</span><span class="sxs-lookup"><span data-stu-id="ec215-588">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="ec215-589">Сообщение или встреча содержат слишком много вложений.</span><span class="sxs-lookup"><span data-stu-id="ec215-589">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="ec215-590">Требования</span><span class="sxs-lookup"><span data-stu-id="ec215-590">Requirements</span></span>

|<span data-ttu-id="ec215-591">Requirement</span><span class="sxs-lookup"><span data-stu-id="ec215-591">Requirement</span></span>| <span data-ttu-id="ec215-592">Значение</span><span class="sxs-lookup"><span data-stu-id="ec215-592">Value</span></span>|
|---|---|
|[<span data-ttu-id="ec215-593">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="ec215-593">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ec215-594">1.1</span><span class="sxs-lookup"><span data-stu-id="ec215-594">1.1</span></span>|
|[<span data-ttu-id="ec215-595">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ec215-595">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ec215-596">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="ec215-596">ReadWriteItem</span></span>|
|[<span data-ttu-id="ec215-597">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ec215-597">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ec215-598">Создание</span><span class="sxs-lookup"><span data-stu-id="ec215-598">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="ec215-599">Пример</span><span class="sxs-lookup"><span data-stu-id="ec215-599">Example</span></span>

<span data-ttu-id="ec215-600">В следующем примере существующий элемент Outlook добавляется в виде вложения с именем `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="ec215-600">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

```JavaScript
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

#### <a name="displayreplyallformformdata"></a><span data-ttu-id="ec215-601">displayReplyAllForm(formData)</span><span class="sxs-lookup"><span data-stu-id="ec215-601">displayReplyAllForm(formData)</span></span>

<span data-ttu-id="ec215-602">Отображает форму ответа, включающую отправителя и всех получателей выбранного сообщения или организатора и всех участников выбранной встречи.</span><span class="sxs-lookup"><span data-stu-id="ec215-602">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="ec215-603">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="ec215-603">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="ec215-604">В Outlook Web App форма ответа отображается в виде всплывающей формы в представлении с 3 либо 1 или 2 колонками.</span><span class="sxs-lookup"><span data-stu-id="ec215-604">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="ec215-605">Если любой строковый параметр превышает указанные для него ограничения, `displayReplyAllForm` возвращает исключение.</span><span class="sxs-lookup"><span data-stu-id="ec215-605">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

> [!NOTE]
> <span data-ttu-id="ec215-606">Набор обязательных элементов 1.1 не поддерживает возможность включения вложений при вызове `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="ec215-606">NOTE: The ability to include attachments in the call to `displayReplyAllForm` is not supported in requirement set 1.1.</span></span> <span data-ttu-id="ec215-607">Поддержка вложений была добавлена для `displayReplyAllForm` в наборе обязательных элементов 1.2 и более поздних версий.</span><span class="sxs-lookup"><span data-stu-id="ec215-607">Attachment support was added to `displayReplyAllForm` in requirement set 1.2 and above.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ec215-608">Параметры:</span><span class="sxs-lookup"><span data-stu-id="ec215-608">Parameters:</span></span>

|<span data-ttu-id="ec215-609">Имя</span><span class="sxs-lookup"><span data-stu-id="ec215-609">Name</span></span>| <span data-ttu-id="ec215-610">Тип</span><span class="sxs-lookup"><span data-stu-id="ec215-610">Type</span></span>| <span data-ttu-id="ec215-611">Описание</span><span class="sxs-lookup"><span data-stu-id="ec215-611">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="ec215-612">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="ec215-612">String &#124; Object</span></span>| |<span data-ttu-id="ec215-p138">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="ec215-p138">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="ec215-615">**ИЛИ**</span><span class="sxs-lookup"><span data-stu-id="ec215-615">**OR**</span></span><br/><span data-ttu-id="ec215-p139">Объект, который содержит текст или данные вложения и функцию обратного вызова. Ниже представлено определение этого объекта.</span><span class="sxs-lookup"><span data-stu-id="ec215-p139">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="ec215-618">String</span><span class="sxs-lookup"><span data-stu-id="ec215-618">String</span></span> | <span data-ttu-id="ec215-619">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="ec215-619">&lt;optional&gt;</span></span> | <span data-ttu-id="ec215-p140">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Длина строки ограничена 32 символами.</span><span class="sxs-lookup"><span data-stu-id="ec215-p140">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `callback` | <span data-ttu-id="ec215-622">function</span><span class="sxs-lookup"><span data-stu-id="ec215-622">function</span></span> | <span data-ttu-id="ec215-623">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="ec215-623">&lt;optional&gt;</span></span> | <span data-ttu-id="ec215-624">По завершении работы метода функция, переданная параметру `callback`, вызывается с помощью одного параметра `asyncResult`, представляющего собой объект [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="ec215-624">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="ec215-625">Требования</span><span class="sxs-lookup"><span data-stu-id="ec215-625">Requirements</span></span>

|<span data-ttu-id="ec215-626">Requirement</span><span class="sxs-lookup"><span data-stu-id="ec215-626">Requirement</span></span>| <span data-ttu-id="ec215-627">Значение</span><span class="sxs-lookup"><span data-stu-id="ec215-627">Value</span></span>|
|---|---|
|[<span data-ttu-id="ec215-628">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="ec215-628">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ec215-629">1.0</span><span class="sxs-lookup"><span data-stu-id="ec215-629">1.0</span></span>|
|[<span data-ttu-id="ec215-630">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ec215-630">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ec215-631">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ec215-631">ReadItem</span></span>|
|[<span data-ttu-id="ec215-632">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ec215-632">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ec215-633">Чтение</span><span class="sxs-lookup"><span data-stu-id="ec215-633">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="ec215-634">Примеры</span><span class="sxs-lookup"><span data-stu-id="ec215-634">Examples</span></span>

<span data-ttu-id="ec215-635">Приведенный ниже код передает строку в функцию `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="ec215-635">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="ec215-636">Ответ с пустым текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="ec215-636">Reply with an empty body.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="ec215-637">Ответ только с текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="ec215-637">Reply with just a body.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="ec215-638">Ответ с текстом сообщения и вложенным файлом.</span><span class="sxs-lookup"><span data-stu-id="ec215-638">Reply with a body and a callback.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi',
  'callback' : function(asyncResult)
  {
    console.log(asyncResult.value);
  }
});
```

#### <a name="displayreplyformformdata"></a><span data-ttu-id="ec215-639">displayReplyForm(formData)</span><span class="sxs-lookup"><span data-stu-id="ec215-639">displayReplyForm(formData)</span></span>

<span data-ttu-id="ec215-640">Отображает форму ответа, включающую только отправителя выбранного сообщения или организатора выбранной встречи.</span><span class="sxs-lookup"><span data-stu-id="ec215-640">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="ec215-641">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="ec215-641">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="ec215-642">В Outlook Web App форма ответа отображается в виде всплывающей формы в представлении с 3 либо 1 или 2 колонками.</span><span class="sxs-lookup"><span data-stu-id="ec215-642">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="ec215-643">Если любой строковый параметр превышает указанные для него ограничения, `displayReplyForm` возвращает исключение.</span><span class="sxs-lookup"><span data-stu-id="ec215-643">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

> [!NOTE]
> <span data-ttu-id="ec215-644">Набор обязательных элементов 1.1 не поддерживает возможность включения вложений при вызове `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="ec215-644">NOTE: The ability to include attachments in the call to `displayReplyForm` is not supported in requirement set 1.1.</span></span> <span data-ttu-id="ec215-645">Поддержка вложений была добавлена для `displayReplyForm` в наборе обязательных элементов 1.2 и более поздних версий.</span><span class="sxs-lookup"><span data-stu-id="ec215-645">Attachment support was added to `displayReplyForm` in requirement set 1.2 and above.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ec215-646">Параметры:</span><span class="sxs-lookup"><span data-stu-id="ec215-646">Parameters:</span></span>

|<span data-ttu-id="ec215-647">Имя</span><span class="sxs-lookup"><span data-stu-id="ec215-647">Name</span></span>| <span data-ttu-id="ec215-648">Тип</span><span class="sxs-lookup"><span data-stu-id="ec215-648">Type</span></span>| <span data-ttu-id="ec215-649">Описание</span><span class="sxs-lookup"><span data-stu-id="ec215-649">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="ec215-650">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="ec215-650">String &#124; Object</span></span>| | <span data-ttu-id="ec215-p142">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="ec215-p142">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="ec215-653">**ИЛИ**</span><span class="sxs-lookup"><span data-stu-id="ec215-653">**OR**</span></span><br/><span data-ttu-id="ec215-p143">Объект, который содержит текст или данные вложения и функцию обратного вызова. Ниже представлено определение этого объекта.</span><span class="sxs-lookup"><span data-stu-id="ec215-p143">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="ec215-656">String</span><span class="sxs-lookup"><span data-stu-id="ec215-656">String</span></span> | <span data-ttu-id="ec215-657">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="ec215-657">&lt;optional&gt;</span></span> | <span data-ttu-id="ec215-p144">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Длина строки ограничена 32 символами.</span><span class="sxs-lookup"><span data-stu-id="ec215-p144">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `callback` | <span data-ttu-id="ec215-660">function</span><span class="sxs-lookup"><span data-stu-id="ec215-660">function</span></span> | <span data-ttu-id="ec215-661">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="ec215-661">&lt;optional&gt;</span></span> | <span data-ttu-id="ec215-662">По завершении работы метода функция, переданная параметру `callback`, вызывается с помощью одного параметра `asyncResult`, представляющего собой объект [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="ec215-662">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="ec215-663">Требования</span><span class="sxs-lookup"><span data-stu-id="ec215-663">Requirements</span></span>

|<span data-ttu-id="ec215-664">Requirement</span><span class="sxs-lookup"><span data-stu-id="ec215-664">Requirement</span></span>| <span data-ttu-id="ec215-665">Значение</span><span class="sxs-lookup"><span data-stu-id="ec215-665">Value</span></span>|
|---|---|
|[<span data-ttu-id="ec215-666">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="ec215-666">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ec215-667">1.0</span><span class="sxs-lookup"><span data-stu-id="ec215-667">1.0</span></span>|
|[<span data-ttu-id="ec215-668">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ec215-668">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ec215-669">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ec215-669">ReadItem</span></span>|
|[<span data-ttu-id="ec215-670">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ec215-670">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ec215-671">Чтение</span><span class="sxs-lookup"><span data-stu-id="ec215-671">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="ec215-672">Примеры</span><span class="sxs-lookup"><span data-stu-id="ec215-672">Examples</span></span>

<span data-ttu-id="ec215-673">Приведенный ниже код передает строку в функцию `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="ec215-673">The following code passes a string to the `displayReplyForm` function.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="ec215-674">Ответ с пустым текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="ec215-674">Reply with an empty body.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="ec215-675">Ответ только с текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="ec215-675">Reply with just a body.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="ec215-676">Ответ с текстом сообщения и вложенным файлом.</span><span class="sxs-lookup"><span data-stu-id="ec215-676">Reply with a body and a callback.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi',
  'callback' : function(asyncResult)
  {
    console.log(asyncResult.value);
  }
});
```

#### <a name="getentities--entitiesjavascriptapioutlook11officeentities"></a><span data-ttu-id="ec215-677">getEntities() → {[Entities](/javascript/api/outlook_1_1/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="ec215-677">getEntities() → {[Entities](/javascript/api/outlook_1_1/office.entities)}</span></span>

<span data-ttu-id="ec215-678">Получает сущности, обнаруженные в теле выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="ec215-678">Gets the entities found in the selected item.</span></span>

> [!NOTE]
> <span data-ttu-id="ec215-679">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="ec215-679">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="ec215-680">Требования</span><span class="sxs-lookup"><span data-stu-id="ec215-680">Requirements</span></span>

|<span data-ttu-id="ec215-681">Requirement</span><span class="sxs-lookup"><span data-stu-id="ec215-681">Requirement</span></span>| <span data-ttu-id="ec215-682">Значение</span><span class="sxs-lookup"><span data-stu-id="ec215-682">Value</span></span>|
|---|---|
|[<span data-ttu-id="ec215-683">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="ec215-683">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ec215-684">1.0</span><span class="sxs-lookup"><span data-stu-id="ec215-684">1.0</span></span>|
|[<span data-ttu-id="ec215-685">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ec215-685">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ec215-686">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ec215-686">ReadItem</span></span>|
|[<span data-ttu-id="ec215-687">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ec215-687">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ec215-688">Чтение</span><span class="sxs-lookup"><span data-stu-id="ec215-688">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="ec215-689">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="ec215-689">Returns:</span></span>

<span data-ttu-id="ec215-690">Тип: [Entities](/javascript/api/outlook_1_1/office.entities)</span><span class="sxs-lookup"><span data-stu-id="ec215-690">Type: [Entities](/javascript/api/outlook_1_1/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="ec215-691">Пример</span><span class="sxs-lookup"><span data-stu-id="ec215-691">Example</span></span>

<span data-ttu-id="ec215-692">Ниже приведен пример получения доступа к сущностям контактов в теле текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="ec215-692">The following example accesses the contacts entities on the current item.</span></span>

```JavaScript
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook11officecontactmeetingsuggestionjavascriptapioutlook11officemeetingsuggestionphonenumberjavascriptapioutlook11officephonenumbertasksuggestionjavascriptapioutlook11officetasksuggestion"></a><span data-ttu-id="ec215-693">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="ec215-693">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))>}</span></span>

<span data-ttu-id="ec215-694">Получает массив всех сущностей указанного типа, обнаруженных в теле выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="ec215-694">Gets an array of all the entities of the specified entity type found in the selected item.</span></span>

> [!NOTE]
> <span data-ttu-id="ec215-695">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="ec215-695">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ec215-696">Параметры</span><span class="sxs-lookup"><span data-stu-id="ec215-696">Parameters:</span></span>

|<span data-ttu-id="ec215-697">Имя</span><span class="sxs-lookup"><span data-stu-id="ec215-697">Name</span></span>| <span data-ttu-id="ec215-698">Тип</span><span class="sxs-lookup"><span data-stu-id="ec215-698">Type</span></span>| <span data-ttu-id="ec215-699">Описание</span><span class="sxs-lookup"><span data-stu-id="ec215-699">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="ec215-700">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="ec215-700">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook_1_1/office.MailboxEnums.entitytype)|<span data-ttu-id="ec215-701">Одно из значений перечисления EntityType.</span><span class="sxs-lookup"><span data-stu-id="ec215-701">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ec215-702">Требования</span><span class="sxs-lookup"><span data-stu-id="ec215-702">Requirements</span></span>

|<span data-ttu-id="ec215-703">Requirement</span><span class="sxs-lookup"><span data-stu-id="ec215-703">Requirement</span></span>| <span data-ttu-id="ec215-704">Значение</span><span class="sxs-lookup"><span data-stu-id="ec215-704">Value</span></span>|
|---|---|
|[<span data-ttu-id="ec215-705">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="ec215-705">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ec215-706">1.0</span><span class="sxs-lookup"><span data-stu-id="ec215-706">1.0</span></span>|
|[<span data-ttu-id="ec215-707">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ec215-707">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ec215-708">Restricted</span><span class="sxs-lookup"><span data-stu-id="ec215-708">Restricted</span></span>|
|[<span data-ttu-id="ec215-709">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ec215-709">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ec215-710">Чтение</span><span class="sxs-lookup"><span data-stu-id="ec215-710">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="ec215-711">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="ec215-711">Returns:</span></span>

<span data-ttu-id="ec215-712">Если значение, переданное в `entityType`, не является допустимым членом перечисления `EntityType`, метод возвращает значение NULL.</span><span class="sxs-lookup"><span data-stu-id="ec215-712">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="ec215-713">Если в теле элемента отсутствуют сущности указанного типа, метод возвращает пустой массив.</span><span class="sxs-lookup"><span data-stu-id="ec215-713">If no entities of the specified type are present on the item, the method returns an empty array.</span></span> <span data-ttu-id="ec215-714">В противном случае тип объектов в возвращаемом массиве зависит от типа сущности, запрошенной в параметре `entityType`.</span><span class="sxs-lookup"><span data-stu-id="ec215-714">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="ec215-715">Хотя минимальный уровень разрешений для использования этого метода — **Restricted**, для некоторых типов сущностей требуется доступ на уровне **ReadItem**, как указано в приведенной ниже таблице.</span><span class="sxs-lookup"><span data-stu-id="ec215-715">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="ec215-716">Значение параметра `entityType`</span><span class="sxs-lookup"><span data-stu-id="ec215-716">Value of `entityType`</span></span> | <span data-ttu-id="ec215-717">Тип объектов в возвращаемом массиве</span><span class="sxs-lookup"><span data-stu-id="ec215-717">Type of objects in returned array</span></span> | <span data-ttu-id="ec215-718">Необходимый уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ec215-718">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="ec215-719">String</span><span class="sxs-lookup"><span data-stu-id="ec215-719">String</span></span> | <span data-ttu-id="ec215-720">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="ec215-720">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="ec215-721">Contact</span><span class="sxs-lookup"><span data-stu-id="ec215-721">Contact</span></span> | <span data-ttu-id="ec215-722">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="ec215-722">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="ec215-723">String</span><span class="sxs-lookup"><span data-stu-id="ec215-723">String</span></span> | <span data-ttu-id="ec215-724">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="ec215-724">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="ec215-725">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="ec215-725">MeetingSuggestion</span></span> | <span data-ttu-id="ec215-726">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="ec215-726">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="ec215-727">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="ec215-727">PhoneNumber</span></span> | <span data-ttu-id="ec215-728">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="ec215-728">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="ec215-729">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="ec215-729">TaskSuggestion</span></span> | <span data-ttu-id="ec215-730">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="ec215-730">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="ec215-731">String</span><span class="sxs-lookup"><span data-stu-id="ec215-731">String</span></span> | <span data-ttu-id="ec215-732">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="ec215-732">**Restricted**</span></span> |

<span data-ttu-id="ec215-733">Тип:  Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="ec215-733">Type:  Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))></span></span>


##### <a name="example"></a><span data-ttu-id="ec215-734">Пример</span><span class="sxs-lookup"><span data-stu-id="ec215-734">Example</span></span>

<span data-ttu-id="ec215-735">В примере ниже показано, как получить доступ к массиву строк, которые представляют собой почтовые адреса в теле текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="ec215-735">The following example shows how to access an array of strings that represent postal addresses in the subject or body of the current item.</span></span>

```JavaScript
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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook11officecontactmeetingsuggestionjavascriptapioutlook11officemeetingsuggestionphonenumberjavascriptapioutlook11officephonenumbertasksuggestionjavascriptapioutlook11officetasksuggestion"></a><span data-ttu-id="ec215-736">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="ec215-736">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))>}</span></span>

<span data-ttu-id="ec215-737">Возвращает известные сущности в выбранном элементе, которые проходят через именованный фильтр, определяемый в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="ec215-737">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="ec215-738">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="ec215-738">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="ec215-739">Метод `getFilteredEntitiesByName` возвращает сущности, соответствующие регулярному выражению, которое определяется в элементе правила [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) в XML-файле манифеста, с использованием указанного значения элемента `FilterName`.</span><span class="sxs-lookup"><span data-stu-id="ec215-739">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ec215-740">Параметры</span><span class="sxs-lookup"><span data-stu-id="ec215-740">Parameters:</span></span>

|<span data-ttu-id="ec215-741">Имя</span><span class="sxs-lookup"><span data-stu-id="ec215-741">Name</span></span>| <span data-ttu-id="ec215-742">Тип</span><span class="sxs-lookup"><span data-stu-id="ec215-742">Type</span></span>| <span data-ttu-id="ec215-743">Описание</span><span class="sxs-lookup"><span data-stu-id="ec215-743">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="ec215-744">String</span><span class="sxs-lookup"><span data-stu-id="ec215-744">String</span></span>|<span data-ttu-id="ec215-745">Имя элемента правила `ItemHasKnownEntity`, определяющее соответствующий фильтр.</span><span class="sxs-lookup"><span data-stu-id="ec215-745">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ec215-746">Требования</span><span class="sxs-lookup"><span data-stu-id="ec215-746">Requirements</span></span>

|<span data-ttu-id="ec215-747">Requirement</span><span class="sxs-lookup"><span data-stu-id="ec215-747">Requirement</span></span>| <span data-ttu-id="ec215-748">Значение</span><span class="sxs-lookup"><span data-stu-id="ec215-748">Value</span></span>|
|---|---|
|[<span data-ttu-id="ec215-749">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="ec215-749">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ec215-750">1.0</span><span class="sxs-lookup"><span data-stu-id="ec215-750">1.0</span></span>|
|[<span data-ttu-id="ec215-751">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ec215-751">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ec215-752">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ec215-752">ReadItem</span></span>|
|[<span data-ttu-id="ec215-753">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ec215-753">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ec215-754">Чтение</span><span class="sxs-lookup"><span data-stu-id="ec215-754">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="ec215-755">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="ec215-755">Returns:</span></span>

<span data-ttu-id="ec215-p146">Если в манифесте нет элемента `ItemHasKnownEntity` со значением `FilterName`, соответствующим параметру `name`, метод возвращает `null`. Если параметр `name` соответствует элементу `ItemHasKnownEntity` в манифесте, но при этом в текущем элементе нет соответствующих сущностей, метод возвращает пустой массив.</span><span class="sxs-lookup"><span data-stu-id="ec215-p146">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>


<span data-ttu-id="ec215-758">Тип: Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="ec215-758">Type: Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))></span></span>


#### <a name="getregexmatches--object"></a><span data-ttu-id="ec215-759">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="ec215-759">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="ec215-760">Возвращает строковые значения в выбранном элементе, которые соответствуют регулярным выражениям, определенным в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="ec215-760">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="ec215-761">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="ec215-761">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="ec215-p147">Метод `getRegExMatches` возвращает строки, соответствующие регулярному выражению, которое определяется в каждом элементе правила `ItemHasRegularExpressionMatch` или `ItemHasKnownEntity` в XML-файле манифеста. Для правила `ItemHasRegularExpressionMatch` соответствующую строку должно содержать свойство элемента, указанного этим правилом. Простой тип `PropertyName` определяет поддерживаемые свойства.</span><span class="sxs-lookup"><span data-stu-id="ec215-p147">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="ec215-765">Например, рассмотрим манифест надстройки, который содержит указанный ниже элемент `Rule`.</span><span class="sxs-lookup"><span data-stu-id="ec215-765">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="ec215-766">Объект, возвращаемый методом `getRegExMatches`, будет содержать два свойства: `fruits` и `veggies`.</span><span class="sxs-lookup"><span data-stu-id="ec215-766">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

> [!NOTE]
> <span data-ttu-id="ec215-p148">Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты.</span><span class="sxs-lookup"><span data-stu-id="ec215-p148">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="requirements"></a><span data-ttu-id="ec215-769">Требования</span><span class="sxs-lookup"><span data-stu-id="ec215-769">Requirements</span></span>

|<span data-ttu-id="ec215-770">Requirement</span><span class="sxs-lookup"><span data-stu-id="ec215-770">Requirement</span></span>| <span data-ttu-id="ec215-771">Значение</span><span class="sxs-lookup"><span data-stu-id="ec215-771">Value</span></span>|
|---|---|
|[<span data-ttu-id="ec215-772">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="ec215-772">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ec215-773">1.0</span><span class="sxs-lookup"><span data-stu-id="ec215-773">1.0</span></span>|
|[<span data-ttu-id="ec215-774">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ec215-774">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ec215-775">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ec215-775">ReadItem</span></span>|
|[<span data-ttu-id="ec215-776">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ec215-776">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ec215-777">Чтение</span><span class="sxs-lookup"><span data-stu-id="ec215-777">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="ec215-778">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="ec215-778">Returns:</span></span>

<span data-ttu-id="ec215-p149">Объект, содержащий массив строк, которые соответствуют регулярным выражениям, определяемым в XML-файле манифеста. Имя каждого массива равно соответствующему значению атрибута `RegExName` подходящего правила `ItemHasRegularExpressionMatch` или атрибута `FilterName` соответствующего правила `ItemHasKnownEntity`.</span><span class="sxs-lookup"><span data-stu-id="ec215-p149">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="ec215-781">

<dt>Тип</dt>

</span><span class="sxs-lookup"><span data-stu-id="ec215-781">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="ec215-782">Object</span><span class="sxs-lookup"><span data-stu-id="ec215-782">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="ec215-783">Пример</span><span class="sxs-lookup"><span data-stu-id="ec215-783">Example</span></span>

<span data-ttu-id="ec215-784">В примере ниже показано, как получить доступ к массиву совпадений для <rule>элементов регулярного выражения `fruits` и `veggies`, которые указаны в манифесте</rule>.</span><span class="sxs-lookup"><span data-stu-id="ec215-784">The following example shows how to access the array of matches for the regular expression <rule>elements `fruits` and `veggies`, which are specified in the manifest.</rule></span></span>

```JavaScript
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veges = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="ec215-785">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="ec215-785">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="ec215-786">Возвращает строковые значения в выбранном элементе, которые соответствуют именованному регулярному выражению, определенному в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="ec215-786">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="ec215-787">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="ec215-787">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="ec215-788">Метод `getRegExMatchesByName` возвращает строки, соответствующие регулярному выражению, которое определяется в элементе правила `ItemHasRegularExpressionMatch` в XML-файле манифеста, с использованием указанного значения элемента `RegExName`.</span><span class="sxs-lookup"><span data-stu-id="ec215-788">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="ec215-p150">Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты.</span><span class="sxs-lookup"><span data-stu-id="ec215-p150">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ec215-791">Параметры</span><span class="sxs-lookup"><span data-stu-id="ec215-791">Parameters:</span></span>

|<span data-ttu-id="ec215-792">Имя</span><span class="sxs-lookup"><span data-stu-id="ec215-792">Name</span></span>| <span data-ttu-id="ec215-793">Тип</span><span class="sxs-lookup"><span data-stu-id="ec215-793">Type</span></span>| <span data-ttu-id="ec215-794">Описание</span><span class="sxs-lookup"><span data-stu-id="ec215-794">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="ec215-795">String</span><span class="sxs-lookup"><span data-stu-id="ec215-795">String</span></span>|<span data-ttu-id="ec215-796">Имя элемента правила `ItemHasRegularExpressionMatch`, определяющее соответствующий фильтр.</span><span class="sxs-lookup"><span data-stu-id="ec215-796">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ec215-797">Требования</span><span class="sxs-lookup"><span data-stu-id="ec215-797">Requirements</span></span>

|<span data-ttu-id="ec215-798">Requirement</span><span class="sxs-lookup"><span data-stu-id="ec215-798">Requirement</span></span>| <span data-ttu-id="ec215-799">Значение</span><span class="sxs-lookup"><span data-stu-id="ec215-799">Value</span></span>|
|---|---|
|[<span data-ttu-id="ec215-800">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="ec215-800">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ec215-801">1.0</span><span class="sxs-lookup"><span data-stu-id="ec215-801">1.0</span></span>|
|[<span data-ttu-id="ec215-802">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ec215-802">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ec215-803">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ec215-803">ReadItem</span></span>|
|[<span data-ttu-id="ec215-804">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ec215-804">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ec215-805">Чтение</span><span class="sxs-lookup"><span data-stu-id="ec215-805">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="ec215-806">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="ec215-806">Returns:</span></span>

<span data-ttu-id="ec215-807">Массив строк, соответствующих регулярному выражению, определяемому в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="ec215-807">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="ec215-808">

<dt>Тип</dt>

</span><span class="sxs-lookup"><span data-stu-id="ec215-808">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="ec215-809">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="ec215-809">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="ec215-810">Пример</span><span class="sxs-lookup"><span data-stu-id="ec215-810">Example</span></span>

```JavaScript
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="ec215-811">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="ec215-811">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="ec215-812">Асинхронно загружает настраиваемые свойства для надстройки для выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="ec215-812">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="ec215-p151">Настраиваемые свойства сохраняются в виде пар "ключ-значение" для каждого приложения и каждого элемента. Этот метод возвращает объект `CustomProperties` при обратном вызове, который предоставляет методы для доступа к настраиваемым свойствам, характерным для текущего элемента и текущей надстройки. Настраиваемые свойства не шифруются для элемента, поэтому этот способ хранения не является безопасным.</span><span class="sxs-lookup"><span data-stu-id="ec215-p151">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ec215-816">Параметры</span><span class="sxs-lookup"><span data-stu-id="ec215-816">Parameters:</span></span>

|<span data-ttu-id="ec215-817">Имя</span><span class="sxs-lookup"><span data-stu-id="ec215-817">Name</span></span>| <span data-ttu-id="ec215-818">Тип</span><span class="sxs-lookup"><span data-stu-id="ec215-818">Type</span></span>| <span data-ttu-id="ec215-819">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="ec215-819">Attributes</span></span>| <span data-ttu-id="ec215-820">Описание</span><span class="sxs-lookup"><span data-stu-id="ec215-820">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="ec215-821">function</span><span class="sxs-lookup"><span data-stu-id="ec215-821">function</span></span>||<span data-ttu-id="ec215-822">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="ec215-822">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="ec215-823">Настраиваемые свойства предоставляются в виде объекта [`CustomProperties`](/javascript/api/outlook_1_1/office.customproperties) в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="ec215-823">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook_1_1/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="ec215-824">Этот объект позволяет получить, задать и удалить настраиваемые свойства из элемента, а также сохранить изменения, внесенные в набор настраиваемых свойств, на сервере.</span><span class="sxs-lookup"><span data-stu-id="ec215-824">The custom properties are provided as a CustomProperties object in the asyncResult.value property. This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="ec215-825">Объект</span><span class="sxs-lookup"><span data-stu-id="ec215-825">Object</span></span>| <span data-ttu-id="ec215-826">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="ec215-826">&lt;optional&gt;</span></span>|<span data-ttu-id="ec215-827">Разработчики могут указать любой объект, к которому необходимо получить доступ, в функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="ec215-827">Developers can provide any object they wish to access in the callback method.</span></span> <span data-ttu-id="ec215-828">Доступ к этому объекту можно получить с помощью свойства `asyncResult.asyncContext` в функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="ec215-828">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.|</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ec215-829">Требования</span><span class="sxs-lookup"><span data-stu-id="ec215-829">Requirements</span></span>

|<span data-ttu-id="ec215-830">Requirement</span><span class="sxs-lookup"><span data-stu-id="ec215-830">Requirement</span></span>| <span data-ttu-id="ec215-831">Значение</span><span class="sxs-lookup"><span data-stu-id="ec215-831">Value</span></span>|
|---|---|
|[<span data-ttu-id="ec215-832">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="ec215-832">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ec215-833">1.0</span><span class="sxs-lookup"><span data-stu-id="ec215-833">1.0</span></span>|
|[<span data-ttu-id="ec215-834">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ec215-834">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ec215-835">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ec215-835">ReadItem</span></span>|
|[<span data-ttu-id="ec215-836">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ec215-836">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ec215-837">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="ec215-837">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="ec215-838">Пример</span><span class="sxs-lookup"><span data-stu-id="ec215-838">Example</span></span>

<span data-ttu-id="ec215-p154">Приведенный ниже пример кода показывает, как асинхронно загружать настраиваемые свойства, характерные для текущего элемента, с помощью метода `loadCustomPropertiesAsync`. Этот пример также показывает, как сохранять эти свойства на сервере с помощью метода `CustomProperties.saveAsync`. После загрузки настраиваемых свойств в этом примере кода метод `CustomProperties.get` используется для считывания настраиваемого свойства `myProp`, метод `CustomProperties.set` — для записи настраиваемого свойства `otherProp`, а метод `saveAsync` — для сохранения настраиваемых свойств.</span><span class="sxs-lookup"><span data-stu-id="ec215-p154">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

```JavaScript
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

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="ec215-842">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="ec215-842">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="ec215-843">Удаляет вложение из сообщения или встречи.</span><span class="sxs-lookup"><span data-stu-id="ec215-843">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="ec215-p155">Метод `removeAttachmentAsync` удаляет из элемента вложение с указанным идентификатором. Идентификатор вложения рекомендуется использовать для удаления вложения, только если оно добавлено тем же почтовым приложением в ходе текущего сеанса. В Outlook Web App и Outlook Web App для устройств идентификатор вложения действителен только в рамках одного сеанса. Сеанс завершается, когда пользователь закрывает приложение или начинает создавать элемент во встроенной форме, а затем переходит из формы в отдельное окно.</span><span class="sxs-lookup"><span data-stu-id="ec215-p155">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ec215-848">Параметры</span><span class="sxs-lookup"><span data-stu-id="ec215-848">Parameters:</span></span>

|<span data-ttu-id="ec215-849">Имя</span><span class="sxs-lookup"><span data-stu-id="ec215-849">Name</span></span>| <span data-ttu-id="ec215-850">Тип</span><span class="sxs-lookup"><span data-stu-id="ec215-850">Type</span></span>| <span data-ttu-id="ec215-851">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="ec215-851">Attributes</span></span>| <span data-ttu-id="ec215-852">Описание</span><span class="sxs-lookup"><span data-stu-id="ec215-852">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="ec215-853">String</span><span class="sxs-lookup"><span data-stu-id="ec215-853">String</span></span>||<span data-ttu-id="ec215-p156">Идентификатор удаляемого вложения. Максимальная длина строки — 100 символов.</span><span class="sxs-lookup"><span data-stu-id="ec215-p156">The identifier of the attachment to remove. The maximum length of the string is 100 characters.</span></span>|
|`options`| <span data-ttu-id="ec215-856">Object</span><span class="sxs-lookup"><span data-stu-id="ec215-856">Object</span></span>| <span data-ttu-id="ec215-857">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="ec215-857">&lt;optional&gt;</span></span>|<span data-ttu-id="ec215-858">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="ec215-858">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="ec215-859">Object</span><span class="sxs-lookup"><span data-stu-id="ec215-859">Object</span></span>| <span data-ttu-id="ec215-860">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="ec215-860">&lt;optional&gt;</span></span>|<span data-ttu-id="ec215-861">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="ec215-861">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="ec215-862">функция</span><span class="sxs-lookup"><span data-stu-id="ec215-862">function</span></span>| <span data-ttu-id="ec215-863">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="ec215-863">&lt;optional&gt;</span></span>|<span data-ttu-id="ec215-864">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="ec215-864">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="ec215-865">Если удалить вложение не удается, свойство `asyncResult.error` содержит код ошибки с указанием ее причины.</span><span class="sxs-lookup"><span data-stu-id="ec215-865">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="ec215-866">Ошибки</span><span class="sxs-lookup"><span data-stu-id="ec215-866">Errors</span></span>

| <span data-ttu-id="ec215-867">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="ec215-867">Error code</span></span> | <span data-ttu-id="ec215-868">Описание</span><span class="sxs-lookup"><span data-stu-id="ec215-868">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="ec215-869">Идентификатор вложения не существует.</span><span class="sxs-lookup"><span data-stu-id="ec215-869">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="ec215-870">Требования</span><span class="sxs-lookup"><span data-stu-id="ec215-870">Requirements</span></span>

|<span data-ttu-id="ec215-871">Requirement</span><span class="sxs-lookup"><span data-stu-id="ec215-871">Requirement</span></span>| <span data-ttu-id="ec215-872">Значение</span><span class="sxs-lookup"><span data-stu-id="ec215-872">Value</span></span>|
|---|---|
|[<span data-ttu-id="ec215-873">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="ec215-873">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ec215-874">1.1</span><span class="sxs-lookup"><span data-stu-id="ec215-874">1.1</span></span>|
|[<span data-ttu-id="ec215-875">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ec215-875">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ec215-876">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="ec215-876">ReadWriteItem</span></span>|
|[<span data-ttu-id="ec215-877">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ec215-877">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ec215-878">Создание</span><span class="sxs-lookup"><span data-stu-id="ec215-878">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="ec215-879">Пример</span><span class="sxs-lookup"><span data-stu-id="ec215-879">Example</span></span>

<span data-ttu-id="ec215-880">Указанный ниже код удаляет вложение с идентификатором "0".</span><span class="sxs-lookup"><span data-stu-id="ec215-880">The following code removes an attachment with an identifier of '0'.</span></span>

```JavaScript
Office.context.mailbox.item.removeAttachmentAsync(
  '0',
  { asyncContext : null },
  function (asyncResult)
  {
    console.log(asyncResult.status);
  }
);
```