
# <a name="item"></a><span data-ttu-id="b621d-101">item</span><span class="sxs-lookup"><span data-stu-id="b621d-101">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="b621d-102">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="b621d-102">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="b621d-p101">Пространство имен `item` используется для доступа к выбранному в данный момент сообщению, приглашению на собрание или описанию встречи. Вы можете определить тип пространства имен `item` с помощью свойства [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook13officemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="b621d-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook13officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="b621d-105">Requirements</span><span class="sxs-lookup"><span data-stu-id="b621d-105">Requirements</span></span>

|<span data-ttu-id="b621d-106">Requirement</span><span class="sxs-lookup"><span data-stu-id="b621d-106">Requirement</span></span>| <span data-ttu-id="b621d-107">Значение</span><span class="sxs-lookup"><span data-stu-id="b621d-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="b621d-108">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="b621d-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b621d-109">1.0</span><span class="sxs-lookup"><span data-stu-id="b621d-109">1.0</span></span>|
|[<span data-ttu-id="b621d-110">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b621d-110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b621d-111">Restricted</span><span class="sxs-lookup"><span data-stu-id="b621d-111">Restricted</span></span>|
|[<span data-ttu-id="b621d-112">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b621d-112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b621d-113">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="b621d-113">Compose or read</span></span>|

### <a name="example"></a><span data-ttu-id="b621d-114">Пример</span><span class="sxs-lookup"><span data-stu-id="b621d-114">Example</span></span>

<span data-ttu-id="b621d-115">В примере кода JavaScript, приведенном ниже, показано, как получить доступ к свойству `subject` текущего элемента в Outlook.</span><span class="sxs-lookup"><span data-stu-id="b621d-115">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="b621d-116">Элементы</span><span class="sxs-lookup"><span data-stu-id="b621d-116">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook13officeattachmentdetails"></a><span data-ttu-id="b621d-117">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_3/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="b621d-117">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_3/office.attachmentdetails)></span></span>

<span data-ttu-id="b621d-p102">Получает массив вложений для элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="b621d-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="b621d-120">Файлы некоторых типов блокируются в Outlook из-за возможных проблем с безопасностью и поэтому не возвращаются.</span><span class="sxs-lookup"><span data-stu-id="b621d-120">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="b621d-121">Дополнительные сведения см. в статье [Блокировка вложений в Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="b621d-121">For more information see [Payments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="b621d-122">Тип:</span><span class="sxs-lookup"><span data-stu-id="b621d-122">Type:</span></span>

*   <span data-ttu-id="b621d-123">Array.<[AttachmentDetails](/javascript/api/outlook_1_3/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="b621d-123">Array.<[AttachmentDetails](/javascript/api/outlook_1_3/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="b621d-124">Требования</span><span class="sxs-lookup"><span data-stu-id="b621d-124">Requirements</span></span>

|<span data-ttu-id="b621d-125">Requirement</span><span class="sxs-lookup"><span data-stu-id="b621d-125">Requirement</span></span>| <span data-ttu-id="b621d-126">Значение</span><span class="sxs-lookup"><span data-stu-id="b621d-126">Value</span></span>|
|---|---|
|[<span data-ttu-id="b621d-127">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="b621d-127">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b621d-128">1.0</span><span class="sxs-lookup"><span data-stu-id="b621d-128">1.0</span></span>|
|[<span data-ttu-id="b621d-129">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b621d-129">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b621d-130">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b621d-130">ReadItem</span></span>|
|[<span data-ttu-id="b621d-131">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b621d-131">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b621d-132">Чтение</span><span class="sxs-lookup"><span data-stu-id="b621d-132">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b621d-133">Пример</span><span class="sxs-lookup"><span data-stu-id="b621d-133">Example</span></span>

<span data-ttu-id="b621d-134">С помощью приведенного ниже кода можно создать HTML-строку с подробными сведениями обо всех вложениях для текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="b621d-134">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

####  <a name="bcc-recipientsjavascriptapioutlook13officerecipients"></a><span data-ttu-id="b621d-135">bcc :[Recipients](/javascript/api/outlook_1_3/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="b621d-135">bcc :[Recipients](/javascript/api/outlook_1_3/office.recipients)</span></span>

<span data-ttu-id="b621d-136">Извлекает объект, предоставляющий методы для получения или обновления получателей, которые указаны в строке СК (скрытая копия) сообщения.</span><span class="sxs-lookup"><span data-stu-id="b621d-136">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="b621d-137">Только в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="b621d-137">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="b621d-138">Тип:</span><span class="sxs-lookup"><span data-stu-id="b621d-138">Type:</span></span>

*   [<span data-ttu-id="b621d-139">Recipients</span><span class="sxs-lookup"><span data-stu-id="b621d-139">Recipients</span></span>](/javascript/api/outlook_1_3/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="b621d-140">Требования</span><span class="sxs-lookup"><span data-stu-id="b621d-140">Requirements</span></span>

|<span data-ttu-id="b621d-141">Requirement</span><span class="sxs-lookup"><span data-stu-id="b621d-141">Requirement</span></span>| <span data-ttu-id="b621d-142">Значение</span><span class="sxs-lookup"><span data-stu-id="b621d-142">Value</span></span>|
|---|---|
|[<span data-ttu-id="b621d-143">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="b621d-143">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b621d-144">1.1</span><span class="sxs-lookup"><span data-stu-id="b621d-144">1.1</span></span>|
|[<span data-ttu-id="b621d-145">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b621d-145">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b621d-146">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b621d-146">ReadItem</span></span>|
|[<span data-ttu-id="b621d-147">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b621d-147">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b621d-148">Создание</span><span class="sxs-lookup"><span data-stu-id="b621d-148">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="b621d-149">Пример</span><span class="sxs-lookup"><span data-stu-id="b621d-149">Example</span></span>

```js
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlook13officebody"></a><span data-ttu-id="b621d-150">body :[Body](/javascript/api/outlook_1_3/office.body)</span><span class="sxs-lookup"><span data-stu-id="b621d-150">body :[Body](/javascript/api/outlook_1_3/office.body)</span></span>

<span data-ttu-id="b621d-151">Получает объект, предоставляющий методы для работы с основным текстом элемента.</span><span class="sxs-lookup"><span data-stu-id="b621d-151">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="b621d-152">Тип:</span><span class="sxs-lookup"><span data-stu-id="b621d-152">Type:</span></span>

*   [<span data-ttu-id="b621d-153">Body</span><span class="sxs-lookup"><span data-stu-id="b621d-153">Body</span></span>](/javascript/api/outlook_1_3/office.body)

##### <a name="requirements"></a><span data-ttu-id="b621d-154">Требования</span><span class="sxs-lookup"><span data-stu-id="b621d-154">Requirements</span></span>

|<span data-ttu-id="b621d-155">Requirement</span><span class="sxs-lookup"><span data-stu-id="b621d-155">Requirement</span></span>| <span data-ttu-id="b621d-156">Значение</span><span class="sxs-lookup"><span data-stu-id="b621d-156">Value</span></span>|
|---|---|
|[<span data-ttu-id="b621d-157">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="b621d-157">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b621d-158">1.1</span><span class="sxs-lookup"><span data-stu-id="b621d-158">1.1</span></span>|
|[<span data-ttu-id="b621d-159">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b621d-159">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b621d-160">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b621d-160">ReadItem</span></span>|
|[<span data-ttu-id="b621d-161">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b621d-161">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b621d-162">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="b621d-162">Compose or read</span></span>|

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlook13officeemailaddressdetailsrecipientsjavascriptapioutlook13officerecipients"></a><span data-ttu-id="b621d-163">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_3/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="b621d-163">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_3/office.recipients)</span></span>

<span data-ttu-id="b621d-164">Предоставляет доступ к получателям копии сообщения.</span><span class="sxs-lookup"><span data-stu-id="b621d-164">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="b621d-165">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="b621d-165">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="b621d-166">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="b621d-166">Read mode</span></span>

<span data-ttu-id="b621d-p106">Свойство `cc` возвращает массив, который содержит объект `EmailAddressDetails` для каждого получателя, указанного в строке **Копия** сообщения. Коллекция может включать не более 100 элементов.</span><span class="sxs-lookup"><span data-stu-id="b621d-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="b621d-169">Режим создания</span><span class="sxs-lookup"><span data-stu-id="b621d-169">Compose mode</span></span>

<span data-ttu-id="b621d-170">Свойство `cc` возвращает объект `Recipients`, предоставляющий методы для получения или обновления получателей, которые указаны в строке **Копия** сообщения.</span><span class="sxs-lookup"><span data-stu-id="b621d-170">The `cc` property returns a `Recipients` object that provides methods for manipulating the recipients on the **Cc** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="b621d-171">Тип:</span><span class="sxs-lookup"><span data-stu-id="b621d-171">Type:</span></span>

*   <span data-ttu-id="b621d-172">Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_3/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="b621d-172">Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_3/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="b621d-173">Требования</span><span class="sxs-lookup"><span data-stu-id="b621d-173">Requirements</span></span>

|<span data-ttu-id="b621d-174">Requirement</span><span class="sxs-lookup"><span data-stu-id="b621d-174">Requirement</span></span>| <span data-ttu-id="b621d-175">Значение</span><span class="sxs-lookup"><span data-stu-id="b621d-175">Value</span></span>|
|---|---|
|[<span data-ttu-id="b621d-176">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="b621d-176">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b621d-177">1.0</span><span class="sxs-lookup"><span data-stu-id="b621d-177">1.0</span></span>|
|[<span data-ttu-id="b621d-178">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b621d-178">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b621d-179">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b621d-179">ReadItem</span></span>|
|[<span data-ttu-id="b621d-180">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b621d-180">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b621d-181">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="b621d-181">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="b621d-182">Пример</span><span class="sxs-lookup"><span data-stu-id="b621d-182">Example</span></span>

```js
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="b621d-183">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="b621d-183">(nullable) conversationId :String</span></span>

<span data-ttu-id="b621d-184">Получает идентификатор разговора по электронной почте, содержащего конкретное сообщение.</span><span class="sxs-lookup"><span data-stu-id="b621d-184">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="b621d-p107">Вы можете получить целочисленное значение этого свойства, если ваше почтовое приложение активируется в формах просмотра или формах создания ответов. Если пользователь изменит тему ответа, после его отправки идентификатор беседы будет изменен, и полученное ранее значение будет недействительным.</span><span class="sxs-lookup"><span data-stu-id="b621d-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="b621d-p108">Это свойство имеет значение NULL для нового элемента в форме создания. Свойство `conversationId` вернет значение, если пользователь задаст тему и сохранит элемент.</span><span class="sxs-lookup"><span data-stu-id="b621d-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="b621d-189">Тип:</span><span class="sxs-lookup"><span data-stu-id="b621d-189">Type:</span></span>

*   <span data-ttu-id="b621d-190">String</span><span class="sxs-lookup"><span data-stu-id="b621d-190">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="b621d-191">Требования</span><span class="sxs-lookup"><span data-stu-id="b621d-191">Requirements</span></span>

|<span data-ttu-id="b621d-192">Requirement</span><span class="sxs-lookup"><span data-stu-id="b621d-192">Requirement</span></span>| <span data-ttu-id="b621d-193">Значение</span><span class="sxs-lookup"><span data-stu-id="b621d-193">Value</span></span>|
|---|---|
|[<span data-ttu-id="b621d-194">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="b621d-194">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b621d-195">1.0</span><span class="sxs-lookup"><span data-stu-id="b621d-195">1.0</span></span>|
|[<span data-ttu-id="b621d-196">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b621d-196">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b621d-197">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b621d-197">ReadItem</span></span>|
|[<span data-ttu-id="b621d-198">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b621d-198">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b621d-199">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="b621d-199">Compose or read</span></span>|

#### <a name="datetimecreated-date"></a><span data-ttu-id="b621d-200">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="b621d-200">dateTimeCreated :Date</span></span>

<span data-ttu-id="b621d-p109">Получает дату и время создания элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="b621d-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="b621d-203">Тип:</span><span class="sxs-lookup"><span data-stu-id="b621d-203">Type:</span></span>

*   <span data-ttu-id="b621d-204">Date</span><span class="sxs-lookup"><span data-stu-id="b621d-204">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="b621d-205">Требования</span><span class="sxs-lookup"><span data-stu-id="b621d-205">Requirements</span></span>

|<span data-ttu-id="b621d-206">Requirement</span><span class="sxs-lookup"><span data-stu-id="b621d-206">Requirement</span></span>| <span data-ttu-id="b621d-207">Значение</span><span class="sxs-lookup"><span data-stu-id="b621d-207">Value</span></span>|
|---|---|
|[<span data-ttu-id="b621d-208">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="b621d-208">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b621d-209">1.0</span><span class="sxs-lookup"><span data-stu-id="b621d-209">1.0</span></span>|
|[<span data-ttu-id="b621d-210">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b621d-210">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b621d-211">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b621d-211">ReadItem</span></span>|
|[<span data-ttu-id="b621d-212">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b621d-212">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b621d-213">Чтение</span><span class="sxs-lookup"><span data-stu-id="b621d-213">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b621d-214">Пример</span><span class="sxs-lookup"><span data-stu-id="b621d-214">Example</span></span>

```js
var created = Office.context.mailbox.item.dateTimeCreated;
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="b621d-215">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="b621d-215">dateTimeModified :Date</span></span>

<span data-ttu-id="b621d-p110">Получает дату и время последнего изменения элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="b621d-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="b621d-218">Этот элемент не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="b621d-218">Note: This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="b621d-219">Тип:</span><span class="sxs-lookup"><span data-stu-id="b621d-219">Type:</span></span>

*   <span data-ttu-id="b621d-220">Date</span><span class="sxs-lookup"><span data-stu-id="b621d-220">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="b621d-221">Требования</span><span class="sxs-lookup"><span data-stu-id="b621d-221">Requirements</span></span>

|<span data-ttu-id="b621d-222">Requirement</span><span class="sxs-lookup"><span data-stu-id="b621d-222">Requirement</span></span>| <span data-ttu-id="b621d-223">Значение</span><span class="sxs-lookup"><span data-stu-id="b621d-223">Value</span></span>|
|---|---|
|[<span data-ttu-id="b621d-224">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="b621d-224">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b621d-225">1.0</span><span class="sxs-lookup"><span data-stu-id="b621d-225">1.0</span></span>|
|[<span data-ttu-id="b621d-226">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b621d-226">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b621d-227">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b621d-227">ReadItem</span></span>|
|[<span data-ttu-id="b621d-228">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b621d-228">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b621d-229">Чтение</span><span class="sxs-lookup"><span data-stu-id="b621d-229">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b621d-230">Пример</span><span class="sxs-lookup"><span data-stu-id="b621d-230">Example</span></span>

```js
var modified = Office.context.mailbox.item.dateTimeModified;
```

####  <a name="end-datetimejavascriptapioutlook13officetime"></a><span data-ttu-id="b621d-231">end :Date|[Time](/javascript/api/outlook_1_3/office.time)</span><span class="sxs-lookup"><span data-stu-id="b621d-231">end :Date|[Time](/javascript/api/outlook_1_3/office.time)</span></span>

<span data-ttu-id="b621d-232">Получает или задает дату и время окончания встречи.</span><span class="sxs-lookup"><span data-stu-id="b621d-232">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="b621d-p111">Свойство `end` представлено в виде значения даты и времени в формате UTC. Преобразовать значение свойства end в местные значения даты и времени клиента можно с помощью метода [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook13officelocalclienttime).</span><span class="sxs-lookup"><span data-stu-id="b621d-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook13officelocalclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="b621d-235">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="b621d-235">Read mode</span></span>

<span data-ttu-id="b621d-236">Свойство `end` возвращает объект `Date`.</span><span class="sxs-lookup"><span data-stu-id="b621d-236">The `end` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="b621d-237">Режим создания</span><span class="sxs-lookup"><span data-stu-id="b621d-237">Compose mode</span></span>

<span data-ttu-id="b621d-238">Свойство `end` возвращает объект `Time`.</span><span class="sxs-lookup"><span data-stu-id="b621d-238">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="b621d-239">Если вы задаете время окончания с помощью метода [`Time.setAsync`](/javascript/api/outlook_1_3/office.time#setasync-datetime--options--callback-), необходимо использовать метод [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) для преобразования местного времени на клиенте в формат UTC для сервера.</span><span class="sxs-lookup"><span data-stu-id="b621d-239">When you use the [`Time.setAsync`](/javascript/api/outlook_1_3/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="b621d-240">Тип:</span><span class="sxs-lookup"><span data-stu-id="b621d-240">Type:</span></span>

*   <span data-ttu-id="b621d-241">Date | [Time](/javascript/api/outlook_1_3/office.time)</span><span class="sxs-lookup"><span data-stu-id="b621d-241">Date | [Time](/javascript/api/outlook_1_3/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="b621d-242">Требования</span><span class="sxs-lookup"><span data-stu-id="b621d-242">Requirements</span></span>

|<span data-ttu-id="b621d-243">Requirement</span><span class="sxs-lookup"><span data-stu-id="b621d-243">Requirement</span></span>| <span data-ttu-id="b621d-244">Значение</span><span class="sxs-lookup"><span data-stu-id="b621d-244">Value</span></span>|
|---|---|
|[<span data-ttu-id="b621d-245">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="b621d-245">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b621d-246">1.0</span><span class="sxs-lookup"><span data-stu-id="b621d-246">1.0</span></span>|
|[<span data-ttu-id="b621d-247">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b621d-247">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b621d-248">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b621d-248">ReadItem</span></span>|
|[<span data-ttu-id="b621d-249">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b621d-249">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b621d-250">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="b621d-250">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="b621d-251">Пример</span><span class="sxs-lookup"><span data-stu-id="b621d-251">Example</span></span>

<span data-ttu-id="b621d-252">В примере ниже показано, как с помощью метода [`setAsync`](/javascript/api/outlook_1_3/office.time#setasync-datetime--options--callback-) объекта `Time` задать время окончания встречи в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="b621d-252">The following example sets the end time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_3/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

#### <a name="from-emailaddressdetailsjavascriptapioutlook13officeemailaddressdetails"></a><span data-ttu-id="b621d-253">from :[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="b621d-253">from :[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)</span></span>

<span data-ttu-id="b621d-p112">Получает электронный адрес отправителя сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="b621d-p112">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="b621d-p113">Свойства `from` и [`sender`](#sender-emailaddressdetailsjavascriptapioutlook13officeemailaddressdetails) представляют одно лицо, если сообщение не отправлено представителем. В противном случае свойство `from` представляет лицо, делегировавшее полномочия, а свойство sender — представителя.</span><span class="sxs-lookup"><span data-stu-id="b621d-p113">The `from` and [`sender`](#sender-emailaddressdetailsjavascriptapioutlook13officeemailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="b621d-258">Свойству `recipientType`, принадлежащему объекту `EmailAddressDetails` в свойстве `from`, задано значение `undefined`.</span><span class="sxs-lookup"><span data-stu-id="b621d-258">Note: The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="b621d-259">Тип:</span><span class="sxs-lookup"><span data-stu-id="b621d-259">Type:</span></span>

*   [<span data-ttu-id="b621d-260">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="b621d-260">EmailAddressDetails</span></span>](/javascript/api/outlook_1_3/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="b621d-261">Требования</span><span class="sxs-lookup"><span data-stu-id="b621d-261">Requirements</span></span>

|<span data-ttu-id="b621d-262">Requirement</span><span class="sxs-lookup"><span data-stu-id="b621d-262">Requirement</span></span>| <span data-ttu-id="b621d-263">Значение</span><span class="sxs-lookup"><span data-stu-id="b621d-263">Value</span></span>|
|---|---|
|[<span data-ttu-id="b621d-264">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="b621d-264">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b621d-265">1.0</span><span class="sxs-lookup"><span data-stu-id="b621d-265">1.0</span></span>|
|[<span data-ttu-id="b621d-266">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b621d-266">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b621d-267">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b621d-267">ReadItem</span></span>|
|[<span data-ttu-id="b621d-268">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b621d-268">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b621d-269">Чтение</span><span class="sxs-lookup"><span data-stu-id="b621d-269">Read</span></span>|

#### <a name="internetmessageid-string"></a><span data-ttu-id="b621d-270">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="b621d-270">internetMessageId :String</span></span>

<span data-ttu-id="b621d-p114">Получает идентификатор интернет-сообщения для электронного сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="b621d-p114">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="b621d-273">Тип:</span><span class="sxs-lookup"><span data-stu-id="b621d-273">Type:</span></span>

*   <span data-ttu-id="b621d-274">String</span><span class="sxs-lookup"><span data-stu-id="b621d-274">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="b621d-275">Требования</span><span class="sxs-lookup"><span data-stu-id="b621d-275">Requirements</span></span>

|<span data-ttu-id="b621d-276">Requirement</span><span class="sxs-lookup"><span data-stu-id="b621d-276">Requirement</span></span>| <span data-ttu-id="b621d-277">Значение</span><span class="sxs-lookup"><span data-stu-id="b621d-277">Value</span></span>|
|---|---|
|[<span data-ttu-id="b621d-278">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="b621d-278">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b621d-279">1.0</span><span class="sxs-lookup"><span data-stu-id="b621d-279">1.0</span></span>|
|[<span data-ttu-id="b621d-280">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b621d-280">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b621d-281">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b621d-281">ReadItem</span></span>|
|[<span data-ttu-id="b621d-282">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b621d-282">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b621d-283">Чтение</span><span class="sxs-lookup"><span data-stu-id="b621d-283">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b621d-284">Пример</span><span class="sxs-lookup"><span data-stu-id="b621d-284">Example</span></span>

```js
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="b621d-285">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="b621d-285">itemClass :String</span></span>

<span data-ttu-id="b621d-p115">Получает класс элемента веб-служб Exchange для выбранного элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="b621d-p115">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="b621d-p116">Свойство `itemClass` указывает класс сообщения выбранного элемента. Ниже приводятся классы сообщения по умолчанию для элемента сообщения или встречи.</span><span class="sxs-lookup"><span data-stu-id="b621d-p116">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="b621d-290">Тип</span><span class="sxs-lookup"><span data-stu-id="b621d-290">Type</span></span> | <span data-ttu-id="b621d-291">Описание</span><span class="sxs-lookup"><span data-stu-id="b621d-291">Description</span></span> | <span data-ttu-id="b621d-292">Класс элемента</span><span class="sxs-lookup"><span data-stu-id="b621d-292">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="b621d-293">Элементы встречи</span><span class="sxs-lookup"><span data-stu-id="b621d-293">Appointment items</span></span> | <span data-ttu-id="b621d-294">Это элементы календаря для класса элемента `IPM.Appointment` или `IPM.Appointment.Occurence`.</span><span class="sxs-lookup"><span data-stu-id="b621d-294">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurence` |
| <span data-ttu-id="b621d-295">Элементы сообщения</span><span class="sxs-lookup"><span data-stu-id="b621d-295">Message items</span></span> | <span data-ttu-id="b621d-296">Сюда входят электронные сообщения, для которых по умолчанию задан класс сообщения `IPM.Note`, а также приглашения на собрания, ответы на них и уведомления об их отмене, использующие `IPM.Schedule.Meeting` в качестве базового класса сообщения.</span><span class="sxs-lookup"><span data-stu-id="b621d-296">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="b621d-297">Можно создавать настраиваемые классы сообщения, расширяющие классы сообщения по умолчанию, например настраиваемый класс сообщения о встрече `IPM.Appointment.Contoso`.</span><span class="sxs-lookup"><span data-stu-id="b621d-297">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="b621d-298">Тип:</span><span class="sxs-lookup"><span data-stu-id="b621d-298">Type:</span></span>

*   <span data-ttu-id="b621d-299">String</span><span class="sxs-lookup"><span data-stu-id="b621d-299">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="b621d-300">Требования</span><span class="sxs-lookup"><span data-stu-id="b621d-300">Requirements</span></span>

|<span data-ttu-id="b621d-301">Requirement</span><span class="sxs-lookup"><span data-stu-id="b621d-301">Requirement</span></span>| <span data-ttu-id="b621d-302">Значение</span><span class="sxs-lookup"><span data-stu-id="b621d-302">Value</span></span>|
|---|---|
|[<span data-ttu-id="b621d-303">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="b621d-303">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b621d-304">1.0</span><span class="sxs-lookup"><span data-stu-id="b621d-304">1.0</span></span>|
|[<span data-ttu-id="b621d-305">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b621d-305">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b621d-306">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b621d-306">ReadItem</span></span>|
|[<span data-ttu-id="b621d-307">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b621d-307">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b621d-308">Чтение</span><span class="sxs-lookup"><span data-stu-id="b621d-308">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b621d-309">Пример</span><span class="sxs-lookup"><span data-stu-id="b621d-309">Example</span></span>

```js
var itemClass = Office.context.mailbox.item.itemClass;
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="b621d-310">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="b621d-310">(nullable) itemId :String</span></span>

<span data-ttu-id="b621d-p117">Получает идентификатор элемента веб-служб Exchange для текущего элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="b621d-p117">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="b621d-313">Идентификатор, возвращаемый свойством `itemId`, совпадает с идентификатором элемента веб-служб Exchange.</span><span class="sxs-lookup"><span data-stu-id="b621d-313">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="b621d-314">Свойство `itemId` не совпадает с идентификатором записи Outlook, а также идентификатором, который используется REST API Outlook.</span><span class="sxs-lookup"><span data-stu-id="b621d-314">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="b621d-315">Прежде чем совершать вызовы REST API, используя это значение, его необходимо преобразовать с помощью [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="b621d-315">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="b621d-316">Дополнительные сведения см. в статье [Использование REST API Outlook из надстройки Outlook](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="b621d-316">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="b621d-p119">Свойство `itemId` недоступно в режиме создания. Если требуется идентификатор элемента, с помощью метода [`saveAsync`](#saveasyncoptions-callback) можно сохранить элемент в хранилище. При этом в параметре [`AsyncResult.value`](/javascript/api/office/office.asyncresult) функции обратного вызова возвращается идентификатор элемента.</span><span class="sxs-lookup"><span data-stu-id="b621d-p119">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="b621d-319">Тип:</span><span class="sxs-lookup"><span data-stu-id="b621d-319">Type:</span></span>

*   <span data-ttu-id="b621d-320">String</span><span class="sxs-lookup"><span data-stu-id="b621d-320">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="b621d-321">Требования</span><span class="sxs-lookup"><span data-stu-id="b621d-321">Requirements</span></span>

|<span data-ttu-id="b621d-322">Requirement</span><span class="sxs-lookup"><span data-stu-id="b621d-322">Requirement</span></span>| <span data-ttu-id="b621d-323">Значение</span><span class="sxs-lookup"><span data-stu-id="b621d-323">Value</span></span>|
|---|---|
|[<span data-ttu-id="b621d-324">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="b621d-324">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b621d-325">1.0</span><span class="sxs-lookup"><span data-stu-id="b621d-325">1.0</span></span>|
|[<span data-ttu-id="b621d-326">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b621d-326">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b621d-327">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b621d-327">ReadItem</span></span>|
|[<span data-ttu-id="b621d-328">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b621d-328">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b621d-329">Чтение</span><span class="sxs-lookup"><span data-stu-id="b621d-329">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b621d-330">Пример</span><span class="sxs-lookup"><span data-stu-id="b621d-330">Example</span></span>

<span data-ttu-id="b621d-p120">Указанный ниже код проверяет наличие идентификатора элемента. Если свойство `itemId` возвращает значение `null` или `undefined`, элемент будет сохранен в хранилище, а из асинхронного результата будет получен идентификатор элемента.</span><span class="sxs-lookup"><span data-stu-id="b621d-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```js
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result){
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook13officemailboxenumsitemtype"></a><span data-ttu-id="b621d-333">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_3/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="b621d-333">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_3/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="b621d-334">Получает тип элемента, который представляет экземпляр.</span><span class="sxs-lookup"><span data-stu-id="b621d-334">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="b621d-335">Свойство `itemType` возвращает одно из значений перечисления `ItemType`, которое указывает, является ли экземпляр объекта `item` сообщением или встречей.</span><span class="sxs-lookup"><span data-stu-id="b621d-335">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="b621d-336">Тип:</span><span class="sxs-lookup"><span data-stu-id="b621d-336">Type:</span></span>

*   [<span data-ttu-id="b621d-337">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="b621d-337">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook_1_3/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="b621d-338">Требования</span><span class="sxs-lookup"><span data-stu-id="b621d-338">Requirements</span></span>

|<span data-ttu-id="b621d-339">Requirement</span><span class="sxs-lookup"><span data-stu-id="b621d-339">Requirement</span></span>| <span data-ttu-id="b621d-340">Значение</span><span class="sxs-lookup"><span data-stu-id="b621d-340">Value</span></span>|
|---|---|
|[<span data-ttu-id="b621d-341">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="b621d-341">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b621d-342">1.0</span><span class="sxs-lookup"><span data-stu-id="b621d-342">1.0</span></span>|
|[<span data-ttu-id="b621d-343">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b621d-343">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b621d-344">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b621d-344">ReadItem</span></span>|
|[<span data-ttu-id="b621d-345">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b621d-345">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b621d-346">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="b621d-346">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="b621d-347">Пример</span><span class="sxs-lookup"><span data-stu-id="b621d-347">Example</span></span>

```js
if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Message)
  // do something
else
  // do something else
```

####  <a name="location-stringlocationjavascriptapioutlook13officelocation"></a><span data-ttu-id="b621d-348">location :String|[Location](/javascript/api/outlook_1_3/office.location)</span><span class="sxs-lookup"><span data-stu-id="b621d-348">location :String|[Location](/javascript/api/outlook_1_3/office.location)</span></span>

<span data-ttu-id="b621d-349">Получает или задает место встречи.</span><span class="sxs-lookup"><span data-stu-id="b621d-349">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="b621d-350">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="b621d-350">Read mode</span></span>

<span data-ttu-id="b621d-351">Свойство `location` возвращает строку, содержащую сведения о месте встречи.</span><span class="sxs-lookup"><span data-stu-id="b621d-351">The `location` property returns a string that contains the location of the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="b621d-352">Режим создания</span><span class="sxs-lookup"><span data-stu-id="b621d-352">Compose mode</span></span>

<span data-ttu-id="b621d-353">Свойство `location` возвращает объект `Location`, предоставляющий методы, которые используются для получения и задания места встречи.</span><span class="sxs-lookup"><span data-stu-id="b621d-353">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="b621d-354">Тип:</span><span class="sxs-lookup"><span data-stu-id="b621d-354">Type:</span></span>

*   <span data-ttu-id="b621d-355">String | [Location](/javascript/api/outlook_1_3/office.location)</span><span class="sxs-lookup"><span data-stu-id="b621d-355">String | [Location](/javascript/api/outlook_1_3/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="b621d-356">Требования</span><span class="sxs-lookup"><span data-stu-id="b621d-356">Requirements</span></span>

|<span data-ttu-id="b621d-357">Requirement</span><span class="sxs-lookup"><span data-stu-id="b621d-357">Requirement</span></span>| <span data-ttu-id="b621d-358">Значение</span><span class="sxs-lookup"><span data-stu-id="b621d-358">Value</span></span>|
|---|---|
|[<span data-ttu-id="b621d-359">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="b621d-359">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b621d-360">1.0</span><span class="sxs-lookup"><span data-stu-id="b621d-360">1.0</span></span>|
|[<span data-ttu-id="b621d-361">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b621d-361">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b621d-362">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b621d-362">ReadItem</span></span>|
|[<span data-ttu-id="b621d-363">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b621d-363">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b621d-364">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="b621d-364">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="b621d-365">Пример</span><span class="sxs-lookup"><span data-stu-id="b621d-365">Example</span></span>

```js
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

#### <a name="normalizedsubject-string"></a><span data-ttu-id="b621d-366">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="b621d-366">normalizedSubject :String</span></span>

<span data-ttu-id="b621d-p121">Получает тему элемента со всеми удаленными префиксами (включая `RE:` и `FWD:`). Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="b621d-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="b621d-p122">Свойство normalizedSubject получает тему элемента со стандартными префиксами (такими как `RE:` и `FW:`), добавляемыми почтовыми программами. Для получения темы элемента с неизмененными префиксами используйте свойство [`subject`](#subject-stringsubjectjavascriptapioutlook13officesubject).</span><span class="sxs-lookup"><span data-stu-id="b621d-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubjectjavascriptapioutlook13officesubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="b621d-371">Тип:</span><span class="sxs-lookup"><span data-stu-id="b621d-371">Type:</span></span>

*   <span data-ttu-id="b621d-372">String</span><span class="sxs-lookup"><span data-stu-id="b621d-372">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="b621d-373">Требования</span><span class="sxs-lookup"><span data-stu-id="b621d-373">Requirements</span></span>

|<span data-ttu-id="b621d-374">Requirement</span><span class="sxs-lookup"><span data-stu-id="b621d-374">Requirement</span></span>| <span data-ttu-id="b621d-375">Значение</span><span class="sxs-lookup"><span data-stu-id="b621d-375">Value</span></span>|
|---|---|
|[<span data-ttu-id="b621d-376">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="b621d-376">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b621d-377">1.0</span><span class="sxs-lookup"><span data-stu-id="b621d-377">1.0</span></span>|
|[<span data-ttu-id="b621d-378">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b621d-378">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b621d-379">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b621d-379">ReadItem</span></span>|
|[<span data-ttu-id="b621d-380">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b621d-380">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b621d-381">Чтение</span><span class="sxs-lookup"><span data-stu-id="b621d-381">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b621d-382">Пример</span><span class="sxs-lookup"><span data-stu-id="b621d-382">Example</span></span>

```js
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
```

####  <a name="notificationmessages-notificationmessagesjavascriptapioutlook13officenotificationmessages"></a><span data-ttu-id="b621d-383">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_3/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="b621d-383">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_3/office.notificationmessages)</span></span>

<span data-ttu-id="b621d-384">Получает сообщения уведомления для элемента.</span><span class="sxs-lookup"><span data-stu-id="b621d-384">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="b621d-385">Тип:</span><span class="sxs-lookup"><span data-stu-id="b621d-385">Type:</span></span>

*   [<span data-ttu-id="b621d-386">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="b621d-386">NotificationMessages</span></span>](/javascript/api/outlook_1_3/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="b621d-387">Требования</span><span class="sxs-lookup"><span data-stu-id="b621d-387">Requirements</span></span>

|<span data-ttu-id="b621d-388">Requirement</span><span class="sxs-lookup"><span data-stu-id="b621d-388">Requirement</span></span>| <span data-ttu-id="b621d-389">Значение</span><span class="sxs-lookup"><span data-stu-id="b621d-389">Value</span></span>|
|---|---|
|[<span data-ttu-id="b621d-390">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="b621d-390">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b621d-391">1.3</span><span class="sxs-lookup"><span data-stu-id="b621d-391">1.3</span></span>|
|[<span data-ttu-id="b621d-392">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b621d-392">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b621d-393">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b621d-393">ReadItem</span></span>|
|[<span data-ttu-id="b621d-394">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b621d-394">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b621d-395">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="b621d-395">Compose or read</span></span>|

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook13officeemailaddressdetailsrecipientsjavascriptapioutlook13officerecipients"></a><span data-ttu-id="b621d-396">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_3/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="b621d-396">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_3/office.recipients)</span></span>

<span data-ttu-id="b621d-397">Предоставляет доступ к необязательным участникам события.</span><span class="sxs-lookup"><span data-stu-id="b621d-397">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="b621d-398">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="b621d-398">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="b621d-399">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="b621d-399">Read mode</span></span>

<span data-ttu-id="b621d-400">Свойство `optionalAttendees` возвращает массив, содержащий объект `EmailAddressDetails` для каждого необязательного участника собрания.</span><span class="sxs-lookup"><span data-stu-id="b621d-400">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="b621d-401">Режим создания</span><span class="sxs-lookup"><span data-stu-id="b621d-401">Compose mode</span></span>

<span data-ttu-id="b621d-402">Свойство `optionalAttendees` возвращает объект `Recipients`, который предоставляет методы для получения или обновления необязательных участников собрания.</span><span class="sxs-lookup"><span data-stu-id="b621d-402">The `optionalAttendees` property returns a `Recipients` object that provides methods to get and set the optional attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="b621d-403">Тип:</span><span class="sxs-lookup"><span data-stu-id="b621d-403">Type:</span></span>

*   <span data-ttu-id="b621d-404">Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_3/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="b621d-404">Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_3/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="b621d-405">Требования</span><span class="sxs-lookup"><span data-stu-id="b621d-405">Requirements</span></span>

|<span data-ttu-id="b621d-406">Requirement</span><span class="sxs-lookup"><span data-stu-id="b621d-406">Requirement</span></span>| <span data-ttu-id="b621d-407">Значение</span><span class="sxs-lookup"><span data-stu-id="b621d-407">Value</span></span>|
|---|---|
|[<span data-ttu-id="b621d-408">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="b621d-408">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b621d-409">1.0</span><span class="sxs-lookup"><span data-stu-id="b621d-409">1.0</span></span>|
|[<span data-ttu-id="b621d-410">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b621d-410">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b621d-411">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b621d-411">ReadItem</span></span>|
|[<span data-ttu-id="b621d-412">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b621d-412">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b621d-413">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="b621d-413">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="b621d-414">Пример</span><span class="sxs-lookup"><span data-stu-id="b621d-414">Example</span></span>

```js
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook13officeemailaddressdetails"></a><span data-ttu-id="b621d-415">organizer :[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="b621d-415">organizer :[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)</span></span>

<span data-ttu-id="b621d-p124">Получает электронный адрес организатора указанного собрания. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="b621d-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="b621d-418">Тип:</span><span class="sxs-lookup"><span data-stu-id="b621d-418">Type:</span></span>

*   [<span data-ttu-id="b621d-419">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="b621d-419">EmailAddressDetails</span></span>](/javascript/api/outlook_1_3/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="b621d-420">Требования</span><span class="sxs-lookup"><span data-stu-id="b621d-420">Requirements</span></span>

|<span data-ttu-id="b621d-421">Requirement</span><span class="sxs-lookup"><span data-stu-id="b621d-421">Requirement</span></span>| <span data-ttu-id="b621d-422">Значение</span><span class="sxs-lookup"><span data-stu-id="b621d-422">Value</span></span>|
|---|---|
|[<span data-ttu-id="b621d-423">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="b621d-423">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b621d-424">1.0</span><span class="sxs-lookup"><span data-stu-id="b621d-424">1.0</span></span>|
|[<span data-ttu-id="b621d-425">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b621d-425">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b621d-426">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b621d-426">ReadItem</span></span>|
|[<span data-ttu-id="b621d-427">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b621d-427">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b621d-428">Чтение</span><span class="sxs-lookup"><span data-stu-id="b621d-428">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b621d-429">Пример</span><span class="sxs-lookup"><span data-stu-id="b621d-429">Example</span></span>

```js
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
```

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook13officeemailaddressdetailsrecipientsjavascriptapioutlook13officerecipients"></a><span data-ttu-id="b621d-430">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_3/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="b621d-430">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_3/office.recipients)</span></span>

<span data-ttu-id="b621d-431">Предоставляет доступ к обязательным участникам события.</span><span class="sxs-lookup"><span data-stu-id="b621d-431">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="b621d-432">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="b621d-432">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="b621d-433">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="b621d-433">Read mode</span></span>

<span data-ttu-id="b621d-434">Свойство `requiredAttendees` возвращает массив, содержащий объект `EmailAddressDetails` для каждого обязательного участника собрания.</span><span class="sxs-lookup"><span data-stu-id="b621d-434">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="b621d-435">Режим создания</span><span class="sxs-lookup"><span data-stu-id="b621d-435">Compose mode</span></span>

<span data-ttu-id="b621d-436">Свойство `requiredAttendees` возвращает объект `Recipients`, предоставляющий методы, с помощью которых можно получить или обновить сведения об обязательных участниках собрания.</span><span class="sxs-lookup"><span data-stu-id="b621d-436">The `requiredAttendees` property returns a `Recipients` object that provides methods to get and set the required attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="b621d-437">Тип:</span><span class="sxs-lookup"><span data-stu-id="b621d-437">Type:</span></span>

*   <span data-ttu-id="b621d-438">Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_3/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="b621d-438">Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_3/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="b621d-439">Требования</span><span class="sxs-lookup"><span data-stu-id="b621d-439">Requirements</span></span>

|<span data-ttu-id="b621d-440">Requirement</span><span class="sxs-lookup"><span data-stu-id="b621d-440">Requirement</span></span>| <span data-ttu-id="b621d-441">Значение</span><span class="sxs-lookup"><span data-stu-id="b621d-441">Value</span></span>|
|---|---|
|[<span data-ttu-id="b621d-442">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="b621d-442">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b621d-443">1.0</span><span class="sxs-lookup"><span data-stu-id="b621d-443">1.0</span></span>|
|[<span data-ttu-id="b621d-444">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b621d-444">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b621d-445">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b621d-445">ReadItem</span></span>|
|[<span data-ttu-id="b621d-446">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b621d-446">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b621d-447">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="b621d-447">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="b621d-448">Пример</span><span class="sxs-lookup"><span data-stu-id="b621d-448">Example</span></span>

```js
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
}
```

#### <a name="sender-emailaddressdetailsjavascriptapioutlook13officeemailaddressdetails"></a><span data-ttu-id="b621d-449">sender :[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="b621d-449">sender :[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)</span></span>

<span data-ttu-id="b621d-p126">Получает электронный адрес отправителя электронного сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="b621d-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="b621d-p127">Свойства [`from`](#from-emailaddressdetailsjavascriptapioutlook13officeemailaddressdetails) и `sender` представляют одно лицо, если сообщение не отправлено представителем. В противном случае свойство `from` представляет лицо, делегировавшее полномочия, а свойство sender — представителя.</span><span class="sxs-lookup"><span data-stu-id="b621d-p127">The [`from`](#from-emailaddressdetailsjavascriptapioutlook13officeemailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="b621d-454">Свойству `recipientType`, принадлежащему объекту `EmailAddressDetails` в свойстве `sender`, задано значение `undefined`.</span><span class="sxs-lookup"><span data-stu-id="b621d-454">Note: The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="b621d-455">Тип:</span><span class="sxs-lookup"><span data-stu-id="b621d-455">Type:</span></span>

*   [<span data-ttu-id="b621d-456">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="b621d-456">EmailAddressDetails</span></span>](/javascript/api/outlook_1_3/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="b621d-457">Требования</span><span class="sxs-lookup"><span data-stu-id="b621d-457">Requirements</span></span>

|<span data-ttu-id="b621d-458">Requirement</span><span class="sxs-lookup"><span data-stu-id="b621d-458">Requirement</span></span>| <span data-ttu-id="b621d-459">Значение</span><span class="sxs-lookup"><span data-stu-id="b621d-459">Value</span></span>|
|---|---|
|[<span data-ttu-id="b621d-460">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="b621d-460">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b621d-461">1.0</span><span class="sxs-lookup"><span data-stu-id="b621d-461">1.0</span></span>|
|[<span data-ttu-id="b621d-462">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b621d-462">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b621d-463">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b621d-463">ReadItem</span></span>|
|[<span data-ttu-id="b621d-464">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b621d-464">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b621d-465">Чтение</span><span class="sxs-lookup"><span data-stu-id="b621d-465">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b621d-466">Пример</span><span class="sxs-lookup"><span data-stu-id="b621d-466">Example</span></span>

```js
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
```

####  <a name="start-datetimejavascriptapioutlook13officetime"></a><span data-ttu-id="b621d-467">start :Date|[Time](/javascript/api/outlook_1_3/office.time)</span><span class="sxs-lookup"><span data-stu-id="b621d-467">start :Date|[Time](/javascript/api/outlook_1_3/office.time)</span></span>

<span data-ttu-id="b621d-468">Получает или задает дату и время начала встречи.</span><span class="sxs-lookup"><span data-stu-id="b621d-468">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="b621d-p128">Свойство `start` представлено в виде значения даты и времени в формате UTC. Это значение можно преобразовать в местные значения даты и времени клиента с помощью метода [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook13officelocalclienttime).</span><span class="sxs-lookup"><span data-stu-id="b621d-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook13officelocalclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="b621d-471">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="b621d-471">Read mode</span></span>

<span data-ttu-id="b621d-472">Свойство `start` возвращает объект `Date`.</span><span class="sxs-lookup"><span data-stu-id="b621d-472">The `start` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="b621d-473">Режим создания</span><span class="sxs-lookup"><span data-stu-id="b621d-473">Compose mode</span></span>

<span data-ttu-id="b621d-474">Свойство `start` возвращает объект `Time`.</span><span class="sxs-lookup"><span data-stu-id="b621d-474">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="b621d-475">Если вы задаете время начала с помощью метода [`Time.setAsync`](/javascript/api/outlook_1_3/office.time#setasync-datetime--options--callback-), необходимо использовать метод [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) для преобразования местного времени на клиенте в формат UTC для сервера.</span><span class="sxs-lookup"><span data-stu-id="b621d-475">When you use the [`Time.setAsync`](/javascript/api/outlook_1_3/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="b621d-476">Тип:</span><span class="sxs-lookup"><span data-stu-id="b621d-476">Type:</span></span>

*   <span data-ttu-id="b621d-477">Date | [Time](/javascript/api/outlook_1_3/office.time)</span><span class="sxs-lookup"><span data-stu-id="b621d-477">Date | [Time](/javascript/api/outlook_1_3/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="b621d-478">Требования</span><span class="sxs-lookup"><span data-stu-id="b621d-478">Requirements</span></span>

|<span data-ttu-id="b621d-479">Requirement</span><span class="sxs-lookup"><span data-stu-id="b621d-479">Requirement</span></span>| <span data-ttu-id="b621d-480">Значение</span><span class="sxs-lookup"><span data-stu-id="b621d-480">Value</span></span>|
|---|---|
|[<span data-ttu-id="b621d-481">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="b621d-481">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b621d-482">1.0</span><span class="sxs-lookup"><span data-stu-id="b621d-482">1.0</span></span>|
|[<span data-ttu-id="b621d-483">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b621d-483">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b621d-484">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b621d-484">ReadItem</span></span>|
|[<span data-ttu-id="b621d-485">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b621d-485">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b621d-486">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="b621d-486">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="b621d-487">Пример</span><span class="sxs-lookup"><span data-stu-id="b621d-487">Example</span></span>

<span data-ttu-id="b621d-488">В примере ниже с помощью метода [`setAsync`](/javascript/api/outlook_1_3/office.time#setasync-datetime--options--callback-) объекта `Time` задается время начала встречи в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="b621d-488">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_3/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

####  <a name="subject-stringsubjectjavascriptapioutlook13officesubject"></a><span data-ttu-id="b621d-489">subject :String|[Subject](/javascript/api/outlook_1_3/office.subject)</span><span class="sxs-lookup"><span data-stu-id="b621d-489">subject :String|[Subject](/javascript/api/outlook_1_3/office.subject)</span></span>

<span data-ttu-id="b621d-490">Получает или задает описание, которое отображается в поле темы элемента.</span><span class="sxs-lookup"><span data-stu-id="b621d-490">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="b621d-491">Свойство `subject` получает или задает всю тему элемента для отправки с почтового сервера.</span><span class="sxs-lookup"><span data-stu-id="b621d-491">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="b621d-492">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="b621d-492">Read mode</span></span>

<span data-ttu-id="b621d-p129">Свойство `subject` возвращает строку. С помощью свойства [`normalizedSubject`](#normalizedsubject-string) можно получить тему без начальных префиксов, таких как `RE:` и `FW:`.</span><span class="sxs-lookup"><span data-stu-id="b621d-p129">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```js
var subject = Office.context.mailbox.item.subject;
```

##### <a name="compose-mode"></a><span data-ttu-id="b621d-495">Режим создания</span><span class="sxs-lookup"><span data-stu-id="b621d-495">Compose mode</span></span>

<span data-ttu-id="b621d-496">Свойство `subject` возвращает объект `Subject`, который предоставляет методы для получения и задания темы.</span><span class="sxs-lookup"><span data-stu-id="b621d-496">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```js
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="b621d-497">Тип:</span><span class="sxs-lookup"><span data-stu-id="b621d-497">Type:</span></span>

*   <span data-ttu-id="b621d-498">String | [Subject](/javascript/api/outlook_1_3/office.subject)</span><span class="sxs-lookup"><span data-stu-id="b621d-498">String | [Subject](/javascript/api/outlook_1_3/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="b621d-499">Требования</span><span class="sxs-lookup"><span data-stu-id="b621d-499">Requirements</span></span>

|<span data-ttu-id="b621d-500">Requirement</span><span class="sxs-lookup"><span data-stu-id="b621d-500">Requirement</span></span>| <span data-ttu-id="b621d-501">Значение</span><span class="sxs-lookup"><span data-stu-id="b621d-501">Value</span></span>|
|---|---|
|[<span data-ttu-id="b621d-502">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="b621d-502">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b621d-503">1.0</span><span class="sxs-lookup"><span data-stu-id="b621d-503">1.0</span></span>|
|[<span data-ttu-id="b621d-504">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b621d-504">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b621d-505">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b621d-505">ReadItem</span></span>|
|[<span data-ttu-id="b621d-506">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b621d-506">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b621d-507">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="b621d-507">Compose or read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlook13officeemailaddressdetailsrecipientsjavascriptapioutlook13officerecipients"></a><span data-ttu-id="b621d-508">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_3/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="b621d-508">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_3/office.recipients)</span></span>

<span data-ttu-id="b621d-509">Предоставляет доступ к получателям, указанным в строке **Кому** сообщения.</span><span class="sxs-lookup"><span data-stu-id="b621d-509">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="b621d-510">Тип объекта и уровень доступа зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="b621d-510">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="b621d-511">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="b621d-511">Read mode</span></span>

<span data-ttu-id="b621d-p131">Свойство `to` возвращает массив, содержащий объект `EmailAddressDetails` для каждого получателя в строке **Кому** сообщения. Коллекция может включать не более 100 элементов.</span><span class="sxs-lookup"><span data-stu-id="b621d-p131">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="b621d-514">Режим создания</span><span class="sxs-lookup"><span data-stu-id="b621d-514">Compose mode</span></span>

<span data-ttu-id="b621d-515">Свойство `to` возвращает объект `Recipients`, предоставляющий методы для получения или обновления получателей, которые указаны в строке **Кому** сообщения.</span><span class="sxs-lookup"><span data-stu-id="b621d-515">The to`to` property returns a Recipients`Recipients` object that provides methods for manipulating the recipients on the **To** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="b621d-516">Тип:</span><span class="sxs-lookup"><span data-stu-id="b621d-516">Type:</span></span>

*   <span data-ttu-id="b621d-517">Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_3/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="b621d-517">Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_3/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="b621d-518">Требования</span><span class="sxs-lookup"><span data-stu-id="b621d-518">Requirements</span></span>

|<span data-ttu-id="b621d-519">Requirement</span><span class="sxs-lookup"><span data-stu-id="b621d-519">Requirement</span></span>| <span data-ttu-id="b621d-520">Значение</span><span class="sxs-lookup"><span data-stu-id="b621d-520">Value</span></span>|
|---|---|
|[<span data-ttu-id="b621d-521">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="b621d-521">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b621d-522">1.0</span><span class="sxs-lookup"><span data-stu-id="b621d-522">1.0</span></span>|
|[<span data-ttu-id="b621d-523">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b621d-523">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b621d-524">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b621d-524">ReadItem</span></span>|
|[<span data-ttu-id="b621d-525">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b621d-525">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b621d-526">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="b621d-526">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="b621d-527">Пример</span><span class="sxs-lookup"><span data-stu-id="b621d-527">Example</span></span>

```js
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

### <a name="methods"></a><span data-ttu-id="b621d-528">Методы</span><span class="sxs-lookup"><span data-stu-id="b621d-528">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="b621d-529">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="b621d-529">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="b621d-530">Добавляет файл в сообщение или встречу в качестве вложения.</span><span class="sxs-lookup"><span data-stu-id="b621d-530">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="b621d-531">Метод `addFileAttachmentAsync` передает файл по указанному универсальному коду ресурса (URI) и вкладывает его в элемент в форме создания.</span><span class="sxs-lookup"><span data-stu-id="b621d-531">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="b621d-532">Идентификатор можно последовательно использовать с методом [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback), чтобы удалить вложение, добавленное во время текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="b621d-532">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b621d-533">Параметры</span><span class="sxs-lookup"><span data-stu-id="b621d-533">Parameters:</span></span>

|<span data-ttu-id="b621d-534">Имя</span><span class="sxs-lookup"><span data-stu-id="b621d-534">Name</span></span>| <span data-ttu-id="b621d-535">Тип</span><span class="sxs-lookup"><span data-stu-id="b621d-535">Type</span></span>| <span data-ttu-id="b621d-536">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="b621d-536">Attributes</span></span>| <span data-ttu-id="b621d-537">Описание</span><span class="sxs-lookup"><span data-stu-id="b621d-537">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="b621d-538">String</span><span class="sxs-lookup"><span data-stu-id="b621d-538">String</span></span>||<span data-ttu-id="b621d-p132">Универсальный код ресурса (URI), представляющий расположение файла, который нужно вложить в сообщение или встречу. Максимальная длина — 2048 символов.</span><span class="sxs-lookup"><span data-stu-id="b621d-p132">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="b621d-541">String</span><span class="sxs-lookup"><span data-stu-id="b621d-541">String</span></span>||<span data-ttu-id="b621d-p133">Имя вложения, которое отображается при передаче вложения. Максимальная длина — 255 символов.</span><span class="sxs-lookup"><span data-stu-id="b621d-p133">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="b621d-544">Object</span><span class="sxs-lookup"><span data-stu-id="b621d-544">Object</span></span>| <span data-ttu-id="b621d-545">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="b621d-545">&lt;optional&gt;</span></span>|<span data-ttu-id="b621d-546">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="b621d-546">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="b621d-547">Object</span><span class="sxs-lookup"><span data-stu-id="b621d-547">Object</span></span>| <span data-ttu-id="b621d-548">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="b621d-548">&lt;optional&gt;</span></span>|<span data-ttu-id="b621d-549">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="b621d-549">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="b621d-550">функция</span><span class="sxs-lookup"><span data-stu-id="b621d-550">function</span></span>| <span data-ttu-id="b621d-551">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="b621d-551">&lt;optional&gt;</span></span>|<span data-ttu-id="b621d-552">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="b621d-552">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="b621d-553">После успешного выполнения идентификатор вложения будет представлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="b621d-553">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="b621d-554">Если добавить вложение не удастся, объект `asyncResult` будет содержать объект `Error` с описанием ошибки.</span><span class="sxs-lookup"><span data-stu-id="b621d-554">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="b621d-555">Ошибки</span><span class="sxs-lookup"><span data-stu-id="b621d-555">Errors</span></span>

| <span data-ttu-id="b621d-556">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="b621d-556">Error code</span></span> | <span data-ttu-id="b621d-557">Описание</span><span class="sxs-lookup"><span data-stu-id="b621d-557">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="b621d-558">Вложение превышает максимальный размер.</span><span class="sxs-lookup"><span data-stu-id="b621d-558">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="b621d-559">Расширение вложения не поддерживается.</span><span class="sxs-lookup"><span data-stu-id="b621d-559">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="b621d-560">Сообщение или встреча содержат слишком много вложений.</span><span class="sxs-lookup"><span data-stu-id="b621d-560">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="b621d-561">Требования</span><span class="sxs-lookup"><span data-stu-id="b621d-561">Requirements</span></span>

|<span data-ttu-id="b621d-562">Requirement</span><span class="sxs-lookup"><span data-stu-id="b621d-562">Requirement</span></span>| <span data-ttu-id="b621d-563">Значение</span><span class="sxs-lookup"><span data-stu-id="b621d-563">Value</span></span>|
|---|---|
|[<span data-ttu-id="b621d-564">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="b621d-564">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b621d-565">1.1</span><span class="sxs-lookup"><span data-stu-id="b621d-565">1.1</span></span>|
|[<span data-ttu-id="b621d-566">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b621d-566">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b621d-567">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="b621d-567">ReadWriteItem</span></span>|
|[<span data-ttu-id="b621d-568">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b621d-568">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b621d-569">Создание</span><span class="sxs-lookup"><span data-stu-id="b621d-569">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="b621d-570">Пример</span><span class="sxs-lookup"><span data-stu-id="b621d-570">Example</span></span>

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

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="b621d-571">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="b621d-571">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="b621d-572">Добавляет к сообщению элемент Exchange, например сообщение, в виде вложения.</span><span class="sxs-lookup"><span data-stu-id="b621d-572">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="b621d-p134">С помощью метода `addItemAttachmentAsync` можно в элемент формы создания вложить элемент с указанным идентификатором Exchange. Если указать метод обратного вызова, то этот метод вызывается с помощью параметра `asyncResult`, который содержит идентификатор вложения или код, указывающий на ошибки, которые произошли при вложении элемента. При необходимости можно использовать параметр `options` для передачи сведений о состоянии методу обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="b621d-p134">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="b621d-576">Идентификатор можно последовательно использовать с методом [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback), чтобы удалить вложение, добавленное во время текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="b621d-576">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="b621d-577">Если ваша надстройка Office выполняется в Outlook Web App, метод `addItemAttachmentAsync` обеспечивает вложение элементов в элементы, отличные от редактируемого. Однако мы не рекомендуем выполнять это действие, так как оно не поддерживается.</span><span class="sxs-lookup"><span data-stu-id="b621d-577">If your Office add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b621d-578">Параметры</span><span class="sxs-lookup"><span data-stu-id="b621d-578">Parameters:</span></span>

|<span data-ttu-id="b621d-579">Имя</span><span class="sxs-lookup"><span data-stu-id="b621d-579">Name</span></span>| <span data-ttu-id="b621d-580">Тип</span><span class="sxs-lookup"><span data-stu-id="b621d-580">Type</span></span>| <span data-ttu-id="b621d-581">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="b621d-581">Attributes</span></span>| <span data-ttu-id="b621d-582">Описание</span><span class="sxs-lookup"><span data-stu-id="b621d-582">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="b621d-583">String</span><span class="sxs-lookup"><span data-stu-id="b621d-583">String</span></span>||<span data-ttu-id="b621d-p135">Идентификатор Exchange для вкладываемого элемента. Максимальная длина — 100 символов.</span><span class="sxs-lookup"><span data-stu-id="b621d-p135">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="b621d-586">String</span><span class="sxs-lookup"><span data-stu-id="b621d-586">String</span></span>||<span data-ttu-id="b621d-p136">Тема вкладываемого элемента. Максимальная длина — 255 символов.</span><span class="sxs-lookup"><span data-stu-id="b621d-p136">The sujbect of the item to be attached. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="b621d-589">Object</span><span class="sxs-lookup"><span data-stu-id="b621d-589">Object</span></span>| <span data-ttu-id="b621d-590">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="b621d-590">&lt;optional&gt;</span></span>|<span data-ttu-id="b621d-591">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="b621d-591">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="b621d-592">Object</span><span class="sxs-lookup"><span data-stu-id="b621d-592">Object</span></span>| <span data-ttu-id="b621d-593">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="b621d-593">&lt;optional&gt;</span></span>|<span data-ttu-id="b621d-594">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="b621d-594">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="b621d-595">функция</span><span class="sxs-lookup"><span data-stu-id="b621d-595">function</span></span>| <span data-ttu-id="b621d-596">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="b621d-596">&lt;optional&gt;</span></span>|<span data-ttu-id="b621d-597">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="b621d-597">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="b621d-598">После успешного выполнения идентификатор вложения будет представлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="b621d-598">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="b621d-599">Если добавить вложение не удастся, объект `asyncResult` будет содержать объект `Error` с описанием ошибки.</span><span class="sxs-lookup"><span data-stu-id="b621d-599">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="b621d-600">Ошибки</span><span class="sxs-lookup"><span data-stu-id="b621d-600">Errors</span></span>

| <span data-ttu-id="b621d-601">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="b621d-601">Error code</span></span> | <span data-ttu-id="b621d-602">Описание</span><span class="sxs-lookup"><span data-stu-id="b621d-602">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="b621d-603">Сообщение или встреча содержат слишком много вложений.</span><span class="sxs-lookup"><span data-stu-id="b621d-603">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="b621d-604">Требования</span><span class="sxs-lookup"><span data-stu-id="b621d-604">Requirements</span></span>

|<span data-ttu-id="b621d-605">Requirement</span><span class="sxs-lookup"><span data-stu-id="b621d-605">Requirement</span></span>| <span data-ttu-id="b621d-606">Значение</span><span class="sxs-lookup"><span data-stu-id="b621d-606">Value</span></span>|
|---|---|
|[<span data-ttu-id="b621d-607">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="b621d-607">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b621d-608">1.1</span><span class="sxs-lookup"><span data-stu-id="b621d-608">1.1</span></span>|
|[<span data-ttu-id="b621d-609">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b621d-609">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b621d-610">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="b621d-610">ReadWriteItem</span></span>|
|[<span data-ttu-id="b621d-611">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b621d-611">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b621d-612">Создание</span><span class="sxs-lookup"><span data-stu-id="b621d-612">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="b621d-613">Пример</span><span class="sxs-lookup"><span data-stu-id="b621d-613">Example</span></span>

<span data-ttu-id="b621d-614">В следующем примере существующий элемент Outlook добавляется в виде вложения с именем `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="b621d-614">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

####  <a name="close"></a><span data-ttu-id="b621d-615">close()</span><span class="sxs-lookup"><span data-stu-id="b621d-615">close()</span></span>

<span data-ttu-id="b621d-616">Закрывает текущий создаваемый элемент.</span><span class="sxs-lookup"><span data-stu-id="b621d-616">Closes the current item that is being composed.</span></span>

<span data-ttu-id="b621d-p137">Работа метода `close` зависит от текущего состояния создаваемого элемента. Если элемент содержит несохраненные изменения, клиент предложит пользователю сохранить или отклонить их либо отменить действие закрытия.</span><span class="sxs-lookup"><span data-stu-id="b621d-p137">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="b621d-619">Если в Outlook в Интернете элемент представляет собой встречу, ранее сохраненную с помощью метода `saveAsync`, пользователю предлагается сохранить, отклонить или отменить действие, даже если с момента последнего сохранения элемента не вносилось никаких изменений.</span><span class="sxs-lookup"><span data-stu-id="b621d-619">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="b621d-620">Если в клиенте Outlook для настольных ПК сообщение представляет собой ответ в тексте, метод `close` не приносит результатов.</span><span class="sxs-lookup"><span data-stu-id="b621d-620">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="b621d-621">Требования</span><span class="sxs-lookup"><span data-stu-id="b621d-621">Requirements</span></span>

|<span data-ttu-id="b621d-622">Requirement</span><span class="sxs-lookup"><span data-stu-id="b621d-622">Requirement</span></span>| <span data-ttu-id="b621d-623">Значение</span><span class="sxs-lookup"><span data-stu-id="b621d-623">Value</span></span>|
|---|---|
|[<span data-ttu-id="b621d-624">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="b621d-624">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b621d-625">1.3</span><span class="sxs-lookup"><span data-stu-id="b621d-625">1.3</span></span>|
|[<span data-ttu-id="b621d-626">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b621d-626">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b621d-627">Restricted</span><span class="sxs-lookup"><span data-stu-id="b621d-627">Restricted</span></span>|
|[<span data-ttu-id="b621d-628">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b621d-628">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b621d-629">Создание</span><span class="sxs-lookup"><span data-stu-id="b621d-629">Compose</span></span>|

#### <a name="displayreplyallformformdata"></a><span data-ttu-id="b621d-630">displayReplyAllForm(formData)</span><span class="sxs-lookup"><span data-stu-id="b621d-630">displayReplyAllForm(formData)</span></span>

<span data-ttu-id="b621d-631">Отображает форму ответа, включающую отправителя и всех получателей выбранного сообщения или организатора и всех участников выбранной встречи.</span><span class="sxs-lookup"><span data-stu-id="b621d-631">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="b621d-632">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="b621d-632">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="b621d-633">В Outlook Web App форма ответа отображается в виде всплывающей формы в представлении с 3 либо 1 или 2 колонками.</span><span class="sxs-lookup"><span data-stu-id="b621d-633">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="b621d-634">Если любой строковый параметр превышает указанные для него ограничения, `displayReplyAllForm` возвращает исключение.</span><span class="sxs-lookup"><span data-stu-id="b621d-634">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="b621d-p138">Если в параметре `formData.attachments` указаны вложения, Outlook и Outlook Web App пытаются скачать их и вложить в форму ответа. Если какие-либо вложения добавить не удается, в форме отображается сообщение об ошибке. Если сообщения об ошибках не предусмотрены, то они не отображаются.</span><span class="sxs-lookup"><span data-stu-id="b621d-p138">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b621d-638">Параметры</span><span class="sxs-lookup"><span data-stu-id="b621d-638">Parameters:</span></span>

|<span data-ttu-id="b621d-639">Имя</span><span class="sxs-lookup"><span data-stu-id="b621d-639">Name</span></span>| <span data-ttu-id="b621d-640">Тип</span><span class="sxs-lookup"><span data-stu-id="b621d-640">Type</span></span>| <span data-ttu-id="b621d-641">Описание</span><span class="sxs-lookup"><span data-stu-id="b621d-641">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="b621d-642">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="b621d-642">String &#124; Object</span></span>| |<span data-ttu-id="b621d-p139">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="b621d-p139">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="b621d-645">**ИЛИ**</span><span class="sxs-lookup"><span data-stu-id="b621d-645">**OR**</span></span><br/><span data-ttu-id="b621d-p140">Объект, который содержит текст или данные вложения и функцию обратного вызова. Ниже представлено определение этого объекта.</span><span class="sxs-lookup"><span data-stu-id="b621d-p140">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="b621d-648">String</span><span class="sxs-lookup"><span data-stu-id="b621d-648">String</span></span> | <span data-ttu-id="b621d-649">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="b621d-649">&lt;optional&gt;</span></span> | <span data-ttu-id="b621d-p141">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="b621d-p141">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="b621d-652">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="b621d-652">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="b621d-653">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="b621d-653">&lt;optional&gt;</span></span> | <span data-ttu-id="b621d-654">Массив объектов JSON, представляющих собой вложенные файлы или элементы.</span><span class="sxs-lookup"><span data-stu-id="b621d-654">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="b621d-655">String</span><span class="sxs-lookup"><span data-stu-id="b621d-655">String</span></span> | | <span data-ttu-id="b621d-p142">Указывает тип вложения. Допустимые значения: `file` для вложенного файла и `item` для вложенного элемента.</span><span class="sxs-lookup"><span data-stu-id="b621d-p142">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="b621d-658">String</span><span class="sxs-lookup"><span data-stu-id="b621d-658">String</span></span> | | <span data-ttu-id="b621d-659">Строка, содержащая имя вложения, длиной до 255 символов.</span><span class="sxs-lookup"><span data-stu-id="b621d-659">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="b621d-660">String</span><span class="sxs-lookup"><span data-stu-id="b621d-660">String</span></span> | | <span data-ttu-id="b621d-p143">Используется, только если свойству `type` задано значение `file`. URI расположения файла.</span><span class="sxs-lookup"><span data-stu-id="b621d-p143">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="b621d-663">String</span><span class="sxs-lookup"><span data-stu-id="b621d-663">String</span></span> | | <span data-ttu-id="b621d-p144">Используется, только если свойству `type` задано значение `item`. Идентификатор вложения EWS. Это строка длиной до 100 символов.</span><span class="sxs-lookup"><span data-stu-id="b621d-p144">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="b621d-667">function</span><span class="sxs-lookup"><span data-stu-id="b621d-667">function</span></span> | <span data-ttu-id="b621d-668">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="b621d-668">&lt;optional&gt;</span></span> | <span data-ttu-id="b621d-669">По завершении работы метода функция, переданная параметру `callback`, вызывается с помощью одного параметра `asyncResult`, представляющего собой объект [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="b621d-669">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="b621d-670">Требования</span><span class="sxs-lookup"><span data-stu-id="b621d-670">Requirements</span></span>

|<span data-ttu-id="b621d-671">Requirement</span><span class="sxs-lookup"><span data-stu-id="b621d-671">Requirement</span></span>| <span data-ttu-id="b621d-672">Значение</span><span class="sxs-lookup"><span data-stu-id="b621d-672">Value</span></span>|
|---|---|
|[<span data-ttu-id="b621d-673">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="b621d-673">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b621d-674">1.0</span><span class="sxs-lookup"><span data-stu-id="b621d-674">1.0</span></span>|
|[<span data-ttu-id="b621d-675">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b621d-675">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b621d-676">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b621d-676">ReadItem</span></span>|
|[<span data-ttu-id="b621d-677">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b621d-677">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b621d-678">Чтение</span><span class="sxs-lookup"><span data-stu-id="b621d-678">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="b621d-679">Примеры</span><span class="sxs-lookup"><span data-stu-id="b621d-679">Examples</span></span>

<span data-ttu-id="b621d-680">Приведенный ниже код передает строку в функцию `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="b621d-680">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="b621d-681">Ответ с пустым текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="b621d-681">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="b621d-682">Ответ только с текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="b621d-682">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="b621d-683">Ответ с текстом сообщения и вложенным файлом.</span><span class="sxs-lookup"><span data-stu-id="b621d-683">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="b621d-684">Ответ с текстом сообщения и вложенным элементом.</span><span class="sxs-lookup"><span data-stu-id="b621d-684">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="b621d-685">Ответ с текстом сообщения, вложенным файлом, вложенным элементом и обратным вызовом.</span><span class="sxs-lookup"><span data-stu-id="b621d-685">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata"></a><span data-ttu-id="b621d-686">displayReplyForm(formData)</span><span class="sxs-lookup"><span data-stu-id="b621d-686">displayReplyForm(formData)</span></span>

<span data-ttu-id="b621d-687">Отображает форму ответа, включающую только отправителя выбранного сообщения или организатора выбранной встречи.</span><span class="sxs-lookup"><span data-stu-id="b621d-687">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="b621d-688">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="b621d-688">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="b621d-689">В Outlook Web App форма ответа отображается в виде всплывающей формы в представлении с 3 либо 1 или 2 колонками.</span><span class="sxs-lookup"><span data-stu-id="b621d-689">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="b621d-690">Если любой строковый параметр превышает указанные для него ограничения, `displayReplyForm` возвращает исключение.</span><span class="sxs-lookup"><span data-stu-id="b621d-690">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="b621d-p145">Если в параметре `formData.attachments` указаны вложения, Outlook и Outlook Web App пытаются скачать их и вложить в форму ответа. Если какие-либо вложения добавить не удается, в форме отображается сообщение об ошибке. Если сообщения об ошибках не предусмотрены, то они не отображаются.</span><span class="sxs-lookup"><span data-stu-id="b621d-p145">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b621d-694">Параметры</span><span class="sxs-lookup"><span data-stu-id="b621d-694">Parameters:</span></span>

|<span data-ttu-id="b621d-695">Имя</span><span class="sxs-lookup"><span data-stu-id="b621d-695">Name</span></span>| <span data-ttu-id="b621d-696">Тип</span><span class="sxs-lookup"><span data-stu-id="b621d-696">Type</span></span>| <span data-ttu-id="b621d-697">Описание</span><span class="sxs-lookup"><span data-stu-id="b621d-697">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="b621d-698">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="b621d-698">String &#124; Object</span></span>| | <span data-ttu-id="b621d-p146">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="b621d-p146">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="b621d-701">**ИЛИ**</span><span class="sxs-lookup"><span data-stu-id="b621d-701">**OR**</span></span><br/><span data-ttu-id="b621d-p147">Объект, который содержит текст или данные вложения и функцию обратного вызова. Ниже представлено определение этого объекта.</span><span class="sxs-lookup"><span data-stu-id="b621d-p147">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="b621d-704">String</span><span class="sxs-lookup"><span data-stu-id="b621d-704">String</span></span> | <span data-ttu-id="b621d-705">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="b621d-705">&lt;optional&gt;</span></span> | <span data-ttu-id="b621d-p148">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="b621d-p148">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="b621d-708">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="b621d-708">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="b621d-709">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="b621d-709">&lt;optional&gt;</span></span> | <span data-ttu-id="b621d-710">Массив объектов JSON, представляющих собой вложенные файлы или элементы.</span><span class="sxs-lookup"><span data-stu-id="b621d-710">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="b621d-711">String</span><span class="sxs-lookup"><span data-stu-id="b621d-711">String</span></span> | | <span data-ttu-id="b621d-p149">Указывает тип вложения. Допустимые значения: `file` для вложенного файла и `item` для вложенного элемента.</span><span class="sxs-lookup"><span data-stu-id="b621d-p149">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="b621d-714">String</span><span class="sxs-lookup"><span data-stu-id="b621d-714">String</span></span> | | <span data-ttu-id="b621d-715">Строка, содержащая имя вложения, длиной до 255 символов.</span><span class="sxs-lookup"><span data-stu-id="b621d-715">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="b621d-716">String</span><span class="sxs-lookup"><span data-stu-id="b621d-716">String</span></span> | | <span data-ttu-id="b621d-p150">Используется, только если свойству `type` задано значение `file`. URI расположения файла.</span><span class="sxs-lookup"><span data-stu-id="b621d-p150">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="b621d-719">String</span><span class="sxs-lookup"><span data-stu-id="b621d-719">String</span></span> | | <span data-ttu-id="b621d-p151">Используется, только если свойству `type` задано значение `item`. Идентификатор вложения EWS. Это строка длиной до 100 символов.</span><span class="sxs-lookup"><span data-stu-id="b621d-p151">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="b621d-723">function</span><span class="sxs-lookup"><span data-stu-id="b621d-723">function</span></span> | <span data-ttu-id="b621d-724">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="b621d-724">&lt;optional&gt;</span></span> | <span data-ttu-id="b621d-725">По завершении работы метода функция, переданная параметру `callback`, вызывается с помощью одного параметра `asyncResult`, представляющего собой объект [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="b621d-725">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="b621d-726">Требования</span><span class="sxs-lookup"><span data-stu-id="b621d-726">Requirements</span></span>

|<span data-ttu-id="b621d-727">Requirement</span><span class="sxs-lookup"><span data-stu-id="b621d-727">Requirement</span></span>| <span data-ttu-id="b621d-728">Значение</span><span class="sxs-lookup"><span data-stu-id="b621d-728">Value</span></span>|
|---|---|
|[<span data-ttu-id="b621d-729">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="b621d-729">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b621d-730">1.0</span><span class="sxs-lookup"><span data-stu-id="b621d-730">1.0</span></span>|
|[<span data-ttu-id="b621d-731">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b621d-731">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b621d-732">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b621d-732">ReadItem</span></span>|
|[<span data-ttu-id="b621d-733">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b621d-733">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b621d-734">Чтение</span><span class="sxs-lookup"><span data-stu-id="b621d-734">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="b621d-735">Примеры</span><span class="sxs-lookup"><span data-stu-id="b621d-735">Examples</span></span>

<span data-ttu-id="b621d-736">Приведенный ниже код передает строку в функцию `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="b621d-736">The following code passes a string to the `displayReplyForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="b621d-737">Ответ с пустым текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="b621d-737">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="b621d-738">Ответ только с текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="b621d-738">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="b621d-739">Ответ с текстом сообщения и вложенным файлом.</span><span class="sxs-lookup"><span data-stu-id="b621d-739">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="b621d-740">Ответ с текстом сообщения и вложенным элементом.</span><span class="sxs-lookup"><span data-stu-id="b621d-740">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="b621d-741">Ответ с текстом сообщения, вложенным файлом, вложенным элементом и обратным вызовом.</span><span class="sxs-lookup"><span data-stu-id="b621d-741">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlook13officeentities"></a><span data-ttu-id="b621d-742">getEntities() → {[Entities](/javascript/api/outlook_1_3/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="b621d-742">getEntities() → {[Entities](/javascript/api/outlook_1_3/office.entities)}</span></span>

<span data-ttu-id="b621d-743">Получает сущности, обнаруженные в тексте выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="b621d-743">Gets the entities found in the selected item.</span></span>

> [!NOTE]
> <span data-ttu-id="b621d-744">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="b621d-744">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="b621d-745">Требования</span><span class="sxs-lookup"><span data-stu-id="b621d-745">Requirements</span></span>

|<span data-ttu-id="b621d-746">Requirement</span><span class="sxs-lookup"><span data-stu-id="b621d-746">Requirement</span></span>| <span data-ttu-id="b621d-747">Значение</span><span class="sxs-lookup"><span data-stu-id="b621d-747">Value</span></span>|
|---|---|
|[<span data-ttu-id="b621d-748">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="b621d-748">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b621d-749">1.0</span><span class="sxs-lookup"><span data-stu-id="b621d-749">1.0</span></span>|
|[<span data-ttu-id="b621d-750">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b621d-750">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b621d-751">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b621d-751">ReadItem</span></span>|
|[<span data-ttu-id="b621d-752">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b621d-752">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b621d-753">Чтение</span><span class="sxs-lookup"><span data-stu-id="b621d-753">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="b621d-754">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="b621d-754">Returns:</span></span>

<span data-ttu-id="b621d-755">Тип: [Entities](/javascript/api/outlook_1_3/office.entities)</span><span class="sxs-lookup"><span data-stu-id="b621d-755">Type: [Entities](/javascript/api/outlook_1_3/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="b621d-756">Пример</span><span class="sxs-lookup"><span data-stu-id="b621d-756">Example</span></span>

<span data-ttu-id="b621d-757">Ниже приведен пример получения доступа к сущностям контактов в тексте текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="b621d-757">The following example accesses the contacts entities on the current item.</span></span>

```js
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook13officecontactmeetingsuggestionjavascriptapioutlook13officemeetingsuggestionphonenumberjavascriptapioutlook13officephonenumbertasksuggestionjavascriptapioutlook13officetasksuggestion"></a><span data-ttu-id="b621d-758">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_3/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_3/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_3/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_3/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="b621d-758">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_3/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_3/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_3/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_3/office.tasksuggestion))>}</span></span>

<span data-ttu-id="b621d-759">Получает массив всех сущностей указанного типа, обнаруженных в тексте выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="b621d-759">Gets an array of all the entities of the specified entity type found in the selected item.</span></span>

> [!NOTE]
> <span data-ttu-id="b621d-760">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="b621d-760">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b621d-761">Параметры</span><span class="sxs-lookup"><span data-stu-id="b621d-761">Parameters:</span></span>

|<span data-ttu-id="b621d-762">Имя</span><span class="sxs-lookup"><span data-stu-id="b621d-762">Name</span></span>| <span data-ttu-id="b621d-763">Тип</span><span class="sxs-lookup"><span data-stu-id="b621d-763">Type</span></span>| <span data-ttu-id="b621d-764">Описание</span><span class="sxs-lookup"><span data-stu-id="b621d-764">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="b621d-765">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="b621d-765">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook_1_3/office.mailboxenums.entitytype)|<span data-ttu-id="b621d-766">Одно из значений перечисления EntityType.</span><span class="sxs-lookup"><span data-stu-id="b621d-766">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b621d-767">Требования</span><span class="sxs-lookup"><span data-stu-id="b621d-767">Requirements</span></span>

|<span data-ttu-id="b621d-768">Requirement</span><span class="sxs-lookup"><span data-stu-id="b621d-768">Requirement</span></span>| <span data-ttu-id="b621d-769">Значение</span><span class="sxs-lookup"><span data-stu-id="b621d-769">Value</span></span>|
|---|---|
|[<span data-ttu-id="b621d-770">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="b621d-770">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b621d-771">1.0</span><span class="sxs-lookup"><span data-stu-id="b621d-771">1.0</span></span>|
|[<span data-ttu-id="b621d-772">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b621d-772">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b621d-773">Restricted</span><span class="sxs-lookup"><span data-stu-id="b621d-773">Restricted</span></span>|
|[<span data-ttu-id="b621d-774">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b621d-774">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b621d-775">Чтение</span><span class="sxs-lookup"><span data-stu-id="b621d-775">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="b621d-776">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="b621d-776">Returns:</span></span>

<span data-ttu-id="b621d-777">Если значение, переданное в `entityType`, не является допустимым членом перечисления `EntityType`, метод возвращает значение NULL.</span><span class="sxs-lookup"><span data-stu-id="b621d-777">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="b621d-778">Если в тексте элемента отсутствуют сущности указанного типа, метод возвращает пустой массив.</span><span class="sxs-lookup"><span data-stu-id="b621d-778">If no entities of the specified type are present on the item, the method returns an empty array.</span></span> <span data-ttu-id="b621d-779">В противном случае тип объектов в возвращаемом массиве зависит от типа сущности, запрошенной в параметре `entityType`.</span><span class="sxs-lookup"><span data-stu-id="b621d-779">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="b621d-780">Хотя минимальный уровень разрешений для использования этого метода — **Restricted**, для некоторых типов сущностей требуется доступ на уровне **ReadItem**, как указано в приведенной ниже таблице.</span><span class="sxs-lookup"><span data-stu-id="b621d-780">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="b621d-781">Значение параметра `entityType`</span><span class="sxs-lookup"><span data-stu-id="b621d-781">Value of `entityType`</span></span> | <span data-ttu-id="b621d-782">Тип объектов в возвращаемом массиве</span><span class="sxs-lookup"><span data-stu-id="b621d-782">Type of objects in returned array</span></span> | <span data-ttu-id="b621d-783">Необходимый уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b621d-783">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="b621d-784">String</span><span class="sxs-lookup"><span data-stu-id="b621d-784">String</span></span> | <span data-ttu-id="b621d-785">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="b621d-785">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="b621d-786">Contact</span><span class="sxs-lookup"><span data-stu-id="b621d-786">Contact</span></span> | <span data-ttu-id="b621d-787">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="b621d-787">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="b621d-788">String</span><span class="sxs-lookup"><span data-stu-id="b621d-788">String</span></span> | <span data-ttu-id="b621d-789">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="b621d-789">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="b621d-790">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="b621d-790">MeetingSuggestion</span></span> | <span data-ttu-id="b621d-791">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="b621d-791">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="b621d-792">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="b621d-792">PhoneNumber</span></span> | <span data-ttu-id="b621d-793">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="b621d-793">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="b621d-794">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="b621d-794">TaskSuggestion</span></span> | <span data-ttu-id="b621d-795">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="b621d-795">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="b621d-796">String</span><span class="sxs-lookup"><span data-stu-id="b621d-796">String</span></span> | <span data-ttu-id="b621d-797">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="b621d-797">**Restricted**</span></span> |

<span data-ttu-id="b621d-798">Тип: Array.<(String|[Contact](/javascript/api/outlook_1_3/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_3/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_3/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_3/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="b621d-798">Type: Array.<(String|[Contact](/javascript/api/outlook_1_3/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_3/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_3/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_3/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="b621d-799">Пример</span><span class="sxs-lookup"><span data-stu-id="b621d-799">Example</span></span>

<span data-ttu-id="b621d-800">В примере ниже показано, как получить доступ к массиву строк, которые представляют собой почтовые адреса в тексте текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="b621d-800">The following example shows how to access an array of strings that represent postal addresses in the subject or body of the current item.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook13officecontactmeetingsuggestionjavascriptapioutlook13officemeetingsuggestionphonenumberjavascriptapioutlook13officephonenumbertasksuggestionjavascriptapioutlook13officetasksuggestion"></a><span data-ttu-id="b621d-801">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_3/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_3/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_3/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_3/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="b621d-801">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_3/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_3/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_3/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_3/office.tasksuggestion))>}</span></span>

<span data-ttu-id="b621d-802">Возвращает известные сущности в выбранном элементе, которые проходят через именованный фильтр, определяемый в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="b621d-802">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="b621d-803">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="b621d-803">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="b621d-804">Метод `getFilteredEntitiesByName` возвращает сущности, соответствующие регулярному выражению, которое определяется в элементе правила [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) в XML-файле манифеста, с использованием указанного значения элемента `FilterName`.</span><span class="sxs-lookup"><span data-stu-id="b621d-804">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b621d-805">Параметры</span><span class="sxs-lookup"><span data-stu-id="b621d-805">Parameters:</span></span>

|<span data-ttu-id="b621d-806">Имя</span><span class="sxs-lookup"><span data-stu-id="b621d-806">Name</span></span>| <span data-ttu-id="b621d-807">Тип</span><span class="sxs-lookup"><span data-stu-id="b621d-807">Type</span></span>| <span data-ttu-id="b621d-808">Описание</span><span class="sxs-lookup"><span data-stu-id="b621d-808">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="b621d-809">String</span><span class="sxs-lookup"><span data-stu-id="b621d-809">String</span></span>|<span data-ttu-id="b621d-810">Имя элемента правила `ItemHasKnownEntity`, определяющее соответствующий фильтр.</span><span class="sxs-lookup"><span data-stu-id="b621d-810">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b621d-811">Требования</span><span class="sxs-lookup"><span data-stu-id="b621d-811">Requirements</span></span>

|<span data-ttu-id="b621d-812">Requirement</span><span class="sxs-lookup"><span data-stu-id="b621d-812">Requirement</span></span>| <span data-ttu-id="b621d-813">Значение</span><span class="sxs-lookup"><span data-stu-id="b621d-813">Value</span></span>|
|---|---|
|[<span data-ttu-id="b621d-814">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="b621d-814">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b621d-815">1.0</span><span class="sxs-lookup"><span data-stu-id="b621d-815">1.0</span></span>|
|[<span data-ttu-id="b621d-816">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b621d-816">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b621d-817">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b621d-817">ReadItem</span></span>|
|[<span data-ttu-id="b621d-818">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b621d-818">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b621d-819">Чтение</span><span class="sxs-lookup"><span data-stu-id="b621d-819">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="b621d-820">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="b621d-820">Returns:</span></span>

<span data-ttu-id="b621d-p153">Если в манифесте нет элемента `ItemHasKnownEntity` со значением `FilterName`, соответствующим параметру `name`, метод возвращает `null`. Если параметр `name` соответствует элементу `ItemHasKnownEntity` в манифесте, но при этом в текущем элементе нет соответствующих сущностей, метод возвращает пустой массив.</span><span class="sxs-lookup"><span data-stu-id="b621d-p153">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="b621d-823">Тип: Array.<(String|[Contact](/javascript/api/outlook_1_3/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_3/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_3/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_3/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="b621d-823">Type: Array.<(String|[Contact](/javascript/api/outlook_1_3/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_3/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_3/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_3/office.tasksuggestion))></span></span>

#### <a name="getregexmatches--object"></a><span data-ttu-id="b621d-824">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="b621d-824">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="b621d-825">Возвращает строковые значения в выбранном элементе, которые соответствуют регулярным выражениям, определенным в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="b621d-825">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="b621d-826">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="b621d-826">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="b621d-p154">Метод `getRegExMatches` возвращает строки, соответствующие регулярному выражению, которое определяется в каждом элементе правила `ItemHasRegularExpressionMatch` или `ItemHasKnownEntity` в XML-файле манифеста. Для правила `ItemHasRegularExpressionMatch` соответствующую строку должно содержать свойство элемента, указанного этим правилом. Простой тип `PropertyName` определяет поддерживаемые свойства.</span><span class="sxs-lookup"><span data-stu-id="b621d-p154">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="b621d-830">Например, рассмотрим манифест надстройки, который содержит указанный ниже элемент `Rule`.</span><span class="sxs-lookup"><span data-stu-id="b621d-830">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="b621d-831">Объект, возвращаемый методом `getRegExMatches`, будет содержать два свойства: `fruits` и `veggies`.</span><span class="sxs-lookup"><span data-stu-id="b621d-831">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="b621d-p155">Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты. Лучше используйте метод [`Body.getAsync`](/javascript/api/outlook_1_3/office.body#getasync-coerciontype--options--callback-) для этого.</span><span class="sxs-lookup"><span data-stu-id="b621d-p155">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook_1_3/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="b621d-835">Requirements</span><span class="sxs-lookup"><span data-stu-id="b621d-835">Requirements</span></span>

|<span data-ttu-id="b621d-836">Requirement</span><span class="sxs-lookup"><span data-stu-id="b621d-836">Requirement</span></span>| <span data-ttu-id="b621d-837">Значение</span><span class="sxs-lookup"><span data-stu-id="b621d-837">Value</span></span>|
|---|---|
|[<span data-ttu-id="b621d-838">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="b621d-838">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b621d-839">1.0</span><span class="sxs-lookup"><span data-stu-id="b621d-839">1.0</span></span>|
|[<span data-ttu-id="b621d-840">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b621d-840">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b621d-841">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b621d-841">ReadItem</span></span>|
|[<span data-ttu-id="b621d-842">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b621d-842">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b621d-843">Чтение</span><span class="sxs-lookup"><span data-stu-id="b621d-843">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="b621d-844">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="b621d-844">Returns:</span></span>

<span data-ttu-id="b621d-p156">Объект, содержащий массив строк, которые соответствуют регулярным выражениям, определяемым в XML-файле манифеста. Имя каждого массива равно соответствующему значению атрибута `RegExName` подходящего правила `ItemHasRegularExpressionMatch` или атрибута `FilterName` соответствующего правила `ItemHasKnownEntity`.</span><span class="sxs-lookup"><span data-stu-id="b621d-p156">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="b621d-847">

<dt>Тип</dt>

</span><span class="sxs-lookup"><span data-stu-id="b621d-847">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="b621d-848">Object</span><span class="sxs-lookup"><span data-stu-id="b621d-848">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="b621d-849">Пример</span><span class="sxs-lookup"><span data-stu-id="b621d-849">Example</span></span>

<span data-ttu-id="b621d-850">В примере ниже показано, как получить доступ к массиву совпадений для элементов `fruits` и `veggies` регулярного выражения <rule>, которые указаны в манифесте.</rule></span><span class="sxs-lookup"><span data-stu-id="b621d-850">The following example shows how to access the array of matches for the regular expression <rule>elements `fruits` and `veggies`, which are specified in the manifest.</rule></span></span>

```js
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veges = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="b621d-851">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="b621d-851">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="b621d-852">Возвращает строковые значения в выбранном элементе, которые соответствуют именованному регулярному выражению, определенному в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="b621d-852">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="b621d-853">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="b621d-853">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="b621d-854">Метод `getRegExMatchesByName` возвращает строки, соответствующие регулярному выражению, которое определяется в элементе правила `ItemHasRegularExpressionMatch` в XML-файле манифеста, с использованием указанного значения элемента `RegExName`.</span><span class="sxs-lookup"><span data-stu-id="b621d-854">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="b621d-p157">Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты.</span><span class="sxs-lookup"><span data-stu-id="b621d-p157">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b621d-857">Параметры</span><span class="sxs-lookup"><span data-stu-id="b621d-857">Parameters:</span></span>

|<span data-ttu-id="b621d-858">Имя</span><span class="sxs-lookup"><span data-stu-id="b621d-858">Name</span></span>| <span data-ttu-id="b621d-859">Тип</span><span class="sxs-lookup"><span data-stu-id="b621d-859">Type</span></span>| <span data-ttu-id="b621d-860">Описание</span><span class="sxs-lookup"><span data-stu-id="b621d-860">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="b621d-861">String</span><span class="sxs-lookup"><span data-stu-id="b621d-861">String</span></span>|<span data-ttu-id="b621d-862">Имя элемента правила `ItemHasRegularExpressionMatch`, определяющее соответствующий фильтр.</span><span class="sxs-lookup"><span data-stu-id="b621d-862">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b621d-863">Требования</span><span class="sxs-lookup"><span data-stu-id="b621d-863">Requirements</span></span>

|<span data-ttu-id="b621d-864">Requirement</span><span class="sxs-lookup"><span data-stu-id="b621d-864">Requirement</span></span>| <span data-ttu-id="b621d-865">Значение</span><span class="sxs-lookup"><span data-stu-id="b621d-865">Value</span></span>|
|---|---|
|[<span data-ttu-id="b621d-866">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="b621d-866">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b621d-867">1.0</span><span class="sxs-lookup"><span data-stu-id="b621d-867">1.0</span></span>|
|[<span data-ttu-id="b621d-868">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b621d-868">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b621d-869">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b621d-869">ReadItem</span></span>|
|[<span data-ttu-id="b621d-870">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b621d-870">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b621d-871">Чтение</span><span class="sxs-lookup"><span data-stu-id="b621d-871">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="b621d-872">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="b621d-872">Returns:</span></span>

<span data-ttu-id="b621d-873">Массив строк, соответствующих регулярному выражению, определяемому в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="b621d-873">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="b621d-874">

<dt>Тип</dt>

</span><span class="sxs-lookup"><span data-stu-id="b621d-874">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="b621d-875">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="b621d-875">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="b621d-876">Пример</span><span class="sxs-lookup"><span data-stu-id="b621d-876">Example</span></span>

```js
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="b621d-877">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="b621d-877">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="b621d-878">Асинхронно возвращает данные, выбранные в теме или тексте сообщения.</span><span class="sxs-lookup"><span data-stu-id="b621d-878">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="b621d-p158">Если выделенный фрагмент отсутствует, но курсор находится в тексте или теме, метод возвращает значение NULL для выбранных данных. Если выбраны не текст и не тема, метод возвращает ошибку `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="b621d-p158">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b621d-881">Параметры</span><span class="sxs-lookup"><span data-stu-id="b621d-881">Parameters:</span></span>

|<span data-ttu-id="b621d-882">Имя</span><span class="sxs-lookup"><span data-stu-id="b621d-882">Name</span></span>| <span data-ttu-id="b621d-883">Тип</span><span class="sxs-lookup"><span data-stu-id="b621d-883">Type</span></span>| <span data-ttu-id="b621d-884">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="b621d-884">Attributes</span></span>| <span data-ttu-id="b621d-885">Описание</span><span class="sxs-lookup"><span data-stu-id="b621d-885">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="b621d-886">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="b621d-886">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="b621d-p159">Запрашивает формат данных. Если задано значение Text, метод возвращает обычный текст как строку, удаляя все имеющиеся HTML-теги. Если задано значение HTML, метод возвращает выделенный текст (обычный текст или HTML).</span><span class="sxs-lookup"><span data-stu-id="b621d-p159">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="b621d-890">Object</span><span class="sxs-lookup"><span data-stu-id="b621d-890">Object</span></span>| <span data-ttu-id="b621d-891">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="b621d-891">&lt;optional&gt;</span></span>|<span data-ttu-id="b621d-892">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="b621d-892">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="b621d-893">Object</span><span class="sxs-lookup"><span data-stu-id="b621d-893">Object</span></span>| <span data-ttu-id="b621d-894">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="b621d-894">&lt;optional&gt;</span></span>|<span data-ttu-id="b621d-895">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="b621d-895">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="b621d-896">функция</span><span class="sxs-lookup"><span data-stu-id="b621d-896">function</span></span>||<span data-ttu-id="b621d-897">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="b621d-897">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="b621d-898">Чтобы получить доступ к выбранным данным из метода обратного вызова, вызовите `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="b621d-898">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="b621d-899">Для доступа к исходному свойству, представляющему собой источник выбранных данных, вызовите параметр `asyncResult.value.sourceProperty`, который может иметь значение `body` или `subject`.</span><span class="sxs-lookup"><span data-stu-id="b621d-899">To access the source property that the selection comes from, call , which will be either  or .|</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b621d-900">Требования</span><span class="sxs-lookup"><span data-stu-id="b621d-900">Requirements</span></span>

|<span data-ttu-id="b621d-901">Requirement</span><span class="sxs-lookup"><span data-stu-id="b621d-901">Requirement</span></span>| <span data-ttu-id="b621d-902">Значение</span><span class="sxs-lookup"><span data-stu-id="b621d-902">Value</span></span>|
|---|---|
|[<span data-ttu-id="b621d-903">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="b621d-903">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b621d-904">1.2</span><span class="sxs-lookup"><span data-stu-id="b621d-904">1.2</span></span>|
|[<span data-ttu-id="b621d-905">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b621d-905">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b621d-906">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="b621d-906">ReadWriteItem</span></span>|
|[<span data-ttu-id="b621d-907">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b621d-907">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b621d-908">Создание</span><span class="sxs-lookup"><span data-stu-id="b621d-908">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="b621d-909">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="b621d-909">Returns:</span></span>

<span data-ttu-id="b621d-910">Выбранные данные в виде строки с форматом, определенным в параметре `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="b621d-910">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="b621d-911">

<dt>Тип</dt>

</span><span class="sxs-lookup"><span data-stu-id="b621d-911">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="b621d-912">String</span><span class="sxs-lookup"><span data-stu-id="b621d-912">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="b621d-913">Пример</span><span class="sxs-lookup"><span data-stu-id="b621d-913">Example</span></span>

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

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="b621d-914">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="b621d-914">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="b621d-915">Асинхронно загружает настраиваемые свойства для надстройки для выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="b621d-915">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="b621d-p161">Настраиваемые свойства сохраняются в виде пар "ключ-значение" для каждого приложения и каждого элемента. Этот метод возвращает объект `CustomProperties` при обратном вызове, который предоставляет методы для доступа к настраиваемым свойствам, характерным для текущего элемента и текущей надстройки. Настраиваемые свойства не шифруются для элемента, поэтому этот способ хранения не является безопасным.</span><span class="sxs-lookup"><span data-stu-id="b621d-p161">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b621d-919">Параметры</span><span class="sxs-lookup"><span data-stu-id="b621d-919">Parameters:</span></span>

|<span data-ttu-id="b621d-920">Имя</span><span class="sxs-lookup"><span data-stu-id="b621d-920">Name</span></span>| <span data-ttu-id="b621d-921">Тип</span><span class="sxs-lookup"><span data-stu-id="b621d-921">Type</span></span>| <span data-ttu-id="b621d-922">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="b621d-922">Attributes</span></span>| <span data-ttu-id="b621d-923">Описание</span><span class="sxs-lookup"><span data-stu-id="b621d-923">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="b621d-924">function</span><span class="sxs-lookup"><span data-stu-id="b621d-924">function</span></span>||<span data-ttu-id="b621d-925">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="b621d-925">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="b621d-926">Настраиваемые свойства предоставляются в виде объекта [`CustomProperties`](/javascript/api/outlook_1_3/office.customproperties) в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="b621d-926">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook_1_3/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="b621d-927">Этот объект позволяет получить, задать и удалить настраиваемые свойства из элемента, а также сохранить изменения, внесенные в набор настраиваемых свойств, на сервере.</span><span class="sxs-lookup"><span data-stu-id="b621d-927">The custom properties are provided as a CustomProperties object in the asyncResult.value property. This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="b621d-928">Object</span><span class="sxs-lookup"><span data-stu-id="b621d-928">Object</span></span>| <span data-ttu-id="b621d-929">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="b621d-929">&lt;optional&gt;</span></span>|<span data-ttu-id="b621d-930">Разработчики могут указать любой объект, к которому необходимо получить доступ, в функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="b621d-930">Developers can provide any object they wish to access in the callback method.</span></span> <span data-ttu-id="b621d-931">Доступ к этому объекту можно получить с помощью свойства `asyncResult.asyncContext` в функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="b621d-931">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.|</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b621d-932">Требования</span><span class="sxs-lookup"><span data-stu-id="b621d-932">Requirements</span></span>

|<span data-ttu-id="b621d-933">Requirement</span><span class="sxs-lookup"><span data-stu-id="b621d-933">Requirement</span></span>| <span data-ttu-id="b621d-934">Значение</span><span class="sxs-lookup"><span data-stu-id="b621d-934">Value</span></span>|
|---|---|
|[<span data-ttu-id="b621d-935">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="b621d-935">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b621d-936">1.0</span><span class="sxs-lookup"><span data-stu-id="b621d-936">1.0</span></span>|
|[<span data-ttu-id="b621d-937">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b621d-937">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b621d-938">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b621d-938">ReadItem</span></span>|
|[<span data-ttu-id="b621d-939">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b621d-939">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b621d-940">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="b621d-940">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="b621d-941">Пример</span><span class="sxs-lookup"><span data-stu-id="b621d-941">Example</span></span>

<span data-ttu-id="b621d-p164">Приведенный ниже пример кода показывает, как асинхронно загружать настраиваемые свойства, характерные для текущего элемента, с помощью метода `loadCustomPropertiesAsync`. Этот пример также показывает, как сохранять эти свойства на сервере с помощью метода `CustomProperties.saveAsync`. После загрузки настраиваемых свойств в этом примере кода метод `CustomProperties.get` используется для считывания настраиваемого свойства `myProp`, метод `CustomProperties.set` — для записи настраиваемого свойства `otherProp`, а метод `saveAsync` — для сохранения настраиваемых свойств.</span><span class="sxs-lookup"><span data-stu-id="b621d-p164">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="b621d-945">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="b621d-945">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="b621d-946">Удаляет вложение из сообщения или встречи.</span><span class="sxs-lookup"><span data-stu-id="b621d-946">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="b621d-p165">Метод `removeAttachmentAsync` удаляет из элемента вложение с указанным идентификатором. Идентификатор вложения рекомендуется использовать для удаления вложения, только если оно добавлено тем же почтовым приложением в ходе текущего сеанса. В Outlook Web App и Outlook Web App для устройств идентификатор вложения действителен только в рамках одного сеанса. Сеанс завершается, когда пользователь закрывает приложение или начинает создавать элемент во встроенной форме, а затем переходит из формы в отдельное окно.</span><span class="sxs-lookup"><span data-stu-id="b621d-p165">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b621d-951">Параметры</span><span class="sxs-lookup"><span data-stu-id="b621d-951">Parameters:</span></span>

|<span data-ttu-id="b621d-952">Имя</span><span class="sxs-lookup"><span data-stu-id="b621d-952">Name</span></span>| <span data-ttu-id="b621d-953">Тип</span><span class="sxs-lookup"><span data-stu-id="b621d-953">Type</span></span>| <span data-ttu-id="b621d-954">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="b621d-954">Attributes</span></span>| <span data-ttu-id="b621d-955">Описание</span><span class="sxs-lookup"><span data-stu-id="b621d-955">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="b621d-956">String</span><span class="sxs-lookup"><span data-stu-id="b621d-956">String</span></span>||<span data-ttu-id="b621d-p166">Идентификатор удаляемого вложения. Максимальная длина строки — 100 символов.</span><span class="sxs-lookup"><span data-stu-id="b621d-p166">The identifier of the attachment to remove. The maximum length of the string is 100 characters.</span></span>|
|`options`| <span data-ttu-id="b621d-959">Object</span><span class="sxs-lookup"><span data-stu-id="b621d-959">Object</span></span>| <span data-ttu-id="b621d-960">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="b621d-960">&lt;optional&gt;</span></span>|<span data-ttu-id="b621d-961">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="b621d-961">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="b621d-962">Object</span><span class="sxs-lookup"><span data-stu-id="b621d-962">Object</span></span>| <span data-ttu-id="b621d-963">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="b621d-963">&lt;optional&gt;</span></span>|<span data-ttu-id="b621d-964">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="b621d-964">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="b621d-965">функция</span><span class="sxs-lookup"><span data-stu-id="b621d-965">function</span></span>| <span data-ttu-id="b621d-966">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="b621d-966">&lt;optional&gt;</span></span>|<span data-ttu-id="b621d-967">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="b621d-967">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="b621d-968">Если удалить вложение не удается, свойство `asyncResult.error` содержит код ошибки с указанием ее причины.</span><span class="sxs-lookup"><span data-stu-id="b621d-968">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="b621d-969">Ошибки</span><span class="sxs-lookup"><span data-stu-id="b621d-969">Errors</span></span>

| <span data-ttu-id="b621d-970">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="b621d-970">Error code</span></span> | <span data-ttu-id="b621d-971">Описание</span><span class="sxs-lookup"><span data-stu-id="b621d-971">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="b621d-972">Идентификатор вложения не существует.</span><span class="sxs-lookup"><span data-stu-id="b621d-972">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="b621d-973">Требования</span><span class="sxs-lookup"><span data-stu-id="b621d-973">Requirements</span></span>

|<span data-ttu-id="b621d-974">Requirement</span><span class="sxs-lookup"><span data-stu-id="b621d-974">Requirement</span></span>| <span data-ttu-id="b621d-975">Значение</span><span class="sxs-lookup"><span data-stu-id="b621d-975">Value</span></span>|
|---|---|
|[<span data-ttu-id="b621d-976">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="b621d-976">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b621d-977">1.1</span><span class="sxs-lookup"><span data-stu-id="b621d-977">1.1</span></span>|
|[<span data-ttu-id="b621d-978">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b621d-978">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b621d-979">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="b621d-979">ReadWriteItem</span></span>|
|[<span data-ttu-id="b621d-980">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b621d-980">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b621d-981">Создание</span><span class="sxs-lookup"><span data-stu-id="b621d-981">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="b621d-982">Пример</span><span class="sxs-lookup"><span data-stu-id="b621d-982">Example</span></span>

<span data-ttu-id="b621d-983">Указанный ниже код удаляет вложение с идентификатором "0".</span><span class="sxs-lookup"><span data-stu-id="b621d-983">The following code removes an attachment with an identifier of '0'.</span></span>

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

####  <a name="saveasyncoptions-callback"></a><span data-ttu-id="b621d-984">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="b621d-984">saveAsync([options], callback)</span></span>

<span data-ttu-id="b621d-985">Асинхронно сохраняет элемент.</span><span class="sxs-lookup"><span data-stu-id="b621d-985">Asynchronously saves an item.</span></span>

<span data-ttu-id="b621d-p167">При вызове этот метод сохраняет текущее сообщение в виде черновика и возвращает идентификатор элемента с помощью метода обратного вызова. В Outlook Web App или интерактивном режиме Outlook этот элемент сохраняется на сервере. В Outlook в режиме кэширования этот элемент сохраняется в локальном кэше.</span><span class="sxs-lookup"><span data-stu-id="b621d-p167">When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="b621d-989">Если в режиме создания надстройка вызывает для элемента метод `saveAsync`, чтобы получить параметр `itemId` для использования с EWS или REST API, помните, что синхронизация элемента с сервером может занять много времени, если Outlook работает в режиме кэширования данных.</span><span class="sxs-lookup"><span data-stu-id="b621d-989">Note: If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server. Until the item is synced, using the  will return an error.</span></span> <span data-ttu-id="b621d-990">До окончания синхронизации применение параметра `itemId` будет приводить к ошибке.</span><span class="sxs-lookup"><span data-stu-id="b621d-990">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="b621d-p169">Если метод `saveAsync` вызывается для встречи в режиме создания, она сохраняется как обычная встреча в календаре пользователя, а не как черновик. При сохранении новой встречи приглашения не отправляются. При сохранении существующей встречи уведомления отправляются добавленным или удаленным участникам.</span><span class="sxs-lookup"><span data-stu-id="b621d-p169">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="b621d-994">Следующие клиенты отличаются другим поведением метода `saveAsync` в отношении встреч в режиме создания:</span><span class="sxs-lookup"><span data-stu-id="b621d-994">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="b621d-995">Outlook для Mac не поддерживает метод `saveAsync` для собраний в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="b621d-995">Note: Mac Outlook does not support `saveAsync` on a meeting in compose mode. Calling  on a meeting in Mac Outlook will return an error.</span></span> <span data-ttu-id="b621d-996">Метод `saveAsync`, вызванный для собрания в Outlook для Mac, возвращает ошибку.</span><span class="sxs-lookup"><span data-stu-id="b621d-996">Note: Mac Outlook does not support  on a meeting in compose mode. Calling `saveAsync` on a meeting in Mac Outlook will return an error.</span></span>
> - <span data-ttu-id="b621d-997">Outlook в Интернете всегда отправляет приглашение или обновление при вызове метода `saveAsync` для встречи в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="b621d-997">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b621d-998">Параметры</span><span class="sxs-lookup"><span data-stu-id="b621d-998">Parameters:</span></span>

|<span data-ttu-id="b621d-999">Имя</span><span class="sxs-lookup"><span data-stu-id="b621d-999">Name</span></span>| <span data-ttu-id="b621d-1000">Тип</span><span class="sxs-lookup"><span data-stu-id="b621d-1000">Type</span></span>| <span data-ttu-id="b621d-1001">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="b621d-1001">Attributes</span></span>| <span data-ttu-id="b621d-1002">Описание</span><span class="sxs-lookup"><span data-stu-id="b621d-1002">Description</span></span>|
|---|---|---|---|
|`options`| <span data-ttu-id="b621d-1003">Object</span><span class="sxs-lookup"><span data-stu-id="b621d-1003">Object</span></span>| <span data-ttu-id="b621d-1004">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="b621d-1004">&lt;optional&gt;</span></span>|<span data-ttu-id="b621d-1005">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="b621d-1005">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="b621d-1006">Object</span><span class="sxs-lookup"><span data-stu-id="b621d-1006">Object</span></span>| <span data-ttu-id="b621d-1007">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="b621d-1007">&lt;optional&gt;</span></span>|<span data-ttu-id="b621d-1008">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="b621d-1008">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="b621d-1009">функция</span><span class="sxs-lookup"><span data-stu-id="b621d-1009">function</span></span>||<span data-ttu-id="b621d-1010">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="b621d-1010">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="b621d-1011">После успешного выполнения идентификатор элемента будет предоставлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="b621d-1011">On success, the item identifier is provided in the `asyncResult.value` property.|</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b621d-1012">Требования</span><span class="sxs-lookup"><span data-stu-id="b621d-1012">Requirements</span></span>

|<span data-ttu-id="b621d-1013">Requirement</span><span class="sxs-lookup"><span data-stu-id="b621d-1013">Requirement</span></span>| <span data-ttu-id="b621d-1014">Значение</span><span class="sxs-lookup"><span data-stu-id="b621d-1014">Value</span></span>|
|---|---|
|[<span data-ttu-id="b621d-1015">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="b621d-1015">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b621d-1016">1.3</span><span class="sxs-lookup"><span data-stu-id="b621d-1016">1.3</span></span>|
|[<span data-ttu-id="b621d-1017">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b621d-1017">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b621d-1018">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="b621d-1018">ReadWriteItem</span></span>|
|[<span data-ttu-id="b621d-1019">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b621d-1019">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b621d-1020">Создание</span><span class="sxs-lookup"><span data-stu-id="b621d-1020">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="b621d-1021">Примеры</span><span class="sxs-lookup"><span data-stu-id="b621d-1021">Examples</span></span>

```js
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result
  });
```

<span data-ttu-id="b621d-p171">Ниже приведен пример параметра `result`, переданного функции обратного вызова. Свойство `value` содержит идентификатор элемента.</span><span class="sxs-lookup"><span data-stu-id="b621d-p171">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

####  <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="b621d-1024">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="b621d-1024">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="b621d-1025">Асинхронно вставляет данные в текст или тему сообщения.</span><span class="sxs-lookup"><span data-stu-id="b621d-1025">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="b621d-p172">Метод `setSelectedDataAsync` вставляет указанную строку в местоположение курсора в теме или тексте элемента либо, если текст выделен в редакторе, он заменяет выделенный текст. Если курсор находится вне текста или темы элемента, возвращается ошибка. После вставки курсор помещается в конец вставленного содержимого.</span><span class="sxs-lookup"><span data-stu-id="b621d-p172">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b621d-1029">Параметры:</span><span class="sxs-lookup"><span data-stu-id="b621d-1029">Parameters:</span></span>

|<span data-ttu-id="b621d-1030">Имя</span><span class="sxs-lookup"><span data-stu-id="b621d-1030">Name</span></span>| <span data-ttu-id="b621d-1031">Тип</span><span class="sxs-lookup"><span data-stu-id="b621d-1031">Type</span></span>| <span data-ttu-id="b621d-1032">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="b621d-1032">Attributes</span></span>| <span data-ttu-id="b621d-1033">Описание</span><span class="sxs-lookup"><span data-stu-id="b621d-1033">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="b621d-1034">String</span><span class="sxs-lookup"><span data-stu-id="b621d-1034">String</span></span>||<span data-ttu-id="b621d-p173">Вставляемые данные. Объем данных не должен превышать 1 000 000 символов. Если передано больше 1 000 000 символов, возвращается исключение `ArgumentOutOfRange`.</span><span class="sxs-lookup"><span data-stu-id="b621d-p173">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="b621d-1038">Object</span><span class="sxs-lookup"><span data-stu-id="b621d-1038">Object</span></span>| <span data-ttu-id="b621d-1039">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="b621d-1039">&lt;optional&gt;</span></span>|<span data-ttu-id="b621d-1040">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="b621d-1040">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="b621d-1041">Object</span><span class="sxs-lookup"><span data-stu-id="b621d-1041">Object</span></span>| <span data-ttu-id="b621d-1042">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="b621d-1042">&lt;optional&gt;</span></span>|<span data-ttu-id="b621d-1043">В методе обратного вызова разработчики могут указать любой объект, к которому необходимо получить доступ.</span><span class="sxs-lookup"><span data-stu-id="b621d-1043">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`| [<span data-ttu-id="b621d-1044">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="b621d-1044">Office.CoercionType</span></span>](office.md#coerciontype-string)| <span data-ttu-id="b621d-1045">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="b621d-1045">&lt;optional&gt;</span></span>|<span data-ttu-id="b621d-p174">Если задано значение `text`, текущий стиль применяется в Outlook Web App и Outlook. Если поле представляет собой редактор HTML, вставляются только текстовые данные, даже если они имеют формат HTML.</span><span class="sxs-lookup"><span data-stu-id="b621d-p174">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="b621d-p175">Если задано значение `html` и поле (не тема) поддерживает HTML, в Outlook Web App применяется текущий стиль, а в Outlook — стиль по умолчанию. Если поле является текстовым, возвращается ошибка `InvalidDataFormat`.</span><span class="sxs-lookup"><span data-stu-id="b621d-p175">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="b621d-1050">Если свойство `coercionType` не задано, результат зависит от поля: если поле имеет формат HTML, используется текст в формате HTML, а если поле текстовое, применяется обычный текст.</span><span class="sxs-lookup"><span data-stu-id="b621d-1050">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="b621d-1051">функция</span><span class="sxs-lookup"><span data-stu-id="b621d-1051">function</span></span>||<span data-ttu-id="b621d-1052">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="b621d-1052">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="b621d-1053">Требования</span><span class="sxs-lookup"><span data-stu-id="b621d-1053">Requirements</span></span>

|<span data-ttu-id="b621d-1054">Requirement</span><span class="sxs-lookup"><span data-stu-id="b621d-1054">Requirement</span></span>| <span data-ttu-id="b621d-1055">Значение</span><span class="sxs-lookup"><span data-stu-id="b621d-1055">Value</span></span>|
|---|---|
|[<span data-ttu-id="b621d-1056">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="b621d-1056">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b621d-1057">1.2</span><span class="sxs-lookup"><span data-stu-id="b621d-1057">1.2</span></span>|
|[<span data-ttu-id="b621d-1058">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b621d-1058">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b621d-1059">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="b621d-1059">ReadWriteItem</span></span>|
|[<span data-ttu-id="b621d-1060">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b621d-1060">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b621d-1061">Создание</span><span class="sxs-lookup"><span data-stu-id="b621d-1061">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="b621d-1062">Пример</span><span class="sxs-lookup"><span data-stu-id="b621d-1062">Example</span></span>

```js
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```