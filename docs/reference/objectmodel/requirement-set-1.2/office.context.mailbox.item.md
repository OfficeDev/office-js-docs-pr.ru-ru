
# <a name="item"></a><span data-ttu-id="1dab6-101">item</span><span class="sxs-lookup"><span data-stu-id="1dab6-101">item</span></span>

### <span data-ttu-id="1dab6-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md). item</span><span class="sxs-lookup"><span data-stu-id="1dab6-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md). item</span></span>

<span data-ttu-id="1dab6-p102">Пространство имен `item` используется для доступа к выбранному в данный момент сообщению, приглашению на собрание или описанию встречи. Вы можете определить тип пространства имен `item` с помощью свойства [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook12officemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="1dab6-p102">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook12officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="1dab6-106">Требования</span><span class="sxs-lookup"><span data-stu-id="1dab6-106">Requirements</span></span>

|<span data-ttu-id="1dab6-107">Требование</span><span class="sxs-lookup"><span data-stu-id="1dab6-107">Requirement</span></span>| <span data-ttu-id="1dab6-108">Значение</span><span class="sxs-lookup"><span data-stu-id="1dab6-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="1dab6-109">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="1dab6-109">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1dab6-110">1.0</span><span class="sxs-lookup"><span data-stu-id="1dab6-110">1.0</span></span>|
|[<span data-ttu-id="1dab6-111">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="1dab6-111">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1dab6-112">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="1dab6-112">Restricted</span></span>|
|[<span data-ttu-id="1dab6-113">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="1dab6-113">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1dab6-114">Compose (создание) или read (чтение)</span><span class="sxs-lookup"><span data-stu-id="1dab6-114">Compose or read</span></span>|

### <a name="example"></a><span data-ttu-id="1dab6-115">Пример</span><span class="sxs-lookup"><span data-stu-id="1dab6-115">Example</span></span>

<span data-ttu-id="1dab6-116">В примере кода JavaScript, приведенном ниже, показано, как получить доступ к свойству `subject` текущего элемента в Outlook.</span><span class="sxs-lookup"><span data-stu-id="1dab6-116">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="1dab6-117">Члены</span><span class="sxs-lookup"><span data-stu-id="1dab6-117">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook12officeattachmentdetails"></a><span data-ttu-id="1dab6-118">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_2/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="1dab6-118">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_2/office.attachmentdetails)></span></span>

<span data-ttu-id="1dab6-p103">Получает массив вложений для элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="1dab6-p103">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="1dab6-121">Некоторые типы файлов блокируются Outlook из-за потенциальных проблем безопасности и поэтому не возвращаются.</span><span class="sxs-lookup"><span data-stu-id="1dab6-121">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="1dab6-122">Для получения дополнительных сведений см. [Блокированные вложения в Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="1dab6-122">For more information see [Payments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="1dab6-123">Тип:</span><span class="sxs-lookup"><span data-stu-id="1dab6-123">Type:</span></span>

*   <span data-ttu-id="1dab6-124">Array.<[AttachmentDetails](/javascript/api/outlook_1_2/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="1dab6-124">Array.<[AttachmentDetails](/javascript/api/outlook_1_2/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="1dab6-125">Требования</span><span class="sxs-lookup"><span data-stu-id="1dab6-125">Requirements</span></span>

|<span data-ttu-id="1dab6-126">Требование</span><span class="sxs-lookup"><span data-stu-id="1dab6-126">Requirement</span></span>| <span data-ttu-id="1dab6-127">Значение</span><span class="sxs-lookup"><span data-stu-id="1dab6-127">Value</span></span>|
|---|---|
|[<span data-ttu-id="1dab6-128">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="1dab6-128">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1dab6-129">1.0</span><span class="sxs-lookup"><span data-stu-id="1dab6-129">1.0</span></span>|
|[<span data-ttu-id="1dab6-130">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="1dab6-130">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1dab6-131">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1dab6-131">ReadItem</span></span>|
|[<span data-ttu-id="1dab6-132">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="1dab6-132">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1dab6-133">Чтение</span><span class="sxs-lookup"><span data-stu-id="1dab6-133">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1dab6-134">Пример</span><span class="sxs-lookup"><span data-stu-id="1dab6-134">Example</span></span>

<span data-ttu-id="1dab6-135">С помощью приведенного ниже кода можно создать HTML-строку с подробными сведениями обо всех вложениях для текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="1dab6-135">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

####  <a name="bcc-recipientsjavascriptapioutlook12officerecipients"></a><span data-ttu-id="1dab6-136">bcc :[Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="1dab6-136">bcc :[Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

<span data-ttu-id="1dab6-137">Получает объект, который предоставляет методы для получения или обновления получателей в строке Bcc (скрытой копии) сообщения.</span><span class="sxs-lookup"><span data-stu-id="1dab6-137">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="1dab6-138">Только в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="1dab6-138">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="1dab6-139">Тип:</span><span class="sxs-lookup"><span data-stu-id="1dab6-139">Type:</span></span>

*   [<span data-ttu-id="1dab6-140">Recipients</span><span class="sxs-lookup"><span data-stu-id="1dab6-140">Recipients</span></span>](/javascript/api/outlook_1_2/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="1dab6-141">Требования</span><span class="sxs-lookup"><span data-stu-id="1dab6-141">Requirements</span></span>

|<span data-ttu-id="1dab6-142">Требование</span><span class="sxs-lookup"><span data-stu-id="1dab6-142">Requirement</span></span>| <span data-ttu-id="1dab6-143">Значение</span><span class="sxs-lookup"><span data-stu-id="1dab6-143">Value</span></span>|
|---|---|
|[<span data-ttu-id="1dab6-144">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="1dab6-144">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1dab6-145">1.1</span><span class="sxs-lookup"><span data-stu-id="1dab6-145">1.1</span></span>|
|[<span data-ttu-id="1dab6-146">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="1dab6-146">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1dab6-147">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1dab6-147">ReadItem</span></span>|
|[<span data-ttu-id="1dab6-148">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="1dab6-148">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1dab6-149">Создание</span><span class="sxs-lookup"><span data-stu-id="1dab6-149">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="1dab6-150">Пример</span><span class="sxs-lookup"><span data-stu-id="1dab6-150">Example</span></span>

```JavaScript
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlook12officebody"></a><span data-ttu-id="1dab6-151">body :[Body](/javascript/api/outlook_1_2/office.body)</span><span class="sxs-lookup"><span data-stu-id="1dab6-151">body :[Body](/javascript/api/outlook_1_2/office.body)</span></span>

<span data-ttu-id="1dab6-152">Получает объект, предоставляющий методы для работы с основным текстом элемента.</span><span class="sxs-lookup"><span data-stu-id="1dab6-152">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="1dab6-153">Тип:</span><span class="sxs-lookup"><span data-stu-id="1dab6-153">Type:</span></span>

*   [<span data-ttu-id="1dab6-154">Body</span><span class="sxs-lookup"><span data-stu-id="1dab6-154">Body</span></span>](/javascript/api/outlook_1_2/office.body)

##### <a name="requirements"></a><span data-ttu-id="1dab6-155">Требования</span><span class="sxs-lookup"><span data-stu-id="1dab6-155">Requirements</span></span>

|<span data-ttu-id="1dab6-156">Требование</span><span class="sxs-lookup"><span data-stu-id="1dab6-156">Requirement</span></span>| <span data-ttu-id="1dab6-157">Значение</span><span class="sxs-lookup"><span data-stu-id="1dab6-157">Value</span></span>|
|---|---|
|[<span data-ttu-id="1dab6-158">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="1dab6-158">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1dab6-159">1.1</span><span class="sxs-lookup"><span data-stu-id="1dab6-159">1.1</span></span>|
|[<span data-ttu-id="1dab6-160">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="1dab6-160">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1dab6-161">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1dab6-161">ReadItem</span></span>|
|[<span data-ttu-id="1dab6-162">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="1dab6-162">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1dab6-163">Compose (создание) или read (чтение)</span><span class="sxs-lookup"><span data-stu-id="1dab6-163">Compose or read</span></span>|

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlook12officeemailaddressdetailsrecipientsjavascriptapioutlook12officerecipients"></a><span data-ttu-id="1dab6-164">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="1dab6-164">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

<span data-ttu-id="1dab6-165">Предоставляет доступ к получателям Cc (копия) сообщения.</span><span class="sxs-lookup"><span data-stu-id="1dab6-165">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="1dab6-166">Уровень доступа и тип объекта зависит от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="1dab6-166">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="1dab6-167">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="1dab6-167">Read mode</span></span>

<span data-ttu-id="1dab6-p107">Свойство `cc` возвращает массив, который содержит объект `EmailAddressDetails` для каждого получателя, указанного в строке **Cc (Копия)** сообщения. Коллекция может включать не более 100 членов.</span><span class="sxs-lookup"><span data-stu-id="1dab6-p107">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="1dab6-170">Режим создания</span><span class="sxs-lookup"><span data-stu-id="1dab6-170">Compose mode</span></span>

<span data-ttu-id="1dab6-171">Свойство `cc` возвращает объект `Recipients`, предоставляющий методы для получения или обновления получателей, которые указаны в строке **Cc (Копия)** сообщения.</span><span class="sxs-lookup"><span data-stu-id="1dab6-171">The `cc` property returns a `Recipients` object that provides methods for manipulating the recipients on the **Cc** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="1dab6-172">Тип:</span><span class="sxs-lookup"><span data-stu-id="1dab6-172">Type:</span></span>

*   <span data-ttu-id="1dab6-173">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="1dab6-173">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="1dab6-174">Требования</span><span class="sxs-lookup"><span data-stu-id="1dab6-174">Requirements</span></span>

|<span data-ttu-id="1dab6-175">Требование</span><span class="sxs-lookup"><span data-stu-id="1dab6-175">Requirement</span></span>| <span data-ttu-id="1dab6-176">Значение</span><span class="sxs-lookup"><span data-stu-id="1dab6-176">Value</span></span>|
|---|---|
|[<span data-ttu-id="1dab6-177">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="1dab6-177">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1dab6-178">1.0</span><span class="sxs-lookup"><span data-stu-id="1dab6-178">1.0</span></span>|
|[<span data-ttu-id="1dab6-179">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="1dab6-179">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1dab6-180">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1dab6-180">ReadItem</span></span>|
|[<span data-ttu-id="1dab6-181">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="1dab6-181">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1dab6-182">Compose (создание) или read (чтение)</span><span class="sxs-lookup"><span data-stu-id="1dab6-182">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="1dab6-183">Пример</span><span class="sxs-lookup"><span data-stu-id="1dab6-183">Example</span></span>

```JavaScript
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="1dab6-184">(допускающий значение NULL) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="1dab6-184">(nullable) conversationId :String</span></span>

<span data-ttu-id="1dab6-185">Получает идентификатор беседы по электронной почте, содержащей конкретное сообщение.</span><span class="sxs-lookup"><span data-stu-id="1dab6-185">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="1dab6-p108">Вы можете получить целочисленное значение этого свойства, если ваше почтовое приложение активируется в формах просмотра или формах создания ответов. Если пользователь изменит тему ответа, после его отправки идентификатор беседы будет изменен, и полученное ранее значение будет недействительным.</span><span class="sxs-lookup"><span data-stu-id="1dab6-p108">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="1dab6-p109">Это свойство имеет значение NULL для нового элемента в форме создания. Свойство `conversationId` вернет значение, если пользователь задаст тему и сохранит элемент.</span><span class="sxs-lookup"><span data-stu-id="1dab6-p109">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="1dab6-190">Тип:</span><span class="sxs-lookup"><span data-stu-id="1dab6-190">Type:</span></span>

*   <span data-ttu-id="1dab6-191">String</span><span class="sxs-lookup"><span data-stu-id="1dab6-191">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="1dab6-192">Требования</span><span class="sxs-lookup"><span data-stu-id="1dab6-192">Requirements</span></span>

|<span data-ttu-id="1dab6-193">Требование</span><span class="sxs-lookup"><span data-stu-id="1dab6-193">Requirement</span></span>| <span data-ttu-id="1dab6-194">Значение</span><span class="sxs-lookup"><span data-stu-id="1dab6-194">Value</span></span>|
|---|---|
|[<span data-ttu-id="1dab6-195">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="1dab6-195">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1dab6-196">1.0</span><span class="sxs-lookup"><span data-stu-id="1dab6-196">1.0</span></span>|
|[<span data-ttu-id="1dab6-197">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="1dab6-197">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1dab6-198">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1dab6-198">ReadItem</span></span>|
|[<span data-ttu-id="1dab6-199">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="1dab6-199">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1dab6-200">Compose (создание) или read (чтение)</span><span class="sxs-lookup"><span data-stu-id="1dab6-200">Compose or read</span></span>|

#### <a name="datetimecreated-date"></a><span data-ttu-id="1dab6-201">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="1dab6-201">dateTimeCreated :Date</span></span>

<span data-ttu-id="1dab6-p110">Получает дату и время создания элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="1dab6-p110">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="1dab6-204">Тип:</span><span class="sxs-lookup"><span data-stu-id="1dab6-204">Type:</span></span>

*   <span data-ttu-id="1dab6-205">Date</span><span class="sxs-lookup"><span data-stu-id="1dab6-205">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="1dab6-206">Требования</span><span class="sxs-lookup"><span data-stu-id="1dab6-206">Requirements</span></span>

|<span data-ttu-id="1dab6-207">Требование</span><span class="sxs-lookup"><span data-stu-id="1dab6-207">Requirement</span></span>| <span data-ttu-id="1dab6-208">Значение</span><span class="sxs-lookup"><span data-stu-id="1dab6-208">Value</span></span>|
|---|---|
|[<span data-ttu-id="1dab6-209">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="1dab6-209">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1dab6-210">1.0</span><span class="sxs-lookup"><span data-stu-id="1dab6-210">1.0</span></span>|
|[<span data-ttu-id="1dab6-211">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="1dab6-211">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1dab6-212">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1dab6-212">ReadItem</span></span>|
|[<span data-ttu-id="1dab6-213">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="1dab6-213">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1dab6-214">Чтение</span><span class="sxs-lookup"><span data-stu-id="1dab6-214">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1dab6-215">Пример</span><span class="sxs-lookup"><span data-stu-id="1dab6-215">Example</span></span>

```JavaScript
var created = Office.context.mailbox.item.dateTimeCreated;
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="1dab6-216">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="1dab6-216">dateTimeModified :Date</span></span>

<span data-ttu-id="1dab6-p111">Получает дату и время последнего изменения элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="1dab6-p111">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="1dab6-219">Этот элемент не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="1dab6-219">Note: This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="1dab6-220">Тип:</span><span class="sxs-lookup"><span data-stu-id="1dab6-220">Type:</span></span>

*   <span data-ttu-id="1dab6-221">Date</span><span class="sxs-lookup"><span data-stu-id="1dab6-221">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="1dab6-222">Требования</span><span class="sxs-lookup"><span data-stu-id="1dab6-222">Requirements</span></span>

|<span data-ttu-id="1dab6-223">Требование</span><span class="sxs-lookup"><span data-stu-id="1dab6-223">Requirement</span></span>| <span data-ttu-id="1dab6-224">Значение</span><span class="sxs-lookup"><span data-stu-id="1dab6-224">Value</span></span>|
|---|---|
|[<span data-ttu-id="1dab6-225">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="1dab6-225">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1dab6-226">1.0</span><span class="sxs-lookup"><span data-stu-id="1dab6-226">1.0</span></span>|
|[<span data-ttu-id="1dab6-227">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="1dab6-227">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1dab6-228">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1dab6-228">ReadItem</span></span>|
|[<span data-ttu-id="1dab6-229">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="1dab6-229">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1dab6-230">Чтение</span><span class="sxs-lookup"><span data-stu-id="1dab6-230">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1dab6-231">Пример</span><span class="sxs-lookup"><span data-stu-id="1dab6-231">Example</span></span>

```JavaScript
var modified = Office.context.mailbox.item.dateTimeModified;
```

####  <a name="end-datetimejavascriptapioutlook12officetime"></a><span data-ttu-id="1dab6-232">end :Date|[Time](/javascript/api/outlook_1_2/office.time)</span><span class="sxs-lookup"><span data-stu-id="1dab6-232">end :Date|[Time](/javascript/api/outlook_1_2/office.time)</span></span>

<span data-ttu-id="1dab6-233">Получает или задает дату и время окончания встречи.</span><span class="sxs-lookup"><span data-stu-id="1dab6-233">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="1dab6-p112">Свойство `end` представлено в виде значения даты и времени в формате UTC. Преобразовать значение свойства end в местные значения даты и времени клиента можно с помощью метода [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook12officelocalclienttime).</span><span class="sxs-lookup"><span data-stu-id="1dab6-p112">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook12officelocalclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="1dab6-236">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="1dab6-236">Read mode</span></span>

<span data-ttu-id="1dab6-237">Свойство `end` возвращает объект `Date`.</span><span class="sxs-lookup"><span data-stu-id="1dab6-237">The `end` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="1dab6-238">Режим создания</span><span class="sxs-lookup"><span data-stu-id="1dab6-238">Compose mode</span></span>

<span data-ttu-id="1dab6-239">Свойство `end` возвращает объект `Time`.</span><span class="sxs-lookup"><span data-stu-id="1dab6-239">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="1dab6-240">Если вы задаете время окончания с помощью метода [`Time.setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-), необходимо использовать метод [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) для преобразования местного времени на клиенте в формат UTC для сервера.</span><span class="sxs-lookup"><span data-stu-id="1dab6-240">When you use the [`Time.setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="1dab6-241">Тип:</span><span class="sxs-lookup"><span data-stu-id="1dab6-241">Type:</span></span>

*   <span data-ttu-id="1dab6-242">Date | [Time](/javascript/api/outlook_1_2/office.time)</span><span class="sxs-lookup"><span data-stu-id="1dab6-242">Date | [Time](/javascript/api/outlook_1_2/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="1dab6-243">Требования</span><span class="sxs-lookup"><span data-stu-id="1dab6-243">Requirements</span></span>

|<span data-ttu-id="1dab6-244">Требование</span><span class="sxs-lookup"><span data-stu-id="1dab6-244">Requirement</span></span>| <span data-ttu-id="1dab6-245">Значение</span><span class="sxs-lookup"><span data-stu-id="1dab6-245">Value</span></span>|
|---|---|
|[<span data-ttu-id="1dab6-246">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="1dab6-246">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1dab6-247">1.0</span><span class="sxs-lookup"><span data-stu-id="1dab6-247">1.0</span></span>|
|[<span data-ttu-id="1dab6-248">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="1dab6-248">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1dab6-249">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1dab6-249">ReadItem</span></span>|
|[<span data-ttu-id="1dab6-250">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="1dab6-250">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1dab6-251">Compose (создание) или read (чтение)</span><span class="sxs-lookup"><span data-stu-id="1dab6-251">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="1dab6-252">Пример</span><span class="sxs-lookup"><span data-stu-id="1dab6-252">Example</span></span>

<span data-ttu-id="1dab6-253">В примере ниже показано, как с помощью метода [`setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) объекта `Time` задать время окончания встречи в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="1dab6-253">The following example sets the end time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

#### <a name="from-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails"></a><span data-ttu-id="1dab6-254">from :[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="1dab6-254">from :[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)</span></span>

<span data-ttu-id="1dab6-p113">Получает электронный адрес отправителя сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="1dab6-p113">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="1dab6-p114">Свойства `from` и [`sender`](#sender-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails) представляют одно лицо, если сообщение не отправлено представителем. В противном случае свойство `from` представляет лицо, делегировавшее полномочия, а свойство sender — представителя.</span><span class="sxs-lookup"><span data-stu-id="1dab6-p114">The `from` and [`sender`](#sender-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="1dab6-259">Свойство `recipientType` объекта `EmailAddressDetails` в свойстве `from` — `undefined`.</span><span class="sxs-lookup"><span data-stu-id="1dab6-259">Note: The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="1dab6-260">Тип:</span><span class="sxs-lookup"><span data-stu-id="1dab6-260">Type:</span></span>

*   [<span data-ttu-id="1dab6-261">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="1dab6-261">EmailAddressDetails</span></span>](/javascript/api/outlook_1_2/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="1dab6-262">Требования</span><span class="sxs-lookup"><span data-stu-id="1dab6-262">Requirements</span></span>

|<span data-ttu-id="1dab6-263">Требование</span><span class="sxs-lookup"><span data-stu-id="1dab6-263">Requirement</span></span>| <span data-ttu-id="1dab6-264">Значение</span><span class="sxs-lookup"><span data-stu-id="1dab6-264">Value</span></span>|
|---|---|
|[<span data-ttu-id="1dab6-265">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="1dab6-265">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1dab6-266">1.0</span><span class="sxs-lookup"><span data-stu-id="1dab6-266">1.0</span></span>|
|[<span data-ttu-id="1dab6-267">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="1dab6-267">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1dab6-268">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1dab6-268">ReadItem</span></span>|
|[<span data-ttu-id="1dab6-269">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="1dab6-269">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1dab6-270">Чтение</span><span class="sxs-lookup"><span data-stu-id="1dab6-270">Read</span></span>|

#### <a name="internetmessageid-string"></a><span data-ttu-id="1dab6-271">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="1dab6-271">internetMessageId :String</span></span>

<span data-ttu-id="1dab6-p115">Получает идентификатор интернет-сообщения для электронного сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="1dab6-p115">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="1dab6-274">Тип:</span><span class="sxs-lookup"><span data-stu-id="1dab6-274">Type:</span></span>

*   <span data-ttu-id="1dab6-275">String</span><span class="sxs-lookup"><span data-stu-id="1dab6-275">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="1dab6-276">Требования</span><span class="sxs-lookup"><span data-stu-id="1dab6-276">Requirements</span></span>

|<span data-ttu-id="1dab6-277">Требование</span><span class="sxs-lookup"><span data-stu-id="1dab6-277">Requirement</span></span>| <span data-ttu-id="1dab6-278">Значение</span><span class="sxs-lookup"><span data-stu-id="1dab6-278">Value</span></span>|
|---|---|
|[<span data-ttu-id="1dab6-279">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="1dab6-279">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1dab6-280">1.0</span><span class="sxs-lookup"><span data-stu-id="1dab6-280">1.0</span></span>|
|[<span data-ttu-id="1dab6-281">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="1dab6-281">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1dab6-282">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1dab6-282">ReadItem</span></span>|
|[<span data-ttu-id="1dab6-283">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="1dab6-283">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1dab6-284">Чтение</span><span class="sxs-lookup"><span data-stu-id="1dab6-284">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1dab6-285">Пример</span><span class="sxs-lookup"><span data-stu-id="1dab6-285">Example</span></span>

```JavaScript
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="1dab6-286">itemClass: String</span><span class="sxs-lookup"><span data-stu-id="1dab6-286">itemClass :String</span></span>

<span data-ttu-id="1dab6-p116">Получает класс элемента веб-служб Exchange для выбранного элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="1dab6-p116">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="1dab6-p117">Свойство `itemClass` указывает класс сообщения выбранного элемента. Ниже приводятся классы сообщения по умолчанию для элемента сообщения или встречи.</span><span class="sxs-lookup"><span data-stu-id="1dab6-p117">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="1dab6-291">Тип</span><span class="sxs-lookup"><span data-stu-id="1dab6-291">Type</span></span> | <span data-ttu-id="1dab6-292">Описание</span><span class="sxs-lookup"><span data-stu-id="1dab6-292">Description</span></span> | <span data-ttu-id="1dab6-293">класс элемента</span><span class="sxs-lookup"><span data-stu-id="1dab6-293">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="1dab6-294">Элементы встречи</span><span class="sxs-lookup"><span data-stu-id="1dab6-294">Appointment items</span></span> | <span data-ttu-id="1dab6-295">Это элементы календаря для класса элемента `IPM.Appointment` или `IPM.Appointment.Occurence`.</span><span class="sxs-lookup"><span data-stu-id="1dab6-295">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurence` |
| <span data-ttu-id="1dab6-296">Элементы сообщения</span><span class="sxs-lookup"><span data-stu-id="1dab6-296">Message items</span></span> | <span data-ttu-id="1dab6-297">Сюда входят электронные сообщения, для которых по умолчанию задан класс сообщения `IPM.Note`, а также приглашения на собрания, ответы на них и уведомления об их отмене, использующие `IPM.Schedule.Meeting` в качестве базового класса сообщения.</span><span class="sxs-lookup"><span data-stu-id="1dab6-297">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="1dab6-298">Можно создавать настраиваемые классы сообщения, расширяющие классы сообщения по умолчанию, например настраиваемый класс сообщения о встрече `IPM.Appointment.Contoso`.</span><span class="sxs-lookup"><span data-stu-id="1dab6-298">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="1dab6-299">Тип:</span><span class="sxs-lookup"><span data-stu-id="1dab6-299">Type:</span></span>

*   <span data-ttu-id="1dab6-300">String</span><span class="sxs-lookup"><span data-stu-id="1dab6-300">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="1dab6-301">Требования</span><span class="sxs-lookup"><span data-stu-id="1dab6-301">Requirements</span></span>

|<span data-ttu-id="1dab6-302">Требование</span><span class="sxs-lookup"><span data-stu-id="1dab6-302">Requirement</span></span>| <span data-ttu-id="1dab6-303">Значение</span><span class="sxs-lookup"><span data-stu-id="1dab6-303">Value</span></span>|
|---|---|
|[<span data-ttu-id="1dab6-304">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="1dab6-304">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1dab6-305">1.0</span><span class="sxs-lookup"><span data-stu-id="1dab6-305">1.0</span></span>|
|[<span data-ttu-id="1dab6-306">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="1dab6-306">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1dab6-307">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1dab6-307">ReadItem</span></span>|
|[<span data-ttu-id="1dab6-308">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="1dab6-308">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1dab6-309">Чтение</span><span class="sxs-lookup"><span data-stu-id="1dab6-309">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1dab6-310">Пример</span><span class="sxs-lookup"><span data-stu-id="1dab6-310">Example</span></span>

```JavaScript
var itemClass = Office.context.mailbox.item.itemClass;
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="1dab6-311">(допускающий значение NULL) itemId :String</span><span class="sxs-lookup"><span data-stu-id="1dab6-311">(nullable) itemId :String</span></span>

<span data-ttu-id="1dab6-p118">Получает идентификатор элемента веб-служб Exchange для текущего элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="1dab6-p118">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="1dab6-314">Идентификатор, возвращаемый свойством `itemId`, совпадает с идентификатором элемента веб-служб Exchange.</span><span class="sxs-lookup"><span data-stu-id="1dab6-314">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="1dab6-315">Свойство  `itemId` не совпадает с идентификатором записи Outlook или идентификатором, используемым интерфейсом  API REST Outlook.</span><span class="sxs-lookup"><span data-stu-id="1dab6-315">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="1dab6-316">Прежде чем осуществлять вызовы интерфейса REST API с помощью этого значения, его следует преобразовать с помощью `Office.context.mailbox.convertToRestId`, который доступен согласно набору требований 1.3.</span><span class="sxs-lookup"><span data-stu-id="1dab6-316">Before making REST API calls using this value, it should be converted using `Office.context.mailbox.convertToRestId`, which is available starting in requirement set 1.3.</span></span> <span data-ttu-id="1dab6-317">Для получения дополнительных сведений см. [Использование REST API Outlook из надстройки Outlook](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="1dab6-317">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

##### <a name="type"></a><span data-ttu-id="1dab6-318">Тип:</span><span class="sxs-lookup"><span data-stu-id="1dab6-318">Type:</span></span>

*   <span data-ttu-id="1dab6-319">String</span><span class="sxs-lookup"><span data-stu-id="1dab6-319">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="1dab6-320">Требования</span><span class="sxs-lookup"><span data-stu-id="1dab6-320">Requirements</span></span>

|<span data-ttu-id="1dab6-321">Требование</span><span class="sxs-lookup"><span data-stu-id="1dab6-321">Requirement</span></span>| <span data-ttu-id="1dab6-322">Значение</span><span class="sxs-lookup"><span data-stu-id="1dab6-322">Value</span></span>|
|---|---|
|[<span data-ttu-id="1dab6-323">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="1dab6-323">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1dab6-324">1.0</span><span class="sxs-lookup"><span data-stu-id="1dab6-324">1.0</span></span>|
|[<span data-ttu-id="1dab6-325">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="1dab6-325">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1dab6-326">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1dab6-326">ReadItem</span></span>|
|[<span data-ttu-id="1dab6-327">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="1dab6-327">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1dab6-328">Чтение</span><span class="sxs-lookup"><span data-stu-id="1dab6-328">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1dab6-329">Пример</span><span class="sxs-lookup"><span data-stu-id="1dab6-329">Example</span></span>

<span data-ttu-id="1dab6-p120">Указанный ниже код проверяет наличие идентификатора элемента. Если свойство `itemId` возвращает значение `null` или `undefined`, элемент будет сохранен в хранилище, а из асинхронного результата будет получен идентификатор элемента.</span><span class="sxs-lookup"><span data-stu-id="1dab6-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```JavaScript
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result){
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook12officemailboxenumsitemtype"></a><span data-ttu-id="1dab6-332">itemType:[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_2/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="1dab6-332">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_2/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="1dab6-333">Получает тип элемента, который представляет экземпляр.</span><span class="sxs-lookup"><span data-stu-id="1dab6-333">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="1dab6-334">Свойство `itemType` возвращает одно из значений перечисления `ItemType`, которое указывает, является ли экземпляр объекта `item` сообщением или встречей.</span><span class="sxs-lookup"><span data-stu-id="1dab6-334">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="1dab6-335">Тип:</span><span class="sxs-lookup"><span data-stu-id="1dab6-335">Type:</span></span>

*   [<span data-ttu-id="1dab6-336">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="1dab6-336">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook_1_2/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="1dab6-337">Требования</span><span class="sxs-lookup"><span data-stu-id="1dab6-337">Requirements</span></span>

|<span data-ttu-id="1dab6-338">Требование</span><span class="sxs-lookup"><span data-stu-id="1dab6-338">Requirement</span></span>| <span data-ttu-id="1dab6-339">Значение</span><span class="sxs-lookup"><span data-stu-id="1dab6-339">Value</span></span>|
|---|---|
|[<span data-ttu-id="1dab6-340">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="1dab6-340">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1dab6-341">1.0</span><span class="sxs-lookup"><span data-stu-id="1dab6-341">1.0</span></span>|
|[<span data-ttu-id="1dab6-342">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="1dab6-342">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1dab6-343">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1dab6-343">ReadItem</span></span>|
|[<span data-ttu-id="1dab6-344">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="1dab6-344">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1dab6-345">Compose (создание) или read (чтение)</span><span class="sxs-lookup"><span data-stu-id="1dab6-345">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="1dab6-346">Пример</span><span class="sxs-lookup"><span data-stu-id="1dab6-346">Example</span></span>

```JavaScript
if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Message)
  // do something
else
  // do something else
```

####  <a name="location-stringlocationjavascriptapioutlook12officelocation"></a><span data-ttu-id="1dab6-347">location :String|[Location](/javascript/api/outlook_1_2/office.location)</span><span class="sxs-lookup"><span data-stu-id="1dab6-347">location :String|[Location](/javascript/api/outlook_1_2/office.location)</span></span>

<span data-ttu-id="1dab6-348">Получает или задает место встречи.</span><span class="sxs-lookup"><span data-stu-id="1dab6-348">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="1dab6-349">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="1dab6-349">Read mode</span></span>

<span data-ttu-id="1dab6-350">Свойство `location` возвращает строку, содержащую сведения о месте встречи.</span><span class="sxs-lookup"><span data-stu-id="1dab6-350">The `location` property returns a string that contains the location of the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="1dab6-351">Режим создания</span><span class="sxs-lookup"><span data-stu-id="1dab6-351">Compose mode</span></span>

<span data-ttu-id="1dab6-352">Свойство `location` возвращает объект `Location`, предоставляющий методы, которые используются для получения и задания места встречи.</span><span class="sxs-lookup"><span data-stu-id="1dab6-352">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="1dab6-353">Тип:</span><span class="sxs-lookup"><span data-stu-id="1dab6-353">Type:</span></span>

*   <span data-ttu-id="1dab6-354">String | [Location](/javascript/api/outlook_1_2/office.location)</span><span class="sxs-lookup"><span data-stu-id="1dab6-354">String | [Location](/javascript/api/outlook_1_2/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="1dab6-355">Требования</span><span class="sxs-lookup"><span data-stu-id="1dab6-355">Requirements</span></span>

|<span data-ttu-id="1dab6-356">Требование</span><span class="sxs-lookup"><span data-stu-id="1dab6-356">Requirement</span></span>| <span data-ttu-id="1dab6-357">Значение</span><span class="sxs-lookup"><span data-stu-id="1dab6-357">Value</span></span>|
|---|---|
|[<span data-ttu-id="1dab6-358">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="1dab6-358">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1dab6-359">1.0</span><span class="sxs-lookup"><span data-stu-id="1dab6-359">1.0</span></span>|
|[<span data-ttu-id="1dab6-360">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="1dab6-360">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1dab6-361">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1dab6-361">ReadItem</span></span>|
|[<span data-ttu-id="1dab6-362">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="1dab6-362">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1dab6-363">Compose (создание) или read (чтение)</span><span class="sxs-lookup"><span data-stu-id="1dab6-363">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="1dab6-364">Пример</span><span class="sxs-lookup"><span data-stu-id="1dab6-364">Example</span></span>

```JavaScript
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

#### <a name="normalizedsubject-string"></a><span data-ttu-id="1dab6-365">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="1dab6-365">normalizedSubject :String</span></span>

<span data-ttu-id="1dab6-p121">Получает тему элемента со всеми удаленными префиксами (включая `RE:` и `FWD:`). Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="1dab6-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="1dab6-p122">Свойство normalizedSubject получает тему элемента со стандартными префиксами (такими как `RE:` и `FW:`), добавляемыми почтовыми программами. Для получения темы элемента с неизмененными префиксами используйте свойство [`subject`](#subject-stringsubjectjavascriptapioutlook12officesubject).</span><span class="sxs-lookup"><span data-stu-id="1dab6-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubjectjavascriptapioutlook12officesubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="1dab6-370">Тип:</span><span class="sxs-lookup"><span data-stu-id="1dab6-370">Type:</span></span>

*   <span data-ttu-id="1dab6-371">String</span><span class="sxs-lookup"><span data-stu-id="1dab6-371">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="1dab6-372">Требования</span><span class="sxs-lookup"><span data-stu-id="1dab6-372">Requirements</span></span>

|<span data-ttu-id="1dab6-373">Требование</span><span class="sxs-lookup"><span data-stu-id="1dab6-373">Requirement</span></span>| <span data-ttu-id="1dab6-374">Значение</span><span class="sxs-lookup"><span data-stu-id="1dab6-374">Value</span></span>|
|---|---|
|[<span data-ttu-id="1dab6-375">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="1dab6-375">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1dab6-376">1.0</span><span class="sxs-lookup"><span data-stu-id="1dab6-376">1.0</span></span>|
|[<span data-ttu-id="1dab6-377">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="1dab6-377">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1dab6-378">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1dab6-378">ReadItem</span></span>|
|[<span data-ttu-id="1dab6-379">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="1dab6-379">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1dab6-380">Чтение</span><span class="sxs-lookup"><span data-stu-id="1dab6-380">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1dab6-381">Пример</span><span class="sxs-lookup"><span data-stu-id="1dab6-381">Example</span></span>

```JavaScript
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
```

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook12officeemailaddressdetailsrecipientsjavascriptapioutlook12officerecipients"></a><span data-ttu-id="1dab6-382">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="1dab6-382">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

<span data-ttu-id="1dab6-383">Предоставляет доступ к необязательным участникам события.</span><span class="sxs-lookup"><span data-stu-id="1dab6-383">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="1dab6-384">Уровень доступа и тип объекта зависит от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="1dab6-384">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="1dab6-385">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="1dab6-385">Read mode</span></span>

<span data-ttu-id="1dab6-386">Свойство `optionalAttendees` возвращает массив, содержащий объект `EmailAddressDetails` для каждого необязательного участника собрания.</span><span class="sxs-lookup"><span data-stu-id="1dab6-386">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="1dab6-387">Режим создания</span><span class="sxs-lookup"><span data-stu-id="1dab6-387">Compose mode</span></span>

<span data-ttu-id="1dab6-388">Свойство `optionalAttendees` возвращает объект `Recipients`, который предоставляет методы для получения и обновления необязательных участников собрания.</span><span class="sxs-lookup"><span data-stu-id="1dab6-388">The `optionalAttendees` property returns a `Recipients` object that provides methods to get and set the optional attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="1dab6-389">Тип:</span><span class="sxs-lookup"><span data-stu-id="1dab6-389">Type:</span></span>

*   <span data-ttu-id="1dab6-390">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="1dab6-390">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="1dab6-391">Требования</span><span class="sxs-lookup"><span data-stu-id="1dab6-391">Requirements</span></span>

|<span data-ttu-id="1dab6-392">Требование</span><span class="sxs-lookup"><span data-stu-id="1dab6-392">Requirement</span></span>| <span data-ttu-id="1dab6-393">Значение</span><span class="sxs-lookup"><span data-stu-id="1dab6-393">Value</span></span>|
|---|---|
|[<span data-ttu-id="1dab6-394">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="1dab6-394">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1dab6-395">1.0</span><span class="sxs-lookup"><span data-stu-id="1dab6-395">1.0</span></span>|
|[<span data-ttu-id="1dab6-396">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="1dab6-396">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1dab6-397">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1dab6-397">ReadItem</span></span>|
|[<span data-ttu-id="1dab6-398">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="1dab6-398">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1dab6-399">Compose (создание) или read (чтение)</span><span class="sxs-lookup"><span data-stu-id="1dab6-399">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="1dab6-400">Пример</span><span class="sxs-lookup"><span data-stu-id="1dab6-400">Example</span></span>

```JavaScript
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails"></a><span data-ttu-id="1dab6-401">organizer :[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="1dab6-401">organizer :[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)</span></span>

<span data-ttu-id="1dab6-p124">Получает электронный адрес организатора указанного собрания. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="1dab6-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="1dab6-404">Тип:</span><span class="sxs-lookup"><span data-stu-id="1dab6-404">Type:</span></span>

*   [<span data-ttu-id="1dab6-405">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="1dab6-405">EmailAddressDetails</span></span>](/javascript/api/outlook_1_2/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="1dab6-406">Требования</span><span class="sxs-lookup"><span data-stu-id="1dab6-406">Requirements</span></span>

|<span data-ttu-id="1dab6-407">Требование</span><span class="sxs-lookup"><span data-stu-id="1dab6-407">Requirement</span></span>| <span data-ttu-id="1dab6-408">Значение</span><span class="sxs-lookup"><span data-stu-id="1dab6-408">Value</span></span>|
|---|---|
|[<span data-ttu-id="1dab6-409">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="1dab6-409">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1dab6-410">1.0</span><span class="sxs-lookup"><span data-stu-id="1dab6-410">1.0</span></span>|
|[<span data-ttu-id="1dab6-411">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="1dab6-411">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1dab6-412">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1dab6-412">ReadItem</span></span>|
|[<span data-ttu-id="1dab6-413">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="1dab6-413">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1dab6-414">Чтение</span><span class="sxs-lookup"><span data-stu-id="1dab6-414">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1dab6-415">Пример</span><span class="sxs-lookup"><span data-stu-id="1dab6-415">Example</span></span>

```JavaScript
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
```

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook12officeemailaddressdetailsrecipientsjavascriptapioutlook12officerecipients"></a><span data-ttu-id="1dab6-416">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="1dab6-416">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

<span data-ttu-id="1dab6-417">Предоставляет доступ к обязательным участникам события.</span><span class="sxs-lookup"><span data-stu-id="1dab6-417">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="1dab6-418">Уровень доступа и тип объекта зависит от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="1dab6-418">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="1dab6-419">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="1dab6-419">Read mode</span></span>

<span data-ttu-id="1dab6-420">Свойство `requiredAttendees` возвращает массив, содержащий объект `EmailAddressDetails` для каждого обязательного участника собрания.</span><span class="sxs-lookup"><span data-stu-id="1dab6-420">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="1dab6-421">Режим создания</span><span class="sxs-lookup"><span data-stu-id="1dab6-421">Compose mode</span></span>

<span data-ttu-id="1dab6-422">Свойство `requiredAttendees` возвращает объект `Recipients`, который предоставляет методы для получения и обновления обязательных участников собрания.</span><span class="sxs-lookup"><span data-stu-id="1dab6-422">The `requiredAttendees` property returns a `Recipients` object that provides methods to get and set the required attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="1dab6-423">Тип:</span><span class="sxs-lookup"><span data-stu-id="1dab6-423">Type:</span></span>

*   <span data-ttu-id="1dab6-424">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="1dab6-424">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="1dab6-425">Требования</span><span class="sxs-lookup"><span data-stu-id="1dab6-425">Requirements</span></span>

|<span data-ttu-id="1dab6-426">Требование</span><span class="sxs-lookup"><span data-stu-id="1dab6-426">Requirement</span></span>| <span data-ttu-id="1dab6-427">Значение</span><span class="sxs-lookup"><span data-stu-id="1dab6-427">Value</span></span>|
|---|---|
|[<span data-ttu-id="1dab6-428">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="1dab6-428">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1dab6-429">1.0</span><span class="sxs-lookup"><span data-stu-id="1dab6-429">1.0</span></span>|
|[<span data-ttu-id="1dab6-430">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="1dab6-430">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1dab6-431">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1dab6-431">ReadItem</span></span>|
|[<span data-ttu-id="1dab6-432">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="1dab6-432">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1dab6-433">Compose (создание) или read (чтение)</span><span class="sxs-lookup"><span data-stu-id="1dab6-433">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="1dab6-434">Пример</span><span class="sxs-lookup"><span data-stu-id="1dab6-434">Example</span></span>

```JavaScript
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
}
```

#### <a name="sender-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails"></a><span data-ttu-id="1dab6-435">sender :[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="1dab6-435">sender :[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)</span></span>

<span data-ttu-id="1dab6-p126">Получает электронный адрес отправителя электронного сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="1dab6-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="1dab6-p127">Свойства [`from`](#from-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails) и `sender` представляют одно лицо, если сообщение не отправлено представителем. В противном случае свойство `from` представляет лицо, делегировавшее полномочия, а свойство sender — представителя.</span><span class="sxs-lookup"><span data-stu-id="1dab6-p127">The [`from`](#from-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="1dab6-440">Свойство `recipientType` объекта `EmailAddressDetails` в свойстве `sender` — `undefined`.</span><span class="sxs-lookup"><span data-stu-id="1dab6-440">Note: The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="1dab6-441">Тип:</span><span class="sxs-lookup"><span data-stu-id="1dab6-441">Type:</span></span>

*   [<span data-ttu-id="1dab6-442">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="1dab6-442">EmailAddressDetails</span></span>](/javascript/api/outlook_1_2/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="1dab6-443">Требования</span><span class="sxs-lookup"><span data-stu-id="1dab6-443">Requirements</span></span>

|<span data-ttu-id="1dab6-444">Требование</span><span class="sxs-lookup"><span data-stu-id="1dab6-444">Requirement</span></span>| <span data-ttu-id="1dab6-445">Значение</span><span class="sxs-lookup"><span data-stu-id="1dab6-445">Value</span></span>|
|---|---|
|[<span data-ttu-id="1dab6-446">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="1dab6-446">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1dab6-447">1.0</span><span class="sxs-lookup"><span data-stu-id="1dab6-447">1.0</span></span>|
|[<span data-ttu-id="1dab6-448">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="1dab6-448">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1dab6-449">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1dab6-449">ReadItem</span></span>|
|[<span data-ttu-id="1dab6-450">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="1dab6-450">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1dab6-451">Чтение</span><span class="sxs-lookup"><span data-stu-id="1dab6-451">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1dab6-452">Пример</span><span class="sxs-lookup"><span data-stu-id="1dab6-452">Example</span></span>

```JavaScript
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
```

####  <a name="start-datetimejavascriptapioutlook12officetime"></a><span data-ttu-id="1dab6-453">start :Date|[Time](/javascript/api/outlook_1_2/office.time)</span><span class="sxs-lookup"><span data-stu-id="1dab6-453">start :Date|[Time](/javascript/api/outlook_1_2/office.time)</span></span>

<span data-ttu-id="1dab6-454">Получает или задает дату и время начала встречи.</span><span class="sxs-lookup"><span data-stu-id="1dab6-454">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="1dab6-p128">Свойство `start` представлено в виде значения даты и времени в формате UTC. Это значение можно преобразовать в местные значения даты и времени клиента с помощью метода [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook12officelocalclienttime).</span><span class="sxs-lookup"><span data-stu-id="1dab6-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook12officelocalclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="1dab6-457">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="1dab6-457">Read mode</span></span>

<span data-ttu-id="1dab6-458">Свойство `start` возвращает объект `Date`.</span><span class="sxs-lookup"><span data-stu-id="1dab6-458">The `start` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="1dab6-459">Режим создания</span><span class="sxs-lookup"><span data-stu-id="1dab6-459">Compose mode</span></span>

<span data-ttu-id="1dab6-460">Свойство `start` возвращает объект `Time`.</span><span class="sxs-lookup"><span data-stu-id="1dab6-460">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="1dab6-461">Если вы задаете время начала с помощью метода [`Time.setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-), необходимо использовать метод [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) для преобразования местного времени на клиенте в формат UTC для сервера.</span><span class="sxs-lookup"><span data-stu-id="1dab6-461">When you use the [`Time.setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="1dab6-462">Тип:</span><span class="sxs-lookup"><span data-stu-id="1dab6-462">Type:</span></span>

*   <span data-ttu-id="1dab6-463">Date | [Time](/javascript/api/outlook_1_2/office.time)</span><span class="sxs-lookup"><span data-stu-id="1dab6-463">Date | [Time](/javascript/api/outlook_1_2/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="1dab6-464">Требования</span><span class="sxs-lookup"><span data-stu-id="1dab6-464">Requirements</span></span>

|<span data-ttu-id="1dab6-465">Требование</span><span class="sxs-lookup"><span data-stu-id="1dab6-465">Requirement</span></span>| <span data-ttu-id="1dab6-466">Значение</span><span class="sxs-lookup"><span data-stu-id="1dab6-466">Value</span></span>|
|---|---|
|[<span data-ttu-id="1dab6-467">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="1dab6-467">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1dab6-468">1.0</span><span class="sxs-lookup"><span data-stu-id="1dab6-468">1.0</span></span>|
|[<span data-ttu-id="1dab6-469">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="1dab6-469">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1dab6-470">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1dab6-470">ReadItem</span></span>|
|[<span data-ttu-id="1dab6-471">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="1dab6-471">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1dab6-472">Compose (создание) или read (чтение)</span><span class="sxs-lookup"><span data-stu-id="1dab6-472">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="1dab6-473">Пример</span><span class="sxs-lookup"><span data-stu-id="1dab6-473">Example</span></span>

<span data-ttu-id="1dab6-474">В примере ниже с помощью метода [`setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) объекта `Time` задается время начала встречи в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="1dab6-474">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

####  <a name="subject-stringsubjectjavascriptapioutlook12officesubject"></a><span data-ttu-id="1dab6-475">subject :String|[Subject](/javascript/api/outlook_1_2/office.subject)</span><span class="sxs-lookup"><span data-stu-id="1dab6-475">subject :String|[Subject](/javascript/api/outlook_1_2/office.subject)</span></span>

<span data-ttu-id="1dab6-476">Получает или задает описание, которое отображается в поле темы элемента.</span><span class="sxs-lookup"><span data-stu-id="1dab6-476">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="1dab6-477">Свойство `subject` получает или задает всю тему элемента для отправки с почтового сервера.</span><span class="sxs-lookup"><span data-stu-id="1dab6-477">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="1dab6-478">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="1dab6-478">Read mode</span></span>

<span data-ttu-id="1dab6-p129">Свойство `subject` возвращает строку. С помощью свойства [`normalizedSubject`](#normalizedsubject-string) можно получить тему без начальных префиксов, таких как `RE:` и `FW:`.</span><span class="sxs-lookup"><span data-stu-id="1dab6-p129">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```
var subject = Office.context.mailbox.item.subject;
```

##### <a name="compose-mode"></a><span data-ttu-id="1dab6-481">Режим создания</span><span class="sxs-lookup"><span data-stu-id="1dab6-481">Compose mode</span></span>

<span data-ttu-id="1dab6-482">Свойство `subject` возвращает объект `Subject`, который предоставляет методы для получения и задания темы.</span><span class="sxs-lookup"><span data-stu-id="1dab6-482">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```JavaScript
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="1dab6-483">Тип:</span><span class="sxs-lookup"><span data-stu-id="1dab6-483">Type:</span></span>

*   <span data-ttu-id="1dab6-484">String | [Subject](/javascript/api/outlook_1_2/office.subject)</span><span class="sxs-lookup"><span data-stu-id="1dab6-484">String | [Subject](/javascript/api/outlook_1_2/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="1dab6-485">Требования</span><span class="sxs-lookup"><span data-stu-id="1dab6-485">Requirements</span></span>

|<span data-ttu-id="1dab6-486">Требование</span><span class="sxs-lookup"><span data-stu-id="1dab6-486">Requirement</span></span>| <span data-ttu-id="1dab6-487">Значение</span><span class="sxs-lookup"><span data-stu-id="1dab6-487">Value</span></span>|
|---|---|
|[<span data-ttu-id="1dab6-488">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="1dab6-488">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1dab6-489">1.0</span><span class="sxs-lookup"><span data-stu-id="1dab6-489">1.0</span></span>|
|[<span data-ttu-id="1dab6-490">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="1dab6-490">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1dab6-491">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1dab6-491">ReadItem</span></span>|
|[<span data-ttu-id="1dab6-492">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="1dab6-492">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1dab6-493">Compose (создание) или read (чтение)</span><span class="sxs-lookup"><span data-stu-id="1dab6-493">Compose or read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlook12officeemailaddressdetailsrecipientsjavascriptapioutlook12officerecipients"></a><span data-ttu-id="1dab6-494">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="1dab6-494">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

<span data-ttu-id="1dab6-495">Предоставляет доступ получателей к строке **To (Кому)** в сообщении.</span><span class="sxs-lookup"><span data-stu-id="1dab6-495">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="1dab6-496">Уровень доступа и тип объекта зависит от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="1dab6-496">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="1dab6-497">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="1dab6-497">Read mode</span></span>

<span data-ttu-id="1dab6-p131">Свойство `to` возвращает массив, содержащий объект `EmailAddressDetails` для каждого получателя в строке **To (Кому)** сообщения. Коллекция может включать не более 100 элементов.</span><span class="sxs-lookup"><span data-stu-id="1dab6-p131">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="1dab6-500">Режим создания</span><span class="sxs-lookup"><span data-stu-id="1dab6-500">Compose mode</span></span>

<span data-ttu-id="1dab6-501">Свойство `to` возвращает объект `Recipients`, предоставляющий методы для получения или обновления получателей, которые указаны в строке **To (Кому)** сообщения.</span><span class="sxs-lookup"><span data-stu-id="1dab6-501">The to`to` property returns a Recipients`Recipients` object that provides methods for manipulating the recipients on the **To** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="1dab6-502">Тип:</span><span class="sxs-lookup"><span data-stu-id="1dab6-502">Type:</span></span>

*   <span data-ttu-id="1dab6-503">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="1dab6-503">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="1dab6-504">Требования</span><span class="sxs-lookup"><span data-stu-id="1dab6-504">Requirements</span></span>

|<span data-ttu-id="1dab6-505">Требование</span><span class="sxs-lookup"><span data-stu-id="1dab6-505">Requirement</span></span>| <span data-ttu-id="1dab6-506">Значение</span><span class="sxs-lookup"><span data-stu-id="1dab6-506">Value</span></span>|
|---|---|
|[<span data-ttu-id="1dab6-507">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="1dab6-507">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1dab6-508">1.0</span><span class="sxs-lookup"><span data-stu-id="1dab6-508">1.0</span></span>|
|[<span data-ttu-id="1dab6-509">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="1dab6-509">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1dab6-510">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1dab6-510">ReadItem</span></span>|
|[<span data-ttu-id="1dab6-511">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="1dab6-511">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1dab6-512">Compose (создание) или read (чтение)</span><span class="sxs-lookup"><span data-stu-id="1dab6-512">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="1dab6-513">Пример</span><span class="sxs-lookup"><span data-stu-id="1dab6-513">Example</span></span>

```JavaScript
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

### <a name="methods"></a><span data-ttu-id="1dab6-514">Методы</span><span class="sxs-lookup"><span data-stu-id="1dab6-514">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="1dab6-515">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="1dab6-515">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="1dab6-516">Добавляет файл в сообщение или встречу в качестве вложения.</span><span class="sxs-lookup"><span data-stu-id="1dab6-516">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="1dab6-517">Метод `addFileAttachmentAsync` передает файл по указанному универсальному коду ресурса (URI) и вкладывает его в элемент в форме создания.</span><span class="sxs-lookup"><span data-stu-id="1dab6-517">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="1dab6-518">Идентификатор можно последовательно использовать с методом [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback), чтобы удалить вложение, добавленное во время текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="1dab6-518">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1dab6-519">Параметры:</span><span class="sxs-lookup"><span data-stu-id="1dab6-519">Parameters:</span></span>

|<span data-ttu-id="1dab6-520">Имя</span><span class="sxs-lookup"><span data-stu-id="1dab6-520">Name</span></span>| <span data-ttu-id="1dab6-521">Тип</span><span class="sxs-lookup"><span data-stu-id="1dab6-521">Type</span></span>| <span data-ttu-id="1dab6-522">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="1dab6-522">Attributes</span></span>| <span data-ttu-id="1dab6-523">Описание</span><span class="sxs-lookup"><span data-stu-id="1dab6-523">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="1dab6-524">String</span><span class="sxs-lookup"><span data-stu-id="1dab6-524">String</span></span>||<span data-ttu-id="1dab6-p132">Универсальный код ресурса (URI), представляющий расположение файла, который нужно вложить в сообщение или встречу. Максимальная длина — 2048 символов.</span><span class="sxs-lookup"><span data-stu-id="1dab6-p132">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="1dab6-527">String</span><span class="sxs-lookup"><span data-stu-id="1dab6-527">String</span></span>||<span data-ttu-id="1dab6-p133">Имя вложения, которое отображается при передаче вложения. Максимальная длина — 255 символов.</span><span class="sxs-lookup"><span data-stu-id="1dab6-p133">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="1dab6-530">Object</span><span class="sxs-lookup"><span data-stu-id="1dab6-530">Object</span></span>| <span data-ttu-id="1dab6-531">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="1dab6-531">&lt;optional&gt;</span></span>|<span data-ttu-id="1dab6-532">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="1dab6-532">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="1dab6-533">Object</span><span class="sxs-lookup"><span data-stu-id="1dab6-533">Object</span></span>| <span data-ttu-id="1dab6-534">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="1dab6-534">&lt;optional&gt;</span></span>|<span data-ttu-id="1dab6-535">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="1dab6-535">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="1dab6-536">function</span><span class="sxs-lookup"><span data-stu-id="1dab6-536">function</span></span>| <span data-ttu-id="1dab6-537">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="1dab6-537">&lt;optional&gt;</span></span>|<span data-ttu-id="1dab6-538">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="1dab6-538">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="1dab6-539">После успешного выполнения идентификатор вложения будет представлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="1dab6-539">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="1dab6-540">Если добавить вложение не удастся, объект `asyncResult` будет содержать объект `Error` с описанием ошибки.</span><span class="sxs-lookup"><span data-stu-id="1dab6-540">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="1dab6-541">Ошибки</span><span class="sxs-lookup"><span data-stu-id="1dab6-541">Errors</span></span>

| <span data-ttu-id="1dab6-542">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="1dab6-542">Error code</span></span> | <span data-ttu-id="1dab6-543">Описание</span><span class="sxs-lookup"><span data-stu-id="1dab6-543">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="1dab6-544">Вложение превышает максимальный размер.</span><span class="sxs-lookup"><span data-stu-id="1dab6-544">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="1dab6-545">Расширение вложения не поддерживается.</span><span class="sxs-lookup"><span data-stu-id="1dab6-545">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="1dab6-546">Сообщение или встреча содержат слишком много вложений.</span><span class="sxs-lookup"><span data-stu-id="1dab6-546">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="1dab6-547">Требования</span><span class="sxs-lookup"><span data-stu-id="1dab6-547">Requirements</span></span>

|<span data-ttu-id="1dab6-548">Требование</span><span class="sxs-lookup"><span data-stu-id="1dab6-548">Requirement</span></span>| <span data-ttu-id="1dab6-549">Значение</span><span class="sxs-lookup"><span data-stu-id="1dab6-549">Value</span></span>|
|---|---|
|[<span data-ttu-id="1dab6-550">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="1dab6-550">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1dab6-551">1.1</span><span class="sxs-lookup"><span data-stu-id="1dab6-551">1.1</span></span>|
|[<span data-ttu-id="1dab6-552">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="1dab6-552">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1dab6-553">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="1dab6-553">ReadWriteItem</span></span>|
|[<span data-ttu-id="1dab6-554">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="1dab6-554">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1dab6-555">Создание</span><span class="sxs-lookup"><span data-stu-id="1dab6-555">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="1dab6-556">Пример</span><span class="sxs-lookup"><span data-stu-id="1dab6-556">Example</span></span>

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

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="1dab6-557">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="1dab6-557">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="1dab6-558">Добавляет к сообщению элемент Exchange, например сообщение, в виде вложения.</span><span class="sxs-lookup"><span data-stu-id="1dab6-558">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="1dab6-p134">С помощью метода `addItemAttachmentAsync` можно в элемент формы создания вложить элемент с указанным идентификатором Exchange. Если указать метод обратного вызова, то этот метод вызывается с помощью параметра `asyncResult`, который содержит идентификатор вложения или код, указывающий на ошибки, которые произошли при вложении элемента. При необходимости можно использовать параметр `options` для передачи сведений о состоянии в метод обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="1dab6-p134">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="1dab6-562">Идентификатор можно последовательно использовать с методом [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback), чтобы удалить вложение, добавленное во время текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="1dab6-562">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="1dab6-563">Если ваша надстройка Office выполняется в веб-приложении Outlook, метод `addItemAttachmentAsync` обеспечивает вложение элементов в элементы, отличные от редактируемого. Однако это действие не рекомендуется выполнять, так как оно не поддерживается.</span><span class="sxs-lookup"><span data-stu-id="1dab6-563">If your Office add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1dab6-564">Параметры:</span><span class="sxs-lookup"><span data-stu-id="1dab6-564">Parameters:</span></span>

|<span data-ttu-id="1dab6-565">Имя</span><span class="sxs-lookup"><span data-stu-id="1dab6-565">Name</span></span>| <span data-ttu-id="1dab6-566">Тип</span><span class="sxs-lookup"><span data-stu-id="1dab6-566">Type</span></span>| <span data-ttu-id="1dab6-567">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="1dab6-567">Attributes</span></span>| <span data-ttu-id="1dab6-568">Описание</span><span class="sxs-lookup"><span data-stu-id="1dab6-568">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="1dab6-569">String</span><span class="sxs-lookup"><span data-stu-id="1dab6-569">String</span></span>||<span data-ttu-id="1dab6-p135">Идентификатор Exchange для вкладываемого элемента. Максимальная длина — 100 символов.</span><span class="sxs-lookup"><span data-stu-id="1dab6-p135">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="1dab6-572">String</span><span class="sxs-lookup"><span data-stu-id="1dab6-572">String</span></span>||<span data-ttu-id="1dab6-p136">Тема вкладываемого элемента. Максимальная длина — 255 символов.</span><span class="sxs-lookup"><span data-stu-id="1dab6-p136">The sujbect of the item to be attached. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="1dab6-575">Object</span><span class="sxs-lookup"><span data-stu-id="1dab6-575">Object</span></span>| <span data-ttu-id="1dab6-576">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="1dab6-576">&lt;optional&gt;</span></span>|<span data-ttu-id="1dab6-577">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="1dab6-577">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="1dab6-578">Object</span><span class="sxs-lookup"><span data-stu-id="1dab6-578">Object</span></span>| <span data-ttu-id="1dab6-579">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="1dab6-579">&lt;optional&gt;</span></span>|<span data-ttu-id="1dab6-580">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="1dab6-580">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="1dab6-581">function</span><span class="sxs-lookup"><span data-stu-id="1dab6-581">function</span></span>| <span data-ttu-id="1dab6-582">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="1dab6-582">&lt;optional&gt;</span></span>|<span data-ttu-id="1dab6-583">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="1dab6-583">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="1dab6-584">После успешного выполнения идентификатор вложения будет представлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="1dab6-584">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="1dab6-585">Если добавить вложение не удастся, объект `asyncResult` будет содержать объект `Error` с описанием ошибки.</span><span class="sxs-lookup"><span data-stu-id="1dab6-585">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="1dab6-586">Ошибки</span><span class="sxs-lookup"><span data-stu-id="1dab6-586">Errors</span></span>

| <span data-ttu-id="1dab6-587">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="1dab6-587">Error code</span></span> | <span data-ttu-id="1dab6-588">Описание</span><span class="sxs-lookup"><span data-stu-id="1dab6-588">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="1dab6-589">Сообщение или встреча содержат слишком много вложений.</span><span class="sxs-lookup"><span data-stu-id="1dab6-589">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="1dab6-590">Требования</span><span class="sxs-lookup"><span data-stu-id="1dab6-590">Requirements</span></span>

|<span data-ttu-id="1dab6-591">Требование</span><span class="sxs-lookup"><span data-stu-id="1dab6-591">Requirement</span></span>| <span data-ttu-id="1dab6-592">Значение</span><span class="sxs-lookup"><span data-stu-id="1dab6-592">Value</span></span>|
|---|---|
|[<span data-ttu-id="1dab6-593">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="1dab6-593">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1dab6-594">1.1</span><span class="sxs-lookup"><span data-stu-id="1dab6-594">1.1</span></span>|
|[<span data-ttu-id="1dab6-595">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="1dab6-595">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1dab6-596">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="1dab6-596">ReadWriteItem</span></span>|
|[<span data-ttu-id="1dab6-597">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="1dab6-597">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1dab6-598">Создание</span><span class="sxs-lookup"><span data-stu-id="1dab6-598">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="1dab6-599">Пример</span><span class="sxs-lookup"><span data-stu-id="1dab6-599">Example</span></span>

<span data-ttu-id="1dab6-600">В следующем примере существующий элемент Outlook добавляется в виде вложения с именем `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="1dab6-600">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

#### <a name="displayreplyallformformdata"></a><span data-ttu-id="1dab6-601">displayReplyAllForm(formData)</span><span class="sxs-lookup"><span data-stu-id="1dab6-601">displayReplyAllForm(formData)</span></span>

<span data-ttu-id="1dab6-602">Отображает форму ответа, включающую отправителя и всех получателей выбранного сообщения или организатора и всех участников выбранной встречи.</span><span class="sxs-lookup"><span data-stu-id="1dab6-602">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="1dab6-603">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="1dab6-603">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="1dab6-604">В веб-приложении Outlook форма ответа отображается в виде всплывающей формы в представлении с 3 либо 1 или 2 колонками.</span><span class="sxs-lookup"><span data-stu-id="1dab6-604">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="1dab6-605">Если любой строковый параметр превышает указанные для него ограничения, `displayReplyAllForm` возвращает исключение.</span><span class="sxs-lookup"><span data-stu-id="1dab6-605">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="1dab6-p137">Если в параметре `formData.attachments` указаны вложения, Outlook и веб-приложение Outlook пытаются скачать их и вложить в форму ответа. Если какие-либо вложения добавить не удается, в форме отображается сообщение об ошибке. Если сообщения об ошибках не предусмотрены, то они не отображаются.</span><span class="sxs-lookup"><span data-stu-id="1dab6-p137">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1dab6-609">Параметры:</span><span class="sxs-lookup"><span data-stu-id="1dab6-609">Parameters:</span></span>

|<span data-ttu-id="1dab6-610">Имя</span><span class="sxs-lookup"><span data-stu-id="1dab6-610">Name</span></span>| <span data-ttu-id="1dab6-611">Тип</span><span class="sxs-lookup"><span data-stu-id="1dab6-611">Type</span></span>| <span data-ttu-id="1dab6-612">Описание</span><span class="sxs-lookup"><span data-stu-id="1dab6-612">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="1dab6-613">String | Object</span><span class="sxs-lookup"><span data-stu-id="1dab6-613">String &#124; Object</span></span>| |<span data-ttu-id="1dab6-p138">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="1dab6-p138">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="1dab6-616">**ИЛИ**</span><span class="sxs-lookup"><span data-stu-id="1dab6-616">**OR**</span></span><br/><span data-ttu-id="1dab6-p139">Объект, который содержит текст или данные вложения и функцию обратного вызова. Ниже представлено определение этого объекта.</span><span class="sxs-lookup"><span data-stu-id="1dab6-p139">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="1dab6-619">String</span><span class="sxs-lookup"><span data-stu-id="1dab6-619">String</span></span> | <span data-ttu-id="1dab6-620">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="1dab6-620">&lt;optional&gt;</span></span> | <span data-ttu-id="1dab6-p140">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="1dab6-p140">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="1dab6-623">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="1dab6-623">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="1dab6-624">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="1dab6-624">&lt;optional&gt;</span></span> | <span data-ttu-id="1dab6-625">Массив объектов JSON, представляющих собой вложенные файлы или элементы.</span><span class="sxs-lookup"><span data-stu-id="1dab6-625">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="1dab6-626">String</span><span class="sxs-lookup"><span data-stu-id="1dab6-626">String</span></span> | | <span data-ttu-id="1dab6-p141">Указывает тип вложения. Допустимые значения: `file` для вложенного файла и `item` для вложенного элемента.</span><span class="sxs-lookup"><span data-stu-id="1dab6-p141">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="1dab6-629">String</span><span class="sxs-lookup"><span data-stu-id="1dab6-629">String</span></span> | | <span data-ttu-id="1dab6-630">Строка, содержащая имя вложения, длиной до 255 символов.</span><span class="sxs-lookup"><span data-stu-id="1dab6-630">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="1dab6-631">String</span><span class="sxs-lookup"><span data-stu-id="1dab6-631">String</span></span> | | <span data-ttu-id="1dab6-p142">Используется, только если свойству `type` задано значение `file`. URI расположения файла.</span><span class="sxs-lookup"><span data-stu-id="1dab6-p142">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="1dab6-634">String</span><span class="sxs-lookup"><span data-stu-id="1dab6-634">String</span></span> | | <span data-ttu-id="1dab6-p143">Используется, только если свойству `type` задано значение `item`. Идентификатор вложения EWS. Это строка длиной до 100 символов.</span><span class="sxs-lookup"><span data-stu-id="1dab6-p143">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="1dab6-638">function</span><span class="sxs-lookup"><span data-stu-id="1dab6-638">function</span></span> | <span data-ttu-id="1dab6-639">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="1dab6-639">&lt;optional&gt;</span></span> | <span data-ttu-id="1dab6-640">По завершении работы метода функция, переданная параметру `callback`, вызывается с помощью одного параметра `asyncResult`, представляющего собой объект [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="1dab6-640">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="1dab6-641">Требования</span><span class="sxs-lookup"><span data-stu-id="1dab6-641">Requirements</span></span>

|<span data-ttu-id="1dab6-642">Требование</span><span class="sxs-lookup"><span data-stu-id="1dab6-642">Requirement</span></span>| <span data-ttu-id="1dab6-643">Значение</span><span class="sxs-lookup"><span data-stu-id="1dab6-643">Value</span></span>|
|---|---|
|[<span data-ttu-id="1dab6-644">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="1dab6-644">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1dab6-645">1.0</span><span class="sxs-lookup"><span data-stu-id="1dab6-645">1.0</span></span>|
|[<span data-ttu-id="1dab6-646">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="1dab6-646">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1dab6-647">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1dab6-647">ReadItem</span></span>|
|[<span data-ttu-id="1dab6-648">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="1dab6-648">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1dab6-649">Чтение</span><span class="sxs-lookup"><span data-stu-id="1dab6-649">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="1dab6-650">Примеры</span><span class="sxs-lookup"><span data-stu-id="1dab6-650">Examples</span></span>

<span data-ttu-id="1dab6-651">Приведенный ниже код передает строку в функцию `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="1dab6-651">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="1dab6-652">Ответ с пустым текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="1dab6-652">Reply with an empty body.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="1dab6-653">Ответ только с текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="1dab6-653">Reply with just a body.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="1dab6-654">Ответ с текстом сообщения и вложенным файлом.</span><span class="sxs-lookup"><span data-stu-id="1dab6-654">Reply with a body and a file attachment.</span></span>

```JavaScript
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

<span data-ttu-id="1dab6-655">Ответ с текстом сообщения и вложенным элементом.</span><span class="sxs-lookup"><span data-stu-id="1dab6-655">Reply with a body and an item attachment.</span></span>

```JavaScript
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

<span data-ttu-id="1dab6-656">Ответ с текстом сообщения, вложенным файлом, вложенным элементом и обратным вызовом.</span><span class="sxs-lookup"><span data-stu-id="1dab6-656">Reply with a body, file attachment, item attachment, and a callback.</span></span>

```JavaScript
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

#### <a name="displayreplyformformdata"></a><span data-ttu-id="1dab6-657">displayReplyForm(formData)</span><span class="sxs-lookup"><span data-stu-id="1dab6-657">displayReplyForm(formData)</span></span>

<span data-ttu-id="1dab6-658">Отображает форму ответа, включающую только отправителя выбранного сообщения или организатора выбранной встречи.</span><span class="sxs-lookup"><span data-stu-id="1dab6-658">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="1dab6-659">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="1dab6-659">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="1dab6-660">В веб-приложении Outlook форма ответа отображается в виде всплывающей формы в представлении с 3 либо 1 или 2 колонками.</span><span class="sxs-lookup"><span data-stu-id="1dab6-660">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="1dab6-661">Если любой строковый параметр превышает указанные для него ограничения, `displayReplyForm` возвращает исключение.</span><span class="sxs-lookup"><span data-stu-id="1dab6-661">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="1dab6-p144">Если в параметре `formData.attachments` указаны вложения, Outlook и веб-приложение Outlook пытаются скачать их и вложить в форму ответа. Если какие-либо вложения добавить не удается, в форме отображается сообщение об ошибке. Если сообщения об ошибках не предусмотрены, то они не отображаются.</span><span class="sxs-lookup"><span data-stu-id="1dab6-p144">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1dab6-665">Параметры:</span><span class="sxs-lookup"><span data-stu-id="1dab6-665">Parameters:</span></span>

|<span data-ttu-id="1dab6-666">Имя</span><span class="sxs-lookup"><span data-stu-id="1dab6-666">Name</span></span>| <span data-ttu-id="1dab6-667">Тип</span><span class="sxs-lookup"><span data-stu-id="1dab6-667">Type</span></span>| <span data-ttu-id="1dab6-668">Описание</span><span class="sxs-lookup"><span data-stu-id="1dab6-668">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="1dab6-669">String | Object</span><span class="sxs-lookup"><span data-stu-id="1dab6-669">String &#124; Object</span></span>| | <span data-ttu-id="1dab6-p145">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="1dab6-p145">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="1dab6-672">**ИЛИ**</span><span class="sxs-lookup"><span data-stu-id="1dab6-672">**OR**</span></span><br/><span data-ttu-id="1dab6-p146">Объект, который содержит текст или данные вложения и функцию обратного вызова. Ниже представлено определение этого объекта.</span><span class="sxs-lookup"><span data-stu-id="1dab6-p146">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="1dab6-675">String</span><span class="sxs-lookup"><span data-stu-id="1dab6-675">String</span></span> | <span data-ttu-id="1dab6-676">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="1dab6-676">&lt;optional&gt;</span></span> | <span data-ttu-id="1dab6-p147">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="1dab6-p147">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="1dab6-679">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="1dab6-679">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="1dab6-680">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="1dab6-680">&lt;optional&gt;</span></span> | <span data-ttu-id="1dab6-681">Массив объектов JSON, представляющих собой вложенные файлы или элементы.</span><span class="sxs-lookup"><span data-stu-id="1dab6-681">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="1dab6-682">String</span><span class="sxs-lookup"><span data-stu-id="1dab6-682">String</span></span> | | <span data-ttu-id="1dab6-p148">Указывает тип вложения. Допустимые значения: `file` для вложенного файла и `item` для вложенного элемента.</span><span class="sxs-lookup"><span data-stu-id="1dab6-p148">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="1dab6-685">String</span><span class="sxs-lookup"><span data-stu-id="1dab6-685">String</span></span> | | <span data-ttu-id="1dab6-686">Строка, содержащая имя вложения, длиной до 255 символов.</span><span class="sxs-lookup"><span data-stu-id="1dab6-686">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="1dab6-687">String</span><span class="sxs-lookup"><span data-stu-id="1dab6-687">String</span></span> | | <span data-ttu-id="1dab6-p149">Используется, только если свойству `type` задано значение `file`. URI расположения файла.</span><span class="sxs-lookup"><span data-stu-id="1dab6-p149">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="1dab6-690">String</span><span class="sxs-lookup"><span data-stu-id="1dab6-690">String</span></span> | | <span data-ttu-id="1dab6-p150">Используется, только если свойству `type` задано значение `item`. Идентификатор вложения EWS. Это строка длиной до 100 символов.</span><span class="sxs-lookup"><span data-stu-id="1dab6-p150">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="1dab6-694">function</span><span class="sxs-lookup"><span data-stu-id="1dab6-694">function</span></span> | <span data-ttu-id="1dab6-695">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="1dab6-695">&lt;optional&gt;</span></span> | <span data-ttu-id="1dab6-696">По завершении работы метода функция, переданная параметру `callback`, вызывается с помощью одного параметра `asyncResult`, представляющего собой объект [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="1dab6-696">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="1dab6-697">Требования</span><span class="sxs-lookup"><span data-stu-id="1dab6-697">Requirements</span></span>

|<span data-ttu-id="1dab6-698">Требование</span><span class="sxs-lookup"><span data-stu-id="1dab6-698">Requirement</span></span>| <span data-ttu-id="1dab6-699">Значение</span><span class="sxs-lookup"><span data-stu-id="1dab6-699">Value</span></span>|
|---|---|
|[<span data-ttu-id="1dab6-700">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="1dab6-700">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1dab6-701">1.0</span><span class="sxs-lookup"><span data-stu-id="1dab6-701">1.0</span></span>|
|[<span data-ttu-id="1dab6-702">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="1dab6-702">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1dab6-703">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1dab6-703">ReadItem</span></span>|
|[<span data-ttu-id="1dab6-704">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="1dab6-704">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1dab6-705">Чтение</span><span class="sxs-lookup"><span data-stu-id="1dab6-705">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="1dab6-706">Примеры</span><span class="sxs-lookup"><span data-stu-id="1dab6-706">Examples</span></span>

<span data-ttu-id="1dab6-707">Приведенный ниже код передает строку в функцию `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="1dab6-707">The following code passes a string to the `displayReplyForm` function.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="1dab6-708">Ответ с пустым текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="1dab6-708">Reply with an empty body.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="1dab6-709">Ответ только с текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="1dab6-709">Reply with just a body.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="1dab6-710">Ответ с текстом сообщения и вложенным файлом.</span><span class="sxs-lookup"><span data-stu-id="1dab6-710">Reply with a body and a file attachment.</span></span>

```JavaScript
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

<span data-ttu-id="1dab6-711">Ответ с текстом сообщения и вложенным элементом.</span><span class="sxs-lookup"><span data-stu-id="1dab6-711">Reply with a body and an item attachment.</span></span>

```JavaScript
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

<span data-ttu-id="1dab6-712">Ответ с текстом сообщения, вложенным файлом, вложенным элементом и обратным вызовом.</span><span class="sxs-lookup"><span data-stu-id="1dab6-712">Reply with a body, file attachment, item attachment, and a callback.</span></span>

```JavaScript
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

#### <a name="getentities--entitiesjavascriptapioutlook12officeentities"></a><span data-ttu-id="1dab6-713">getEntities() → {[Entities](/javascript/api/outlook_1_2/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="1dab6-713">getEntities() → {[Entities](/javascript/api/outlook_1_2/office.entities)}</span></span>

<span data-ttu-id="1dab6-714">Получает сущности, обнаруженные в выбранном элементе.</span><span class="sxs-lookup"><span data-stu-id="1dab6-714">Gets the entities found in the selected item.</span></span>

> [!NOTE]
> <span data-ttu-id="1dab6-715">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="1dab6-715">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="1dab6-716">Требования</span><span class="sxs-lookup"><span data-stu-id="1dab6-716">Requirements</span></span>

|<span data-ttu-id="1dab6-717">Требование</span><span class="sxs-lookup"><span data-stu-id="1dab6-717">Requirement</span></span>| <span data-ttu-id="1dab6-718">Значение</span><span class="sxs-lookup"><span data-stu-id="1dab6-718">Value</span></span>|
|---|---|
|[<span data-ttu-id="1dab6-719">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="1dab6-719">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1dab6-720">1.0</span><span class="sxs-lookup"><span data-stu-id="1dab6-720">1.0</span></span>|
|[<span data-ttu-id="1dab6-721">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="1dab6-721">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1dab6-722">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1dab6-722">ReadItem</span></span>|
|[<span data-ttu-id="1dab6-723">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="1dab6-723">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1dab6-724">Чтение</span><span class="sxs-lookup"><span data-stu-id="1dab6-724">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="1dab6-725">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="1dab6-725">Returns:</span></span>

<span data-ttu-id="1dab6-726">Тип: [Entities](/javascript/api/outlook_1_2/office.entities)</span><span class="sxs-lookup"><span data-stu-id="1dab6-726">Type: [Entities](/javascript/api/outlook_1_2/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="1dab6-727">Пример</span><span class="sxs-lookup"><span data-stu-id="1dab6-727">Example</span></span>

<span data-ttu-id="1dab6-728">Ниже приведен пример получения доступа к сущностям контактов в тексте текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="1dab6-728">The following example accesses the contacts entities on the current item.</span></span>

```
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook12officecontactmeetingsuggestionjavascriptapioutlook12officemeetingsuggestionphonenumberjavascriptapioutlook12officephonenumbertasksuggestionjavascriptapioutlook12officetasksuggestion"></a><span data-ttu-id="1dab6-729">getEntitiesByType(entityType) → (допускающий значение NULL) {Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="1dab6-729">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))>}</span></span>

<span data-ttu-id="1dab6-730">Получает массив всех сущностей указанного типа, обнаруженных в тексте выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="1dab6-730">Gets an array of all the entities of the specified entity type found in the selected item.</span></span>

> [!NOTE]
> <span data-ttu-id="1dab6-731">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="1dab6-731">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1dab6-732">Параметры:</span><span class="sxs-lookup"><span data-stu-id="1dab6-732">Parameters:</span></span>

|<span data-ttu-id="1dab6-733">Имя</span><span class="sxs-lookup"><span data-stu-id="1dab6-733">Name</span></span>| <span data-ttu-id="1dab6-734">Тип</span><span class="sxs-lookup"><span data-stu-id="1dab6-734">Type</span></span>| <span data-ttu-id="1dab6-735">Описание</span><span class="sxs-lookup"><span data-stu-id="1dab6-735">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="1dab6-736">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="1dab6-736">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook_1_2/office.mailboxenums.entitytype)|<span data-ttu-id="1dab6-737">Одно из значений перечисления EntityType.</span><span class="sxs-lookup"><span data-stu-id="1dab6-737">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1dab6-738">Требования</span><span class="sxs-lookup"><span data-stu-id="1dab6-738">Requirements</span></span>

|<span data-ttu-id="1dab6-739">Требование</span><span class="sxs-lookup"><span data-stu-id="1dab6-739">Requirement</span></span>| <span data-ttu-id="1dab6-740">Значение</span><span class="sxs-lookup"><span data-stu-id="1dab6-740">Value</span></span>|
|---|---|
|[<span data-ttu-id="1dab6-741">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="1dab6-741">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1dab6-742">1.0</span><span class="sxs-lookup"><span data-stu-id="1dab6-742">1.0</span></span>|
|[<span data-ttu-id="1dab6-743">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="1dab6-743">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1dab6-744">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="1dab6-744">Restricted</span></span>|
|[<span data-ttu-id="1dab6-745">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="1dab6-745">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1dab6-746">Чтение</span><span class="sxs-lookup"><span data-stu-id="1dab6-746">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="1dab6-747">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="1dab6-747">Returns:</span></span>

<span data-ttu-id="1dab6-748">Если значение, переданное в `entityType`, не является допустимым членом перечисления `EntityType`, метод возвращает значение NULL.</span><span class="sxs-lookup"><span data-stu-id="1dab6-748">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="1dab6-749">Если в тексте элемента отсутствуют сущности указанного типа, метод возвращает пустой массив.</span><span class="sxs-lookup"><span data-stu-id="1dab6-749">If no entities of the specified type are present on the item, the method returns an empty array.</span></span> <span data-ttu-id="1dab6-750">В противном случае тип объектов в возвращаемом массиве зависит от типа сущности, запрошенной в параметре `entityType`.</span><span class="sxs-lookup"><span data-stu-id="1dab6-750">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="1dab6-751">Хотя минимальный уровень разрешений для использования этого метода — **Restricted**, для некоторых типов сущностей требуется доступ на уровне **ReadItem**, как указано в приведенной ниже таблице.</span><span class="sxs-lookup"><span data-stu-id="1dab6-751">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="1dab6-752">Значение параметра `entityType`</span><span class="sxs-lookup"><span data-stu-id="1dab6-752">Value of `entityType`</span></span> | <span data-ttu-id="1dab6-753">Тип объектов в возвращаемом массиве</span><span class="sxs-lookup"><span data-stu-id="1dab6-753">Type of objects in returned array</span></span> | <span data-ttu-id="1dab6-754">Необходимый уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="1dab6-754">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="1dab6-755">String</span><span class="sxs-lookup"><span data-stu-id="1dab6-755">String</span></span> | <span data-ttu-id="1dab6-756">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="1dab6-756">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="1dab6-757">Contact</span><span class="sxs-lookup"><span data-stu-id="1dab6-757">Contact</span></span> | <span data-ttu-id="1dab6-758">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="1dab6-758">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="1dab6-759">String</span><span class="sxs-lookup"><span data-stu-id="1dab6-759">String</span></span> | <span data-ttu-id="1dab6-760">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="1dab6-760">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="1dab6-761">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="1dab6-761">MeetingSuggestion</span></span> | <span data-ttu-id="1dab6-762">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="1dab6-762">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="1dab6-763">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="1dab6-763">PhoneNumber</span></span> | <span data-ttu-id="1dab6-764">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="1dab6-764">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="1dab6-765">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="1dab6-765">TaskSuggestion</span></span> | <span data-ttu-id="1dab6-766">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="1dab6-766">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="1dab6-767">String</span><span class="sxs-lookup"><span data-stu-id="1dab6-767">String</span></span> | <span data-ttu-id="1dab6-768">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="1dab6-768">**Restricted**</span></span> |

<span data-ttu-id="1dab6-769">Тип: Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="1dab6-769">Type: Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="1dab6-770">Пример</span><span class="sxs-lookup"><span data-stu-id="1dab6-770">Example</span></span>

<span data-ttu-id="1dab6-771">В примере ниже показано, как получить доступ к массиву строк, которые представляют собой почтовые адреса в тексте текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="1dab6-771">The following example shows how to access an array of strings that represent postal addresses in the subject or body of the current item.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook12officecontactmeetingsuggestionjavascriptapioutlook12officemeetingsuggestionphonenumberjavascriptapioutlook12officephonenumbertasksuggestionjavascriptapioutlook12officetasksuggestion"></a><span data-ttu-id="1dab6-772">getFilteredEntitiesByName(name) → (допускающий значение NULL) {Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="1dab6-772">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))>}</span></span>

<span data-ttu-id="1dab6-773">Возвращает известные сущности в выбранном элементе, которые проходят через именованный фильтр, определяемый в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="1dab6-773">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="1dab6-774">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="1dab6-774">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="1dab6-775">Метод `getFilteredEntitiesByName` возвращает сущности, соответствующие регулярному выражению, которое определяется в элементе правила [ItemHasKnownEntity](/javascript/office/manifest/rule#itemhasknownentity-rule) в XML-файле манифеста, с использованием указанного значения элемента `FilterName`.</span><span class="sxs-lookup"><span data-stu-id="1dab6-775">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/javascript/office/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1dab6-776">Параметры:</span><span class="sxs-lookup"><span data-stu-id="1dab6-776">Parameters:</span></span>

|<span data-ttu-id="1dab6-777">Имя</span><span class="sxs-lookup"><span data-stu-id="1dab6-777">Name</span></span>| <span data-ttu-id="1dab6-778">Тип</span><span class="sxs-lookup"><span data-stu-id="1dab6-778">Type</span></span>| <span data-ttu-id="1dab6-779">Описание</span><span class="sxs-lookup"><span data-stu-id="1dab6-779">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="1dab6-780">String</span><span class="sxs-lookup"><span data-stu-id="1dab6-780">String</span></span>|<span data-ttu-id="1dab6-781">Имя элемента правила `ItemHasKnownEntity`, определяющее соответствующий фильтр.</span><span class="sxs-lookup"><span data-stu-id="1dab6-781">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1dab6-782">Требования</span><span class="sxs-lookup"><span data-stu-id="1dab6-782">Requirements</span></span>

|<span data-ttu-id="1dab6-783">Требование</span><span class="sxs-lookup"><span data-stu-id="1dab6-783">Requirement</span></span>| <span data-ttu-id="1dab6-784">Значение</span><span class="sxs-lookup"><span data-stu-id="1dab6-784">Value</span></span>|
|---|---|
|[<span data-ttu-id="1dab6-785">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="1dab6-785">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1dab6-786">1.0</span><span class="sxs-lookup"><span data-stu-id="1dab6-786">1.0</span></span>|
|[<span data-ttu-id="1dab6-787">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="1dab6-787">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1dab6-788">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1dab6-788">ReadItem</span></span>|
|[<span data-ttu-id="1dab6-789">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="1dab6-789">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1dab6-790">Чтение</span><span class="sxs-lookup"><span data-stu-id="1dab6-790">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="1dab6-791">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="1dab6-791">Returns:</span></span>

<span data-ttu-id="1dab6-p152">Если в манифесте нет элемента `ItemHasKnownEntity` со значением `FilterName`, соответствующим параметру `name`, метод возвращает `null`. Если параметр `name` соответствует элементу `ItemHasKnownEntity` в манифесте, но при этом в текущем элементе нет соответствующих сущностей, метод возвращает пустой массив.</span><span class="sxs-lookup"><span data-stu-id="1dab6-p152">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="1dab6-794">Тип: Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="1dab6-794">Type: Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))></span></span>

#### <a name="getregexmatches--object"></a><span data-ttu-id="1dab6-795">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="1dab6-795">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="1dab6-796">Возвращает строковые значения в выбранном элементе, которые соответствуют регулярным выражениям, определенным в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="1dab6-796">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="1dab6-797">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="1dab6-797">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="1dab6-p153">Метод `getRegExMatches` возвращает строки, соответствующие регулярному выражению, которое определяется в каждом элементе правила `ItemHasRegularExpressionMatch` или `ItemHasKnownEntity` в XML-файле манифеста. Для правила `ItemHasRegularExpressionMatch` соответствующую строку должно содержать свойство элемента, указанного этим правилом. Простой тип `PropertyName` определяет поддерживаемые свойства.</span><span class="sxs-lookup"><span data-stu-id="1dab6-p153">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="1dab6-801">Например, рассмотрим манифест надстройки, который содержит указанный ниже элемент `Rule`.</span><span class="sxs-lookup"><span data-stu-id="1dab6-801">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```JavaScript
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="1dab6-802">Объект, возвращаемый методом `getRegExMatches`, будет содержать два свойства: `fruits` и `veggies`.</span><span class="sxs-lookup"><span data-stu-id="1dab6-802">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```JavaScript
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

> [!NOTE]
> <span data-ttu-id="1dab6-p154">Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты.</span><span class="sxs-lookup"><span data-stu-id="1dab6-p154">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="requirements"></a><span data-ttu-id="1dab6-805">Требования</span><span class="sxs-lookup"><span data-stu-id="1dab6-805">Requirements</span></span>

|<span data-ttu-id="1dab6-806">Требование</span><span class="sxs-lookup"><span data-stu-id="1dab6-806">Requirement</span></span>| <span data-ttu-id="1dab6-807">Значение</span><span class="sxs-lookup"><span data-stu-id="1dab6-807">Value</span></span>|
|---|---|
|[<span data-ttu-id="1dab6-808">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="1dab6-808">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1dab6-809">1.0</span><span class="sxs-lookup"><span data-stu-id="1dab6-809">1.0</span></span>|
|[<span data-ttu-id="1dab6-810">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="1dab6-810">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1dab6-811">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1dab6-811">ReadItem</span></span>|
|[<span data-ttu-id="1dab6-812">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="1dab6-812">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1dab6-813">Чтение</span><span class="sxs-lookup"><span data-stu-id="1dab6-813">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="1dab6-814">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="1dab6-814">Returns:</span></span>

<span data-ttu-id="1dab6-p155">Объект, содержащий массив строк, которые соответствуют регулярным выражениям, определяемым в XML-файле манифеста. Имя каждого массива равно соответствующему значению атрибута `RegExName` правила сопоставления `ItemHasRegularExpressionMatch` или атрибута `FilterName` правила сопоставления `ItemHasKnownEntity`.</span><span class="sxs-lookup"><span data-stu-id="1dab6-p155">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="1dab6-817">

<dt>Тип</dt>

</span><span class="sxs-lookup"><span data-stu-id="1dab6-817">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="1dab6-818">Object</span><span class="sxs-lookup"><span data-stu-id="1dab6-818">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="1dab6-819">Пример</span><span class="sxs-lookup"><span data-stu-id="1dab6-819">Example</span></span>

<span data-ttu-id="1dab6-820">В примере ниже показано, как получить доступ к массиву совпадений для элементов <rule> регулярного выражения `fruits` и `veggies`, которые указаны в манифесте.</rule></span><span class="sxs-lookup"><span data-stu-id="1dab6-820">The following example shows how to access the array of matches for the regular expression <rule>elements `fruits` and `veggies`, which are specified in the manifest.</rule></span></span>

```JavaScript
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veges = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="1dab6-821">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="1dab6-821">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="1dab6-822">Возвращает строковые значения в выбранном элементе, которые соответствуют именованному регулярному выражению, определенному в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="1dab6-822">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="1dab6-823">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="1dab6-823">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="1dab6-824">Метод `getRegExMatchesByName` возвращает строки, соответствующие регулярному выражению, которое определяется в элементе правила `ItemHasRegularExpressionMatch` в XML-файле манифеста, с использованием указанного значения элемента `RegExName`.</span><span class="sxs-lookup"><span data-stu-id="1dab6-824">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="1dab6-p156">Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты.</span><span class="sxs-lookup"><span data-stu-id="1dab6-p156">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1dab6-827">Параметры:</span><span class="sxs-lookup"><span data-stu-id="1dab6-827">Parameters:</span></span>

|<span data-ttu-id="1dab6-828">Имя</span><span class="sxs-lookup"><span data-stu-id="1dab6-828">Name</span></span>| <span data-ttu-id="1dab6-829">Тип</span><span class="sxs-lookup"><span data-stu-id="1dab6-829">Type</span></span>| <span data-ttu-id="1dab6-830">Описание</span><span class="sxs-lookup"><span data-stu-id="1dab6-830">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="1dab6-831">String</span><span class="sxs-lookup"><span data-stu-id="1dab6-831">String</span></span>|<span data-ttu-id="1dab6-832">Имя элемента правила `ItemHasRegularExpressionMatch`, определяющее соответствующий фильтр.</span><span class="sxs-lookup"><span data-stu-id="1dab6-832">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1dab6-833">Требования</span><span class="sxs-lookup"><span data-stu-id="1dab6-833">Requirements</span></span>

|<span data-ttu-id="1dab6-834">Требование</span><span class="sxs-lookup"><span data-stu-id="1dab6-834">Requirement</span></span>| <span data-ttu-id="1dab6-835">Значение</span><span class="sxs-lookup"><span data-stu-id="1dab6-835">Value</span></span>|
|---|---|
|[<span data-ttu-id="1dab6-836">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="1dab6-836">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1dab6-837">1.0</span><span class="sxs-lookup"><span data-stu-id="1dab6-837">1.0</span></span>|
|[<span data-ttu-id="1dab6-838">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="1dab6-838">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1dab6-839">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1dab6-839">ReadItem</span></span>|
|[<span data-ttu-id="1dab6-840">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="1dab6-840">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1dab6-841">Чтение</span><span class="sxs-lookup"><span data-stu-id="1dab6-841">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="1dab6-842">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="1dab6-842">Returns:</span></span>

<span data-ttu-id="1dab6-843">Массив строк, соответствующих регулярному выражению, определяемому в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="1dab6-843">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="1dab6-844">

<dt>Тип</dt>

</span><span class="sxs-lookup"><span data-stu-id="1dab6-844">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="1dab6-845">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="1dab6-845">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="1dab6-846">Пример</span><span class="sxs-lookup"><span data-stu-id="1dab6-846">Example</span></span>

```JavaScript
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="1dab6-847">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="1dab6-847">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="1dab6-848">Асинхронно возвращает данные, выбранные в теме или тексте сообщения.</span><span class="sxs-lookup"><span data-stu-id="1dab6-848">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="1dab6-p157">Если выделенный фрагмент отсутствует, но курсор находится в тексте или теме, метод возвращает значение NULL для выбранных данных. Если выбраны не текст и не тема, метод возвращает ошибку `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="1dab6-p157">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1dab6-851">Параметры:</span><span class="sxs-lookup"><span data-stu-id="1dab6-851">Parameters:</span></span>

|<span data-ttu-id="1dab6-852">Имя</span><span class="sxs-lookup"><span data-stu-id="1dab6-852">Name</span></span>| <span data-ttu-id="1dab6-853">Тип</span><span class="sxs-lookup"><span data-stu-id="1dab6-853">Type</span></span>| <span data-ttu-id="1dab6-854">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="1dab6-854">Attributes</span></span>| <span data-ttu-id="1dab6-855">Описание</span><span class="sxs-lookup"><span data-stu-id="1dab6-855">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="1dab6-856">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="1dab6-856">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="1dab6-p158">Запрашивает формат данных. Если задано значение Text, метод возвращает обычный текст как строку, удаляя все имеющиеся HTML-теги. Если задано значение HTML, метод возвращает выделенный текст (обычный текст или HTML).</span><span class="sxs-lookup"><span data-stu-id="1dab6-p158">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="1dab6-860">Object</span><span class="sxs-lookup"><span data-stu-id="1dab6-860">Object</span></span>| <span data-ttu-id="1dab6-861">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="1dab6-861">&lt;optional&gt;</span></span>|<span data-ttu-id="1dab6-862">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="1dab6-862">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="1dab6-863">Object</span><span class="sxs-lookup"><span data-stu-id="1dab6-863">Object</span></span>| <span data-ttu-id="1dab6-864">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="1dab6-864">&lt;optional&gt;</span></span>|<span data-ttu-id="1dab6-865">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="1dab6-865">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="1dab6-866">function</span><span class="sxs-lookup"><span data-stu-id="1dab6-866">function</span></span>||<span data-ttu-id="1dab6-867">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="1dab6-867">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="1dab6-868">Чтобы получить доступ к выбранным данным из метода обратного вызова, вызовите `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="1dab6-868">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="1dab6-869">Для доступа к исходному свойству, на основе которого созданы выбранные данные, вызовите  `asyncResult.value.sourceProperty`, который может иметь значение `body` или `subject`.</span><span class="sxs-lookup"><span data-stu-id="1dab6-869">To access the source property that the selection comes from, call , which will be either  or .|</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1dab6-870">Требования</span><span class="sxs-lookup"><span data-stu-id="1dab6-870">Requirements</span></span>

|<span data-ttu-id="1dab6-871">Требование</span><span class="sxs-lookup"><span data-stu-id="1dab6-871">Requirement</span></span>| <span data-ttu-id="1dab6-872">Значение</span><span class="sxs-lookup"><span data-stu-id="1dab6-872">Value</span></span>|
|---|---|
|[<span data-ttu-id="1dab6-873">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="1dab6-873">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1dab6-874">1.2</span><span class="sxs-lookup"><span data-stu-id="1dab6-874">1.2</span></span>|
|[<span data-ttu-id="1dab6-875">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="1dab6-875">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1dab6-876">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="1dab6-876">ReadWriteItem</span></span>|
|[<span data-ttu-id="1dab6-877">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="1dab6-877">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1dab6-878">Создание</span><span class="sxs-lookup"><span data-stu-id="1dab6-878">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="1dab6-879">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="1dab6-879">Returns:</span></span>

<span data-ttu-id="1dab6-880">Выбранные данные в виде строки с форматом, определенным в параметре `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="1dab6-880">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="1dab6-881">

<dt>Тип</dt>

</span><span class="sxs-lookup"><span data-stu-id="1dab6-881">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="1dab6-882">String</span><span class="sxs-lookup"><span data-stu-id="1dab6-882">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="1dab6-883">Пример</span><span class="sxs-lookup"><span data-stu-id="1dab6-883">Example</span></span>

```JavaScript
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

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="1dab6-884">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="1dab6-884">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="1dab6-885">Асинхронно загружает настраиваемые свойства для надстройки выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="1dab6-885">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="1dab6-p160">Настраиваемые свойства сохраняются в виде пар "ключ-значение" для каждого приложения и каждого элемента. Этот метод возвращает объект `CustomProperties` при обратном вызове, который предоставляет методы для доступа к настраиваемым свойствам, характерным для текущего элемента и текущей надстройки. Настраиваемые свойства не шифруются для элемента, поэтому этот способ хранения не является безопасным.</span><span class="sxs-lookup"><span data-stu-id="1dab6-p160">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1dab6-889">Параметры:</span><span class="sxs-lookup"><span data-stu-id="1dab6-889">Parameters:</span></span>

|<span data-ttu-id="1dab6-890">Имя</span><span class="sxs-lookup"><span data-stu-id="1dab6-890">Name</span></span>| <span data-ttu-id="1dab6-891">Тип</span><span class="sxs-lookup"><span data-stu-id="1dab6-891">Type</span></span>| <span data-ttu-id="1dab6-892">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="1dab6-892">Attributes</span></span>| <span data-ttu-id="1dab6-893">Описание</span><span class="sxs-lookup"><span data-stu-id="1dab6-893">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="1dab6-894">function</span><span class="sxs-lookup"><span data-stu-id="1dab6-894">function</span></span>||<span data-ttu-id="1dab6-895">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="1dab6-895">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="1dab6-896">Настраиваемые свойства предоставляются в виде объекта [`CustomProperties`](/javascript/api/outlook_1_2/office.customproperties) в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="1dab6-896">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook_1_2/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="1dab6-897">Этот объект можно использовать для получения, задания и удаления настраиваемых свойств из элемента и сохранения изменений настраиваемого свойства на сервере.</span><span class="sxs-lookup"><span data-stu-id="1dab6-897">The custom properties are provided as a CustomProperties object in the asyncResult.value property. This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="1dab6-898">Object</span><span class="sxs-lookup"><span data-stu-id="1dab6-898">Object</span></span>| <span data-ttu-id="1dab6-899">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="1dab6-899">&lt;optional&gt;</span></span>|<span data-ttu-id="1dab6-900">Разработчики могут указать любой объект, к которому необходимо получить доступ, в функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="1dab6-900">Developers can provide any object they wish to access in the callback method.</span></span> <span data-ttu-id="1dab6-901">Доступ к этому объекту можно получить с помощью свойства `asyncResult.asyncContext` в функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="1dab6-901">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.|</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1dab6-902">Требования</span><span class="sxs-lookup"><span data-stu-id="1dab6-902">Requirements</span></span>

|<span data-ttu-id="1dab6-903">Требование</span><span class="sxs-lookup"><span data-stu-id="1dab6-903">Requirement</span></span>| <span data-ttu-id="1dab6-904">Значение</span><span class="sxs-lookup"><span data-stu-id="1dab6-904">Value</span></span>|
|---|---|
|[<span data-ttu-id="1dab6-905">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="1dab6-905">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1dab6-906">1.0</span><span class="sxs-lookup"><span data-stu-id="1dab6-906">1.0</span></span>|
|[<span data-ttu-id="1dab6-907">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="1dab6-907">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1dab6-908">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1dab6-908">ReadItem</span></span>|
|[<span data-ttu-id="1dab6-909">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="1dab6-909">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1dab6-910">Compose (создание) или read (чтение)</span><span class="sxs-lookup"><span data-stu-id="1dab6-910">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="1dab6-911">Пример</span><span class="sxs-lookup"><span data-stu-id="1dab6-911">Example</span></span>

<span data-ttu-id="1dab6-p163">Приведенный ниже пример кода показывает, как асинхронно загружать настраиваемые свойства, характерные для текущего элемента, с помощью метода `loadCustomPropertiesAsync`. Этот пример также показывает, как сохранять эти свойства на сервере с помощью метода `CustomProperties.saveAsync`. В этом примере кода, после того как выполнена загрузка настраиваемых свойств, метод `CustomProperties.get` используется для считывания настраиваемого свойства `myProp`, метод `CustomProperties.set` — для записи настраиваемого свойства `otherProp`, а метод `saveAsync` — для сохранения настраиваемых свойств.</span><span class="sxs-lookup"><span data-stu-id="1dab6-p163">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="1dab6-915">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="1dab6-915">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="1dab6-916">Удаляет вложение из сообщения или встречи.</span><span class="sxs-lookup"><span data-stu-id="1dab6-916">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="1dab6-p164">Метод `removeAttachmentAsync` удаляет из элемента вложение с указанным идентификатором. Идентификатор вложения рекомендуется использовать для удаления вложения, только если оно добавлено тем же почтовым приложением в ходе текущего сеанса. В веб-приложении Outlook и веб-приложении Outlook для устройств идентификатор вложения действителен только в рамках одного сеанса. Сеанс завершается, когда пользователь закрывает приложение или начинает создавать элемент во встроенной форме, а затем переходит из формы в отдельное окно.</span><span class="sxs-lookup"><span data-stu-id="1dab6-p164">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1dab6-921">Параметры:</span><span class="sxs-lookup"><span data-stu-id="1dab6-921">Parameters:</span></span>

|<span data-ttu-id="1dab6-922">Имя</span><span class="sxs-lookup"><span data-stu-id="1dab6-922">Name</span></span>| <span data-ttu-id="1dab6-923">Тип</span><span class="sxs-lookup"><span data-stu-id="1dab6-923">Type</span></span>| <span data-ttu-id="1dab6-924">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="1dab6-924">Attributes</span></span>| <span data-ttu-id="1dab6-925">Описание</span><span class="sxs-lookup"><span data-stu-id="1dab6-925">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="1dab6-926">String</span><span class="sxs-lookup"><span data-stu-id="1dab6-926">String</span></span>||<span data-ttu-id="1dab6-p165">Идентификатор удаляемого вложения. Максимальная длина строки — 100 символов.</span><span class="sxs-lookup"><span data-stu-id="1dab6-p165">The identifier of the attachment to remove. The maximum length of the string is 100 characters.</span></span>|
|`options`| <span data-ttu-id="1dab6-929">Object</span><span class="sxs-lookup"><span data-stu-id="1dab6-929">Object</span></span>| <span data-ttu-id="1dab6-930">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="1dab6-930">&lt;optional&gt;</span></span>|<span data-ttu-id="1dab6-931">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="1dab6-931">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="1dab6-932">Object</span><span class="sxs-lookup"><span data-stu-id="1dab6-932">Object</span></span>| <span data-ttu-id="1dab6-933">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="1dab6-933">&lt;optional&gt;</span></span>|<span data-ttu-id="1dab6-934">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="1dab6-934">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="1dab6-935">function</span><span class="sxs-lookup"><span data-stu-id="1dab6-935">function</span></span>| <span data-ttu-id="1dab6-936">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="1dab6-936">&lt;optional&gt;</span></span>|<span data-ttu-id="1dab6-937">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="1dab6-937">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="1dab6-938">Если удалить вложение не удается, свойство `asyncResult.error` содержит код ошибки с указанием ее причины.</span><span class="sxs-lookup"><span data-stu-id="1dab6-938">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="1dab6-939">Ошибки</span><span class="sxs-lookup"><span data-stu-id="1dab6-939">Errors</span></span>

| <span data-ttu-id="1dab6-940">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="1dab6-940">Error code</span></span> | <span data-ttu-id="1dab6-941">Описание</span><span class="sxs-lookup"><span data-stu-id="1dab6-941">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="1dab6-942">Идентификатор вложения не существует.</span><span class="sxs-lookup"><span data-stu-id="1dab6-942">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="1dab6-943">Требования</span><span class="sxs-lookup"><span data-stu-id="1dab6-943">Requirements</span></span>

|<span data-ttu-id="1dab6-944">Требование</span><span class="sxs-lookup"><span data-stu-id="1dab6-944">Requirement</span></span>| <span data-ttu-id="1dab6-945">Значение</span><span class="sxs-lookup"><span data-stu-id="1dab6-945">Value</span></span>|
|---|---|
|[<span data-ttu-id="1dab6-946">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="1dab6-946">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1dab6-947">1.1</span><span class="sxs-lookup"><span data-stu-id="1dab6-947">1.1</span></span>|
|[<span data-ttu-id="1dab6-948">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="1dab6-948">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1dab6-949">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="1dab6-949">ReadWriteItem</span></span>|
|[<span data-ttu-id="1dab6-950">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="1dab6-950">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1dab6-951">Создание</span><span class="sxs-lookup"><span data-stu-id="1dab6-951">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="1dab6-952">Пример</span><span class="sxs-lookup"><span data-stu-id="1dab6-952">Example</span></span>

<span data-ttu-id="1dab6-953">Указанный ниже код удаляет вложение с идентификатором "0".</span><span class="sxs-lookup"><span data-stu-id="1dab6-953">The following code removes an attachment with an identifier of '0'.</span></span>

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

####  <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="1dab6-954">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="1dab6-954">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="1dab6-955">Асинхронно вставляет данные в текст или тему сообщения.</span><span class="sxs-lookup"><span data-stu-id="1dab6-955">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="1dab6-p166">Метод `setSelectedDataAsync` вставляет указанную строку в местоположение курсора в теме или тексте элемента либо, если текст выделен в редакторе, он заменяет выделенный текст. Если курсор находится вне текста или темы элемента, возвращается ошибка. После вставки курсор помещается в конец вставленного содержимого.</span><span class="sxs-lookup"><span data-stu-id="1dab6-p166">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1dab6-959">Параметры:</span><span class="sxs-lookup"><span data-stu-id="1dab6-959">Parameters:</span></span>

|<span data-ttu-id="1dab6-960">Имя</span><span class="sxs-lookup"><span data-stu-id="1dab6-960">Name</span></span>| <span data-ttu-id="1dab6-961">Тип</span><span class="sxs-lookup"><span data-stu-id="1dab6-961">Type</span></span>| <span data-ttu-id="1dab6-962">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="1dab6-962">Attributes</span></span>| <span data-ttu-id="1dab6-963">Описание</span><span class="sxs-lookup"><span data-stu-id="1dab6-963">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="1dab6-964">String</span><span class="sxs-lookup"><span data-stu-id="1dab6-964">String</span></span>||<span data-ttu-id="1dab6-p167">Вставляемые данные. Объем данных не должен превышать 1 000 000 символов. Если передано больше 1 000 000 символов, возвращается исключение `ArgumentOutOfRange`.</span><span class="sxs-lookup"><span data-stu-id="1dab6-p167">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="1dab6-968">Object</span><span class="sxs-lookup"><span data-stu-id="1dab6-968">Object</span></span>| <span data-ttu-id="1dab6-969">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="1dab6-969">&lt;optional&gt;</span></span>|<span data-ttu-id="1dab6-970">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="1dab6-970">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="1dab6-971">Object</span><span class="sxs-lookup"><span data-stu-id="1dab6-971">Object</span></span>| <span data-ttu-id="1dab6-972">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="1dab6-972">&lt;optional&gt;</span></span>|<span data-ttu-id="1dab6-973">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="1dab6-973">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`| [<span data-ttu-id="1dab6-974">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="1dab6-974">Office.CoercionType</span></span>](office.md#coerciontype-string)| <span data-ttu-id="1dab6-975">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="1dab6-975">&lt;optional&gt;</span></span>|<span data-ttu-id="1dab6-p168">Если задано значение `text`, текущий стиль применяется в Outlook и веб-приложении Outlook. Если поле представляет собой редактор HTML, вставляются только текстовые данные, даже если они имеют формат HTML.</span><span class="sxs-lookup"><span data-stu-id="1dab6-p168">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="1dab6-p169">Если `html` и поле поддерживают HTML (тема не поддерживает), в веб-приложении Outlook применяется текущий стиль, а в Outlook — стиль по умолчанию. Если поле является текстовым, возвращается ошибка `InvalidDataFormat`.</span><span class="sxs-lookup"><span data-stu-id="1dab6-p169">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="1dab6-980">Если тип `coercionType` не установлен, результат зависит от поля: если поле имеет формат HTML, то используется HTML; если поле является текстовым, то используется обычный текст.</span><span class="sxs-lookup"><span data-stu-id="1dab6-980">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="1dab6-981">function</span><span class="sxs-lookup"><span data-stu-id="1dab6-981">function</span></span>||<span data-ttu-id="1dab6-982">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="1dab6-982">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="1dab6-983">Требования</span><span class="sxs-lookup"><span data-stu-id="1dab6-983">Requirements</span></span>

|<span data-ttu-id="1dab6-984">Требование</span><span class="sxs-lookup"><span data-stu-id="1dab6-984">Requirement</span></span>| <span data-ttu-id="1dab6-985">Значение</span><span class="sxs-lookup"><span data-stu-id="1dab6-985">Value</span></span>|
|---|---|
|[<span data-ttu-id="1dab6-986">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="1dab6-986">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1dab6-987">1.2</span><span class="sxs-lookup"><span data-stu-id="1dab6-987">1.2</span></span>|
|[<span data-ttu-id="1dab6-988">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="1dab6-988">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1dab6-989">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="1dab6-989">ReadWriteItem</span></span>|
|[<span data-ttu-id="1dab6-990">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="1dab6-990">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1dab6-991">Создание</span><span class="sxs-lookup"><span data-stu-id="1dab6-991">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="1dab6-992">Пример</span><span class="sxs-lookup"><span data-stu-id="1dab6-992">Example</span></span>

```JavaScript
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```