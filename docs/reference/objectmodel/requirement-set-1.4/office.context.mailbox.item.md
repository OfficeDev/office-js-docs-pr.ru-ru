
# <a name="item"></a><span data-ttu-id="d9e46-101">item</span><span class="sxs-lookup"><span data-stu-id="d9e46-101">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="d9e46-102">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="d9e46-102">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="d9e46-p101">Пространство имен `item` используется для доступа к выбранному в данный момент сообщению, приглашению на собрание или описанию встречи. Вы можете определить тип пространства имен `item` с помощью свойства [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook14officemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="d9e46-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook14officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="d9e46-105">Требования</span><span class="sxs-lookup"><span data-stu-id="d9e46-105">Requirements</span></span>

|<span data-ttu-id="d9e46-106">Требование</span><span class="sxs-lookup"><span data-stu-id="d9e46-106">Requirement</span></span>| <span data-ttu-id="d9e46-107">Значение</span><span class="sxs-lookup"><span data-stu-id="d9e46-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="d9e46-108">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="d9e46-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d9e46-109">1.0</span><span class="sxs-lookup"><span data-stu-id="d9e46-109">1.0</span></span>|
|[<span data-ttu-id="d9e46-110">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="d9e46-110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d9e46-111">Ограниченный доступ</span><span class="sxs-lookup"><span data-stu-id="d9e46-111">Restricted</span></span>|
|[<span data-ttu-id="d9e46-112">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="d9e46-112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d9e46-113">Cоздание или чтение</span><span class="sxs-lookup"><span data-stu-id="d9e46-113">Compose or read</span></span>|

### <a name="example"></a><span data-ttu-id="d9e46-114">Пример</span><span class="sxs-lookup"><span data-stu-id="d9e46-114">Example</span></span>

<span data-ttu-id="d9e46-115">В приведенном ниже примере кода JavaScript показано, как получить доступ к свойству `subject` текущего элемента в Outlook.</span><span class="sxs-lookup"><span data-stu-id="d9e46-115">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="d9e46-116">Члены</span><span class="sxs-lookup"><span data-stu-id="d9e46-116">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook14officeattachmentdetails"></a><span data-ttu-id="d9e46-117">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_4/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="d9e46-117">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_4/office.attachmentdetails)></span></span>

<span data-ttu-id="d9e46-p102">Получает массив вложений для элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="d9e46-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="d9e46-p103">Некоторые типы файлов блокируются Outlook из-за потенциальных проблем безопасности и поэтому не возвращаются. Дополнительные сведения см. в статье [Блокированные вложения в Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="d9e46-p103">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned. For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="d9e46-122">Тип:</span><span class="sxs-lookup"><span data-stu-id="d9e46-122">Type:</span></span>

*   <span data-ttu-id="d9e46-123">Array.<[AttachmentDetails](/javascript/api/outlook_1_4/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="d9e46-123">Array.<[AttachmentDetails](/javascript/api/outlook_1_4/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="d9e46-124">Требования</span><span class="sxs-lookup"><span data-stu-id="d9e46-124">Requirements</span></span>

|<span data-ttu-id="d9e46-125">Требование</span><span class="sxs-lookup"><span data-stu-id="d9e46-125">Requirement</span></span>| <span data-ttu-id="d9e46-126">Значение</span><span class="sxs-lookup"><span data-stu-id="d9e46-126">Value</span></span>|
|---|---|
|[<span data-ttu-id="d9e46-127">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="d9e46-127">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d9e46-128">1.0</span><span class="sxs-lookup"><span data-stu-id="d9e46-128">1.0</span></span>|
|[<span data-ttu-id="d9e46-129">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="d9e46-129">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d9e46-130">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d9e46-130">ReadItem</span></span>|
|[<span data-ttu-id="d9e46-131">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="d9e46-131">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d9e46-132">Чтение</span><span class="sxs-lookup"><span data-stu-id="d9e46-132">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d9e46-133">Пример</span><span class="sxs-lookup"><span data-stu-id="d9e46-133">Example</span></span>

<span data-ttu-id="d9e46-134">С помощью приведенного ниже кода можно создать HTML-строку с подробными сведениями обо всех вложениях для текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="d9e46-134">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

####  <a name="bcc-recipientsjavascriptapioutlook14officerecipients"></a><span data-ttu-id="d9e46-135">bcc :[Recipients](/javascript/api/outlook_1_4/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="d9e46-135">bcc :[Recipients](/javascript/api/outlook_1_4/office.recipients)</span></span>

<span data-ttu-id="d9e46-p104">Получает объект, который предоставляет методы для получения или обновления Bcc (скрытой копии) сообщения. Только режим создания.</span><span class="sxs-lookup"><span data-stu-id="d9e46-p104">Gets an object that provides methods to get or update the Bcc (blind carbon copy) line of a message. Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="d9e46-138">Тип:</span><span class="sxs-lookup"><span data-stu-id="d9e46-138">Type:</span></span>

*   [<span data-ttu-id="d9e46-139">Recipients</span><span class="sxs-lookup"><span data-stu-id="d9e46-139">Recipients</span></span>](/javascript/api/outlook_1_4/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="d9e46-140">Требования</span><span class="sxs-lookup"><span data-stu-id="d9e46-140">Requirements</span></span>

|<span data-ttu-id="d9e46-141">Требование</span><span class="sxs-lookup"><span data-stu-id="d9e46-141">Requirement</span></span>| <span data-ttu-id="d9e46-142">Значение</span><span class="sxs-lookup"><span data-stu-id="d9e46-142">Value</span></span>|
|---|---|
|[<span data-ttu-id="d9e46-143">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="d9e46-143">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d9e46-144">1.1</span><span class="sxs-lookup"><span data-stu-id="d9e46-144">1.1</span></span>|
|[<span data-ttu-id="d9e46-145">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="d9e46-145">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d9e46-146">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d9e46-146">ReadItem</span></span>|
|[<span data-ttu-id="d9e46-147">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="d9e46-147">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d9e46-148">Создание</span><span class="sxs-lookup"><span data-stu-id="d9e46-148">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="d9e46-149">Пример</span><span class="sxs-lookup"><span data-stu-id="d9e46-149">Example</span></span>

```
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlook14officebody"></a><span data-ttu-id="d9e46-150">body :[Body](/javascript/api/outlook_1_4/office.body)</span><span class="sxs-lookup"><span data-stu-id="d9e46-150">body :[Body](/javascript/api/outlook_1_4/office.body)</span></span>

<span data-ttu-id="d9e46-151">Получает объект, предоставляющий методы для работы с основным текстом элемента.</span><span class="sxs-lookup"><span data-stu-id="d9e46-151">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="d9e46-152">Тип:</span><span class="sxs-lookup"><span data-stu-id="d9e46-152">Type:</span></span>

*   [<span data-ttu-id="d9e46-153">Body</span><span class="sxs-lookup"><span data-stu-id="d9e46-153">Body</span></span>](/javascript/api/outlook_1_4/office.body)

##### <a name="requirements"></a><span data-ttu-id="d9e46-154">Требования</span><span class="sxs-lookup"><span data-stu-id="d9e46-154">Requirements</span></span>

|<span data-ttu-id="d9e46-155">Требование</span><span class="sxs-lookup"><span data-stu-id="d9e46-155">Requirement</span></span>| <span data-ttu-id="d9e46-156">Значение</span><span class="sxs-lookup"><span data-stu-id="d9e46-156">Value</span></span>|
|---|---|
|[<span data-ttu-id="d9e46-157">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="d9e46-157">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d9e46-158">1.1</span><span class="sxs-lookup"><span data-stu-id="d9e46-158">1.1</span></span>|
|[<span data-ttu-id="d9e46-159">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="d9e46-159">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d9e46-160">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d9e46-160">ReadItem</span></span>|
|[<span data-ttu-id="d9e46-161">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="d9e46-161">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d9e46-162">Cоздание или чтение</span><span class="sxs-lookup"><span data-stu-id="d9e46-162">Compose or read</span></span>|

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlook14officeemailaddressdetailsrecipientsjavascriptapioutlook14officerecipients"></a><span data-ttu-id="d9e46-163">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_4/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="d9e46-163">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_4/office.recipients)</span></span>

<span data-ttu-id="d9e46-p105">Предоставляет доступ к «Cc» (копии) получателей сообщения. Уровень доступа и тип объекта зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="d9e46-p105">Provides access to the Cc (carbon copy) recipients of a message. The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="d9e46-166">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="d9e46-166">Read mode</span></span>

<span data-ttu-id="d9e46-p106">Свойство `cc` возвращает массив, который содержит объект `EmailAddressDetails`, каждому получателю, указанному в строке **Cc (копия)** сообщения. Коллекция может включать не более 100 членов.</span><span class="sxs-lookup"><span data-stu-id="d9e46-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="d9e46-169">Режим создания</span><span class="sxs-lookup"><span data-stu-id="d9e46-169">Compose mode</span></span>

<span data-ttu-id="d9e46-170">Свойство `cc` возвращает объект `Recipients`, предоставляющий методы для получения или обновления получателей, которые указаны в строке **Cc (копия)** сообщения.</span><span class="sxs-lookup"><span data-stu-id="d9e46-170">The `cc` property returns a `Recipients` object that provides methods for manipulating the recipients on the **Cc** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="d9e46-171">Тип:</span><span class="sxs-lookup"><span data-stu-id="d9e46-171">Type:</span></span>

*   <span data-ttu-id="d9e46-172">Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_4/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="d9e46-172">Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_4/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="d9e46-173">Требования</span><span class="sxs-lookup"><span data-stu-id="d9e46-173">Requirements</span></span>

|<span data-ttu-id="d9e46-174">Требование</span><span class="sxs-lookup"><span data-stu-id="d9e46-174">Requirement</span></span>| <span data-ttu-id="d9e46-175">Значение</span><span class="sxs-lookup"><span data-stu-id="d9e46-175">Value</span></span>|
|---|---|
|[<span data-ttu-id="d9e46-176">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="d9e46-176">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d9e46-177">1.0</span><span class="sxs-lookup"><span data-stu-id="d9e46-177">1.0</span></span>|
|[<span data-ttu-id="d9e46-178">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="d9e46-178">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d9e46-179">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d9e46-179">ReadItem</span></span>|
|[<span data-ttu-id="d9e46-180">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="d9e46-180">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d9e46-181">Cоздание или чтение</span><span class="sxs-lookup"><span data-stu-id="d9e46-181">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="d9e46-182">Пример</span><span class="sxs-lookup"><span data-stu-id="d9e46-182">Example</span></span>

```
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="d9e46-183">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="d9e46-183">(nullable) conversationId :String</span></span>

<span data-ttu-id="d9e46-184">Получает идентификатор разговора по электронной почте, содержащего конкретное сообщение.</span><span class="sxs-lookup"><span data-stu-id="d9e46-184">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="d9e46-p107">Вы можете получить целочисленное значение этого свойства, если ваше почтовое приложение активируется в формах просмотра или формах создания ответов. Если пользователь в свою очередь изменит тему ответа, после его отправки идентификатор беседы будет изменен, и полученное ранее значение будет недействительным.</span><span class="sxs-lookup"><span data-stu-id="d9e46-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="d9e46-p108">Это свойство имеет значение NULL для нового элемента в форме создания. Свойство `conversationId` вернет значение, если пользователь задаст тему и сохранит элемент.</span><span class="sxs-lookup"><span data-stu-id="d9e46-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="d9e46-189">Тип:</span><span class="sxs-lookup"><span data-stu-id="d9e46-189">Type:</span></span>

*   <span data-ttu-id="d9e46-190">String</span><span class="sxs-lookup"><span data-stu-id="d9e46-190">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="d9e46-191">Требования</span><span class="sxs-lookup"><span data-stu-id="d9e46-191">Requirements</span></span>

|<span data-ttu-id="d9e46-192">Требование</span><span class="sxs-lookup"><span data-stu-id="d9e46-192">Requirement</span></span>| <span data-ttu-id="d9e46-193">Значение</span><span class="sxs-lookup"><span data-stu-id="d9e46-193">Value</span></span>|
|---|---|
|[<span data-ttu-id="d9e46-194">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="d9e46-194">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d9e46-195">1.0</span><span class="sxs-lookup"><span data-stu-id="d9e46-195">1.0</span></span>|
|[<span data-ttu-id="d9e46-196">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="d9e46-196">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d9e46-197">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d9e46-197">ReadItem</span></span>|
|[<span data-ttu-id="d9e46-198">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="d9e46-198">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d9e46-199">Cоздание или чтение</span><span class="sxs-lookup"><span data-stu-id="d9e46-199">Compose or read</span></span>|

#### <a name="datetimecreated-date"></a><span data-ttu-id="d9e46-200">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="d9e46-200">dateTimeCreated :Date</span></span>

<span data-ttu-id="d9e46-p109">Получает дату и время создания элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="d9e46-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="d9e46-203">Тип:</span><span class="sxs-lookup"><span data-stu-id="d9e46-203">Type:</span></span>

*   <span data-ttu-id="d9e46-204">Date</span><span class="sxs-lookup"><span data-stu-id="d9e46-204">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="d9e46-205">Требования</span><span class="sxs-lookup"><span data-stu-id="d9e46-205">Requirements</span></span>

|<span data-ttu-id="d9e46-206">Требование</span><span class="sxs-lookup"><span data-stu-id="d9e46-206">Requirement</span></span>| <span data-ttu-id="d9e46-207">Значение</span><span class="sxs-lookup"><span data-stu-id="d9e46-207">Value</span></span>|
|---|---|
|[<span data-ttu-id="d9e46-208">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="d9e46-208">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d9e46-209">1.0</span><span class="sxs-lookup"><span data-stu-id="d9e46-209">1.0</span></span>|
|[<span data-ttu-id="d9e46-210">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="d9e46-210">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d9e46-211">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d9e46-211">ReadItem</span></span>|
|[<span data-ttu-id="d9e46-212">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="d9e46-212">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d9e46-213">Чтение</span><span class="sxs-lookup"><span data-stu-id="d9e46-213">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d9e46-214">Пример</span><span class="sxs-lookup"><span data-stu-id="d9e46-214">Example</span></span>

```
var created = Office.context.mailbox.item.dateTimeCreated;
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="d9e46-215">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="d9e46-215">dateTimeModified :Date</span></span>

<span data-ttu-id="d9e46-p110">Получает дату и время последнего изменения элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="d9e46-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="d9e46-218">Этот элемент не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="d9e46-218">Note: This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="d9e46-219">Тип:</span><span class="sxs-lookup"><span data-stu-id="d9e46-219">Type:</span></span>

*   <span data-ttu-id="d9e46-220">Date</span><span class="sxs-lookup"><span data-stu-id="d9e46-220">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="d9e46-221">Требования</span><span class="sxs-lookup"><span data-stu-id="d9e46-221">Requirements</span></span>

|<span data-ttu-id="d9e46-222">Требование</span><span class="sxs-lookup"><span data-stu-id="d9e46-222">Requirement</span></span>| <span data-ttu-id="d9e46-223">Значение</span><span class="sxs-lookup"><span data-stu-id="d9e46-223">Value</span></span>|
|---|---|
|[<span data-ttu-id="d9e46-224">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="d9e46-224">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d9e46-225">1.0</span><span class="sxs-lookup"><span data-stu-id="d9e46-225">1.0</span></span>|
|[<span data-ttu-id="d9e46-226">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="d9e46-226">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d9e46-227">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d9e46-227">ReadItem</span></span>|
|[<span data-ttu-id="d9e46-228">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="d9e46-228">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d9e46-229">Чтение</span><span class="sxs-lookup"><span data-stu-id="d9e46-229">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d9e46-230">Пример</span><span class="sxs-lookup"><span data-stu-id="d9e46-230">Example</span></span>

```
var modified = Office.context.mailbox.item.dateTimeModified;
```

####  <a name="end-datetimejavascriptapioutlook14officetime"></a><span data-ttu-id="d9e46-231">end :Date|[Time](/javascript/api/outlook_1_4/office.time)</span><span class="sxs-lookup"><span data-stu-id="d9e46-231">end :Date|[Time](/javascript/api/outlook_1_4/office.time)</span></span>

<span data-ttu-id="d9e46-232">Получает или задает дату и время окончания встречи.</span><span class="sxs-lookup"><span data-stu-id="d9e46-232">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="d9e46-p111">Свойство `end` представлено в виде значения даты и времени в формате UTC. Преобразовать значение свойства end в местные значения даты и времени клиента можно с помощью метода [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook14officelocalclienttime).</span><span class="sxs-lookup"><span data-stu-id="d9e46-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook14officelocalclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="d9e46-235">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="d9e46-235">Read mode</span></span>

<span data-ttu-id="d9e46-236">Свойство `end` возвращает объект `Date`.</span><span class="sxs-lookup"><span data-stu-id="d9e46-236">The `end` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="d9e46-237">Режим создания</span><span class="sxs-lookup"><span data-stu-id="d9e46-237">Compose mode</span></span>

<span data-ttu-id="d9e46-238">Свойство `end` возвращает объект `Time`.</span><span class="sxs-lookup"><span data-stu-id="d9e46-238">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="d9e46-239">Когда вы используете метод [`Time.setAsync`](/javascript/api/outlook_1_4/office.time#setasync-datetime--options--callback-) для того, чтобы задать время окончания, вы должны использовать метод [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) , чтобы преобразовать местное время на клиенте в формат UTC.</span><span class="sxs-lookup"><span data-stu-id="d9e46-239">When you use the [`Time.setAsync`](/javascript/api/outlook_1_4/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="d9e46-240">Тип:</span><span class="sxs-lookup"><span data-stu-id="d9e46-240">Type:</span></span>

*   <span data-ttu-id="d9e46-241">Date | [Time](/javascript/api/outlook_1_4/office.time)</span><span class="sxs-lookup"><span data-stu-id="d9e46-241">Date | [Time](/javascript/api/outlook_1_4/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="d9e46-242">Требования</span><span class="sxs-lookup"><span data-stu-id="d9e46-242">Requirements</span></span>

|<span data-ttu-id="d9e46-243">Требование</span><span class="sxs-lookup"><span data-stu-id="d9e46-243">Requirement</span></span>| <span data-ttu-id="d9e46-244">Значение</span><span class="sxs-lookup"><span data-stu-id="d9e46-244">Value</span></span>|
|---|---|
|[<span data-ttu-id="d9e46-245">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="d9e46-245">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d9e46-246">1.0</span><span class="sxs-lookup"><span data-stu-id="d9e46-246">1.0</span></span>|
|[<span data-ttu-id="d9e46-247">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="d9e46-247">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d9e46-248">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d9e46-248">ReadItem</span></span>|
|[<span data-ttu-id="d9e46-249">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="d9e46-249">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d9e46-250">Cоздание или чтение</span><span class="sxs-lookup"><span data-stu-id="d9e46-250">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="d9e46-251">Пример</span><span class="sxs-lookup"><span data-stu-id="d9e46-251">Example</span></span>

<span data-ttu-id="d9e46-252">В примере ниже показано, как с помощью метода [`setAsync`](/javascript/api/outlook_1_4/office.time#setasync-datetime--options--callback-) объекта `Time` задать время окончания встречи в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="d9e46-252">The following example sets the end time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_4/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

#### <a name="from-emailaddressdetailsjavascriptapioutlook14officeemailaddressdetails"></a><span data-ttu-id="d9e46-253">from :[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="d9e46-253">from :[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)</span></span>

<span data-ttu-id="d9e46-p112">Получает электронный адрес отправителя сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="d9e46-p112">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="d9e46-p113">Свойства `from` и [`sender`](#sender-emailaddressdetailsjavascriptapioutlook14officeemailaddressdetails) представляют одно лицо, если сообщение не отправлено представителем. В противном случае свойство `from` представляет лицо, делегировавшее полномочия, а свойство sender — представителя.</span><span class="sxs-lookup"><span data-stu-id="d9e46-p113">The `from` and [`sender`](#sender-emailaddressdetailsjavascriptapioutlook14officeemailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="d9e46-258">Свойство `recipientType` объекта `EmailAddressDetails` в свойстве `from` — `undefined`.</span><span class="sxs-lookup"><span data-stu-id="d9e46-258">Note: The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="d9e46-259">Тип:</span><span class="sxs-lookup"><span data-stu-id="d9e46-259">Type:</span></span>

*   [<span data-ttu-id="d9e46-260">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="d9e46-260">EmailAddressDetails</span></span>](/javascript/api/outlook_1_4/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="d9e46-261">Требования</span><span class="sxs-lookup"><span data-stu-id="d9e46-261">Requirements</span></span>

|<span data-ttu-id="d9e46-262">Требование</span><span class="sxs-lookup"><span data-stu-id="d9e46-262">Requirement</span></span>| <span data-ttu-id="d9e46-263">Значение</span><span class="sxs-lookup"><span data-stu-id="d9e46-263">Value</span></span>|
|---|---|
|[<span data-ttu-id="d9e46-264">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="d9e46-264">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d9e46-265">1.0</span><span class="sxs-lookup"><span data-stu-id="d9e46-265">1.0</span></span>|
|[<span data-ttu-id="d9e46-266">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="d9e46-266">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d9e46-267">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d9e46-267">ReadItem</span></span>|
|[<span data-ttu-id="d9e46-268">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="d9e46-268">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d9e46-269">Чтение</span><span class="sxs-lookup"><span data-stu-id="d9e46-269">Read</span></span>|

#### <a name="internetmessageid-string"></a><span data-ttu-id="d9e46-270">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="d9e46-270">internetMessageId :String</span></span>

<span data-ttu-id="d9e46-p114">Получает идентификатор интернет-сообщения для электронного сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="d9e46-p114">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="d9e46-273">Тип:</span><span class="sxs-lookup"><span data-stu-id="d9e46-273">Type:</span></span>

*   <span data-ttu-id="d9e46-274">String</span><span class="sxs-lookup"><span data-stu-id="d9e46-274">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="d9e46-275">Требования</span><span class="sxs-lookup"><span data-stu-id="d9e46-275">Requirements</span></span>

|<span data-ttu-id="d9e46-276">Требование</span><span class="sxs-lookup"><span data-stu-id="d9e46-276">Requirement</span></span>| <span data-ttu-id="d9e46-277">Значение</span><span class="sxs-lookup"><span data-stu-id="d9e46-277">Value</span></span>|
|---|---|
|[<span data-ttu-id="d9e46-278">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="d9e46-278">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d9e46-279">1.0</span><span class="sxs-lookup"><span data-stu-id="d9e46-279">1.0</span></span>|
|[<span data-ttu-id="d9e46-280">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="d9e46-280">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d9e46-281">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d9e46-281">ReadItem</span></span>|
|[<span data-ttu-id="d9e46-282">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="d9e46-282">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d9e46-283">Чтение</span><span class="sxs-lookup"><span data-stu-id="d9e46-283">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d9e46-284">Пример</span><span class="sxs-lookup"><span data-stu-id="d9e46-284">Example</span></span>

```
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="d9e46-285">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="d9e46-285">itemClass :String</span></span>

<span data-ttu-id="d9e46-p115">Получает класс элемента веб-служб Exchange для выбранного элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="d9e46-p115">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="d9e46-p116">Свойство `itemClass` указывает класс сообщения выбранного элемента. Ниже приводятся классы сообщения по умолчанию для элемента сообщения или встречи.</span><span class="sxs-lookup"><span data-stu-id="d9e46-p116">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="d9e46-290">Тип</span><span class="sxs-lookup"><span data-stu-id="d9e46-290">Type</span></span> | <span data-ttu-id="d9e46-291">Описание</span><span class="sxs-lookup"><span data-stu-id="d9e46-291">Description</span></span> | <span data-ttu-id="d9e46-292">item class</span><span class="sxs-lookup"><span data-stu-id="d9e46-292">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="d9e46-293">Элементы встречи</span><span class="sxs-lookup"><span data-stu-id="d9e46-293">Appointment items</span></span> | <span data-ttu-id="d9e46-294">Это элементы календаря для класса элемента `IPM.Appointment` или `IPM.Appointment.Occurence`.</span><span class="sxs-lookup"><span data-stu-id="d9e46-294">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurence` |
| <span data-ttu-id="d9e46-295">Элементы сообщения</span><span class="sxs-lookup"><span data-stu-id="d9e46-295">Message items</span></span> | <span data-ttu-id="d9e46-296">Сюда входят электронные сообщения, для которых по умолчанию задан класс сообщения `IPM.Note`, а также приглашения на собрания, ответы на них и уведомления об их отмене, использующие `IPM.Schedule.Meeting` в качестве базового класса сообщений.</span><span class="sxs-lookup"><span data-stu-id="d9e46-296">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="d9e46-297">Вы можете создавать настраиваемые классы сообщения, расширяющие классы сообщения по умолчанию, например, настраиваемый класс сообщения о встрече `IPM.Appointment.Contoso`.</span><span class="sxs-lookup"><span data-stu-id="d9e46-297">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="d9e46-298">Тип:</span><span class="sxs-lookup"><span data-stu-id="d9e46-298">Type:</span></span>

*   <span data-ttu-id="d9e46-299">String</span><span class="sxs-lookup"><span data-stu-id="d9e46-299">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="d9e46-300">Требования</span><span class="sxs-lookup"><span data-stu-id="d9e46-300">Requirements</span></span>

|<span data-ttu-id="d9e46-301">Требование</span><span class="sxs-lookup"><span data-stu-id="d9e46-301">Requirement</span></span>| <span data-ttu-id="d9e46-302">Значение</span><span class="sxs-lookup"><span data-stu-id="d9e46-302">Value</span></span>|
|---|---|
|[<span data-ttu-id="d9e46-303">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="d9e46-303">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d9e46-304">1.0</span><span class="sxs-lookup"><span data-stu-id="d9e46-304">1.0</span></span>|
|[<span data-ttu-id="d9e46-305">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="d9e46-305">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d9e46-306">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d9e46-306">ReadItem</span></span>|
|[<span data-ttu-id="d9e46-307">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="d9e46-307">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d9e46-308">Чтение</span><span class="sxs-lookup"><span data-stu-id="d9e46-308">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d9e46-309">Пример</span><span class="sxs-lookup"><span data-stu-id="d9e46-309">Example</span></span>

```
var itemClass = Office.context.mailbox.item.itemClass;
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="d9e46-310">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="d9e46-310">(nullable) itemId :String</span></span>

<span data-ttu-id="d9e46-p117">Получает идентификатор элемента веб-служб Exchange для текущего элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="d9e46-p117">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="d9e46-p118">Идентификатор, возвращенный свойством `itemId`, — то же, что идентификатор элемента веб-служб Exchange.  Свойство `itemId` не идентично идентификаторам Outlook, используемым API-Интерфейсом REST Outlook. Прежде чем позволить вызовам API REST использовать это значения, его следует преобразовать с помощью [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string). Дополнительные сведения см. в статье [Использование API REST для Outlook из надстройки Outlook](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="d9e46-p118">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier. The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API. Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string). For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="d9e46-p119">Свойство `itemId` недоступно в режиме создания. Если требуется идентификатор элемента, с помощью метода [`saveAsync`](#saveasyncoptions-callback) можно сохранить элемент в хранилище. При этом в параметре [`AsyncResult.value`](/javascript/api/office/office.asyncresult) функции обратного вызова возвращается идентификатор элемента.</span><span class="sxs-lookup"><span data-stu-id="d9e46-p119">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="d9e46-319">Тип:</span><span class="sxs-lookup"><span data-stu-id="d9e46-319">Type:</span></span>

*   <span data-ttu-id="d9e46-320">String</span><span class="sxs-lookup"><span data-stu-id="d9e46-320">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="d9e46-321">Требования</span><span class="sxs-lookup"><span data-stu-id="d9e46-321">Requirements</span></span>

|<span data-ttu-id="d9e46-322">Требование</span><span class="sxs-lookup"><span data-stu-id="d9e46-322">Requirement</span></span>| <span data-ttu-id="d9e46-323">Значение</span><span class="sxs-lookup"><span data-stu-id="d9e46-323">Value</span></span>|
|---|---|
|[<span data-ttu-id="d9e46-324">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="d9e46-324">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d9e46-325">1.0</span><span class="sxs-lookup"><span data-stu-id="d9e46-325">1.0</span></span>|
|[<span data-ttu-id="d9e46-326">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="d9e46-326">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d9e46-327">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d9e46-327">ReadItem</span></span>|
|[<span data-ttu-id="d9e46-328">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="d9e46-328">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d9e46-329">Чтение</span><span class="sxs-lookup"><span data-stu-id="d9e46-329">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d9e46-330">Пример</span><span class="sxs-lookup"><span data-stu-id="d9e46-330">Example</span></span>

<span data-ttu-id="d9e46-p120">Указанный ниже код проверяет наличие идентификатора элемента. Если свойство `itemId` возвращает значение `null` или `undefined`, элемент будет сохранен в хранилище, а из асинхронного результата будет получен идентификатор элемента.</span><span class="sxs-lookup"><span data-stu-id="d9e46-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result){
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook14officemailboxenumsitemtype"></a><span data-ttu-id="d9e46-333">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_4/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="d9e46-333">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_4/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="d9e46-334">Получает тип элемента, который представляет экземпляр.</span><span class="sxs-lookup"><span data-stu-id="d9e46-334">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="d9e46-335">Свойство `itemType` возвращает одно из значений перечисления `ItemType`, которое указывает, является ли экземпляр объекта `item` сообщением или встречей.</span><span class="sxs-lookup"><span data-stu-id="d9e46-335">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="d9e46-336">Тип:</span><span class="sxs-lookup"><span data-stu-id="d9e46-336">Type:</span></span>

*   [<span data-ttu-id="d9e46-337">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="d9e46-337">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook_1_4/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="d9e46-338">Требования</span><span class="sxs-lookup"><span data-stu-id="d9e46-338">Requirements</span></span>

|<span data-ttu-id="d9e46-339">Требование</span><span class="sxs-lookup"><span data-stu-id="d9e46-339">Requirement</span></span>| <span data-ttu-id="d9e46-340">Значение</span><span class="sxs-lookup"><span data-stu-id="d9e46-340">Value</span></span>|
|---|---|
|[<span data-ttu-id="d9e46-341">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="d9e46-341">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d9e46-342">1.0</span><span class="sxs-lookup"><span data-stu-id="d9e46-342">1.0</span></span>|
|[<span data-ttu-id="d9e46-343">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="d9e46-343">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d9e46-344">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d9e46-344">ReadItem</span></span>|
|[<span data-ttu-id="d9e46-345">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="d9e46-345">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d9e46-346">Cоздание или чтение</span><span class="sxs-lookup"><span data-stu-id="d9e46-346">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="d9e46-347">Пример</span><span class="sxs-lookup"><span data-stu-id="d9e46-347">Example</span></span>

```
if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Message)
  // do something
else
  // do something else
```

####  <a name="location-stringlocationjavascriptapioutlook14officelocation"></a><span data-ttu-id="d9e46-348">location :String|[Location](/javascript/api/outlook_1_4/office.location)</span><span class="sxs-lookup"><span data-stu-id="d9e46-348">location :String|[Location](/javascript/api/outlook_1_4/office.location)</span></span>

<span data-ttu-id="d9e46-349">Получает или задает место встречи.</span><span class="sxs-lookup"><span data-stu-id="d9e46-349">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="d9e46-350">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="d9e46-350">Read mode</span></span>

<span data-ttu-id="d9e46-351">Свойство `location` возвращает строку, содержащую сведения о месте встречи.</span><span class="sxs-lookup"><span data-stu-id="d9e46-351">The `location` property returns a string that contains the location of the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="d9e46-352">Режим создания</span><span class="sxs-lookup"><span data-stu-id="d9e46-352">Compose mode</span></span>

<span data-ttu-id="d9e46-353">Свойство `location` возвращает объект `Location`, предоставляющий методы, которые используются для получения и задания места встречи.</span><span class="sxs-lookup"><span data-stu-id="d9e46-353">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="d9e46-354">Тип:</span><span class="sxs-lookup"><span data-stu-id="d9e46-354">Type:</span></span>

*   <span data-ttu-id="d9e46-355">String | [Location](/javascript/api/outlook_1_4/office.location)</span><span class="sxs-lookup"><span data-stu-id="d9e46-355">String | [Location](/javascript/api/outlook_1_4/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="d9e46-356">Требования</span><span class="sxs-lookup"><span data-stu-id="d9e46-356">Requirements</span></span>

|<span data-ttu-id="d9e46-357">Требование</span><span class="sxs-lookup"><span data-stu-id="d9e46-357">Requirement</span></span>| <span data-ttu-id="d9e46-358">Значение</span><span class="sxs-lookup"><span data-stu-id="d9e46-358">Value</span></span>|
|---|---|
|[<span data-ttu-id="d9e46-359">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="d9e46-359">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d9e46-360">1.0</span><span class="sxs-lookup"><span data-stu-id="d9e46-360">1.0</span></span>|
|[<span data-ttu-id="d9e46-361">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="d9e46-361">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d9e46-362">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d9e46-362">ReadItem</span></span>|
|[<span data-ttu-id="d9e46-363">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="d9e46-363">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d9e46-364">Cоздание или чтение</span><span class="sxs-lookup"><span data-stu-id="d9e46-364">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="d9e46-365">Пример</span><span class="sxs-lookup"><span data-stu-id="d9e46-365">Example</span></span>

```
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

#### <a name="normalizedsubject-string"></a><span data-ttu-id="d9e46-366">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="d9e46-366">normalizedSubject :String</span></span>

<span data-ttu-id="d9e46-p121">Получает тему элемента со всеми удаленными префиксами (включая `RE:` и `FWD:`). Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="d9e46-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="d9e46-p122">Свойство normalizedSubject получает тему элемента со стандартными префиксами (такими как `RE:` и `FW:`), добавляемыми почтовыми программами. Для получения темы элемента с неизмененными префиксами используйте свойство [`subject`](#subject-stringsubjectjavascriptapioutlook14officesubject).</span><span class="sxs-lookup"><span data-stu-id="d9e46-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubjectjavascriptapioutlook14officesubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="d9e46-371">Тип:</span><span class="sxs-lookup"><span data-stu-id="d9e46-371">Type:</span></span>

*   <span data-ttu-id="d9e46-372">String</span><span class="sxs-lookup"><span data-stu-id="d9e46-372">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="d9e46-373">Требования</span><span class="sxs-lookup"><span data-stu-id="d9e46-373">Requirements</span></span>

|<span data-ttu-id="d9e46-374">Требование</span><span class="sxs-lookup"><span data-stu-id="d9e46-374">Requirement</span></span>| <span data-ttu-id="d9e46-375">Значение</span><span class="sxs-lookup"><span data-stu-id="d9e46-375">Value</span></span>|
|---|---|
|[<span data-ttu-id="d9e46-376">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="d9e46-376">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d9e46-377">1.0</span><span class="sxs-lookup"><span data-stu-id="d9e46-377">1.0</span></span>|
|[<span data-ttu-id="d9e46-378">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="d9e46-378">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d9e46-379">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d9e46-379">ReadItem</span></span>|
|[<span data-ttu-id="d9e46-380">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="d9e46-380">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d9e46-381">Чтение</span><span class="sxs-lookup"><span data-stu-id="d9e46-381">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d9e46-382">Пример</span><span class="sxs-lookup"><span data-stu-id="d9e46-382">Example</span></span>

```
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
```

####  <a name="notificationmessages-notificationmessagesjavascriptapioutlook14officenotificationmessages"></a><span data-ttu-id="d9e46-383">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_4/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="d9e46-383">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_4/office.notificationmessages)</span></span>

<span data-ttu-id="d9e46-384">Получает сообщения уведомления для элемента.</span><span class="sxs-lookup"><span data-stu-id="d9e46-384">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="d9e46-385">Тип:</span><span class="sxs-lookup"><span data-stu-id="d9e46-385">Type:</span></span>

*   [<span data-ttu-id="d9e46-386">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="d9e46-386">NotificationMessages</span></span>](/javascript/api/outlook_1_4/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="d9e46-387">Требования</span><span class="sxs-lookup"><span data-stu-id="d9e46-387">Requirements</span></span>

|<span data-ttu-id="d9e46-388">Требование</span><span class="sxs-lookup"><span data-stu-id="d9e46-388">Requirement</span></span>| <span data-ttu-id="d9e46-389">Значение</span><span class="sxs-lookup"><span data-stu-id="d9e46-389">Value</span></span>|
|---|---|
|[<span data-ttu-id="d9e46-390">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="d9e46-390">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d9e46-391">1.3</span><span class="sxs-lookup"><span data-stu-id="d9e46-391">1.3</span></span>|
|[<span data-ttu-id="d9e46-392">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="d9e46-392">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d9e46-393">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d9e46-393">ReadItem</span></span>|
|[<span data-ttu-id="d9e46-394">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="d9e46-394">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d9e46-395">Cоздание или чтение</span><span class="sxs-lookup"><span data-stu-id="d9e46-395">Compose or read</span></span>|

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook14officeemailaddressdetailsrecipientsjavascriptapioutlook14officerecipients"></a><span data-ttu-id="d9e46-396">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_4/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="d9e46-396">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_4/office.recipients)</span></span>

<span data-ttu-id="d9e46-p123">Предоставляет доступ к необязательным участникам события. Уровень доступа и тип объекта зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="d9e46-p123">Provides access to the optional attendees of an event. The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="d9e46-399">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="d9e46-399">Read mode</span></span>

<span data-ttu-id="d9e46-400">Свойство `optionalAttendees` возвращает массив, содержащий объект `EmailAddressDetails` для каждого необязательного участника собрания.</span><span class="sxs-lookup"><span data-stu-id="d9e46-400">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="d9e46-401">Режим создания</span><span class="sxs-lookup"><span data-stu-id="d9e46-401">Compose mode</span></span>

<span data-ttu-id="d9e46-402">Свойство `optionalAttendees` возвращает объект `Recipients`, который предоставляет методы для получения и обновления необязательных участников собрания.</span><span class="sxs-lookup"><span data-stu-id="d9e46-402">The `optionalAttendees` property returns a `Recipients` object that provides methods to get and set the optional attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="d9e46-403">Тип:</span><span class="sxs-lookup"><span data-stu-id="d9e46-403">Type:</span></span>

*   <span data-ttu-id="d9e46-404">Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_4/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="d9e46-404">Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_4/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="d9e46-405">Требования</span><span class="sxs-lookup"><span data-stu-id="d9e46-405">Requirements</span></span>

|<span data-ttu-id="d9e46-406">Требование</span><span class="sxs-lookup"><span data-stu-id="d9e46-406">Requirement</span></span>| <span data-ttu-id="d9e46-407">Значение</span><span class="sxs-lookup"><span data-stu-id="d9e46-407">Value</span></span>|
|---|---|
|[<span data-ttu-id="d9e46-408">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="d9e46-408">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d9e46-409">1.0</span><span class="sxs-lookup"><span data-stu-id="d9e46-409">1.0</span></span>|
|[<span data-ttu-id="d9e46-410">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="d9e46-410">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d9e46-411">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d9e46-411">ReadItem</span></span>|
|[<span data-ttu-id="d9e46-412">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="d9e46-412">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d9e46-413">Cоздание или чтение</span><span class="sxs-lookup"><span data-stu-id="d9e46-413">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="d9e46-414">Пример</span><span class="sxs-lookup"><span data-stu-id="d9e46-414">Example</span></span>

```
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook14officeemailaddressdetails"></a><span data-ttu-id="d9e46-415">organizer :[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="d9e46-415">organizer :[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)</span></span>

<span data-ttu-id="d9e46-p124">Получает электронный адрес организатора указанного собрания. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="d9e46-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="d9e46-418">Тип:</span><span class="sxs-lookup"><span data-stu-id="d9e46-418">Type:</span></span>

*   [<span data-ttu-id="d9e46-419">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="d9e46-419">EmailAddressDetails</span></span>](/javascript/api/outlook_1_4/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="d9e46-420">Требования</span><span class="sxs-lookup"><span data-stu-id="d9e46-420">Requirements</span></span>

|<span data-ttu-id="d9e46-421">Требование</span><span class="sxs-lookup"><span data-stu-id="d9e46-421">Requirement</span></span>| <span data-ttu-id="d9e46-422">Значение</span><span class="sxs-lookup"><span data-stu-id="d9e46-422">Value</span></span>|
|---|---|
|[<span data-ttu-id="d9e46-423">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="d9e46-423">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d9e46-424">1.0</span><span class="sxs-lookup"><span data-stu-id="d9e46-424">1.0</span></span>|
|[<span data-ttu-id="d9e46-425">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="d9e46-425">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d9e46-426">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d9e46-426">ReadItem</span></span>|
|[<span data-ttu-id="d9e46-427">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="d9e46-427">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d9e46-428">Чтение</span><span class="sxs-lookup"><span data-stu-id="d9e46-428">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d9e46-429">Пример</span><span class="sxs-lookup"><span data-stu-id="d9e46-429">Example</span></span>

```
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
```

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook14officeemailaddressdetailsrecipientsjavascriptapioutlook14officerecipients"></a><span data-ttu-id="d9e46-430">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_4/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="d9e46-430">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_4/office.recipients)</span></span>

<span data-ttu-id="d9e46-p125">Предоставляет доступ к обязательным участникам события. Уровень доступа и тип объекта зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="d9e46-p125">Provides access to the required attendees of an event. The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="d9e46-433">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="d9e46-433">Read mode</span></span>

<span data-ttu-id="d9e46-434">Свойство `requiredAttendees` возвращает массив, содержащий объект `EmailAddressDetails`, каждому обязательному участнику собрания.</span><span class="sxs-lookup"><span data-stu-id="d9e46-434">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="d9e46-435">Режим создания</span><span class="sxs-lookup"><span data-stu-id="d9e46-435">Compose mode</span></span>

<span data-ttu-id="d9e46-436">Свойство `requiredAttendees` возвращает объект `Recipients`, который предоставляет методы для получения и обновления обязательных участников собрания.</span><span class="sxs-lookup"><span data-stu-id="d9e46-436">The `requiredAttendees` property returns a `Recipients` object that provides methods to get and set the required attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="d9e46-437">Тип:</span><span class="sxs-lookup"><span data-stu-id="d9e46-437">Type:</span></span>

*   <span data-ttu-id="d9e46-438">Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_4/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="d9e46-438">Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_4/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="d9e46-439">Требования</span><span class="sxs-lookup"><span data-stu-id="d9e46-439">Requirements</span></span>

|<span data-ttu-id="d9e46-440">Требование</span><span class="sxs-lookup"><span data-stu-id="d9e46-440">Requirement</span></span>| <span data-ttu-id="d9e46-441">Значение</span><span class="sxs-lookup"><span data-stu-id="d9e46-441">Value</span></span>|
|---|---|
|[<span data-ttu-id="d9e46-442">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="d9e46-442">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d9e46-443">1.0</span><span class="sxs-lookup"><span data-stu-id="d9e46-443">1.0</span></span>|
|[<span data-ttu-id="d9e46-444">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="d9e46-444">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d9e46-445">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d9e46-445">ReadItem</span></span>|
|[<span data-ttu-id="d9e46-446">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="d9e46-446">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d9e46-447">Cоздание или чтение</span><span class="sxs-lookup"><span data-stu-id="d9e46-447">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="d9e46-448">Пример</span><span class="sxs-lookup"><span data-stu-id="d9e46-448">Example</span></span>

```
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
}
```

#### <a name="sender-emailaddressdetailsjavascriptapioutlook14officeemailaddressdetails"></a><span data-ttu-id="d9e46-449">sender :[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="d9e46-449">sender :[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)</span></span>

<span data-ttu-id="d9e46-p126">Получает электронный адрес отправителя электронного сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="d9e46-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="d9e46-p127">Свойства [`from`](#from-emailaddressdetailsjavascriptapioutlook14officeemailaddressdetails) и `sender` представляют одно и то же лицо, если сообщение не отправлено делегатом. В противном случае свойство `from` представляет лицо, делегировавшее полномочия, а свойство sender — делегата.</span><span class="sxs-lookup"><span data-stu-id="d9e46-p127">The [`from`](#from-emailaddressdetailsjavascriptapioutlook14officeemailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="d9e46-454">Свойство `recipientType` объекта `EmailAddressDetails` в свойстве `sender` — `undefined`.</span><span class="sxs-lookup"><span data-stu-id="d9e46-454">Note: The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="d9e46-455">Тип:</span><span class="sxs-lookup"><span data-stu-id="d9e46-455">Type:</span></span>

*   [<span data-ttu-id="d9e46-456">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="d9e46-456">EmailAddressDetails</span></span>](/javascript/api/outlook_1_4/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="d9e46-457">Требования</span><span class="sxs-lookup"><span data-stu-id="d9e46-457">Requirements</span></span>

|<span data-ttu-id="d9e46-458">Требование</span><span class="sxs-lookup"><span data-stu-id="d9e46-458">Requirement</span></span>| <span data-ttu-id="d9e46-459">Значение</span><span class="sxs-lookup"><span data-stu-id="d9e46-459">Value</span></span>|
|---|---|
|[<span data-ttu-id="d9e46-460">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="d9e46-460">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d9e46-461">1.0</span><span class="sxs-lookup"><span data-stu-id="d9e46-461">1.0</span></span>|
|[<span data-ttu-id="d9e46-462">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="d9e46-462">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d9e46-463">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d9e46-463">ReadItem</span></span>|
|[<span data-ttu-id="d9e46-464">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="d9e46-464">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d9e46-465">Чтение</span><span class="sxs-lookup"><span data-stu-id="d9e46-465">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d9e46-466">Пример</span><span class="sxs-lookup"><span data-stu-id="d9e46-466">Example</span></span>

```
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
```

####  <a name="start-datetimejavascriptapioutlook14officetime"></a><span data-ttu-id="d9e46-467">start :Date|[Time](/javascript/api/outlook_1_4/office.time)</span><span class="sxs-lookup"><span data-stu-id="d9e46-467">start :Date|[Time](/javascript/api/outlook_1_4/office.time)</span></span>

<span data-ttu-id="d9e46-468">Получает или задает дату и время начала встречи.</span><span class="sxs-lookup"><span data-stu-id="d9e46-468">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="d9e46-p128">Свойство `start` представлено в виде значения даты и времени в формате UTC. Это значение можно преобразовать в местные значения даты и времени клиента с помощью метода [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook14officelocalclienttime).</span><span class="sxs-lookup"><span data-stu-id="d9e46-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook14officelocalclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="d9e46-471">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="d9e46-471">Read mode</span></span>

<span data-ttu-id="d9e46-472">Свойство `start` возвращает объект `Date`.</span><span class="sxs-lookup"><span data-stu-id="d9e46-472">The `start` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="d9e46-473">Режим создания</span><span class="sxs-lookup"><span data-stu-id="d9e46-473">Compose mode</span></span>

<span data-ttu-id="d9e46-474">Свойство `start` возвращает объект `Time`.</span><span class="sxs-lookup"><span data-stu-id="d9e46-474">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="d9e46-475">Если вы задаете время начала с помощью метода [`Time.setAsync`](/javascript/api/outlook_1_4/office.time#setasync-datetime--options--callback-), необходимо использовать метод [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) для преобразования местного времени на клиенте в формат UTC для сервера.</span><span class="sxs-lookup"><span data-stu-id="d9e46-475">When you use the [`Time.setAsync`](/javascript/api/outlook_1_4/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="d9e46-476">Тип:</span><span class="sxs-lookup"><span data-stu-id="d9e46-476">Type:</span></span>

*   <span data-ttu-id="d9e46-477">Date | [Time](/javascript/api/outlook_1_4/office.time)</span><span class="sxs-lookup"><span data-stu-id="d9e46-477">Date | [Time](/javascript/api/outlook_1_4/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="d9e46-478">Требования</span><span class="sxs-lookup"><span data-stu-id="d9e46-478">Requirements</span></span>

|<span data-ttu-id="d9e46-479">Требование</span><span class="sxs-lookup"><span data-stu-id="d9e46-479">Requirement</span></span>| <span data-ttu-id="d9e46-480">Значение</span><span class="sxs-lookup"><span data-stu-id="d9e46-480">Value</span></span>|
|---|---|
|[<span data-ttu-id="d9e46-481">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="d9e46-481">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d9e46-482">1.0</span><span class="sxs-lookup"><span data-stu-id="d9e46-482">1.0</span></span>|
|[<span data-ttu-id="d9e46-483">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="d9e46-483">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d9e46-484">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d9e46-484">ReadItem</span></span>|
|[<span data-ttu-id="d9e46-485">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="d9e46-485">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d9e46-486">Cоздание или чтение</span><span class="sxs-lookup"><span data-stu-id="d9e46-486">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="d9e46-487">Пример</span><span class="sxs-lookup"><span data-stu-id="d9e46-487">Example</span></span>

<span data-ttu-id="d9e46-488">В примере ниже с помощью метода [`setAsync`](/javascript/api/outlook_1_4/office.time#setasync-datetime--options--callback-) объекта `Time` задается время начала встречи в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="d9e46-488">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_4/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

####  <a name="subject-stringsubjectjavascriptapioutlook14officesubject"></a><span data-ttu-id="d9e46-489">subject :String|[Subject](/javascript/api/outlook_1_4/office.subject)</span><span class="sxs-lookup"><span data-stu-id="d9e46-489">subject :String|[Subject](/javascript/api/outlook_1_4/office.subject)</span></span>

<span data-ttu-id="d9e46-490">Получает или задает описание, которое отображается в поле темы элемента.</span><span class="sxs-lookup"><span data-stu-id="d9e46-490">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="d9e46-491">Свойство `subject` получает или задает всю тему элемента для отправки с почтового сервера.</span><span class="sxs-lookup"><span data-stu-id="d9e46-491">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="d9e46-492">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="d9e46-492">Read mode</span></span>

<span data-ttu-id="d9e46-p129">Свойство `subject` возвращает строку. С помощью свойства [`normalizedSubject`](#normalizedsubject-string) можно получить тему без начальных префиксов, например, `RE:` и `FW:`.</span><span class="sxs-lookup"><span data-stu-id="d9e46-p129">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```
var subject = Office.context.mailbox.item.subject;
```

##### <a name="compose-mode"></a><span data-ttu-id="d9e46-495">Режим создания</span><span class="sxs-lookup"><span data-stu-id="d9e46-495">Compose mode</span></span>

<span data-ttu-id="d9e46-496">Свойство `subject` возвращает объект `Subject`, который предоставляет методы для получения и задания темы.</span><span class="sxs-lookup"><span data-stu-id="d9e46-496">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="d9e46-497">Тип:</span><span class="sxs-lookup"><span data-stu-id="d9e46-497">Type:</span></span>

*   <span data-ttu-id="d9e46-498">String | [Subject](/javascript/api/outlook_1_4/office.subject)</span><span class="sxs-lookup"><span data-stu-id="d9e46-498">String | [Subject](/javascript/api/outlook_1_4/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="d9e46-499">Требования</span><span class="sxs-lookup"><span data-stu-id="d9e46-499">Requirements</span></span>

|<span data-ttu-id="d9e46-500">Требование</span><span class="sxs-lookup"><span data-stu-id="d9e46-500">Requirement</span></span>| <span data-ttu-id="d9e46-501">Значение</span><span class="sxs-lookup"><span data-stu-id="d9e46-501">Value</span></span>|
|---|---|
|[<span data-ttu-id="d9e46-502">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="d9e46-502">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d9e46-503">1.0</span><span class="sxs-lookup"><span data-stu-id="d9e46-503">1.0</span></span>|
|[<span data-ttu-id="d9e46-504">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="d9e46-504">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d9e46-505">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d9e46-505">ReadItem</span></span>|
|[<span data-ttu-id="d9e46-506">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="d9e46-506">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d9e46-507">Cоздание или чтение</span><span class="sxs-lookup"><span data-stu-id="d9e46-507">Compose or read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlook14officeemailaddressdetailsrecipientsjavascriptapioutlook14officerecipients"></a><span data-ttu-id="d9e46-508">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_4/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="d9e46-508">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_4/office.recipients)</span></span>

<span data-ttu-id="d9e46-p130">Предоставляет доступ к получателям в строке **To (Кому)** сообщения. Уровень доступа и тип объекта зависят от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="d9e46-p130">Provides access to the recipients on the **To** line of a message. The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="d9e46-511">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="d9e46-511">Read mode</span></span>

<span data-ttu-id="d9e46-p131">Свойство `to` возвращает массив, содержащий объект `EmailAddressDetails` для каждого получателя в строке **To (Кому)** сообщения. Коллекция может включать не более 100 элементов.</span><span class="sxs-lookup"><span data-stu-id="d9e46-p131">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="d9e46-514">Режим создания</span><span class="sxs-lookup"><span data-stu-id="d9e46-514">Compose mode</span></span>

<span data-ttu-id="d9e46-515">Свойство `to` возвращает объект `Recipients`, предоставляющий методы для получения или обновления получателей, которые указаны в строке **To (кому)** сообщения.</span><span class="sxs-lookup"><span data-stu-id="d9e46-515">The to`to` property returns a Recipients`Recipients` object that provides methods for manipulating the recipients on the **To** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="d9e46-516">Тип:</span><span class="sxs-lookup"><span data-stu-id="d9e46-516">Type:</span></span>

*   <span data-ttu-id="d9e46-517">Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_4/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="d9e46-517">Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_4/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="d9e46-518">Требования</span><span class="sxs-lookup"><span data-stu-id="d9e46-518">Requirements</span></span>

|<span data-ttu-id="d9e46-519">Требование</span><span class="sxs-lookup"><span data-stu-id="d9e46-519">Requirement</span></span>| <span data-ttu-id="d9e46-520">Значение</span><span class="sxs-lookup"><span data-stu-id="d9e46-520">Value</span></span>|
|---|---|
|[<span data-ttu-id="d9e46-521">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="d9e46-521">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d9e46-522">1.0</span><span class="sxs-lookup"><span data-stu-id="d9e46-522">1.0</span></span>|
|[<span data-ttu-id="d9e46-523">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="d9e46-523">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d9e46-524">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d9e46-524">ReadItem</span></span>|
|[<span data-ttu-id="d9e46-525">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="d9e46-525">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d9e46-526">Cоздание или чтение</span><span class="sxs-lookup"><span data-stu-id="d9e46-526">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="d9e46-527">Пример</span><span class="sxs-lookup"><span data-stu-id="d9e46-527">Example</span></span>

```
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

### <a name="methods"></a><span data-ttu-id="d9e46-528">Методы</span><span class="sxs-lookup"><span data-stu-id="d9e46-528">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="d9e46-529">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="d9e46-529">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="d9e46-530">Добавляет файл в сообщение или встречу в качестве вложения.</span><span class="sxs-lookup"><span data-stu-id="d9e46-530">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="d9e46-531">Метод `addFileAttachmentAsync` передает файл по указанному универсальному коду ресурса (URI) и вкладывает его в элемент в форме создания.</span><span class="sxs-lookup"><span data-stu-id="d9e46-531">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="d9e46-532">Идентификатор можно последовательно использовать с методом [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback), чтобы удалить вложение, добавленное во время текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="d9e46-532">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d9e46-533">Параметры:</span><span class="sxs-lookup"><span data-stu-id="d9e46-533">Parameters:</span></span>

|<span data-ttu-id="d9e46-534">Имя</span><span class="sxs-lookup"><span data-stu-id="d9e46-534">Name</span></span>| <span data-ttu-id="d9e46-535">Тип</span><span class="sxs-lookup"><span data-stu-id="d9e46-535">Type</span></span>| <span data-ttu-id="d9e46-536">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="d9e46-536">Attributes</span></span>| <span data-ttu-id="d9e46-537">Описание</span><span class="sxs-lookup"><span data-stu-id="d9e46-537">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="d9e46-538">String</span><span class="sxs-lookup"><span data-stu-id="d9e46-538">String</span></span>||<span data-ttu-id="d9e46-p132">URI-адрес, представляющий расположение файла, который нужно вложить в сообщение или встречу. Максимальная длина — 2048 символов.</span><span class="sxs-lookup"><span data-stu-id="d9e46-p132">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="d9e46-541">String</span><span class="sxs-lookup"><span data-stu-id="d9e46-541">String</span></span>||<span data-ttu-id="d9e46-p133">Имя вложения, которое отображается при передаче вложения. Максимальная длина — 255 символов.</span><span class="sxs-lookup"><span data-stu-id="d9e46-p133">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="d9e46-544">Объект</span><span class="sxs-lookup"><span data-stu-id="d9e46-544">Object</span></span>| <span data-ttu-id="d9e46-545">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d9e46-545">&lt;optional&gt;</span></span>|<span data-ttu-id="d9e46-546">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="d9e46-546">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="d9e46-547">Объект</span><span class="sxs-lookup"><span data-stu-id="d9e46-547">Object</span></span>| <span data-ttu-id="d9e46-548">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="d9e46-548">&lt;optional&gt;</span></span>|<span data-ttu-id="d9e46-549">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="d9e46-549">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="d9e46-550">function</span><span class="sxs-lookup"><span data-stu-id="d9e46-550">function</span></span>| <span data-ttu-id="d9e46-551">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="d9e46-551">&lt;optional&gt;</span></span>|<span data-ttu-id="d9e46-552">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="d9e46-552">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="d9e46-553">После успешного выполнения идентификатор вложения будет представлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="d9e46-553">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="d9e46-554">Если добавить вложение не удастся, объект `asyncResult` будет содержать объект `Error` с описанием ошибки.</span><span class="sxs-lookup"><span data-stu-id="d9e46-554">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="d9e46-555">Ошибки</span><span class="sxs-lookup"><span data-stu-id="d9e46-555">Errors</span></span>

| <span data-ttu-id="d9e46-556">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="d9e46-556">Error code</span></span> | <span data-ttu-id="d9e46-557">Описание</span><span class="sxs-lookup"><span data-stu-id="d9e46-557">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="d9e46-558">Вложение превышает максимальный размер.</span><span class="sxs-lookup"><span data-stu-id="d9e46-558">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="d9e46-559">Расширение вложения не поддерживается.</span><span class="sxs-lookup"><span data-stu-id="d9e46-559">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="d9e46-560">Сообщение или встреча содержат слишком много вложений.</span><span class="sxs-lookup"><span data-stu-id="d9e46-560">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="d9e46-561">Требования</span><span class="sxs-lookup"><span data-stu-id="d9e46-561">Requirements</span></span>

|<span data-ttu-id="d9e46-562">Требование</span><span class="sxs-lookup"><span data-stu-id="d9e46-562">Requirement</span></span>| <span data-ttu-id="d9e46-563">Значение</span><span class="sxs-lookup"><span data-stu-id="d9e46-563">Value</span></span>|
|---|---|
|[<span data-ttu-id="d9e46-564">Версия минимального набора требований для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="d9e46-564">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d9e46-565">1.1</span><span class="sxs-lookup"><span data-stu-id="d9e46-565">1.1</span></span>|
|[<span data-ttu-id="d9e46-566">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="d9e46-566">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d9e46-567">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="d9e46-567">ReadWriteItem</span></span>|
|[<span data-ttu-id="d9e46-568">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="d9e46-568">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d9e46-569">Создание</span><span class="sxs-lookup"><span data-stu-id="d9e46-569">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="d9e46-570">Пример</span><span class="sxs-lookup"><span data-stu-id="d9e46-570">Example</span></span>

```
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

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="d9e46-571">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="d9e46-571">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="d9e46-572">Добавляет к сообщению или встрече элемент Exchange (например, сообщение) в виде вложения.</span><span class="sxs-lookup"><span data-stu-id="d9e46-572">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="d9e46-p134">С помощью метода `addItemAttachmentAsync` в элемент формы создания можно вложить элемент с указанным идентификатором Exchange. Если указать метод обратного вызова, то этот метод вызывается с помощью параметра `asyncResult`, который содержит идентификатор вложения или код, указывающий на ошибки, которые произошли при вложении элемента. При необходимости можно использовать параметр `options` для передачи сведений о состоянии в метод обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="d9e46-p134">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="d9e46-576">Идентификатор можно последовательно использовать с методом [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback), чтобы удалить вложение, добавленное во время текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="d9e46-576">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="d9e46-577">Если ваша надстройка Office выполняется в веб-приложении Outlook, метод `addItemAttachmentAsync` обеспечивает вложение элементов в элементы, отличные от редактируемого. Однако это действие не рекомендуется выполнять, так как оно не поддерживается.</span><span class="sxs-lookup"><span data-stu-id="d9e46-577">If your Office add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d9e46-578">Параметры:</span><span class="sxs-lookup"><span data-stu-id="d9e46-578">Parameters:</span></span>

|<span data-ttu-id="d9e46-579">Имя</span><span class="sxs-lookup"><span data-stu-id="d9e46-579">Name</span></span>| <span data-ttu-id="d9e46-580">Тип</span><span class="sxs-lookup"><span data-stu-id="d9e46-580">Type</span></span>| <span data-ttu-id="d9e46-581">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="d9e46-581">Attributes</span></span>| <span data-ttu-id="d9e46-582">Описание</span><span class="sxs-lookup"><span data-stu-id="d9e46-582">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="d9e46-583">String</span><span class="sxs-lookup"><span data-stu-id="d9e46-583">String</span></span>||<span data-ttu-id="d9e46-p135">Идентификатор Exchange для вкладываемого элемента. Максимальная длина — 100 символов.</span><span class="sxs-lookup"><span data-stu-id="d9e46-p135">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="d9e46-586">String</span><span class="sxs-lookup"><span data-stu-id="d9e46-586">String</span></span>||<span data-ttu-id="d9e46-p136">Тема вкладываемого элемента. Максимальная длина — 255 символов.</span><span class="sxs-lookup"><span data-stu-id="d9e46-p136">The sujbect of the item to be attached. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="d9e46-589">Объект</span><span class="sxs-lookup"><span data-stu-id="d9e46-589">Object</span></span>| <span data-ttu-id="d9e46-590">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="d9e46-590">&lt;optional&gt;</span></span>|<span data-ttu-id="d9e46-591">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="d9e46-591">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="d9e46-592">Объект</span><span class="sxs-lookup"><span data-stu-id="d9e46-592">Object</span></span>| <span data-ttu-id="d9e46-593">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="d9e46-593">&lt;optional&gt;</span></span>|<span data-ttu-id="d9e46-594">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="d9e46-594">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="d9e46-595">function</span><span class="sxs-lookup"><span data-stu-id="d9e46-595">function</span></span>| <span data-ttu-id="d9e46-596">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="d9e46-596">&lt;optional&gt;</span></span>|<span data-ttu-id="d9e46-597">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="d9e46-597">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="d9e46-598">После успешного выполнения идентификатор вложения будет представлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="d9e46-598">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="d9e46-599">Если добавить вложение не удастся, объект `asyncResult` будет содержать объект `Error` с описанием ошибки.</span><span class="sxs-lookup"><span data-stu-id="d9e46-599">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="d9e46-600">Ошибки</span><span class="sxs-lookup"><span data-stu-id="d9e46-600">Errors</span></span>

| <span data-ttu-id="d9e46-601">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="d9e46-601">Error code</span></span> | <span data-ttu-id="d9e46-602">Описание</span><span class="sxs-lookup"><span data-stu-id="d9e46-602">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="d9e46-603">Сообщение или встреча содержат слишком много вложений.</span><span class="sxs-lookup"><span data-stu-id="d9e46-603">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="d9e46-604">Требования</span><span class="sxs-lookup"><span data-stu-id="d9e46-604">Requirements</span></span>

|<span data-ttu-id="d9e46-605">Требование</span><span class="sxs-lookup"><span data-stu-id="d9e46-605">Requirement</span></span>| <span data-ttu-id="d9e46-606">Значение</span><span class="sxs-lookup"><span data-stu-id="d9e46-606">Value</span></span>|
|---|---|
|[<span data-ttu-id="d9e46-607">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="d9e46-607">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d9e46-608">1.1</span><span class="sxs-lookup"><span data-stu-id="d9e46-608">1.1</span></span>|
|[<span data-ttu-id="d9e46-609">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="d9e46-609">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d9e46-610">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="d9e46-610">ReadWriteItem</span></span>|
|[<span data-ttu-id="d9e46-611">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="d9e46-611">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d9e46-612">Создание</span><span class="sxs-lookup"><span data-stu-id="d9e46-612">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="d9e46-613">Пример</span><span class="sxs-lookup"><span data-stu-id="d9e46-613">Example</span></span>

<span data-ttu-id="d9e46-614">В следующем примере существующий элемент Outlook добавляется в виде вложения с именем `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="d9e46-614">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

####  <a name="close"></a><span data-ttu-id="d9e46-615">close()</span><span class="sxs-lookup"><span data-stu-id="d9e46-615">close()</span></span>

<span data-ttu-id="d9e46-616">Закрывает текущий создаваемый элемент.</span><span class="sxs-lookup"><span data-stu-id="d9e46-616">Closes the current item that is being composed.</span></span>

<span data-ttu-id="d9e46-p137">Работа метода `close` зависит от текущего состояния создаваемого элемента. Если элемент содержит несохраненные изменения, клиент предложит пользователю сохранить или отклонить их либо отменить действие закрытия.</span><span class="sxs-lookup"><span data-stu-id="d9e46-p137">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="d9e46-619">Если элемент является встречей в Outlook в Интернете, и он был ранее сохранен с помощью `saveAsync`, пользователю предлагается сохранить, отменить или удалить его, даже если не произошло каких-либо изменений, поскольку этот элемент был последним сохраненным.</span><span class="sxs-lookup"><span data-stu-id="d9e46-619">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="d9e46-620">Если в клиенте Outlook для настольных ПК сообщение представляет собой ответ в тексте, метод `close` не работает.</span><span class="sxs-lookup"><span data-stu-id="d9e46-620">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="d9e46-621">Требования</span><span class="sxs-lookup"><span data-stu-id="d9e46-621">Requirements</span></span>

|<span data-ttu-id="d9e46-622">Требование</span><span class="sxs-lookup"><span data-stu-id="d9e46-622">Requirement</span></span>| <span data-ttu-id="d9e46-623">Значение</span><span class="sxs-lookup"><span data-stu-id="d9e46-623">Value</span></span>|
|---|---|
|[<span data-ttu-id="d9e46-624">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="d9e46-624">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d9e46-625">1.3</span><span class="sxs-lookup"><span data-stu-id="d9e46-625">1.3</span></span>|
|[<span data-ttu-id="d9e46-626">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="d9e46-626">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d9e46-627">Ограниченный доступ</span><span class="sxs-lookup"><span data-stu-id="d9e46-627">Restricted</span></span>|
|[<span data-ttu-id="d9e46-628">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="d9e46-628">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d9e46-629">Создание</span><span class="sxs-lookup"><span data-stu-id="d9e46-629">Compose</span></span>|

#### <a name="displayreplyallformformdata"></a><span data-ttu-id="d9e46-630">displayReplyAllForm(formData)</span><span class="sxs-lookup"><span data-stu-id="d9e46-630">displayReplyAllForm(formData)</span></span>

<span data-ttu-id="d9e46-631">Отображает форму ответа, включающую отправителя и всех получателей выбранного сообщения или организатора и всех участников выбранной встречи.</span><span class="sxs-lookup"><span data-stu-id="d9e46-631">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="d9e46-632">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="d9e46-632">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="d9e46-633">В веб-приложении Outlook форма ответа отображается в виде всплывающей формы в представлении с 3 колонками либо всплывающей формы в представлении с 1 или 2 колонками.</span><span class="sxs-lookup"><span data-stu-id="d9e46-633">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="d9e46-634">Если любой строчный параметр превышает указанные для него ограничения, `displayReplyAllForm` возвращает исключение.</span><span class="sxs-lookup"><span data-stu-id="d9e46-634">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="d9e46-p138">Если в параметре `formData.attachments` указаны вложения, Outlook и веб-приложение Outlook пытаются скачать их и вложить в форму ответа. Если какие-либо вложения добавить не удается, в форме отображается сообщение об ошибке. Если сообщения об ошибках не предусмотрены, то они не отображаются.</span><span class="sxs-lookup"><span data-stu-id="d9e46-p138">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d9e46-638">Параметры:</span><span class="sxs-lookup"><span data-stu-id="d9e46-638">Parameters:</span></span>

|<span data-ttu-id="d9e46-639">Имя</span><span class="sxs-lookup"><span data-stu-id="d9e46-639">Name</span></span>| <span data-ttu-id="d9e46-640">Тип</span><span class="sxs-lookup"><span data-stu-id="d9e46-640">Type</span></span>| <span data-ttu-id="d9e46-641">Описание</span><span class="sxs-lookup"><span data-stu-id="d9e46-641">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="d9e46-642">String | Object</span><span class="sxs-lookup"><span data-stu-id="d9e46-642">String &#124; Object</span></span>| |<span data-ttu-id="d9e46-p139">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="d9e46-p139">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="d9e46-645">**ИЛИ**</span><span class="sxs-lookup"><span data-stu-id="d9e46-645">**OR**</span></span><br/><span data-ttu-id="d9e46-p140">Объект, который содержит текст или данные вложения и функцию обратного вызова. Ниже представлено определение этого объекта.</span><span class="sxs-lookup"><span data-stu-id="d9e46-p140">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="d9e46-648">String</span><span class="sxs-lookup"><span data-stu-id="d9e46-648">String</span></span> | <span data-ttu-id="d9e46-649">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="d9e46-649">&lt;optional&gt;</span></span> | <span data-ttu-id="d9e46-p141">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="d9e46-p141">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="d9e46-652">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="d9e46-652">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="d9e46-653">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="d9e46-653">&lt;optional&gt;</span></span> | <span data-ttu-id="d9e46-654">Массив объектов JSON, представляющих собой вложенные файлы или элементы.</span><span class="sxs-lookup"><span data-stu-id="d9e46-654">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="d9e46-655">String</span><span class="sxs-lookup"><span data-stu-id="d9e46-655">String</span></span> | | <span data-ttu-id="d9e46-p142">Указывает тип вложения. Допустимые значения: `file` для вложенного файла и `item` для вложенного элемента.</span><span class="sxs-lookup"><span data-stu-id="d9e46-p142">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="d9e46-658">String</span><span class="sxs-lookup"><span data-stu-id="d9e46-658">String</span></span> | | <span data-ttu-id="d9e46-659">Строка, содержащая имя вложения, длиной до 255 символов.</span><span class="sxs-lookup"><span data-stu-id="d9e46-659">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="d9e46-660">String</span><span class="sxs-lookup"><span data-stu-id="d9e46-660">String</span></span> | | <span data-ttu-id="d9e46-p143">Используется, только если свойству `type` задано значение `file`. URI расположения файла.</span><span class="sxs-lookup"><span data-stu-id="d9e46-p143">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="d9e46-663">String</span><span class="sxs-lookup"><span data-stu-id="d9e46-663">String</span></span> | | <span data-ttu-id="d9e46-p144">Используется, только если свойству `type` задано значение `item`. Идентификатор вложения EWS. Это строка длиной до 100 символов.</span><span class="sxs-lookup"><span data-stu-id="d9e46-p144">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="d9e46-667">function</span><span class="sxs-lookup"><span data-stu-id="d9e46-667">function</span></span> | <span data-ttu-id="d9e46-668">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="d9e46-668">&lt;optional&gt;</span></span> | <span data-ttu-id="d9e46-669">По завершении работы метода функция, переданная параметру `callback`, вызывается с помощью одного параметра `asyncResult`, представляющего собой объект [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="d9e46-669">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="d9e46-670">Требования</span><span class="sxs-lookup"><span data-stu-id="d9e46-670">Requirements</span></span>

|<span data-ttu-id="d9e46-671">Требование</span><span class="sxs-lookup"><span data-stu-id="d9e46-671">Requirement</span></span>| <span data-ttu-id="d9e46-672">Значение</span><span class="sxs-lookup"><span data-stu-id="d9e46-672">Value</span></span>|
|---|---|
|[<span data-ttu-id="d9e46-673">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="d9e46-673">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d9e46-674">1.0</span><span class="sxs-lookup"><span data-stu-id="d9e46-674">1.0</span></span>|
|[<span data-ttu-id="d9e46-675">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="d9e46-675">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d9e46-676">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d9e46-676">ReadItem</span></span>|
|[<span data-ttu-id="d9e46-677">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="d9e46-677">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d9e46-678">Чтение</span><span class="sxs-lookup"><span data-stu-id="d9e46-678">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="d9e46-679">Примеры</span><span class="sxs-lookup"><span data-stu-id="d9e46-679">Examples</span></span>

<span data-ttu-id="d9e46-680">Приведенный ниже код передает строку в функцию `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="d9e46-680">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="d9e46-681">Ответ с пустым текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="d9e46-681">Reply with an empty body.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="d9e46-682">Ответ только с текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="d9e46-682">Reply with just a body.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="d9e46-683">Ответ с текстом сообщения и вложенным файлом.</span><span class="sxs-lookup"><span data-stu-id="d9e46-683">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="d9e46-684">Ответ с текстом сообщения и вложенным элементом.</span><span class="sxs-lookup"><span data-stu-id="d9e46-684">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="d9e46-685">Ответ с текстом сообщения, вложенным файлом, вложенным элементом и обратным вызовом.</span><span class="sxs-lookup"><span data-stu-id="d9e46-685">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata"></a><span data-ttu-id="d9e46-686">displayReplyForm(formData)</span><span class="sxs-lookup"><span data-stu-id="d9e46-686">displayReplyForm(formData)</span></span>

<span data-ttu-id="d9e46-687">Отображает форму ответа, включающую только отправителя выбранного сообщения или организатора выбранной встречи.</span><span class="sxs-lookup"><span data-stu-id="d9e46-687">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="d9e46-688">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="d9e46-688">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="d9e46-689">В веб-приложении Outlook форма ответа отображается в виде всплывающей формы в представлении с 3 колонками либо всплывающей формы в представлении с 1 или 2 колонками.</span><span class="sxs-lookup"><span data-stu-id="d9e46-689">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="d9e46-690">Если любой строчный параметр превышает указанные для него ограничения, `displayReplyForm` возвращает исключение.</span><span class="sxs-lookup"><span data-stu-id="d9e46-690">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="d9e46-p145">Если в параметре `formData.attachments` указаны вложения, Outlook и веб-приложение Outlook пытаются скачать их и вложить в форму ответа. Если какие-либо вложения добавить не удается, в форме отображается сообщение об ошибке. Если сообщения об ошибках не предусмотрены, то они не отображаются.</span><span class="sxs-lookup"><span data-stu-id="d9e46-p145">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d9e46-694">Параметры:</span><span class="sxs-lookup"><span data-stu-id="d9e46-694">Parameters:</span></span>

|<span data-ttu-id="d9e46-695">Имя</span><span class="sxs-lookup"><span data-stu-id="d9e46-695">Name</span></span>| <span data-ttu-id="d9e46-696">Тип</span><span class="sxs-lookup"><span data-stu-id="d9e46-696">Type</span></span>| <span data-ttu-id="d9e46-697">Описание</span><span class="sxs-lookup"><span data-stu-id="d9e46-697">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="d9e46-698">String | Object</span><span class="sxs-lookup"><span data-stu-id="d9e46-698">String &#124; Object</span></span>| | <span data-ttu-id="d9e46-p146">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="d9e46-p146">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="d9e46-701">**ИЛИ**</span><span class="sxs-lookup"><span data-stu-id="d9e46-701">**OR**</span></span><br/><span data-ttu-id="d9e46-p147">Объект, который содержит текст или данные вложения и функцию обратного вызова. Ниже представлено определение этого объекта.</span><span class="sxs-lookup"><span data-stu-id="d9e46-p147">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="d9e46-704">String</span><span class="sxs-lookup"><span data-stu-id="d9e46-704">String</span></span> | <span data-ttu-id="d9e46-705">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="d9e46-705">&lt;optional&gt;</span></span> | <span data-ttu-id="d9e46-p148">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="d9e46-p148">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="d9e46-708">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="d9e46-708">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="d9e46-709">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="d9e46-709">&lt;optional&gt;</span></span> | <span data-ttu-id="d9e46-710">Массив объектов JSON, представляющих собой вложенные файлы или элементы.</span><span class="sxs-lookup"><span data-stu-id="d9e46-710">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="d9e46-711">String</span><span class="sxs-lookup"><span data-stu-id="d9e46-711">String</span></span> | | <span data-ttu-id="d9e46-p149">Указывает тип вложения. Допустимые значения: `file` для вложенного файла и `item` для вложенного элемента.</span><span class="sxs-lookup"><span data-stu-id="d9e46-p149">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="d9e46-714">String</span><span class="sxs-lookup"><span data-stu-id="d9e46-714">String</span></span> | | <span data-ttu-id="d9e46-715">Строка, содержащая имя вложения, длиной до 255 символов.</span><span class="sxs-lookup"><span data-stu-id="d9e46-715">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="d9e46-716">String</span><span class="sxs-lookup"><span data-stu-id="d9e46-716">String</span></span> | | <span data-ttu-id="d9e46-p150">Используется, только если свойству `type` задано значение `file`. URI расположения файла.</span><span class="sxs-lookup"><span data-stu-id="d9e46-p150">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="d9e46-719">String</span><span class="sxs-lookup"><span data-stu-id="d9e46-719">String</span></span> | | <span data-ttu-id="d9e46-p151">Используется, только если свойству `type` задано значение `item`. Идентификатор вложения EWS. Это строка длиной до 100 символов.</span><span class="sxs-lookup"><span data-stu-id="d9e46-p151">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="d9e46-723">function</span><span class="sxs-lookup"><span data-stu-id="d9e46-723">function</span></span> | <span data-ttu-id="d9e46-724">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="d9e46-724">&lt;optional&gt;</span></span> | <span data-ttu-id="d9e46-725">По завершении работы метода функция, переданная параметру `callback`, вызывается с помощью одного параметра `asyncResult`, представляющего собой объект [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="d9e46-725">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="d9e46-726">Требования</span><span class="sxs-lookup"><span data-stu-id="d9e46-726">Requirements</span></span>

|<span data-ttu-id="d9e46-727">Требование</span><span class="sxs-lookup"><span data-stu-id="d9e46-727">Requirement</span></span>| <span data-ttu-id="d9e46-728">Значение</span><span class="sxs-lookup"><span data-stu-id="d9e46-728">Value</span></span>|
|---|---|
|[<span data-ttu-id="d9e46-729">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="d9e46-729">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d9e46-730">1.0</span><span class="sxs-lookup"><span data-stu-id="d9e46-730">1.0</span></span>|
|[<span data-ttu-id="d9e46-731">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="d9e46-731">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d9e46-732">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d9e46-732">ReadItem</span></span>|
|[<span data-ttu-id="d9e46-733">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="d9e46-733">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d9e46-734">Чтение</span><span class="sxs-lookup"><span data-stu-id="d9e46-734">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="d9e46-735">Примеры</span><span class="sxs-lookup"><span data-stu-id="d9e46-735">Examples</span></span>

<span data-ttu-id="d9e46-736">Приведенный ниже код передает строку в функцию `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="d9e46-736">The following code passes a string to the `displayReplyForm` function.</span></span>

```
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="d9e46-737">Ответ с пустым текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="d9e46-737">Reply with an empty body.</span></span>

```
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="d9e46-738">Ответ только с текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="d9e46-738">Reply with just a body.</span></span>

```
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="d9e46-739">Ответ с текстом сообщения и вложенным файлом.</span><span class="sxs-lookup"><span data-stu-id="d9e46-739">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="d9e46-740">Ответ с текстом сообщения и вложенным элементом.</span><span class="sxs-lookup"><span data-stu-id="d9e46-740">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="d9e46-741">Ответ с текстом сообщения, вложенным файлом, вложенным элементом и обратным вызовом.</span><span class="sxs-lookup"><span data-stu-id="d9e46-741">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlook14officeentities"></a><span data-ttu-id="d9e46-742">getEntities() → {[Entities](/javascript/api/outlook_1_4/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="d9e46-742">getEntities() → {[Entities](/javascript/api/outlook_1_4/office.entities)}</span></span>

<span data-ttu-id="d9e46-743">Получает сущности, обнаруженные в выбранном тексте элемента.</span><span class="sxs-lookup"><span data-stu-id="d9e46-743">Gets the entities found in the selected item.</span></span>

> [!NOTE]
> <span data-ttu-id="d9e46-744">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="d9e46-744">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="d9e46-745">Требования</span><span class="sxs-lookup"><span data-stu-id="d9e46-745">Requirements</span></span>

|<span data-ttu-id="d9e46-746">Требование</span><span class="sxs-lookup"><span data-stu-id="d9e46-746">Requirement</span></span>| <span data-ttu-id="d9e46-747">Значение</span><span class="sxs-lookup"><span data-stu-id="d9e46-747">Value</span></span>|
|---|---|
|[<span data-ttu-id="d9e46-748">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="d9e46-748">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d9e46-749">1.0</span><span class="sxs-lookup"><span data-stu-id="d9e46-749">1.0</span></span>|
|[<span data-ttu-id="d9e46-750">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="d9e46-750">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d9e46-751">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d9e46-751">ReadItem</span></span>|
|[<span data-ttu-id="d9e46-752">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="d9e46-752">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d9e46-753">Чтение</span><span class="sxs-lookup"><span data-stu-id="d9e46-753">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="d9e46-754">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="d9e46-754">Returns:</span></span>

<span data-ttu-id="d9e46-755">Тип: [Entities](/javascript/api/outlook_1_4/office.entities)</span><span class="sxs-lookup"><span data-stu-id="d9e46-755">Type: [Entities](/javascript/api/outlook_1_4/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="d9e46-756">Пример</span><span class="sxs-lookup"><span data-stu-id="d9e46-756">Example</span></span>

<span data-ttu-id="d9e46-757">Ниже приведен пример получения доступа к сущностям контактов в тексте текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="d9e46-757">The following example accesses the contacts entities on the current item.</span></span>

```
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook14officecontactmeetingsuggestionjavascriptapioutlook14officemeetingsuggestionphonenumberjavascriptapioutlook14officephonenumbertasksuggestionjavascriptapioutlook14officetasksuggestion"></a><span data-ttu-id="d9e46-758">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_4/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_4/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_4/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_4/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="d9e46-758">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_4/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_4/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_4/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_4/office.tasksuggestion))>}</span></span>

<span data-ttu-id="d9e46-759">Получает массив всех сущностей указанного типа, обнаруженных в тексте выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="d9e46-759">Gets an array of all the entities of the specified entity type found in the selected item.</span></span>

> [!NOTE]
> <span data-ttu-id="d9e46-760">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="d9e46-760">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d9e46-761">Параметры:</span><span class="sxs-lookup"><span data-stu-id="d9e46-761">Parameters:</span></span>

|<span data-ttu-id="d9e46-762">Имя</span><span class="sxs-lookup"><span data-stu-id="d9e46-762">Name</span></span>| <span data-ttu-id="d9e46-763">Тип</span><span class="sxs-lookup"><span data-stu-id="d9e46-763">Type</span></span>| <span data-ttu-id="d9e46-764">Описание</span><span class="sxs-lookup"><span data-stu-id="d9e46-764">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="d9e46-765">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="d9e46-765">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook_1_4/office.mailboxenums.entitytype)|<span data-ttu-id="d9e46-766">Одно из значений перечисления EntityType.</span><span class="sxs-lookup"><span data-stu-id="d9e46-766">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d9e46-767">Требования</span><span class="sxs-lookup"><span data-stu-id="d9e46-767">Requirements</span></span>

|<span data-ttu-id="d9e46-768">Требование</span><span class="sxs-lookup"><span data-stu-id="d9e46-768">Requirement</span></span>| <span data-ttu-id="d9e46-769">Значение</span><span class="sxs-lookup"><span data-stu-id="d9e46-769">Value</span></span>|
|---|---|
|[<span data-ttu-id="d9e46-770">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="d9e46-770">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d9e46-771">1.0</span><span class="sxs-lookup"><span data-stu-id="d9e46-771">1.0</span></span>|
|[<span data-ttu-id="d9e46-772">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="d9e46-772">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d9e46-773">Ограниченный доступ</span><span class="sxs-lookup"><span data-stu-id="d9e46-773">Restricted</span></span>|
|[<span data-ttu-id="d9e46-774">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="d9e46-774">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d9e46-775">Чтение</span><span class="sxs-lookup"><span data-stu-id="d9e46-775">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="d9e46-776">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="d9e46-776">Returns:</span></span>

<span data-ttu-id="d9e46-p152">Если значение, переданное в `entityType`, не является допустимым членом перечисления `EntityType`, метод возвращает значение NULL. Если сущности указанного типа отсутствуют в тексте элемента, метод возвращает пустой массив. В противном случае — тип объектов в возвращаемом массиве зависит от типа сущности, запрошенной в параметре `entityType`.</span><span class="sxs-lookup"><span data-stu-id="d9e46-p152">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null. If no entities of the specified type are present in the item's body, the method returns an empty array. Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="d9e46-780">Хотя минимальный уровень разрешений для использования этого метода — **Restricted**, для некоторых типов сущностей требуется доступ на уровне **ReadItem**, как указано в приведенной ниже таблице.</span><span class="sxs-lookup"><span data-stu-id="d9e46-780">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="d9e46-781">Значение параметра `entityType`</span><span class="sxs-lookup"><span data-stu-id="d9e46-781">Value of `entityType`</span></span> | <span data-ttu-id="d9e46-782">Тип объектов в возвращаемом массиве</span><span class="sxs-lookup"><span data-stu-id="d9e46-782">Type of objects in returned array</span></span> | <span data-ttu-id="d9e46-783">Необходимый уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="d9e46-783">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="d9e46-784">String</span><span class="sxs-lookup"><span data-stu-id="d9e46-784">String</span></span> | <span data-ttu-id="d9e46-785">**С ограничениями**</span><span class="sxs-lookup"><span data-stu-id="d9e46-785">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="d9e46-786">Contact</span><span class="sxs-lookup"><span data-stu-id="d9e46-786">Contact</span></span> | <span data-ttu-id="d9e46-787">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="d9e46-787">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="d9e46-788">String</span><span class="sxs-lookup"><span data-stu-id="d9e46-788">String</span></span> | <span data-ttu-id="d9e46-789">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="d9e46-789">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="d9e46-790">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="d9e46-790">MeetingSuggestion</span></span> | <span data-ttu-id="d9e46-791">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="d9e46-791">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="d9e46-792">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="d9e46-792">PhoneNumber</span></span> | <span data-ttu-id="d9e46-793">**С ограничениями**</span><span class="sxs-lookup"><span data-stu-id="d9e46-793">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="d9e46-794">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="d9e46-794">TaskSuggestion</span></span> | <span data-ttu-id="d9e46-795">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="d9e46-795">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="d9e46-796">String</span><span class="sxs-lookup"><span data-stu-id="d9e46-796">String</span></span> | <span data-ttu-id="d9e46-797">**С ограничениями**</span><span class="sxs-lookup"><span data-stu-id="d9e46-797">**Restricted**</span></span> |

<span data-ttu-id="d9e46-798">Type: Array.<(String|[Contact](/javascript/api/outlook_1_4/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_4/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_4/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_4/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="d9e46-798">Type: Array.<(String|[Contact](/javascript/api/outlook_1_4/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_4/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_4/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_4/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="d9e46-799">Пример</span><span class="sxs-lookup"><span data-stu-id="d9e46-799">Example</span></span>

<span data-ttu-id="d9e46-800">В примере ниже показано, как получить доступ к массиву строк, которые представляют собой почтовые адреса в тексте текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="d9e46-800">The following example shows how to access an array of strings that represent postal addresses in the subject or body of the current item.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook14officecontactmeetingsuggestionjavascriptapioutlook14officemeetingsuggestionphonenumberjavascriptapioutlook14officephonenumbertasksuggestionjavascriptapioutlook14officetasksuggestion"></a><span data-ttu-id="d9e46-801">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_4/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_4/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_4/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_4/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="d9e46-801">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_4/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_4/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_4/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_4/office.tasksuggestion))>}</span></span>

<span data-ttu-id="d9e46-802">Возвращает известные сущности в выбранном элементе, которые проходят через именованный фильтр, определяемый в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="d9e46-802">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="d9e46-803">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="d9e46-803">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="d9e46-804">Метод `getFilteredEntitiesByName` возвращает сущности, соответствующие регулярному выражению, которое определяется в элементе правила [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) в XML-файле манифеста, с использованием указанного значения элемента `FilterName`.</span><span class="sxs-lookup"><span data-stu-id="d9e46-804">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d9e46-805">Параметры:</span><span class="sxs-lookup"><span data-stu-id="d9e46-805">Parameters:</span></span>

|<span data-ttu-id="d9e46-806">Имя</span><span class="sxs-lookup"><span data-stu-id="d9e46-806">Name</span></span>| <span data-ttu-id="d9e46-807">Тип</span><span class="sxs-lookup"><span data-stu-id="d9e46-807">Type</span></span>| <span data-ttu-id="d9e46-808">Описание</span><span class="sxs-lookup"><span data-stu-id="d9e46-808">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="d9e46-809">String</span><span class="sxs-lookup"><span data-stu-id="d9e46-809">String</span></span>|<span data-ttu-id="d9e46-810">Имя элемента правила `ItemHasKnownEntity`, определяющее соответствующий фильтр.</span><span class="sxs-lookup"><span data-stu-id="d9e46-810">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d9e46-811">Требования</span><span class="sxs-lookup"><span data-stu-id="d9e46-811">Requirements</span></span>

|<span data-ttu-id="d9e46-812">Требование</span><span class="sxs-lookup"><span data-stu-id="d9e46-812">Requirement</span></span>| <span data-ttu-id="d9e46-813">Значение</span><span class="sxs-lookup"><span data-stu-id="d9e46-813">Value</span></span>|
|---|---|
|[<span data-ttu-id="d9e46-814">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="d9e46-814">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d9e46-815">1.0</span><span class="sxs-lookup"><span data-stu-id="d9e46-815">1.0</span></span>|
|[<span data-ttu-id="d9e46-816">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="d9e46-816">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d9e46-817">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d9e46-817">ReadItem</span></span>|
|[<span data-ttu-id="d9e46-818">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="d9e46-818">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d9e46-819">Чтение</span><span class="sxs-lookup"><span data-stu-id="d9e46-819">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="d9e46-820">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="d9e46-820">Returns:</span></span>

<span data-ttu-id="d9e46-p153">Если в манифесте нет элемента `ItemHasKnownEntity` со значением `FilterName`, соответствующим параметру `name`, метод возвращает `null`. Если параметр `name` соответствует элементу `ItemHasKnownEntity` в манифесте, но при этом в текущем элементе нет соответствующих сущностей, метод возвращает пустой массив.</span><span class="sxs-lookup"><span data-stu-id="d9e46-p153">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="d9e46-823">Тип: Array.<(String|[Contact](/javascript/api/outlook_1_4/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_4/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_4/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_4/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="d9e46-823">Type: Array.<(String|[Contact](/javascript/api/outlook_1_4/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_4/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_4/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_4/office.tasksuggestion))></span></span>

#### <a name="getregexmatches--object"></a><span data-ttu-id="d9e46-824">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="d9e46-824">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="d9e46-825">Возвращает строчные значения в выбранном элементе, которые соответствуют регулярным выражениям, определенным в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="d9e46-825">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="d9e46-826">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="d9e46-826">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="d9e46-p154">Метод `getRegExMatches` возвращает строки, соответствующие регулярному выражению, которое определяется в каждом элементе правила `ItemHasRegularExpressionMatch` или `ItemHasKnownEntity` в XML-файле манифеста. Для правила `ItemHasRegularExpressionMatch` свойство элемента, указанного этим правилом, должно содержать соответствующую строку. Простой тип `PropertyName` определяет поддерживаемые свойства.</span><span class="sxs-lookup"><span data-stu-id="d9e46-p154">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="d9e46-830">Например, рассмотрим манифест надстройки, который содержит указанный ниже элемент `Rule`.</span><span class="sxs-lookup"><span data-stu-id="d9e46-830">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="d9e46-831">Объект, возвращаемый методом `getRegExMatches`, будет содержать два свойства: `fruits` и `veggies`.</span><span class="sxs-lookup"><span data-stu-id="d9e46-831">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="d9e46-p155">Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты. Лучше используйте для этого метод [`Body.getAsync`](/javascript/api/outlook_1_4/office.body#getasync-coerciontype--options--callback-).</span><span class="sxs-lookup"><span data-stu-id="d9e46-p155">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook_1_4/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="d9e46-835">Требования</span><span class="sxs-lookup"><span data-stu-id="d9e46-835">Requirements</span></span>

|<span data-ttu-id="d9e46-836">Требование</span><span class="sxs-lookup"><span data-stu-id="d9e46-836">Requirement</span></span>| <span data-ttu-id="d9e46-837">Значение</span><span class="sxs-lookup"><span data-stu-id="d9e46-837">Value</span></span>|
|---|---|
|[<span data-ttu-id="d9e46-838">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="d9e46-838">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d9e46-839">1.0</span><span class="sxs-lookup"><span data-stu-id="d9e46-839">1.0</span></span>|
|[<span data-ttu-id="d9e46-840">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="d9e46-840">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d9e46-841">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d9e46-841">ReadItem</span></span>|
|[<span data-ttu-id="d9e46-842">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="d9e46-842">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d9e46-843">Чтение</span><span class="sxs-lookup"><span data-stu-id="d9e46-843">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="d9e46-844">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="d9e46-844">Returns:</span></span>

<span data-ttu-id="d9e46-p156">Объект, содержащий массив строк, которые соответствуют регулярным выражениям, определяемым в XML-файле манифеста. Имя каждого массива равно соответствующему значению атрибута `RegExName` правила сопоставления `ItemHasRegularExpressionMatch` или атрибута `FilterName` правила сопоставления `ItemHasKnownEntity`.</span><span class="sxs-lookup"><span data-stu-id="d9e46-p156">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="d9e46-847">

<dt>Тип</dt>

</span><span class="sxs-lookup"><span data-stu-id="d9e46-847">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="d9e46-848">Объект</span><span class="sxs-lookup"><span data-stu-id="d9e46-848">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="d9e46-849">Пример</span><span class="sxs-lookup"><span data-stu-id="d9e46-849">Example</span></span>

<span data-ttu-id="d9e46-850">В примере ниже показано, как получить доступ к массиву совпадений для элементов `fruits` регулярного выражения<rule> и `veggies`, которые указаны в манифесте.</rule></span><span class="sxs-lookup"><span data-stu-id="d9e46-850">The following example shows how to access the array of matches for the regular expression <rule>elements `fruits` and `veggies`, which are specified in the manifest.</rule></span></span>

```
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veges = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="d9e46-851">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="d9e46-851">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="d9e46-852">Возвращает строчные значения в выбранном элементе, которые соответствуют именованному регулярному выражению, определенному в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="d9e46-852">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="d9e46-853">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="d9e46-853">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="d9e46-854">Метод `getRegExMatchesByName` возвращает строки, соответствующие регулярному выражению, которое определяется в элементе правила `ItemHasRegularExpressionMatch` в XML-файле манифеста, с использованием указанного значения элемента `RegExName`.</span><span class="sxs-lookup"><span data-stu-id="d9e46-854">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="d9e46-p157">Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты.</span><span class="sxs-lookup"><span data-stu-id="d9e46-p157">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d9e46-857">Параметры:</span><span class="sxs-lookup"><span data-stu-id="d9e46-857">Parameters:</span></span>

|<span data-ttu-id="d9e46-858">Имя</span><span class="sxs-lookup"><span data-stu-id="d9e46-858">Name</span></span>| <span data-ttu-id="d9e46-859">Тип</span><span class="sxs-lookup"><span data-stu-id="d9e46-859">Type</span></span>| <span data-ttu-id="d9e46-860">Описание</span><span class="sxs-lookup"><span data-stu-id="d9e46-860">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="d9e46-861">String</span><span class="sxs-lookup"><span data-stu-id="d9e46-861">String</span></span>|<span data-ttu-id="d9e46-862">Имя элемента правила `ItemHasRegularExpressionMatch`, определяющее соответствующий фильтр.</span><span class="sxs-lookup"><span data-stu-id="d9e46-862">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d9e46-863">Требования</span><span class="sxs-lookup"><span data-stu-id="d9e46-863">Requirements</span></span>

|<span data-ttu-id="d9e46-864">Требование</span><span class="sxs-lookup"><span data-stu-id="d9e46-864">Requirement</span></span>| <span data-ttu-id="d9e46-865">Значение</span><span class="sxs-lookup"><span data-stu-id="d9e46-865">Value</span></span>|
|---|---|
|[<span data-ttu-id="d9e46-866">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="d9e46-866">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d9e46-867">1.0</span><span class="sxs-lookup"><span data-stu-id="d9e46-867">1.0</span></span>|
|[<span data-ttu-id="d9e46-868">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="d9e46-868">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d9e46-869">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d9e46-869">ReadItem</span></span>|
|[<span data-ttu-id="d9e46-870">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="d9e46-870">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d9e46-871">Чтение</span><span class="sxs-lookup"><span data-stu-id="d9e46-871">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="d9e46-872">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="d9e46-872">Returns:</span></span>

<span data-ttu-id="d9e46-873">Массив строк, соответствующих регулярному выражению, определяемому в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="d9e46-873">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="d9e46-874">

<dt>Тип</dt>

</span><span class="sxs-lookup"><span data-stu-id="d9e46-874">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="d9e46-875">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="d9e46-875">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="d9e46-876">Пример</span><span class="sxs-lookup"><span data-stu-id="d9e46-876">Example</span></span>

```
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="d9e46-877">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="d9e46-877">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="d9e46-878">Асинхронно возвращает данные, выбранные в теме или тексте сообщения.</span><span class="sxs-lookup"><span data-stu-id="d9e46-878">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="d9e46-p158">Если выделенный фрагмент отсутствует, но курсор находится в тексте или теме, метод возвращает значение NULL для выбранных данных. Если выбраны не текст и не тема, метод возвращает ошибку `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="d9e46-p158">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d9e46-881">Параметры:</span><span class="sxs-lookup"><span data-stu-id="d9e46-881">Parameters:</span></span>

|<span data-ttu-id="d9e46-882">Имя</span><span class="sxs-lookup"><span data-stu-id="d9e46-882">Name</span></span>| <span data-ttu-id="d9e46-883">Тип</span><span class="sxs-lookup"><span data-stu-id="d9e46-883">Type</span></span>| <span data-ttu-id="d9e46-884">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="d9e46-884">Attributes</span></span>| <span data-ttu-id="d9e46-885">Описание</span><span class="sxs-lookup"><span data-stu-id="d9e46-885">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="d9e46-886">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="d9e46-886">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="d9e46-p159">Запрашивает формат данных. Если задано значение Text, метод возвращает обычный текст как строку, удаляя все имеющиеся HTML-теги. Если задано значение HTML, метод возвращает выделенный текст (обычный текст или HTML).</span><span class="sxs-lookup"><span data-stu-id="d9e46-p159">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="d9e46-890">Объект</span><span class="sxs-lookup"><span data-stu-id="d9e46-890">Object</span></span>| <span data-ttu-id="d9e46-891">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="d9e46-891">&lt;optional&gt;</span></span>|<span data-ttu-id="d9e46-892">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="d9e46-892">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="d9e46-893">Объект</span><span class="sxs-lookup"><span data-stu-id="d9e46-893">Object</span></span>| <span data-ttu-id="d9e46-894">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="d9e46-894">&lt;optional&gt;</span></span>|<span data-ttu-id="d9e46-895">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="d9e46-895">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="d9e46-896">функция</span><span class="sxs-lookup"><span data-stu-id="d9e46-896">function</span></span>||<span data-ttu-id="d9e46-897">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="d9e46-897">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="d9e46-p160">Чтобы получить доступ к выделенным данным из метода обратного вызова, вызовите `asyncResult.value.data`. Для доступа к исходному свойству, на основе которого созданы выбранные данные, вызовите `asyncResult.value.sourceProperty`, который может иметь значение `body` или `subject`.</span><span class="sxs-lookup"><span data-stu-id="d9e46-p160">To access the selected data from the callback method, call `asyncResult.value.data`. To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d9e46-900">Требования</span><span class="sxs-lookup"><span data-stu-id="d9e46-900">Requirements</span></span>

|<span data-ttu-id="d9e46-901">Требование</span><span class="sxs-lookup"><span data-stu-id="d9e46-901">Requirement</span></span>| <span data-ttu-id="d9e46-902">Значение</span><span class="sxs-lookup"><span data-stu-id="d9e46-902">Value</span></span>|
|---|---|
|[<span data-ttu-id="d9e46-903">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="d9e46-903">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d9e46-904">1.2</span><span class="sxs-lookup"><span data-stu-id="d9e46-904">1.2</span></span>|
|[<span data-ttu-id="d9e46-905">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="d9e46-905">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d9e46-906">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="d9e46-906">ReadWriteItem</span></span>|
|[<span data-ttu-id="d9e46-907">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="d9e46-907">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d9e46-908">Создание</span><span class="sxs-lookup"><span data-stu-id="d9e46-908">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="d9e46-909">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="d9e46-909">Returns:</span></span>

<span data-ttu-id="d9e46-910">Выбранные данные в виде строки с форматом, определенным в параметре `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="d9e46-910">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="d9e46-911">

<dt>Тип</dt>

</span><span class="sxs-lookup"><span data-stu-id="d9e46-911">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="d9e46-912">String</span><span class="sxs-lookup"><span data-stu-id="d9e46-912">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="d9e46-913">Пример</span><span class="sxs-lookup"><span data-stu-id="d9e46-913">Example</span></span>

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

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="d9e46-914">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="d9e46-914">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="d9e46-915">Асинхронно загружает настраиваемые свойства для надстройки выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="d9e46-915">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="d9e46-p161">Настраиваемые свойства сохраняются в виде пар "ключ-значение" для каждого приложения и каждого элемента. Этот метод возвращает объект `CustomProperties` при обратном вызове, который предоставляет методы для доступа к настраиваемым свойствам, характерным для текущего элемента и текущей надстройки. Настраиваемые свойства не шифруются для элемента, поэтому этот способ хранения не является безопасным.</span><span class="sxs-lookup"><span data-stu-id="d9e46-p161">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d9e46-919">Параметры:</span><span class="sxs-lookup"><span data-stu-id="d9e46-919">Parameters:</span></span>

|<span data-ttu-id="d9e46-920">Имя</span><span class="sxs-lookup"><span data-stu-id="d9e46-920">Name</span></span>| <span data-ttu-id="d9e46-921">Тип</span><span class="sxs-lookup"><span data-stu-id="d9e46-921">Type</span></span>| <span data-ttu-id="d9e46-922">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="d9e46-922">Attributes</span></span>| <span data-ttu-id="d9e46-923">Описание</span><span class="sxs-lookup"><span data-stu-id="d9e46-923">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="d9e46-924">function</span><span class="sxs-lookup"><span data-stu-id="d9e46-924">function</span></span>||<span data-ttu-id="d9e46-925">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="d9e46-925">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="d9e46-p162">Настраиваемые свойства предоставляются в виде объекта [`CustomProperties`](/javascript/api/outlook_1_4/office.customproperties) в свойстве `asyncResult.value`. Этот объект позволяет получить, задать и удалить настраиваемые свойства из элемента, а также сохранить изменения, внесенные в настраиваемое свойство, на сервере.</span><span class="sxs-lookup"><span data-stu-id="d9e46-p162">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook_1_4/office.customproperties) object in the `asyncResult.value` property. This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="d9e46-928">Объект</span><span class="sxs-lookup"><span data-stu-id="d9e46-928">Object</span></span>| <span data-ttu-id="d9e46-929">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="d9e46-929">&lt;optional&gt;</span></span>|<span data-ttu-id="d9e46-p163">Разработчики могут указать любой объект, к которому необходимо получить доступ, в функции обратного вызова. Доступ к этому объекту можно получить с помощью свойства `asyncResult.asyncContext` в функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="d9e46-p163">Developers can provide any object they wish to access in the callback function. This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d9e46-932">Требования</span><span class="sxs-lookup"><span data-stu-id="d9e46-932">Requirements</span></span>

|<span data-ttu-id="d9e46-933">Требование</span><span class="sxs-lookup"><span data-stu-id="d9e46-933">Requirement</span></span>| <span data-ttu-id="d9e46-934">Значение</span><span class="sxs-lookup"><span data-stu-id="d9e46-934">Value</span></span>|
|---|---|
|[<span data-ttu-id="d9e46-935">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="d9e46-935">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d9e46-936">1.0</span><span class="sxs-lookup"><span data-stu-id="d9e46-936">1.0</span></span>|
|[<span data-ttu-id="d9e46-937">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="d9e46-937">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d9e46-938">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d9e46-938">ReadItem</span></span>|
|[<span data-ttu-id="d9e46-939">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="d9e46-939">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d9e46-940">Cоздание или чтение</span><span class="sxs-lookup"><span data-stu-id="d9e46-940">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="d9e46-941">Пример</span><span class="sxs-lookup"><span data-stu-id="d9e46-941">Example</span></span>

<span data-ttu-id="d9e46-p164">В приведенном ниже примере кода показано, как асинхронно загружать настраиваемые свойства, характерные для текущего элемента, с помощью метода `loadCustomPropertiesAsync`. Этот пример также показывает, как сохранять эти свойства на сервере с помощью метода `CustomProperties.saveAsync`. В этом примере кода, после того как выполнена загрузка настраиваемых свойств, метод `CustomProperties.get` используется для считывания настраиваемого свойства `myProp`, метод `CustomProperties.set` — для записи настраиваемого свойства `otherProp`, а метод `saveAsync` — для сохранения настраиваемых свойств.</span><span class="sxs-lookup"><span data-stu-id="d9e46-p164">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="d9e46-945">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="d9e46-945">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="d9e46-946">Удаляет вложение из сообщения или встречи.</span><span class="sxs-lookup"><span data-stu-id="d9e46-946">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="d9e46-p165">Метод `removeAttachmentAsync` удаляет из элемента вложение с указанным идентификатором. Идентификатор вложения рекомендуется использовать для удаления вложения, только если оно добавлено тем же почтовым приложением в ходе текущего сеанса. В веб-приложении Outlook и веб-приложении Outlook для устройств идентификатор вложения действителен только в рамках одного сеанса. Сеанс завершается, когда пользователь закрывает приложение или начинает создавать элемент во встроенной форме, а затем переходит из формы в отдельное окно.</span><span class="sxs-lookup"><span data-stu-id="d9e46-p165">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d9e46-951">Параметры:</span><span class="sxs-lookup"><span data-stu-id="d9e46-951">Parameters:</span></span>

|<span data-ttu-id="d9e46-952">Имя</span><span class="sxs-lookup"><span data-stu-id="d9e46-952">Name</span></span>| <span data-ttu-id="d9e46-953">Тип</span><span class="sxs-lookup"><span data-stu-id="d9e46-953">Type</span></span>| <span data-ttu-id="d9e46-954">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="d9e46-954">Attributes</span></span>| <span data-ttu-id="d9e46-955">Описание</span><span class="sxs-lookup"><span data-stu-id="d9e46-955">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="d9e46-956">String</span><span class="sxs-lookup"><span data-stu-id="d9e46-956">String</span></span>||<span data-ttu-id="d9e46-p166">Идентификатор удаляемого вложения. Максимальная длина строки — 100 символов.</span><span class="sxs-lookup"><span data-stu-id="d9e46-p166">The identifier of the attachment to remove. The maximum length of the string is 100 characters.</span></span>|
|`options`| <span data-ttu-id="d9e46-959">Объект</span><span class="sxs-lookup"><span data-stu-id="d9e46-959">Object</span></span>| <span data-ttu-id="d9e46-960">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d9e46-960">&lt;optional&gt;</span></span>|<span data-ttu-id="d9e46-961">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="d9e46-961">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="d9e46-962">Объект</span><span class="sxs-lookup"><span data-stu-id="d9e46-962">Object</span></span>| <span data-ttu-id="d9e46-963">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="d9e46-963">&lt;optional&gt;</span></span>|<span data-ttu-id="d9e46-964">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="d9e46-964">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="d9e46-965">function</span><span class="sxs-lookup"><span data-stu-id="d9e46-965">function</span></span>| <span data-ttu-id="d9e46-966">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="d9e46-966">&lt;optional&gt;</span></span>|<span data-ttu-id="d9e46-967">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="d9e46-967">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="d9e46-968">Если удалить вложение не удается, свойство `asyncResult.error` содержит код ошибки с указанием ее причины.</span><span class="sxs-lookup"><span data-stu-id="d9e46-968">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="d9e46-969">Ошибки</span><span class="sxs-lookup"><span data-stu-id="d9e46-969">Errors</span></span>

| <span data-ttu-id="d9e46-970">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="d9e46-970">Error code</span></span> | <span data-ttu-id="d9e46-971">Описание</span><span class="sxs-lookup"><span data-stu-id="d9e46-971">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="d9e46-972">Идентификатор вложения не существует.</span><span class="sxs-lookup"><span data-stu-id="d9e46-972">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="d9e46-973">Требования</span><span class="sxs-lookup"><span data-stu-id="d9e46-973">Requirements</span></span>

|<span data-ttu-id="d9e46-974">Требование</span><span class="sxs-lookup"><span data-stu-id="d9e46-974">Requirement</span></span>| <span data-ttu-id="d9e46-975">Значение</span><span class="sxs-lookup"><span data-stu-id="d9e46-975">Value</span></span>|
|---|---|
|[<span data-ttu-id="d9e46-976">Версия минимального набора требований для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="d9e46-976">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d9e46-977">1.1</span><span class="sxs-lookup"><span data-stu-id="d9e46-977">1.1</span></span>|
|[<span data-ttu-id="d9e46-978">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="d9e46-978">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d9e46-979">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="d9e46-979">ReadWriteItem</span></span>|
|[<span data-ttu-id="d9e46-980">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="d9e46-980">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d9e46-981">Создание</span><span class="sxs-lookup"><span data-stu-id="d9e46-981">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="d9e46-982">Пример</span><span class="sxs-lookup"><span data-stu-id="d9e46-982">Example</span></span>

<span data-ttu-id="d9e46-983">Указанный ниже код удаляет вложение с идентификатором "0".</span><span class="sxs-lookup"><span data-stu-id="d9e46-983">The following code removes an attachment with an identifier of '0'.</span></span>

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

####  <a name="saveasyncoptions-callback"></a><span data-ttu-id="d9e46-984">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="d9e46-984">saveAsync([options], callback)</span></span>

<span data-ttu-id="d9e46-985">Асинхронно сохраняет элемент.</span><span class="sxs-lookup"><span data-stu-id="d9e46-985">Asynchronously saves an item.</span></span>

<span data-ttu-id="d9e46-p167">При вызове этот метод сохраняет текущее сообщение в виде черновика и возвращает идентификатор элемента с помощью метода обратного вызова. В веб-приложении Outlook или интерактивном режиме Outlook этот элемент сохраняется на сервере. В Outlook в режиме кэширования этот элемент сохраняется в локальном кэше.</span><span class="sxs-lookup"><span data-stu-id="d9e46-p167">When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="d9e46-p168">Если в режиме создания надстройка вызывает для элемента метод `saveAsync`, чтобы получить параметр `itemId` для использования с EWS или REST API, необходимо помнить, что синхронизация элемента с сервером может занять много времени, если Outlook работает в режиме кэширования данных. До окончания синхронизации применение параметра `itemId`  будет приводить к ошибке.</span><span class="sxs-lookup"><span data-stu-id="d9e46-p168">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server. Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="d9e46-p169">Так как для встреч не предусмотрено состояние черновика, если `saveAsync` вызывается для встречи в режиме создания, она сохраняется как обычная встреча в календаре пользователя. При сохранении новой встречи приглашения не отправляются. При сохранении существующей встречи уведомления отправляются добавленным или удаленным участникам.</span><span class="sxs-lookup"><span data-stu-id="d9e46-p169">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="d9e46-994">Следующие клиенты имеют разную реакцию на событие для `saveAsync` для встреч в режиме создания:</span><span class="sxs-lookup"><span data-stu-id="d9e46-994">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="d9e46-p170">Outlook для Mac не поддерживает `saveAsync` для собраний в режиме создания. Метод `saveAsync`, вызванный для собрания в Outlook для Mac, возвращает ошибку.</span><span class="sxs-lookup"><span data-stu-id="d9e46-p170">Mac Outlook does not support `saveAsync` on a meeting in compose mode. Calling `saveAsync` on a meeting in Mac Outlook will return an error.</span></span>
> - <span data-ttu-id="d9e46-997">Outlook в Интернете всегда отправляет приглашение или обновления при вызове `saveAsync` на встрече в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="d9e46-997">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d9e46-998">Параметры:</span><span class="sxs-lookup"><span data-stu-id="d9e46-998">Parameters:</span></span>

|<span data-ttu-id="d9e46-999">Имя</span><span class="sxs-lookup"><span data-stu-id="d9e46-999">Name</span></span>| <span data-ttu-id="d9e46-1000">Тип</span><span class="sxs-lookup"><span data-stu-id="d9e46-1000">Type</span></span>| <span data-ttu-id="d9e46-1001">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="d9e46-1001">Attributes</span></span>| <span data-ttu-id="d9e46-1002">Описание</span><span class="sxs-lookup"><span data-stu-id="d9e46-1002">Description</span></span>|
|---|---|---|---|
|`options`| <span data-ttu-id="d9e46-1003">Объект</span><span class="sxs-lookup"><span data-stu-id="d9e46-1003">Object</span></span>| <span data-ttu-id="d9e46-1004">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="d9e46-1004">&lt;optional&gt;</span></span>|<span data-ttu-id="d9e46-1005">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="d9e46-1005">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="d9e46-1006">Объект</span><span class="sxs-lookup"><span data-stu-id="d9e46-1006">Object</span></span>| <span data-ttu-id="d9e46-1007">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="d9e46-1007">&lt;optional&gt;</span></span>|<span data-ttu-id="d9e46-1008">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="d9e46-1008">Developers can provide any object they wish to access in the callback method.</span></span>||
|`callback`| <span data-ttu-id="d9e46-1009">функция</span><span class="sxs-lookup"><span data-stu-id="d9e46-1009">function</span></span>||<span data-ttu-id="d9e46-1010">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="d9e46-1010">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="d9e46-1011">После успешного выполнения идентификатор элемента будет предоставлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="d9e46-1011">On success, the item identifier is provided in the `asyncResult.value` property.|</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d9e46-1012">Требования</span><span class="sxs-lookup"><span data-stu-id="d9e46-1012">Requirements</span></span>

|<span data-ttu-id="d9e46-1013">Требование</span><span class="sxs-lookup"><span data-stu-id="d9e46-1013">Requirement</span></span>| <span data-ttu-id="d9e46-1014">Значение</span><span class="sxs-lookup"><span data-stu-id="d9e46-1014">Value</span></span>|
|---|---|
|[<span data-ttu-id="d9e46-1015">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="d9e46-1015">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d9e46-1016">1.3</span><span class="sxs-lookup"><span data-stu-id="d9e46-1016">1.3</span></span>|
|[<span data-ttu-id="d9e46-1017">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="d9e46-1017">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d9e46-1018">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="d9e46-1018">ReadWriteItem</span></span>|
|[<span data-ttu-id="d9e46-1019">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="d9e46-1019">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d9e46-1020">Создание</span><span class="sxs-lookup"><span data-stu-id="d9e46-1020">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="d9e46-1021">Примеры</span><span class="sxs-lookup"><span data-stu-id="d9e46-1021">Examples</span></span>

```
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result
  });
```

<span data-ttu-id="d9e46-p171">Ниже приведен пример параметра `result`, переданного функции обратного вызова. Свойство `value` содержит идентификатор элемента.</span><span class="sxs-lookup"><span data-stu-id="d9e46-p171">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

####  <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="d9e46-1024">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="d9e46-1024">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="d9e46-1025">Асинхронно вставляет данные в текст или тему сообщения.</span><span class="sxs-lookup"><span data-stu-id="d9e46-1025">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="d9e46-p172">Метод `setSelectedDataAsync` вставляет указанную строку в местоположение курсора в теме или тексте элемента либо, если текст выделен в редакторе, он заменяет выделенный текст. Если курсор находится вне текста или темы элемента, возвращается ошибка. После вставки курсор помещается в конец вставленного содержимого.</span><span class="sxs-lookup"><span data-stu-id="d9e46-p172">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d9e46-1029">Параметры:</span><span class="sxs-lookup"><span data-stu-id="d9e46-1029">Parameters:</span></span>

|<span data-ttu-id="d9e46-1030">Имя</span><span class="sxs-lookup"><span data-stu-id="d9e46-1030">Name</span></span>| <span data-ttu-id="d9e46-1031">Тип</span><span class="sxs-lookup"><span data-stu-id="d9e46-1031">Type</span></span>| <span data-ttu-id="d9e46-1032">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="d9e46-1032">Attributes</span></span>| <span data-ttu-id="d9e46-1033">Описание</span><span class="sxs-lookup"><span data-stu-id="d9e46-1033">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="d9e46-1034">String</span><span class="sxs-lookup"><span data-stu-id="d9e46-1034">String</span></span>||<span data-ttu-id="d9e46-p173">Вставляемые данные. Объем данных не должен превышать 1 000 000 символов. Если передано больше 1 000 000 символов, возвращается исключение `ArgumentOutOfRange`.</span><span class="sxs-lookup"><span data-stu-id="d9e46-p173">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="d9e46-1038">Объект</span><span class="sxs-lookup"><span data-stu-id="d9e46-1038">Object</span></span>| <span data-ttu-id="d9e46-1039">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d9e46-1039">&lt;optional&gt;</span></span>|<span data-ttu-id="d9e46-1040">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="d9e46-1040">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="d9e46-1041">Объект</span><span class="sxs-lookup"><span data-stu-id="d9e46-1041">Object</span></span>| <span data-ttu-id="d9e46-1042">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="d9e46-1042">&lt;optional&gt;</span></span>|<span data-ttu-id="d9e46-1043">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="d9e46-1043">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`| [<span data-ttu-id="d9e46-1044">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="d9e46-1044">Office.CoercionType</span></span>](office.md#coerciontype-string)| <span data-ttu-id="d9e46-1045">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="d9e46-1045">&lt;optional&gt;</span></span>|<span data-ttu-id="d9e46-p174">Если задано значение `text`, текущий стиль применяется в Outlook и веб-приложении Outlook. Если поле представляет собой редактор HTML, вставляются только текстовые данные, даже если они имеют формат HTML.</span><span class="sxs-lookup"><span data-stu-id="d9e46-p174">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="d9e46-p175">Если `html` и поле поддерживают HTML (а тема не поддерживает), в веб-приложении Outlook применяется текущий стиль, а в Outlook — стиль по умолчанию. Если поле является текстовым, возвращается ошибка `InvalidDataFormat`.</span><span class="sxs-lookup"><span data-stu-id="d9e46-p175">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="d9e46-1050">Если тип `coercionType` не установлен, результат зависит от поля: если поле имеет формат HTML, то используется HTML; если поле является текстовым, то используется обычный текст.</span><span class="sxs-lookup"><span data-stu-id="d9e46-1050">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="d9e46-1051">function</span><span class="sxs-lookup"><span data-stu-id="d9e46-1051">function</span></span>||<span data-ttu-id="d9e46-1052">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="d9e46-1052">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="d9e46-1053">Требования</span><span class="sxs-lookup"><span data-stu-id="d9e46-1053">Requirements</span></span>

|<span data-ttu-id="d9e46-1054">Требование</span><span class="sxs-lookup"><span data-stu-id="d9e46-1054">Requirement</span></span>| <span data-ttu-id="d9e46-1055">Значение</span><span class="sxs-lookup"><span data-stu-id="d9e46-1055">Value</span></span>|
|---|---|
|[<span data-ttu-id="d9e46-1056">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="d9e46-1056">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d9e46-1057">1.2</span><span class="sxs-lookup"><span data-stu-id="d9e46-1057">1.2</span></span>|
|[<span data-ttu-id="d9e46-1058">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="d9e46-1058">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d9e46-1059">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="d9e46-1059">ReadWriteItem</span></span>|
|[<span data-ttu-id="d9e46-1060">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="d9e46-1060">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d9e46-1061">Создание</span><span class="sxs-lookup"><span data-stu-id="d9e46-1061">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="d9e46-1062">Пример</span><span class="sxs-lookup"><span data-stu-id="d9e46-1062">Example</span></span>

```
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```