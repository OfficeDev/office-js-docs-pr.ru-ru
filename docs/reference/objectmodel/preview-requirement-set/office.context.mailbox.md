
# <a name="mailbox"></a><span data-ttu-id="d0b52-101">mailbox</span><span class="sxs-lookup"><span data-stu-id="d0b52-101">mailbox</span></span>

### <span data-ttu-id="d0b52-p101">[Office](Office.md)[.context](Office.context.md). mailbox</span><span class="sxs-lookup"><span data-stu-id="d0b52-p101">[Office](Office.md)[.context](Office.context.md). mailbox</span></span>

<span data-ttu-id="d0b52-104">Предоставляет для Microsoft Outlook и Microsoft Outlook в Интернете доступ к объектной модели надстройки Outlook.</span><span class="sxs-lookup"><span data-stu-id="d0b52-104">Provides access to the Outlook Add-in object model for Microsoft Outlook and Microsoft Outlook on the web.</span></span>

##### <a name="requirements"></a><span data-ttu-id="d0b52-105">Требования</span><span class="sxs-lookup"><span data-stu-id="d0b52-105">Requirements</span></span>

|<span data-ttu-id="d0b52-106">Требование</span><span class="sxs-lookup"><span data-stu-id="d0b52-106">Requirement</span></span>| <span data-ttu-id="d0b52-107">Значение</span><span class="sxs-lookup"><span data-stu-id="d0b52-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="d0b52-108">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="d0b52-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d0b52-109">1.0</span><span class="sxs-lookup"><span data-stu-id="d0b52-109">1.0</span></span>|
|[<span data-ttu-id="d0b52-110">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="d0b52-110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d0b52-111">Ограниченный доступ</span><span class="sxs-lookup"><span data-stu-id="d0b52-111">Restricted</span></span>|
|[<span data-ttu-id="d0b52-112">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="d0b52-112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d0b52-113">Compose или read</span><span class="sxs-lookup"><span data-stu-id="d0b52-113">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="d0b52-114">Члены и методы</span><span class="sxs-lookup"><span data-stu-id="d0b52-114">Members and methods</span></span>

| <span data-ttu-id="d0b52-115">Член</span><span class="sxs-lookup"><span data-stu-id="d0b52-115">Member</span></span> | <span data-ttu-id="d0b52-116">Тип</span><span class="sxs-lookup"><span data-stu-id="d0b52-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="d0b52-117">ewsUrl</span><span class="sxs-lookup"><span data-stu-id="d0b52-117">ewsUrl</span></span>](#ewsurl-string) | <span data-ttu-id="d0b52-118">Член</span><span class="sxs-lookup"><span data-stu-id="d0b52-118">Member</span></span> |
| [<span data-ttu-id="d0b52-119">restUrl</span><span class="sxs-lookup"><span data-stu-id="d0b52-119">restUrl</span></span>](#resturl-string) | <span data-ttu-id="d0b52-120">Член</span><span class="sxs-lookup"><span data-stu-id="d0b52-120">Member</span></span> |
| [<span data-ttu-id="d0b52-121">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="d0b52-121">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="d0b52-122">Метод</span><span class="sxs-lookup"><span data-stu-id="d0b52-122">Method</span></span> |
| [<span data-ttu-id="d0b52-123">convertToEwsId</span><span class="sxs-lookup"><span data-stu-id="d0b52-123">convertToEwsId</span></span>](#converttoewsiditemid-restversion--string) | <span data-ttu-id="d0b52-124">Метод</span><span class="sxs-lookup"><span data-stu-id="d0b52-124">Method</span></span> |
| [<span data-ttu-id="d0b52-125">convertToLocalClientTime</span><span class="sxs-lookup"><span data-stu-id="d0b52-125">convertToLocalClientTime</span></span>](#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime) | <span data-ttu-id="d0b52-126">Метод</span><span class="sxs-lookup"><span data-stu-id="d0b52-126">Method</span></span> |
| [<span data-ttu-id="d0b52-127">convertToRestId</span><span class="sxs-lookup"><span data-stu-id="d0b52-127">convertToRestId</span></span>](#converttorestiditemid-restversion--string) | <span data-ttu-id="d0b52-128">Метод</span><span class="sxs-lookup"><span data-stu-id="d0b52-128">Method</span></span> |
| [<span data-ttu-id="d0b52-129">convertToUtcClientTime</span><span class="sxs-lookup"><span data-stu-id="d0b52-129">convertToUtcClientTime</span></span>](#converttoutcclienttimeinput--date) | <span data-ttu-id="d0b52-130">Метод</span><span class="sxs-lookup"><span data-stu-id="d0b52-130">Method</span></span> |
| [<span data-ttu-id="d0b52-131">displayAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="d0b52-131">displayAppointmentForm</span></span>](#displayappointmentformitemid) | <span data-ttu-id="d0b52-132">Метод</span><span class="sxs-lookup"><span data-stu-id="d0b52-132">Method</span></span> |
| [<span data-ttu-id="d0b52-133">displayMessageForm</span><span class="sxs-lookup"><span data-stu-id="d0b52-133">displayMessageForm</span></span>](#displaymessageformitemid) | <span data-ttu-id="d0b52-134">Метод</span><span class="sxs-lookup"><span data-stu-id="d0b52-134">Method</span></span> |
| [<span data-ttu-id="d0b52-135">displayNewAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="d0b52-135">displayNewAppointmentForm</span></span>](#displaynewappointmentformparameters) | <span data-ttu-id="d0b52-136">Метод</span><span class="sxs-lookup"><span data-stu-id="d0b52-136">Method</span></span> |
| [<span data-ttu-id="d0b52-137">displayNewMessageForm</span><span class="sxs-lookup"><span data-stu-id="d0b52-137">displayNewMessageForm</span></span>](#displaynewmessageformparameters) | <span data-ttu-id="d0b52-138">Метод</span><span class="sxs-lookup"><span data-stu-id="d0b52-138">Method</span></span> |
| [<span data-ttu-id="d0b52-139">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="d0b52-139">getCallbackTokenAsync</span></span>](#getcallbacktokenasyncoptions-callback) | <span data-ttu-id="d0b52-140">Метод</span><span class="sxs-lookup"><span data-stu-id="d0b52-140">Method</span></span> |
| [<span data-ttu-id="d0b52-141">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="d0b52-141">getCallbackTokenAsync</span></span>](#getcallbacktokenasynccallback-usercontext) | <span data-ttu-id="d0b52-142">Метод</span><span class="sxs-lookup"><span data-stu-id="d0b52-142">Method</span></span> |
| [<span data-ttu-id="d0b52-143">getUserIdentityTokenAsync</span><span class="sxs-lookup"><span data-stu-id="d0b52-143">getUserIdentityTokenAsync</span></span>](#getuseridentitytokenasynccallback-usercontext) | <span data-ttu-id="d0b52-144">Метод</span><span class="sxs-lookup"><span data-stu-id="d0b52-144">Method</span></span> |
| [<span data-ttu-id="d0b52-145">makeEwsRequestAsync</span><span class="sxs-lookup"><span data-stu-id="d0b52-145">makeEwsRequestAsync</span></span>](#makeewsrequestasyncdata-callback-usercontext) | <span data-ttu-id="d0b52-146">Метод</span><span class="sxs-lookup"><span data-stu-id="d0b52-146">Method</span></span> |

### <a name="namespaces"></a><span data-ttu-id="d0b52-147">Пространства имен</span><span class="sxs-lookup"><span data-stu-id="d0b52-147">Namespaces</span></span>

<span data-ttu-id="d0b52-148">[diagnostics](Office.context.mailbox.diagnostics.md): предоставляет надстройке Outlook диагностические сведения.</span><span class="sxs-lookup"><span data-stu-id="d0b52-148">[diagnostics](Office.context.mailbox.diagnostics.md): Provides diagnostic information to an Outlook add-in.</span></span>

<span data-ttu-id="d0b52-149">[item](Office.context.mailbox.item.md). Предоставляет методы и свойства для доступа к сообщению или встрече в надстройке Outlook.</span><span class="sxs-lookup"><span data-stu-id="d0b52-149">[item](Office.context.mailbox.item.md): Provides methods and properties for accessing a message or appointment in an Outlook add-in.</span></span>

<span data-ttu-id="d0b52-150">[userProfile](Office.context.mailbox.userProfile.md): предоставляет сведения о пользователе в надстройке Outlook.</span><span class="sxs-lookup"><span data-stu-id="d0b52-150">[userProfile](Office.context.mailbox.userProfile.md): Provides information about the user in an Outlook add-in.</span></span>

### <a name="members"></a><span data-ttu-id="d0b52-151">Члены</span><span class="sxs-lookup"><span data-stu-id="d0b52-151">Members</span></span>

#### <a name="ewsurl-string"></a><span data-ttu-id="d0b52-152">ewsUrl :String</span><span class="sxs-lookup"><span data-stu-id="d0b52-152">ewsUrl :String</span></span>

<span data-ttu-id="d0b52-p102">Получает URL-адрес конечной точки веб-служб Exchange (EWS) для конкретной учетной записи электронной почты. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="d0b52-p102">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="d0b52-155">Этот элемент не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="d0b52-155">Note: This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="d0b52-p103">Удаленная служба может использовать значение `ewsUrl`, чтобы выполнять вызовы EWS для почтового ящика пользователя. Например, вы можете создать удаленную службу, чтобы [получить вложения из выбранного элемента](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="d0b52-p103">The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox. For example, you can create a remote service to [get attachments from the selected item](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="d0b52-158">В манифесте приложения должно быть указано разрешение **ReadItem** для вызова члена `ewsUrl` в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="d0b52-158">Your app must have the **ReadItem** permission specified in its manifest to call the `ewsUrl` member in read mode.</span></span>

<span data-ttu-id="d0b52-p104">Перед использованием члена `ewsUrl` в режиме создания необходимо вызвать метод [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback). Для вызова метода `saveAsync` приложение должно иметь разрешения **ReadWriteItem**.</span><span class="sxs-lookup"><span data-stu-id="d0b52-p104">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `ewsUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="d0b52-161">Тип:</span><span class="sxs-lookup"><span data-stu-id="d0b52-161">Type:</span></span>

*   <span data-ttu-id="d0b52-162">String</span><span class="sxs-lookup"><span data-stu-id="d0b52-162">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="d0b52-163">Требования</span><span class="sxs-lookup"><span data-stu-id="d0b52-163">Requirements</span></span>

|<span data-ttu-id="d0b52-164">Требование</span><span class="sxs-lookup"><span data-stu-id="d0b52-164">Requirement</span></span>| <span data-ttu-id="d0b52-165">Значение</span><span class="sxs-lookup"><span data-stu-id="d0b52-165">Value</span></span>|
|---|---|
|[<span data-ttu-id="d0b52-166">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="d0b52-166">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d0b52-167">1.0</span><span class="sxs-lookup"><span data-stu-id="d0b52-167">1.0</span></span>|
|[<span data-ttu-id="d0b52-168">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="d0b52-168">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d0b52-169">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d0b52-169">ReadItem</span></span>|
|[<span data-ttu-id="d0b52-170">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="d0b52-170">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d0b52-171">Создание или чтение​</span><span class="sxs-lookup"><span data-stu-id="d0b52-171">Compose or read</span></span>|

#### <a name="resturl-string"></a><span data-ttu-id="d0b52-172">restUrl :String</span><span class="sxs-lookup"><span data-stu-id="d0b52-172">restUrl :String</span></span>

<span data-ttu-id="d0b52-173">Возвращает URL-адрес конечной точки REST для этой учетной записи электронной почты.</span><span class="sxs-lookup"><span data-stu-id="d0b52-173">Gets the URL of the REST endpoint for this email account.</span></span>

<span data-ttu-id="d0b52-174">С помощью значения `restUrl` можно выполнять вызовы [REST API](https://docs.microsoft.com/outlook/rest/) для почтового ящика пользователя.</span><span class="sxs-lookup"><span data-stu-id="d0b52-174">The `restUrl` value can be used to make [REST API](https://docs.microsoft.com/outlook/rest/) calls to the user's mailbox.</span></span>

<span data-ttu-id="d0b52-175">В манифесте приложения должно быть указано разрешение **ReadItem** для вызова члена `restUrl` в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="d0b52-175">Your app must have the **ReadItem** permission specified in its manifest to call the `restUrl` member in read mode.</span></span>

<span data-ttu-id="d0b52-p105">Перед использованием члена `restUrl` в режиме создания необходимо вызвать метод [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback). Для вызова метода `saveAsync` приложение должно иметь разрешения **ReadWriteItem**.</span><span class="sxs-lookup"><span data-stu-id="d0b52-p105">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `restUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="d0b52-178">Тип:</span><span class="sxs-lookup"><span data-stu-id="d0b52-178">Type:</span></span>

*   <span data-ttu-id="d0b52-179">String</span><span class="sxs-lookup"><span data-stu-id="d0b52-179">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="d0b52-180">Требования</span><span class="sxs-lookup"><span data-stu-id="d0b52-180">Requirements</span></span>

|<span data-ttu-id="d0b52-181">Требование</span><span class="sxs-lookup"><span data-stu-id="d0b52-181">Requirement</span></span>| <span data-ttu-id="d0b52-182">Значение</span><span class="sxs-lookup"><span data-stu-id="d0b52-182">Value</span></span>|
|---|---|
|[<span data-ttu-id="d0b52-183">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="d0b52-183">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d0b52-184">1.5</span><span class="sxs-lookup"><span data-stu-id="d0b52-184">1.5</span></span> |
|[<span data-ttu-id="d0b52-185">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="d0b52-185">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d0b52-186">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d0b52-186">ReadItem</span></span>|
|[<span data-ttu-id="d0b52-187">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="d0b52-187">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d0b52-188">Создание или чтение​</span><span class="sxs-lookup"><span data-stu-id="d0b52-188">Compose or read</span></span>|

### <a name="methods"></a><span data-ttu-id="d0b52-189">Методы</span><span class="sxs-lookup"><span data-stu-id="d0b52-189">Methods</span></span>

####  <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="d0b52-190">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="d0b52-190">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="d0b52-191">Добавляет обработчик для поддерживаемого события.</span><span class="sxs-lookup"><span data-stu-id="d0b52-191">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="d0b52-192">В настоящее время поддерживаемые типы событий: `Office.EventType.ItemChanged` и `Office.EventType.OfficeThemeChanged`.</span><span class="sxs-lookup"><span data-stu-id="d0b52-192">Currently, the supported event types are `Office.EventType.ItemChanged` and `Office.EventType.OfficeThemeChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d0b52-193">Параметры:</span><span class="sxs-lookup"><span data-stu-id="d0b52-193">Parameters:</span></span>

| <span data-ttu-id="d0b52-194">Имя</span><span class="sxs-lookup"><span data-stu-id="d0b52-194">Name</span></span> | <span data-ttu-id="d0b52-195">Тип</span><span class="sxs-lookup"><span data-stu-id="d0b52-195">Type</span></span> | <span data-ttu-id="d0b52-196">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="d0b52-196">Attributes</span></span> | <span data-ttu-id="d0b52-197">Описание</span><span class="sxs-lookup"><span data-stu-id="d0b52-197">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="d0b52-198">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="d0b52-198">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="d0b52-199">Событие, которое должно вызвать обработчик.</span><span class="sxs-lookup"><span data-stu-id="d0b52-199">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="d0b52-200">Функция</span><span class="sxs-lookup"><span data-stu-id="d0b52-200">Function</span></span> || <span data-ttu-id="d0b52-p106">Функция для обработки события. Функция должна принимать один параметр, представляющий собой объектный литерал. Значение свойства `type` параметра совпадет со значением параметра `eventType`, переданного методу `addHandlerAsync`.</span><span class="sxs-lookup"><span data-stu-id="d0b52-p106">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="d0b52-204">Object</span><span class="sxs-lookup"><span data-stu-id="d0b52-204">Object</span></span> | <span data-ttu-id="d0b52-205">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="d0b52-205">&lt;optional&gt;</span></span> | <span data-ttu-id="d0b52-206">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="d0b52-206">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="d0b52-207">Object</span><span class="sxs-lookup"><span data-stu-id="d0b52-207">Object</span></span> | <span data-ttu-id="d0b52-208">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="d0b52-208">&lt;optional&gt;</span></span> | <span data-ttu-id="d0b52-209">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="d0b52-209">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="d0b52-210">function</span><span class="sxs-lookup"><span data-stu-id="d0b52-210">function</span></span>| <span data-ttu-id="d0b52-211">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="d0b52-211">&lt;optional&gt;</span></span>|<span data-ttu-id="d0b52-212">Когда метод завершает выполнение, переданная в параметре `callback` функция вызывается с единственным параметром `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="d0b52-212">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d0b52-213">Требования</span><span class="sxs-lookup"><span data-stu-id="d0b52-213">Requirements</span></span>

|<span data-ttu-id="d0b52-214">Требование</span><span class="sxs-lookup"><span data-stu-id="d0b52-214">Requirement</span></span>| <span data-ttu-id="d0b52-215">Значение</span><span class="sxs-lookup"><span data-stu-id="d0b52-215">Value</span></span>|
|---|---|
|[<span data-ttu-id="d0b52-216">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="d0b52-216">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d0b52-217">1.5</span><span class="sxs-lookup"><span data-stu-id="d0b52-217">1.5</span></span> |
|[<span data-ttu-id="d0b52-218">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="d0b52-218">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d0b52-219">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d0b52-219">ReadItem</span></span> |
|[<span data-ttu-id="d0b52-220">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="d0b52-220">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d0b52-221">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="d0b52-221">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="d0b52-222">Пример</span><span class="sxs-lookup"><span data-stu-id="d0b52-222">Example</span></span>

```
Office.initialize = function (reason) {
  $(document).ready(function () {
    Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, loadNewItem, function (result) {
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

####  <a name="converttoewsiditemid-restversion--string"></a><span data-ttu-id="d0b52-223">convertToEwsId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="d0b52-223">convertToEwsId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="d0b52-224">Преобразует идентификатор элемента из формата REST в формат EWS.</span><span class="sxs-lookup"><span data-stu-id="d0b52-224">Converts an item ID formatted for REST into EWS format.</span></span>

> [!NOTE]
> <span data-ttu-id="d0b52-225">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="d0b52-225">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="d0b52-p107">Формат идентификаторов, извлекаемых через API REST (например, [API Почты Outlook](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) или [Microsoft Graph](http://graph.microsoft.io/)), отличается от формата веб-служб Exchange (EWS). Метод `convertToEwsId` преобразует идентификатор из формата REST в формат EWS.</span><span class="sxs-lookup"><span data-stu-id="d0b52-p107">Item IDs retrieved via a REST API (such as the [Outlook Mail API](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](http://graph.microsoft.io/)) use a different format than the format used by Exchange Web Services (EWS). The `convertToEwsId` method converts a REST-formatted ID into the proper format for EWS.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d0b52-228">Параметры:</span><span class="sxs-lookup"><span data-stu-id="d0b52-228">Parameters:</span></span>

|<span data-ttu-id="d0b52-229">Имя</span><span class="sxs-lookup"><span data-stu-id="d0b52-229">Name</span></span>| <span data-ttu-id="d0b52-230">Тип</span><span class="sxs-lookup"><span data-stu-id="d0b52-230">Type</span></span>| <span data-ttu-id="d0b52-231">Описание</span><span class="sxs-lookup"><span data-stu-id="d0b52-231">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="d0b52-232">String</span><span class="sxs-lookup"><span data-stu-id="d0b52-232">String</span></span>|<span data-ttu-id="d0b52-233">Идентификатор элемента в формате REST API для Outlook</span><span class="sxs-lookup"><span data-stu-id="d0b52-233">An item ID formatted for the Outlook REST APIs</span></span>|
|`restVersion`| [<span data-ttu-id="d0b52-234">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="d0b52-234">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion)|<span data-ttu-id="d0b52-235">Значение, определяющее версию REST API для Outlook, которая используется для извлечения идентификатора элемента.</span><span class="sxs-lookup"><span data-stu-id="d0b52-235">A value indicating the version of the Outlook REST API used to retrieve the item ID.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d0b52-236">Требования</span><span class="sxs-lookup"><span data-stu-id="d0b52-236">Requirements</span></span>

|<span data-ttu-id="d0b52-237">Требование</span><span class="sxs-lookup"><span data-stu-id="d0b52-237">Requirement</span></span>| <span data-ttu-id="d0b52-238">Значение</span><span class="sxs-lookup"><span data-stu-id="d0b52-238">Value</span></span>|
|---|---|
|[<span data-ttu-id="d0b52-239">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="d0b52-239">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d0b52-240">1.3</span><span class="sxs-lookup"><span data-stu-id="d0b52-240">1.3</span></span>|
|[<span data-ttu-id="d0b52-241">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="d0b52-241">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d0b52-242">Ограниченный доступ</span><span class="sxs-lookup"><span data-stu-id="d0b52-242">Restricted</span></span>|
|[<span data-ttu-id="d0b52-243">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="d0b52-243">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d0b52-244">Compose или read</span><span class="sxs-lookup"><span data-stu-id="d0b52-244">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="d0b52-245">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="d0b52-245">Returns:</span></span>

<span data-ttu-id="d0b52-246">Тип: Строка</span><span class="sxs-lookup"><span data-stu-id="d0b52-246">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="d0b52-247">Пример</span><span class="sxs-lookup"><span data-stu-id="d0b52-247">Example</span></span>

```
// Get an item's ID from a REST API
var restId = 'AAMkAGVlOTZjNTM3LW...';

// Treat restId as coming from the v2.0 version of the
// Outlook Mail API
var ewsId = Office.context.mailbox.convertToEwsId(restId, Office.MailboxEnums.RestVersion.v2_0);
```

####  <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime"></a><span data-ttu-id="d0b52-248">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime)}</span><span class="sxs-lookup"><span data-stu-id="d0b52-248">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime)}</span></span>

<span data-ttu-id="d0b52-249">Получает словарь, содержащий информацию о времени в локальном времени клиента.</span><span class="sxs-lookup"><span data-stu-id="d0b52-249">Gets a dictionary containing time information in local client time.</span></span>

<span data-ttu-id="d0b52-p108">Для даты и времени в почтовом приложении для Outlook или Outlook Web App могут использоваться разные часовые пояса. Outlook использует часовой пояс клиентского компьютера. Outlook Web App использует часовой пояс, заданный в Центре администрирования Exchange (EAC). Значения даты и времени должны обрабатываться так, чтобы значения в пользовательском интерфейсе всегда согласовывались с часовым поясом, ожидаемым пользователем.</span><span class="sxs-lookup"><span data-stu-id="d0b52-p108">The dates and times used by a mail app for Outlook or Outlook Web App can use different time zones. Outlook uses the client computer time zone; Outlook Web App uses the time zone set on the Exchange Admin Center (EAC). You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that the user expects.</span></span>

<span data-ttu-id="d0b52-p109">Если почтовое приложение работает в Outlook, метод `convertToLocalClientTime` вернет объект словаря со значениями часового пояса клиентского компьютера. Если почтовое приложение работает в Outlook Web App, метод `convertToLocalClientTime` вернет объект словаря со значениями часового пояса, заданного в Центре администрирования Exchange.</span><span class="sxs-lookup"><span data-stu-id="d0b52-p109">If the mail app is running in Outlook, the `convertToLocalClientTime` method will return a dictionary object with the values set to the client computer time zone. If the mail app is running in Outlook Web App, the `convertToLocalClientTime` method will return a dictionary object with the values set to the time zone specified in the EAC.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d0b52-255">Параметры:</span><span class="sxs-lookup"><span data-stu-id="d0b52-255">Parameters:</span></span>

|<span data-ttu-id="d0b52-256">Имя</span><span class="sxs-lookup"><span data-stu-id="d0b52-256">Name</span></span>| <span data-ttu-id="d0b52-257">Тип</span><span class="sxs-lookup"><span data-stu-id="d0b52-257">Type</span></span>| <span data-ttu-id="d0b52-258">Описание</span><span class="sxs-lookup"><span data-stu-id="d0b52-258">Description</span></span>|
|---|---|---|
|`timeValue`| <span data-ttu-id="d0b52-259">Date</span><span class="sxs-lookup"><span data-stu-id="d0b52-259">Date</span></span>|<span data-ttu-id="d0b52-260">Объект Date</span><span class="sxs-lookup"><span data-stu-id="d0b52-260">A Date object</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d0b52-261">Требования</span><span class="sxs-lookup"><span data-stu-id="d0b52-261">Requirements</span></span>

|<span data-ttu-id="d0b52-262">Требование</span><span class="sxs-lookup"><span data-stu-id="d0b52-262">Requirement</span></span>| <span data-ttu-id="d0b52-263">Значение</span><span class="sxs-lookup"><span data-stu-id="d0b52-263">Value</span></span>|
|---|---|
|[<span data-ttu-id="d0b52-264">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="d0b52-264">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d0b52-265">1.0</span><span class="sxs-lookup"><span data-stu-id="d0b52-265">1.0</span></span>|
|[<span data-ttu-id="d0b52-266">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="d0b52-266">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d0b52-267">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d0b52-267">ReadItem</span></span>|
|[<span data-ttu-id="d0b52-268">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="d0b52-268">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d0b52-269">Compose или read</span><span class="sxs-lookup"><span data-stu-id="d0b52-269">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="d0b52-270">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="d0b52-270">Returns:</span></span>

<span data-ttu-id="d0b52-271">Тип: [LocalClientTime](/javascript/api/outlook/office.LocalClientTime)</span><span class="sxs-lookup"><span data-stu-id="d0b52-271">Type: [LocalClientTime](/javascript/api/outlook/office.LocalClientTime)</span></span>

####  <a name="converttorestiditemid-restversion--string"></a><span data-ttu-id="d0b52-272">convertToRestId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="d0b52-272">convertToRestId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="d0b52-273">Преобразует идентификатор элемента из формата EWS в формат REST.</span><span class="sxs-lookup"><span data-stu-id="d0b52-273">Converts an item ID formatted for EWS into REST format.</span></span>

> [!NOTE]
> <span data-ttu-id="d0b52-274">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="d0b52-274">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="d0b52-p110">Формат идентификаторов, извлекаемых через EWS или через свойство `itemId`, отличается от формата API REST (таких как [API почты Outlook](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) или [Microsoft Graph](http://graph.microsoft.io/)). Метод `convertToRestId` преобразует идентификатор из формата EWS в формат REST.</span><span class="sxs-lookup"><span data-stu-id="d0b52-p110">Item IDs retrieved via EWS or via the `itemId` property use a different format than the format used by REST APIs (such as the [Outlook Mail API](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](http://graph.microsoft.io/)). The `convertToRestId` method converts an EWS-formatted ID into the proper format for REST.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d0b52-277">Параметры:</span><span class="sxs-lookup"><span data-stu-id="d0b52-277">Parameters:</span></span>

|<span data-ttu-id="d0b52-278">Имя</span><span class="sxs-lookup"><span data-stu-id="d0b52-278">Name</span></span>| <span data-ttu-id="d0b52-279">Тип</span><span class="sxs-lookup"><span data-stu-id="d0b52-279">Type</span></span>| <span data-ttu-id="d0b52-280">Описание</span><span class="sxs-lookup"><span data-stu-id="d0b52-280">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="d0b52-281">String</span><span class="sxs-lookup"><span data-stu-id="d0b52-281">String</span></span>|<span data-ttu-id="d0b52-282">Идентификатор элемента в формате EWS</span><span class="sxs-lookup"><span data-stu-id="d0b52-282">An item ID formatted for Exchange Web Services (EWS)</span></span>|
|`restVersion`| [<span data-ttu-id="d0b52-283">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="d0b52-283">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion)|<span data-ttu-id="d0b52-284">Значение, определяющее версию REST API для Outlook, с которой будет использоваться преобразованный идентификатор.</span><span class="sxs-lookup"><span data-stu-id="d0b52-284">A value indicating the version of the Outlook REST API that the converted ID will be used with.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d0b52-285">Требования</span><span class="sxs-lookup"><span data-stu-id="d0b52-285">Requirements</span></span>

|<span data-ttu-id="d0b52-286">Требование</span><span class="sxs-lookup"><span data-stu-id="d0b52-286">Requirement</span></span>| <span data-ttu-id="d0b52-287">Значение</span><span class="sxs-lookup"><span data-stu-id="d0b52-287">Value</span></span>|
|---|---|
|[<span data-ttu-id="d0b52-288">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="d0b52-288">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d0b52-289">1.3</span><span class="sxs-lookup"><span data-stu-id="d0b52-289">1.3</span></span>|
|[<span data-ttu-id="d0b52-290">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="d0b52-290">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d0b52-291">Ограниченный доступ</span><span class="sxs-lookup"><span data-stu-id="d0b52-291">Restricted</span></span>|
|[<span data-ttu-id="d0b52-292">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="d0b52-292">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d0b52-293">Compose или read</span><span class="sxs-lookup"><span data-stu-id="d0b52-293">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="d0b52-294">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="d0b52-294">Returns:</span></span>

<span data-ttu-id="d0b52-295">Тип: Строка</span><span class="sxs-lookup"><span data-stu-id="d0b52-295">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="d0b52-296">Пример</span><span class="sxs-lookup"><span data-stu-id="d0b52-296">Example</span></span>

```
// Get the currently selected item's ID
var ewsId = Office.context.mailbox.item.itemId;

// Convert to a REST ID for the v2.0 version of the
// Outlook Mail API
var restId = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
```

####  <a name="converttoutcclienttimeinput--date"></a><span data-ttu-id="d0b52-297">convertToUtcClientTime(input) → {Date}</span><span class="sxs-lookup"><span data-stu-id="d0b52-297">convertToUtcClientTime(input) → {Date}</span></span>

<span data-ttu-id="d0b52-298">Получает объект Date из словаря, содержащего сведения о времени.</span><span class="sxs-lookup"><span data-stu-id="d0b52-298">Gets a Date object from a dictionary containing time information.</span></span>

<span data-ttu-id="d0b52-299">Метод `convertToUtcClientTime` преобразует словарь, содержащий локальную дату и время, в объект Date с правильными значениями локальной даты и времени.</span><span class="sxs-lookup"><span data-stu-id="d0b52-299">The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a Date object with the correct values for the local date and time.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d0b52-300">Параметры:</span><span class="sxs-lookup"><span data-stu-id="d0b52-300">Parameters:</span></span>

|<span data-ttu-id="d0b52-301">Имя</span><span class="sxs-lookup"><span data-stu-id="d0b52-301">Name</span></span>| <span data-ttu-id="d0b52-302">Тип</span><span class="sxs-lookup"><span data-stu-id="d0b52-302">Type</span></span>| <span data-ttu-id="d0b52-303">Описание</span><span class="sxs-lookup"><span data-stu-id="d0b52-303">Description</span></span>|
|---|---|---|
|`input`| [<span data-ttu-id="d0b52-304">LocalClientTime</span><span class="sxs-lookup"><span data-stu-id="d0b52-304">LocalClientTime</span></span>](/javascript/api/outlook/office.LocalClientTime)|<span data-ttu-id="d0b52-305">Значение локального времени для преобразования.</span><span class="sxs-lookup"><span data-stu-id="d0b52-305">The local time value to convert.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d0b52-306">Требования</span><span class="sxs-lookup"><span data-stu-id="d0b52-306">Requirements</span></span>

|<span data-ttu-id="d0b52-307">Требование</span><span class="sxs-lookup"><span data-stu-id="d0b52-307">Requirement</span></span>| <span data-ttu-id="d0b52-308">Значение</span><span class="sxs-lookup"><span data-stu-id="d0b52-308">Value</span></span>|
|---|---|
|[<span data-ttu-id="d0b52-309">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="d0b52-309">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d0b52-310">1.0</span><span class="sxs-lookup"><span data-stu-id="d0b52-310">1.0</span></span>|
|[<span data-ttu-id="d0b52-311">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="d0b52-311">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d0b52-312">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d0b52-312">ReadItem</span></span>|
|[<span data-ttu-id="d0b52-313">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="d0b52-313">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d0b52-314">Compose или read</span><span class="sxs-lookup"><span data-stu-id="d0b52-314">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="d0b52-315">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="d0b52-315">Returns:</span></span>

<span data-ttu-id="d0b52-316">Объект Date со временем в формате UTC.</span><span class="sxs-lookup"><span data-stu-id="d0b52-316">A Date object with the time expressed in UTC.</span></span>

<dl class="param-type"><span data-ttu-id="d0b52-317">

<dt>Тип</dt>

</span><span class="sxs-lookup"><span data-stu-id="d0b52-317">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="d0b52-318">Date</span><span class="sxs-lookup"><span data-stu-id="d0b52-318">Date</span></span></dd>

</dl>

####  <a name="displayappointmentformitemid"></a><span data-ttu-id="d0b52-319">displayAppointmentForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="d0b52-319">displayAppointmentForm(itemId)</span></span>

<span data-ttu-id="d0b52-320">Отображает имеющуюся встречу из календаря.</span><span class="sxs-lookup"><span data-stu-id="d0b52-320">Displays an existing calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="d0b52-321">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="d0b52-321">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="d0b52-322">Метод `displayAppointmentForm` открывает новое окно на компьютере или диалоговое окно на мобильном устройстве, содержащее сведения календаря о существующей встрече.</span><span class="sxs-lookup"><span data-stu-id="d0b52-322">The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="d0b52-p111">В Outlook для Mac с помощью этого метода можно отобразить одну встречу, которая не является частью повторяющегося ряда, или образец встречи из повторяющегося ряда, но не экземпляр ряда, так как в Outlook для Mac невозможно получить доступ к свойствам экземпляра повторяющегося ряда (в том числе к идентификатору элемента).</span><span class="sxs-lookup"><span data-stu-id="d0b52-p111">In Outlook for Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment of a recurring series, but you cannot display an instance of the series. This is because in Outlook for Mac, you cannot access the properties (including the item ID) of instances of a recurring series.</span></span>

<span data-ttu-id="d0b52-325">В Outlook Web App этот метод открывает указанную форму, только если текст формы содержит не более 32 КБ символов.</span><span class="sxs-lookup"><span data-stu-id="d0b52-325">In Outlook Web App, this method opens the specified form only if the body of the form is less than or equal to 32KB number of characters.</span></span>

<span data-ttu-id="d0b52-326">Если указанный идентификатор элемента не определяет существующую встречу, на клиентском компьютере или устройстве открывается пустая страница, и сообщение об ошибке не возвращается.</span><span class="sxs-lookup"><span data-stu-id="d0b52-326">If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and no error message will be returned.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d0b52-327">Параметры:</span><span class="sxs-lookup"><span data-stu-id="d0b52-327">Parameters:</span></span>

|<span data-ttu-id="d0b52-328">Имя</span><span class="sxs-lookup"><span data-stu-id="d0b52-328">Name</span></span>| <span data-ttu-id="d0b52-329">Тип</span><span class="sxs-lookup"><span data-stu-id="d0b52-329">Type</span></span>| <span data-ttu-id="d0b52-330">Описание</span><span class="sxs-lookup"><span data-stu-id="d0b52-330">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="d0b52-331">String</span><span class="sxs-lookup"><span data-stu-id="d0b52-331">String</span></span>|<span data-ttu-id="d0b52-332">Идентификатор веб-служб Exchange (EWS) для существующей встречи в календаре.</span><span class="sxs-lookup"><span data-stu-id="d0b52-332">The Exchange Web Services (EWS) identifier for an existing calendar appointment.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d0b52-333">Требования</span><span class="sxs-lookup"><span data-stu-id="d0b52-333">Requirements</span></span>

|<span data-ttu-id="d0b52-334">Требование</span><span class="sxs-lookup"><span data-stu-id="d0b52-334">Requirement</span></span>| <span data-ttu-id="d0b52-335">Значение</span><span class="sxs-lookup"><span data-stu-id="d0b52-335">Value</span></span>|
|---|---|
|[<span data-ttu-id="d0b52-336">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="d0b52-336">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d0b52-337">1.0</span><span class="sxs-lookup"><span data-stu-id="d0b52-337">1.0</span></span>|
|[<span data-ttu-id="d0b52-338">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="d0b52-338">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d0b52-339">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d0b52-339">ReadItem</span></span>|
|[<span data-ttu-id="d0b52-340">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="d0b52-340">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d0b52-341">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="d0b52-341">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="d0b52-342">Пример</span><span class="sxs-lookup"><span data-stu-id="d0b52-342">Example</span></span>

```
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

####  <a name="displaymessageformitemid"></a><span data-ttu-id="d0b52-343">displayMessageForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="d0b52-343">displayMessageForm(itemId)</span></span>

<span data-ttu-id="d0b52-344">Отображает существующее сообщение.</span><span class="sxs-lookup"><span data-stu-id="d0b52-344">Displays an existing message.</span></span>

> [!NOTE]
> <span data-ttu-id="d0b52-345">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="d0b52-345">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="d0b52-346">Метод `displayMessageForm` открывает существующее сообщение в новом окне на компьютере или в диалоговом окне на мобильных устройствах.</span><span class="sxs-lookup"><span data-stu-id="d0b52-346">The `displayMessageForm` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="d0b52-347">В Outlook Web App этот метод открывает указанную форму, только если текст формы содержит не более 32 КБ символов.</span><span class="sxs-lookup"><span data-stu-id="d0b52-347">In Outlook Web App, this method opens the specified form only if the body of the form is less than or equal to 32 KB number of characters.</span></span>

<span data-ttu-id="d0b52-348">Если указанный идентификатор элемента не определяет существующее сообщение, окно на клиентском компьютере не открывается и сообщение об ошибке не возвращается.</span><span class="sxs-lookup"><span data-stu-id="d0b52-348">If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and no error message will be returned.</span></span>

<span data-ttu-id="d0b52-p112">Не используйте `displayMessageForm` с параметром `itemId`, который представляет собой встречу. Используйте метод `displayAppointmentForm`, чтобы отобразить сведения о существующей встрече, а метод `displayNewAppointmentForm` — для отображения формы создания встречи.</span><span class="sxs-lookup"><span data-stu-id="d0b52-p112">Do not use the `displayMessageForm` with an `itemId` that represents an appointment. Use the `displayAppointmentForm` method to display an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d0b52-351">Параметры:</span><span class="sxs-lookup"><span data-stu-id="d0b52-351">Parameters:</span></span>

|<span data-ttu-id="d0b52-352">Имя</span><span class="sxs-lookup"><span data-stu-id="d0b52-352">Name</span></span>| <span data-ttu-id="d0b52-353">Тип</span><span class="sxs-lookup"><span data-stu-id="d0b52-353">Type</span></span>| <span data-ttu-id="d0b52-354">Описание</span><span class="sxs-lookup"><span data-stu-id="d0b52-354">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="d0b52-355">String</span><span class="sxs-lookup"><span data-stu-id="d0b52-355">String</span></span>|<span data-ttu-id="d0b52-356">Идентификатор веб-служб Exchange (EWS) для существующего сообщения.</span><span class="sxs-lookup"><span data-stu-id="d0b52-356">The Exchange Web Services (EWS) identifier for an existing message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d0b52-357">Требования</span><span class="sxs-lookup"><span data-stu-id="d0b52-357">Requirements</span></span>

|<span data-ttu-id="d0b52-358">Требование</span><span class="sxs-lookup"><span data-stu-id="d0b52-358">Requirement</span></span>| <span data-ttu-id="d0b52-359">Значение</span><span class="sxs-lookup"><span data-stu-id="d0b52-359">Value</span></span>|
|---|---|
|[<span data-ttu-id="d0b52-360">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="d0b52-360">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d0b52-361">1.0</span><span class="sxs-lookup"><span data-stu-id="d0b52-361">1.0</span></span>|
|[<span data-ttu-id="d0b52-362">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="d0b52-362">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d0b52-363">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d0b52-363">ReadItem</span></span>|
|[<span data-ttu-id="d0b52-364">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="d0b52-364">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d0b52-365">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="d0b52-365">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="d0b52-366">Пример</span><span class="sxs-lookup"><span data-stu-id="d0b52-366">Example</span></span>

```
Office.context.mailbox.displayMessageForm(messageId);
```

#### <a name="displaynewappointmentformparameters"></a><span data-ttu-id="d0b52-367">displayNewAppointmentForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="d0b52-367">displayNewAppointmentForm(parameters)</span></span>

<span data-ttu-id="d0b52-368">Отображает форму для создания новой встречи в календаре.</span><span class="sxs-lookup"><span data-stu-id="d0b52-368">Displays a form for creating a new calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="d0b52-369">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="d0b52-369">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="d0b52-p113">Метод `displayNewAppointmentForm` открывает форму, в которой пользователь может создать встречу или собрание. Если параметры заданы, поля формы встречи автоматически заполняются их содержимым.</span><span class="sxs-lookup"><span data-stu-id="d0b52-p113">The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting. If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="d0b52-p114">В Outlook Web App и Outlook Web App для устройств этот метод всегда отображает форму с полем участников. Если вы не укажете участников в качестве входных аргументов, метод отображает форму с кнопкой **Сохранить**. Если вы укажете участников, форма будет включать участников и кнопку **Отправить**.</span><span class="sxs-lookup"><span data-stu-id="d0b52-p114">In Outlook Web App and OWA for Devices, this method always displays a form with an attendees field. If you do not specify any attendees as input arguments, the method displays a form with a **Save** button. If you have specified attendees, the form would include the attendees and a **Send** button.</span></span>

<span data-ttu-id="d0b52-p115">Если вы укажете участников или ресурсы с помощью параметра `requiredAttendees`, `optionalAttendees` или `resources` в клиенте Outlook с расширенными возможностями и Outlook RT, этот метод отобразит форму собрания с кнопкой **Отправить**. Если не указать получателей, этот метод отобразит форму встречи с кнопкой **Сохранить и закрыть**.</span><span class="sxs-lookup"><span data-stu-id="d0b52-p115">In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or `resources` parameter, this method displays a meeting form with a **Send** button. If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.</span></span>

<span data-ttu-id="d0b52-377">Если параметры превышают указанные ограничения размера или если указано неизвестное имя параметра, создается исключение.</span><span class="sxs-lookup"><span data-stu-id="d0b52-377">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d0b52-378">Параметры:</span><span class="sxs-lookup"><span data-stu-id="d0b52-378">Parameters:</span></span>

> [!NOTE]
> <span data-ttu-id="d0b52-379">Примечание. Все параметры являются необязательными.</span><span class="sxs-lookup"><span data-stu-id="d0b52-379">Note: All parameters are optional.</span></span>

|<span data-ttu-id="d0b52-380">Имя</span><span class="sxs-lookup"><span data-stu-id="d0b52-380">Name</span></span>| <span data-ttu-id="d0b52-381">Тип</span><span class="sxs-lookup"><span data-stu-id="d0b52-381">Type</span></span>| <span data-ttu-id="d0b52-382">Описание</span><span class="sxs-lookup"><span data-stu-id="d0b52-382">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="d0b52-383">Object</span><span class="sxs-lookup"><span data-stu-id="d0b52-383">Object</span></span> | <span data-ttu-id="d0b52-384">Словарь параметров, описывающий новую встречу.</span><span class="sxs-lookup"><span data-stu-id="d0b52-384">A dictionary of parameters describing the new appointment.</span></span> |
| `parameters.requiredAttendees` | <span data-ttu-id="d0b52-385">Array.&lt;String&gt; | Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="d0b52-385">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="d0b52-p116">Массив строк, содержащий электронные адреса, или массив, содержащий объекты `EmailAddressDetails` для каждого из обязательных участников встречи. Массив может включать не более 100 записей.</span><span class="sxs-lookup"><span data-stu-id="d0b52-p116">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the required attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.optionalAttendees` | <span data-ttu-id="d0b52-388">Array.&lt;String&gt; | Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="d0b52-388">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="d0b52-p117">Массив строк, содержащий электронные адреса, или массив, содержащий объекты `EmailAddressDetails` для каждого из необязательных участников встречи. Массив может включать не более 100 записей.</span><span class="sxs-lookup"><span data-stu-id="d0b52-p117">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the optional attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.start` | <span data-ttu-id="d0b52-391">Date</span><span class="sxs-lookup"><span data-stu-id="d0b52-391">Date</span></span> | <span data-ttu-id="d0b52-392">Объект `Date`, указывающий дату и время начала встречи.</span><span class="sxs-lookup"><span data-stu-id="d0b52-392">A `Date` object specifying the start date and time of the appointment.</span></span> |
| `parameters.end` | <span data-ttu-id="d0b52-393">Date</span><span class="sxs-lookup"><span data-stu-id="d0b52-393">Date</span></span> | <span data-ttu-id="d0b52-394">Объект `Date`, указывающий дату и время окончания встречи.</span><span class="sxs-lookup"><span data-stu-id="d0b52-394">A `Date` object specifying the end date and time of the appointment.</span></span> |
| `parameters.location` | <span data-ttu-id="d0b52-395">String</span><span class="sxs-lookup"><span data-stu-id="d0b52-395">String</span></span> | <span data-ttu-id="d0b52-p118">Строка со сведениями о месте встречи. Максимальное количество символов в строке — 255.</span><span class="sxs-lookup"><span data-stu-id="d0b52-p118">A string containing the location of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.resources` | <span data-ttu-id="d0b52-398">Array.&lt;String&gt;</span><span class="sxs-lookup"><span data-stu-id="d0b52-398">Array.&lt;String&gt;</span></span> | <span data-ttu-id="d0b52-p119">Массив строк, содержащий необходимые для встречи ресурсы. Массив может включать не более 100 записей.</span><span class="sxs-lookup"><span data-stu-id="d0b52-p119">An array of strings containing the resources required for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="d0b52-401">String</span><span class="sxs-lookup"><span data-stu-id="d0b52-401">String</span></span> | <span data-ttu-id="d0b52-p120">Строка с темой встречи. Максимальное количество символов в строке — 255.</span><span class="sxs-lookup"><span data-stu-id="d0b52-p120">A string containing the subject of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.body` | <span data-ttu-id="d0b52-404">String</span><span class="sxs-lookup"><span data-stu-id="d0b52-404">String</span></span> | <span data-ttu-id="d0b52-p121">Текст сообщения о встрече. Максимальный размер содержимого сообщения — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="d0b52-p121">The body of the appointment. The body content is limited to a maximum size of 32 KB.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="d0b52-407">Требования</span><span class="sxs-lookup"><span data-stu-id="d0b52-407">Requirements</span></span>

|<span data-ttu-id="d0b52-408">Требование</span><span class="sxs-lookup"><span data-stu-id="d0b52-408">Requirement</span></span>| <span data-ttu-id="d0b52-409">Значение</span><span class="sxs-lookup"><span data-stu-id="d0b52-409">Value</span></span>|
|---|---|
|[<span data-ttu-id="d0b52-410">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="d0b52-410">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d0b52-411">1.0</span><span class="sxs-lookup"><span data-stu-id="d0b52-411">1.0</span></span>|
|[<span data-ttu-id="d0b52-412">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="d0b52-412">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d0b52-413">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d0b52-413">ReadItem</span></span>|
|[<span data-ttu-id="d0b52-414">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="d0b52-414">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d0b52-415">Чтение</span><span class="sxs-lookup"><span data-stu-id="d0b52-415">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d0b52-416">Пример</span><span class="sxs-lookup"><span data-stu-id="d0b52-416">Example</span></span>

```
var start = new Date();
var end = new Date();
end.setHours(start.getHours() + 1);

Office.context.mailbox.displayNewAppointmentForm(
  {
    requiredAttendees: ['bob@contoso.com'],
    optionalAttendees: ['sam@contoso.com'],
    start: start,
    end: end,
    location: 'Home',
    resources: ['projector@contoso.com'],
    subject: 'meeting',
    body: 'Hello World!'
  });
```

#### <a name="displaynewmessageformparameters"></a><span data-ttu-id="d0b52-417">displayNewMessageForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="d0b52-417">displayNewMessageForm(parameters)</span></span>

<span data-ttu-id="d0b52-418">Тип: String</span><span class="sxs-lookup"><span data-stu-id="d0b52-418">Displays a form for creating a new message.</span></span>

<span data-ttu-id="d0b52-419">Метод `displayNewMessageForm` открывает форму, в которой пользователь может создать сообщение.</span><span class="sxs-lookup"><span data-stu-id="d0b52-419">The `displayNewMessageForm` method opens a form that enables the user to create a new message.</span></span> <span data-ttu-id="d0b52-420">Если параметры заданы, поля формы сообщения автоматически заполняются их содержимым.</span><span class="sxs-lookup"><span data-stu-id="d0b52-420">If parameters are specified, the message form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="d0b52-421">Если параметры превышают указанные ограничения размера или если указано неизвестное имя параметра, создается исключение.</span><span class="sxs-lookup"><span data-stu-id="d0b52-421">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d0b52-422">Параметры:</span><span class="sxs-lookup"><span data-stu-id="d0b52-422">Parameters:</span></span>

> [!NOTE]
> <span data-ttu-id="d0b52-423">Примечание. Все параметры являются необязательными.</span><span class="sxs-lookup"><span data-stu-id="d0b52-423">Note: All parameters are optional.</span></span>

|<span data-ttu-id="d0b52-424">Имя</span><span class="sxs-lookup"><span data-stu-id="d0b52-424">Name</span></span>| <span data-ttu-id="d0b52-425">Тип</span><span class="sxs-lookup"><span data-stu-id="d0b52-425">Type</span></span>| <span data-ttu-id="d0b52-426">Описание</span><span class="sxs-lookup"><span data-stu-id="d0b52-426">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="d0b52-427">Объект</span><span class="sxs-lookup"><span data-stu-id="d0b52-427">Object</span></span> | <span data-ttu-id="d0b52-428">Словарь параметров, описывающий новое сообщение.</span><span class="sxs-lookup"><span data-stu-id="d0b52-428">A dictionary of parameters describing the new message.</span></span> |
| `parameters.toRecipients` | <span data-ttu-id="d0b52-429">Array.&lt;String&gt; | Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="d0b52-429">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="d0b52-430">Массив строк, содержащий электронные адреса, или массив, содержащий объекты `EmailAddressDetails` для каждого из получателей, указанных в строке "Кому".</span><span class="sxs-lookup"><span data-stu-id="d0b52-430">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the To line.</span></span> <span data-ttu-id="d0b52-431">Массив может включать не более 100 записей.</span><span class="sxs-lookup"><span data-stu-id="d0b52-431">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.ccRecipients` | <span data-ttu-id="d0b52-432">Array.&lt;String&gt; | Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="d0b52-432">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="d0b52-433">Массив строк, содержащий электронные адреса, или массив, содержащий объекты `EmailAddressDetails` для каждого из получателей, указанных в строке Cc (Копия).</span><span class="sxs-lookup"><span data-stu-id="d0b52-433">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Cc line.</span></span> <span data-ttu-id="d0b52-434">Массив может включать не более 100 записей.</span><span class="sxs-lookup"><span data-stu-id="d0b52-434">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.bccRecipients` | <span data-ttu-id="d0b52-435">Array.&lt;String&gt; | Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="d0b52-435">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="d0b52-436">Массив строк, содержащий электронные адреса, или массив, содержащий объекты `EmailAddressDetails` для каждого из получателей, указанных в строке "Скрытая копия".</span><span class="sxs-lookup"><span data-stu-id="d0b52-436">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Bcc line.</span></span> <span data-ttu-id="d0b52-437">Массив может включать не более 100 записей.</span><span class="sxs-lookup"><span data-stu-id="d0b52-437">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="d0b52-438">String</span><span class="sxs-lookup"><span data-stu-id="d0b52-438">String</span></span> | <span data-ttu-id="d0b52-439">Строка с темой сообщения.</span><span class="sxs-lookup"><span data-stu-id="d0b52-439">A string containing the subject of the message.</span></span> <span data-ttu-id="d0b52-440">Максимальное количество символов в строке — 255.</span><span class="sxs-lookup"><span data-stu-id="d0b52-440">The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.htmlBody` | <span data-ttu-id="d0b52-441">String</span><span class="sxs-lookup"><span data-stu-id="d0b52-441">String</span></span> | <span data-ttu-id="d0b52-442">Текст сообщения в формате HTML.</span><span class="sxs-lookup"><span data-stu-id="d0b52-442">The HTML body of the message.</span></span> <span data-ttu-id="d0b52-443">Максимальный размер содержимого сообщения — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="d0b52-443">The body content is limited to a maximum size of 32 KB.</span></span> |
| `parameters.attachments` | <span data-ttu-id="d0b52-444">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="d0b52-444">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="d0b52-445">Массив объектов JSON, представляющих собой вложенные файлы или элементы.</span><span class="sxs-lookup"><span data-stu-id="d0b52-445">An array of JSON objects that are either file or item attachments.</span></span> |
| `parameters.attachments.type` | <span data-ttu-id="d0b52-446">String</span><span class="sxs-lookup"><span data-stu-id="d0b52-446">String</span></span> | <span data-ttu-id="d0b52-p128">Указывает тип вложения. Допустимые значения: `file` для вложенного файла и `item` для вложенного элемента.</span><span class="sxs-lookup"><span data-stu-id="d0b52-p128">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `parameters.attachments.name` | <span data-ttu-id="d0b52-449">String</span><span class="sxs-lookup"><span data-stu-id="d0b52-449">String</span></span> | <span data-ttu-id="d0b52-450">Строка, содержащая имя вложения, длиной до 255 символов.</span><span class="sxs-lookup"><span data-stu-id="d0b52-450">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `parameters.attachments.url` | <span data-ttu-id="d0b52-451">String</span><span class="sxs-lookup"><span data-stu-id="d0b52-451">String</span></span> | <span data-ttu-id="d0b52-p129">Используется, только если свойству `type` задано значение `file`. URI расположения файла.</span><span class="sxs-lookup"><span data-stu-id="d0b52-p129">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `parameters.attachments.isInline` | <span data-ttu-id="d0b52-454">Логическое значение</span><span class="sxs-lookup"><span data-stu-id="d0b52-454">Boolean</span></span> | <span data-ttu-id="d0b52-p130">Используется только в том случае, если свойству `type` задано значение `file`. Значение `true` указывает на то, что вложение будет встроено в текст сообщения и не должно отображаться в списке вложений.</span><span class="sxs-lookup"><span data-stu-id="d0b52-p130">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `parameters.attachments.itemId` | <span data-ttu-id="d0b52-457">String</span><span class="sxs-lookup"><span data-stu-id="d0b52-457">String</span></span> | <span data-ttu-id="d0b52-458">Используется только в том случае, если свойству `type` присвоено значение `item`.</span><span class="sxs-lookup"><span data-stu-id="d0b52-458">Only used if `type` is set to `item`.</span></span> <span data-ttu-id="d0b52-459">Идентификатор элемента веб-служб Exchange существующего сообщения электронной почты, которые необходимо присоединить к новому сообщению.</span><span class="sxs-lookup"><span data-stu-id="d0b52-459">The EWS item id of the existing e-mail you want to attach to the new message.</span></span> <span data-ttu-id="d0b52-460">Это строка длиной до 100 символов.</span><span class="sxs-lookup"><span data-stu-id="d0b52-460">This is a string up to 100 characters.</span></span> |


##### <a name="requirements"></a><span data-ttu-id="d0b52-461">Требования</span><span class="sxs-lookup"><span data-stu-id="d0b52-461">Requirements</span></span>

|<span data-ttu-id="d0b52-462">Требование</span><span class="sxs-lookup"><span data-stu-id="d0b52-462">Requirement</span></span>| <span data-ttu-id="d0b52-463">Значение</span><span class="sxs-lookup"><span data-stu-id="d0b52-463">Value</span></span>|
|---|---|
|[<span data-ttu-id="d0b52-464">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="d0b52-464">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d0b52-465">1.6</span><span class="sxs-lookup"><span data-stu-id="d0b52-465">1.6</span></span> |
|[<span data-ttu-id="d0b52-466">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="d0b52-466">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d0b52-467">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d0b52-467">ReadItem</span></span>|
|[<span data-ttu-id="d0b52-468">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="d0b52-468">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d0b52-469">Чтение</span><span class="sxs-lookup"><span data-stu-id="d0b52-469">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d0b52-470">Пример</span><span class="sxs-lookup"><span data-stu-id="d0b52-470">Example</span></span>

```
Office.context.mailbox.displayNewMessageForm(
  {
    toRecipients: Office.context.mailbox.item.to, // Copy the To line from current item
    ccRecipients: ['sam@contoso.com'],
    subject: 'Outlook add-ins are cool!',
    htmlBody: 'Hello <b>World</b>!<br/><img src="cid:image.png"></i>',
    attachments: [
      {
        type: 'file',
        name: 'image.png',
        url: 'http://contoso.com/image.png',
        isInline: true
      }
    ]
  });
```

#### <a name="getcallbacktokenasyncoptions-callback"></a><span data-ttu-id="d0b52-471">getCallbackTokenAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="d0b52-471">getCallbackTokenAsync([options], callback)</span></span>

<span data-ttu-id="d0b52-472">Возвращает строку, содержащую маркер, который используется для вызова API REST или веб-служб Exchange.</span><span class="sxs-lookup"><span data-stu-id="d0b52-472">Gets a string that contains a token used to call REST APIs or Exchange Web Services.</span></span>

<span data-ttu-id="d0b52-p132">Метод `getCallbackTokenAsync` совершает асинхронный вызов, чтобы получить непрозрачный маркер с сервера Exchange Server, на котором размещен почтовый ящик пользователя. Время существования маркера обратного вызова составляет 5 минут.</span><span class="sxs-lookup"><span data-stu-id="d0b52-p132">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

> [!NOTE]
> <span data-ttu-id="d0b52-475">Рекомендуем сделать так, чтобы по мере возможности надстройки использовали интерфейсы REST API, а не веб-службы Exchange.</span><span class="sxs-lookup"><span data-stu-id="d0b52-475">Note: It is recommended that add-ins use the REST APIs instead of Exchange Web Services whenever possible.</span></span> 

<span data-ttu-id="d0b52-476">**Маркеры REST**</span><span class="sxs-lookup"><span data-stu-id="d0b52-476">**REST Tokens**</span></span>

<span data-ttu-id="d0b52-p133">Если запрашивается маркер REST (`options.isRest = true`), полученный маркер не подойдет для проверки подлинности при вызовах веб-служб Exchange. Область действия маркера будет ограничена доступом только для чтения к текущему элементу и его вложениям, если в манифесте надстройки не указано разрешение [`ReadWriteMailbox`](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission). Если указано разрешение `ReadWriteMailbox`, полученный маркер предоставит доступ на чтение и запись к почте, календарю и контактам, включая возможность отправки почты.</span><span class="sxs-lookup"><span data-stu-id="d0b52-p133">When a REST token is requested (`options.isRest = true`), the resulting token will not work to authenticate Exchange Web Services calls. The token will be limited in scope to read-only access to the current item and its attachments, unless the add-in has specified the [`ReadWriteMailbox`](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) permission in its manifest. If the `ReadWriteMailbox` permission is specified, the resulting token will grant read/write access to mail, calendar, and contacts, including the ability to send mail.</span></span>

<span data-ttu-id="d0b52-480">С помощью свойства `restUrl` надстройка должна определить правильный URL-адрес для вызовов REST API.</span><span class="sxs-lookup"><span data-stu-id="d0b52-480">The add-in should use the `restUrl` property to determine the correct URL to use when making REST API calls.</span></span>

<span data-ttu-id="d0b52-481">**Маркеры EWS**</span><span class="sxs-lookup"><span data-stu-id="d0b52-481">**EWS Tokens**</span></span>

<span data-ttu-id="d0b52-p134">Если запрашивается маркер EWS (`options.isRest = false`), полученный маркер не подойдет для проверки подлинности при вызовах REST API. Область действия маркера будет ограничена доступом к текущему элементу.</span><span class="sxs-lookup"><span data-stu-id="d0b52-p134">When an EWS token is requested (`options.isRest = false`), the resulting token will not work to authenticate REST API calls. The token will be limited in scope to accessing the current item.</span></span>

<span data-ttu-id="d0b52-484">С помощью свойства `ewsUrl` надстройка должна определить правильный URL-адрес для вызовов EWS.</span><span class="sxs-lookup"><span data-stu-id="d0b52-484">The add-in should use the `ewsUrl` property to determine the correct URL to use when making EWS calls.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d0b52-485">Параметры:</span><span class="sxs-lookup"><span data-stu-id="d0b52-485">Parameters:</span></span>

|<span data-ttu-id="d0b52-486">Имя</span><span class="sxs-lookup"><span data-stu-id="d0b52-486">Name</span></span>| <span data-ttu-id="d0b52-487">Тип</span><span class="sxs-lookup"><span data-stu-id="d0b52-487">Type</span></span>| <span data-ttu-id="d0b52-488">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="d0b52-488">Attributes</span></span>| <span data-ttu-id="d0b52-489">Описание</span><span class="sxs-lookup"><span data-stu-id="d0b52-489">Description</span></span>|
|---|---|---|---|
| `options` | <span data-ttu-id="d0b52-490">Object</span><span class="sxs-lookup"><span data-stu-id="d0b52-490">Object</span></span> | <span data-ttu-id="d0b52-491">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="d0b52-491">&lt;optional&gt;</span></span> | <span data-ttu-id="d0b52-492">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="d0b52-492">An object literal that contains one or more of the following properties.</span></span> |
| `options.isRest` | <span data-ttu-id="d0b52-493">Логическое значение</span><span class="sxs-lookup"><span data-stu-id="d0b52-493">Boolean</span></span> |  <span data-ttu-id="d0b52-494">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="d0b52-494">&lt;optional&gt;</span></span> | <span data-ttu-id="d0b52-p135">Определяет, будет ли предоставленный маркер использоваться для интерфейсов REST API Outlook или веб-служб Exchange. Значение по умолчанию — `false`.</span><span class="sxs-lookup"><span data-stu-id="d0b52-p135">Determines if the token provided will be used for the Outlook REST APIs or Exchange Web Services. Default value is `false`.</span></span> |
| `options.asyncContext` | <span data-ttu-id="d0b52-497">Объект</span><span class="sxs-lookup"><span data-stu-id="d0b52-497">Object</span></span> |  <span data-ttu-id="d0b52-498">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="d0b52-498">&lt;optional&gt;</span></span> | <span data-ttu-id="d0b52-499">Данные о состоянии, передаваемые в асинхронный метод.</span><span class="sxs-lookup"><span data-stu-id="d0b52-499">Any state data that is passed to the asynchronous method.</span></span> |
|`callback`| <span data-ttu-id="d0b52-500">функция</span><span class="sxs-lookup"><span data-stu-id="d0b52-500">function</span></span>||<span data-ttu-id="d0b52-p136">Когда метод завершает выполнение, переданная в параметре `callback` функция вызывается с единственным параметром `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult). Маркер указывается в виде строки в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="d0b52-p136">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object. The token is provided as a string in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d0b52-503">Требования</span><span class="sxs-lookup"><span data-stu-id="d0b52-503">Requirements</span></span>

|<span data-ttu-id="d0b52-504">Требование</span><span class="sxs-lookup"><span data-stu-id="d0b52-504">Requirement</span></span>| <span data-ttu-id="d0b52-505">Значение</span><span class="sxs-lookup"><span data-stu-id="d0b52-505">Value</span></span>|
|---|---|
|[<span data-ttu-id="d0b52-506">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="d0b52-506">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d0b52-507">1.5</span><span class="sxs-lookup"><span data-stu-id="d0b52-507">1.5</span></span> |
|[<span data-ttu-id="d0b52-508">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="d0b52-508">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d0b52-509">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d0b52-509">ReadItem</span></span>|
|[<span data-ttu-id="d0b52-510">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="d0b52-510">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d0b52-511">Создание и чтение</span><span class="sxs-lookup"><span data-stu-id="d0b52-511">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="d0b52-512">Пример</span><span class="sxs-lookup"><span data-stu-id="d0b52-512">Example</span></span>

```js
function getCallbackToken() {
  var options = {
    isRest: true,
    asyncContext: { message: 'Hello World!' }
  };

  Office.context.mailbox.getCallbackTokenAsync(options, cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

#### <a name="getcallbacktokenasynccallback-usercontext"></a><span data-ttu-id="d0b52-513">getCallbackTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="d0b52-513">getCallbackTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="d0b52-514">Получает строку, содержащую маркер, используемый для получения вложения или элемента с Exchange Server.</span><span class="sxs-lookup"><span data-stu-id="d0b52-514">Gets a string that contains a token used to get an attachment or item from an Exchange Server.</span></span>

<span data-ttu-id="d0b52-p137">Метод `getCallbackTokenAsync` совершает асинхронный вызов, чтобы получить непрозрачный маркер с сервера Exchange Server, на котором размещен почтовый ящик пользователя. Время существования маркера обратного вызова составляет 5 минут.</span><span class="sxs-lookup"><span data-stu-id="d0b52-p137">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

<span data-ttu-id="d0b52-p138">Вы можете передать сторонней системе маркер и идентификатор вложения или элемента. Сторонняя система использует этот маркер как маркер авторизации, чтобы вызвать операцию [GetAttachment](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getattachment-operation) или [GetItem](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getitem-operation) веб-служб Exchange для возврата вложения или элемента. Например, вы можете создать удаленную службу, чтобы [получить вложения из выбранного элемента](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="d0b52-p138">You can pass the token and an attachment identifier or item identifier to a third-party system. The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getattachment-operation) or [GetItem](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item. For example, you can create a remote service to [get attachments from the selected item](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="d0b52-520">В манифесте приложения должно быть указано разрешение **ReadItem** для вызова метода `getCallbackTokenAsync` в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="d0b52-520">Your app must have the **ReadItem** permission specified in its manifest to call the `getCallbackTokenAsync` method in read mode.</span></span>

<span data-ttu-id="d0b52-p139">Чтобы получить идентификатор элемента для передачи в метод `getCallbackTokenAsync`, в режиме создания необходимо вызвать метод [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback). Для вызова метода `saveAsync` приложение должно иметь разрешения **ReadWriteItem**.</span><span class="sxs-lookup"><span data-stu-id="d0b52-p139">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method to get an item identifier to pass to the `getCallbackTokenAsync` method. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d0b52-523">Параметры:</span><span class="sxs-lookup"><span data-stu-id="d0b52-523">Parameters:</span></span>

|<span data-ttu-id="d0b52-524">Имя</span><span class="sxs-lookup"><span data-stu-id="d0b52-524">Name</span></span>| <span data-ttu-id="d0b52-525">Тип</span><span class="sxs-lookup"><span data-stu-id="d0b52-525">Type</span></span>| <span data-ttu-id="d0b52-526">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="d0b52-526">Attributes</span></span>| <span data-ttu-id="d0b52-527">Описание</span><span class="sxs-lookup"><span data-stu-id="d0b52-527">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="d0b52-528">function</span><span class="sxs-lookup"><span data-stu-id="d0b52-528">function</span></span>||<span data-ttu-id="d0b52-p140">Когда метод завершает выполнение, переданная в параметре `callback` функция вызывается с единственным параметром `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult). Маркер указывается в виде строки в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="d0b52-p140">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object. The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="d0b52-531">Object</span><span class="sxs-lookup"><span data-stu-id="d0b52-531">Object</span></span>| <span data-ttu-id="d0b52-532">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="d0b52-532">&lt;optional&gt;</span></span>|<span data-ttu-id="d0b52-533">Данные о состоянии, передаваемые в асинхронный метод.</span><span class="sxs-lookup"><span data-stu-id="d0b52-533">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d0b52-534">Требования</span><span class="sxs-lookup"><span data-stu-id="d0b52-534">Requirements</span></span>

|<span data-ttu-id="d0b52-535">Требование</span><span class="sxs-lookup"><span data-stu-id="d0b52-535">Requirement</span></span>| <span data-ttu-id="d0b52-536">Значение</span><span class="sxs-lookup"><span data-stu-id="d0b52-536">Value</span></span>|
|---|---|
|[<span data-ttu-id="d0b52-537">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="d0b52-537">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d0b52-538">1.3</span><span class="sxs-lookup"><span data-stu-id="d0b52-538">1.3</span></span>|
|[<span data-ttu-id="d0b52-539">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="d0b52-539">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d0b52-540">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d0b52-540">ReadItem</span></span>|
|[<span data-ttu-id="d0b52-541">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="d0b52-541">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d0b52-542">Создание и чтение</span><span class="sxs-lookup"><span data-stu-id="d0b52-542">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="d0b52-543">Пример</span><span class="sxs-lookup"><span data-stu-id="d0b52-543">Example</span></span>

```js
function getCallbackToken() {
  Office.context.mailbox.getCallbackTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

####  <a name="getuseridentitytokenasynccallback-usercontext"></a><span data-ttu-id="d0b52-544">getUserIdentityTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="d0b52-544">getUserIdentityTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="d0b52-545">Получает маркер, идентифицирующий пользователя и надстройку Office.</span><span class="sxs-lookup"><span data-stu-id="d0b52-545">Gets a token identifying the user and the Office Add-in.</span></span>

<span data-ttu-id="d0b52-546">Метод `getUserIdentityTokenAsync` возвращает маркер, который можно использовать для идентификации и [проверки подлинности надстройки и пользователя в сторонней системе](https://docs.microsoft.com/outlook/add-ins/authentication).</span><span class="sxs-lookup"><span data-stu-id="d0b52-546">The `getUserIdentityTokenAsync` method returns a token that you can use to identify and [authenticate the add-in and user with a third-party system](https://docs.microsoft.com/outlook/add-ins/authentication).</span></span>

##### <a name="parameters"></a><span data-ttu-id="d0b52-547">Параметры:</span><span class="sxs-lookup"><span data-stu-id="d0b52-547">Parameters:</span></span>

|<span data-ttu-id="d0b52-548">Имя</span><span class="sxs-lookup"><span data-stu-id="d0b52-548">Name</span></span>| <span data-ttu-id="d0b52-549">Тип</span><span class="sxs-lookup"><span data-stu-id="d0b52-549">Type</span></span>| <span data-ttu-id="d0b52-550">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="d0b52-550">Attributes</span></span>| <span data-ttu-id="d0b52-551">Описание</span><span class="sxs-lookup"><span data-stu-id="d0b52-551">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="d0b52-552">function</span><span class="sxs-lookup"><span data-stu-id="d0b52-552">function</span></span>||<span data-ttu-id="d0b52-553">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="d0b52-553">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="d0b52-554">Маркер указывается в виде строки в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="d0b52-554">The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="d0b52-555">Объект</span><span class="sxs-lookup"><span data-stu-id="d0b52-555">Object</span></span>| <span data-ttu-id="d0b52-556">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="d0b52-556">&lt;optional&gt;</span></span>|<span data-ttu-id="d0b52-557">Данные о состоянии, передаваемые в асинхронный метод.</span><span class="sxs-lookup"><span data-stu-id="d0b52-557">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d0b52-558">Требования</span><span class="sxs-lookup"><span data-stu-id="d0b52-558">Requirements</span></span>

|<span data-ttu-id="d0b52-559">Требование</span><span class="sxs-lookup"><span data-stu-id="d0b52-559">Requirement</span></span>| <span data-ttu-id="d0b52-560">Значение</span><span class="sxs-lookup"><span data-stu-id="d0b52-560">Value</span></span>|
|---|---|
|[<span data-ttu-id="d0b52-561">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="d0b52-561">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d0b52-562">1.0</span><span class="sxs-lookup"><span data-stu-id="d0b52-562">1.0</span></span>|
|[<span data-ttu-id="d0b52-563">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="d0b52-563">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d0b52-564">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d0b52-564">ReadItem</span></span>|
|[<span data-ttu-id="d0b52-565">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="d0b52-565">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d0b52-566">Compose (создание) или read (чтение)</span><span class="sxs-lookup"><span data-stu-id="d0b52-566">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="d0b52-567">Пример</span><span class="sxs-lookup"><span data-stu-id="d0b52-567">Example</span></span>

```js
function getIdentityToken() {
  Office.context.mailbox.getUserIdentityTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

####  <a name="makeewsrequestasyncdata-callback-usercontext"></a><span data-ttu-id="d0b52-568">makeEwsRequestAsync(data, callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="d0b52-568">makeEwsRequestAsync(data, callback, [userContext])</span></span>

<span data-ttu-id="d0b52-569">Выполняет асинхронный запрос для веб-служб Exchange (EWS) на сервере Exchange Server, на котором размещен почтовый ящик пользователя.</span><span class="sxs-lookup"><span data-stu-id="d0b52-569">Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="d0b52-570">Этот метод не поддерживается в следующих сценариях.</span><span class="sxs-lookup"><span data-stu-id="d0b52-570">This method is not supported in the following scenarios.</span></span>
> - <span data-ttu-id="d0b52-571">В Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="d0b52-571">In Outlook for iOS or Outlook for Android</span></span>
> - <span data-ttu-id="d0b52-572">Когда надстройка загружается в почтовом ящике Gmail</span><span class="sxs-lookup"><span data-stu-id="d0b52-572">When the add-in is loaded in a Gmail mailbox</span></span>
> 
> <span data-ttu-id="d0b52-573">Вместо этого надстройкам следует использовать [API-интерфейсы REST](https://docs.microsoft.com/outlook/add-ins/use-rest-api) для доступа к почтовому ящику пользователя.</span><span class="sxs-lookup"><span data-stu-id="d0b52-573">In these cases, add-ins should [use REST APIs](https://docs.microsoft.com/outlook/add-ins/use-rest-api) to access the user's mailbox instead.</span></span>

<span data-ttu-id="d0b52-574">Метод `makeEwsRequestAsync` отправляет к Exchange EWS-запрос от имени надстройки.</span><span class="sxs-lookup"><span data-stu-id="d0b52-574">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange.</span></span> <span data-ttu-id="d0b52-575">См. [Вызов веб-служб из надстройки Outlook](https://docs.microsoft.com/outlook/add-ins/web-services#ews-operations-that-add-ins-support) для ознакомления с информацией о списке поддерживаемых операций веб-служб Exchange.</span><span class="sxs-lookup"><span data-stu-id="d0b52-575">See [Call web services from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span></span>

<span data-ttu-id="d0b52-576">С помощью метода `makeEwsRequestAsync` невозможно запрашивать элементы, связанные с папкой.</span><span class="sxs-lookup"><span data-stu-id="d0b52-576">You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.</span></span>

<span data-ttu-id="d0b52-577">XML-запрос должен указывать кодировку UTF-8.</span><span class="sxs-lookup"><span data-stu-id="d0b52-577">The XML request must specify UTF-8 encoding.</span></span>

```
<?xml version="1.0" encoding="utf-8"?>
```

<span data-ttu-id="d0b52-p142">У вашей надстройки должно быть разрешение **ReadWriteMailbox** для использования метода `makeEwsRequestAsync`. Сведения об использовании разрешения **ReadWriteMailbox** и операций EWS, которые можно вызывать с помощью метода `makeEwsRequestAsync`, см. в статье [Указание разрешений для доступа почтовой надстройки к почтовому ящику пользователя](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions).</span><span class="sxs-lookup"><span data-stu-id="d0b52-p142">Your add-in must have the **ReadWriteMailbox** permission to use the `makeEwsRequestAsync` method. For information about using the **ReadWriteMailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method, see [Specify permissions for mail add-in access to the user's mailbox](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>

> [!NOTE]
> <span data-ttu-id="d0b52-580">Администратор сервера должен установить для `OAuthAuthentication` значение true в каталоге сервера клиентского доступа EWS, чтобы включить метод  `makeEwsRequestAsync` для запросов служб EWS.</span><span class="sxs-lookup"><span data-stu-id="d0b52-580">NOTE: The server administrator must set `OAuthAuthentication` to true on the Client Access Server EWS directory to enable the `makeEwsRequestAsync` method to make EWS requests.</span></span>

##### <a name="version-differences"></a><span data-ttu-id="d0b52-581">Различия версий</span><span class="sxs-lookup"><span data-stu-id="d0b52-581">Version differences</span></span>

<span data-ttu-id="d0b52-582">Если вы используете метод `makeEwsRequestAsync` в почтовых приложениях, которые выполняются в версии Outlook, предшествующей 15.0.4535.1004, указывайте кодировку `ISO-8859-1`.</span><span class="sxs-lookup"><span data-stu-id="d0b52-582">When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set the encoding value to `ISO-8859-1`.</span></span>

```
<?xml version="1.0" encoding="iso-8859-1"?>
```

<span data-ttu-id="d0b52-p143">Значение кодировки не нужно указывать, если почтовое приложение выполняется в Outlook в Интернете. Чтобы определить, выполняется ли приложение в Outlook или Outlook в Интернете, используйте свойство mailbox.diagnostics.hostName. Используемую версию Outlook можно определить с помощью свойства mailbox.diagnostics.hostVersion.</span><span class="sxs-lookup"><span data-stu-id="d0b52-p143">You do not need to set the encoding value when your mail app is running in Outlook on the web. You can determine whether your mail app is running in Outlook or Outlook on the web by using the mailbox.diagnostics.hostName property. You can determine what version of Outlook is running by using the mailbox.diagnostics.hostVersion property.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d0b52-586">Параметры:</span><span class="sxs-lookup"><span data-stu-id="d0b52-586">Parameters:</span></span>

|<span data-ttu-id="d0b52-587">Имя</span><span class="sxs-lookup"><span data-stu-id="d0b52-587">Name</span></span>| <span data-ttu-id="d0b52-588">Тип</span><span class="sxs-lookup"><span data-stu-id="d0b52-588">Type</span></span>| <span data-ttu-id="d0b52-589">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="d0b52-589">Attributes</span></span>| <span data-ttu-id="d0b52-590">Описание</span><span class="sxs-lookup"><span data-stu-id="d0b52-590">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="d0b52-591">String</span><span class="sxs-lookup"><span data-stu-id="d0b52-591">String</span></span>||<span data-ttu-id="d0b52-592">Запрос EWS.</span><span class="sxs-lookup"><span data-stu-id="d0b52-592">The EWS request.</span></span>|
|`callback`| <span data-ttu-id="d0b52-593">function</span><span class="sxs-lookup"><span data-stu-id="d0b52-593">function</span></span>||<span data-ttu-id="d0b52-594">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="d0b52-594">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="d0b52-595">Результат XML вызова EWS указывается в виде строки в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="d0b52-595">The XML result of the EWS call is provided as a string in the `asyncResult.value` property.</span></span> <span data-ttu-id="d0b52-596">Если размер результата превышает 1 МБ, возвращается сообщение об ошибке.</span><span class="sxs-lookup"><span data-stu-id="d0b52-596">If the result exceeds 1 MB in size, an error message is returned instead.| | Object| optional|Any state data that is passed to the asynchronous method.|</span></span>|
|`userContext`| <span data-ttu-id="d0b52-597">Объект</span><span class="sxs-lookup"><span data-stu-id="d0b52-597">Object</span></span>| <span data-ttu-id="d0b52-598">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="d0b52-598">&lt;optional&gt;</span></span>|<span data-ttu-id="d0b52-599">Данные о состоянии, передаваемые в асинхронный метод.</span><span class="sxs-lookup"><span data-stu-id="d0b52-599">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d0b52-600">Требования</span><span class="sxs-lookup"><span data-stu-id="d0b52-600">Requirements</span></span>

|<span data-ttu-id="d0b52-601">Требование</span><span class="sxs-lookup"><span data-stu-id="d0b52-601">Requirement</span></span>| <span data-ttu-id="d0b52-602">Значение</span><span class="sxs-lookup"><span data-stu-id="d0b52-602">Value</span></span>|
|---|---|
|[<span data-ttu-id="d0b52-603">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="d0b52-603">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d0b52-604">1.0</span><span class="sxs-lookup"><span data-stu-id="d0b52-604">1.0</span></span>|
|[<span data-ttu-id="d0b52-605">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="d0b52-605">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d0b52-606">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="d0b52-606">ReadWriteMailbox</span></span>|
|[<span data-ttu-id="d0b52-607">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="d0b52-607">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d0b52-608">Compose (создание) или read (чтение)</span><span class="sxs-lookup"><span data-stu-id="d0b52-608">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="d0b52-609">Пример</span><span class="sxs-lookup"><span data-stu-id="d0b52-609">Example</span></span>

<span data-ttu-id="d0b52-610">В следующем примере вызывается `makeEwsRequestAsync` для получения темы элемента с помощью операции `GetItem`.</span><span class="sxs-lookup"><span data-stu-id="d0b52-610">The following example calls `makeEwsRequestAsync` to use the `GetItem` operation to get the subject of an item.</span></span>

```js
function getSubjectRequest(id) {
   // Return a GetItem operation request for the subject of the specified item.
   var request =
    '<?xml version="1.0" encoding="utf-8"?>' +
    '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"' +
    '               xmlns:xsd="http://www.w3.org/2001/XMLSchema"' +
    '               xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"' +
    '               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">' +
    '  <soap:Header>' +
    '    <RequestServerVersion Version="Exchange2013" xmlns="http://schemas.microsoft.com/exchange/services/2006/types" soap:mustUnderstand="0" />' +
    '  </soap:Header>' +
    '  <soap:Body>' +
    '    <GetItem xmlns="http://schemas.microsoft.com/exchange/services/2006/messages">' +
    '      <ItemShape>' +
    '        <t:BaseShape>IdOnly</t:BaseShape>' +
    '        <t:AdditionalProperties>' +
    '            <t:FieldURI FieldURI="item:Subject"/>' +
    '        </t:AdditionalProperties>' +
    '      </ItemShape>' +
    '      <ItemIds><t:ItemId Id="' + id + '"/></ItemIds>' +
    '    </GetItem>' +
    '  </soap:Body>' +
    '</soap:Envelope>';

   return request;
}

function sendRequest() {
   // Create a local variable that contains the mailbox.
   Office.context.mailbox.makeEwsRequestAsync(
    getSubjectRequest(mailbox.item.itemId), callback);
}

function callback(asyncResult)  {
   var result = asyncResult.value;
   var context = asyncResult.asyncContext;

   // Process the returned response here.
}
```