---
title: Office. Context. Mailbox — набор обязательных элементов 1,6
description: ''
ms.date: 10/21/2019
localization_priority: Normal
ms.openlocfilehash: b4bc64aa1ff836408a8b8b1efdaed7ddc8ce5725
ms.sourcegitcommit: 499bf49b41205f8034c501d4db5fe4b02dab205e
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 10/22/2019
ms.locfileid: "37627085"
---
# <a name="mailbox"></a><span data-ttu-id="26f37-102">mailbox</span><span class="sxs-lookup"><span data-stu-id="26f37-102">mailbox</span></span>

### <a name="officeofficemdcontextofficecontextmdmailbox"></a><span data-ttu-id="26f37-103">[Office](Office.md)[.context](Office.context.md).mailbox</span><span class="sxs-lookup"><span data-stu-id="26f37-103">[Office](Office.md)[.context](Office.context.md).mailbox</span></span>

<span data-ttu-id="26f37-104">Предоставляет для Microsoft Outlook доступ к объектной модели надстройки Outlook.</span><span class="sxs-lookup"><span data-stu-id="26f37-104">Provides access to the Outlook add-in object model for Microsoft Outlook.</span></span>

##### <a name="requirements"></a><span data-ttu-id="26f37-105">Требования</span><span class="sxs-lookup"><span data-stu-id="26f37-105">Requirements</span></span>

|<span data-ttu-id="26f37-106">Требование</span><span class="sxs-lookup"><span data-stu-id="26f37-106">Requirement</span></span>| <span data-ttu-id="26f37-107">Значение</span><span class="sxs-lookup"><span data-stu-id="26f37-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="26f37-108">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="26f37-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="26f37-109">1.0</span><span class="sxs-lookup"><span data-stu-id="26f37-109">1.0</span></span>|
|[<span data-ttu-id="26f37-110">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="26f37-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="26f37-111">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="26f37-111">Restricted</span></span>|
|[<span data-ttu-id="26f37-112">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="26f37-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="26f37-113">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="26f37-113">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="26f37-114">Элементы и методы</span><span class="sxs-lookup"><span data-stu-id="26f37-114">Members and methods</span></span>

| <span data-ttu-id="26f37-115">Элемент</span><span class="sxs-lookup"><span data-stu-id="26f37-115">Member</span></span> | <span data-ttu-id="26f37-116">Тип</span><span class="sxs-lookup"><span data-stu-id="26f37-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="26f37-117">ewsUrl</span><span class="sxs-lookup"><span data-stu-id="26f37-117">ewsUrl</span></span>](#ewsurl-string) | <span data-ttu-id="26f37-118">Элемент</span><span class="sxs-lookup"><span data-stu-id="26f37-118">Member</span></span> |
| [<span data-ttu-id="26f37-119">restUrl</span><span class="sxs-lookup"><span data-stu-id="26f37-119">restUrl</span></span>](#resturl-string) | <span data-ttu-id="26f37-120">Элемент</span><span class="sxs-lookup"><span data-stu-id="26f37-120">Member</span></span> |
| [<span data-ttu-id="26f37-121">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="26f37-121">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="26f37-122">Метод</span><span class="sxs-lookup"><span data-stu-id="26f37-122">Method</span></span> |
| [<span data-ttu-id="26f37-123">convertToEwsId</span><span class="sxs-lookup"><span data-stu-id="26f37-123">convertToEwsId</span></span>](#converttoewsiditemid-restversion--string) | <span data-ttu-id="26f37-124">Метод</span><span class="sxs-lookup"><span data-stu-id="26f37-124">Method</span></span> |
| [<span data-ttu-id="26f37-125">convertToLocalClientTime</span><span class="sxs-lookup"><span data-stu-id="26f37-125">convertToLocalClientTime</span></span>](#converttolocalclienttimetimevalue--localclienttime) | <span data-ttu-id="26f37-126">Метод</span><span class="sxs-lookup"><span data-stu-id="26f37-126">Method</span></span> |
| [<span data-ttu-id="26f37-127">convertToRestId</span><span class="sxs-lookup"><span data-stu-id="26f37-127">convertToRestId</span></span>](#converttorestiditemid-restversion--string) | <span data-ttu-id="26f37-128">Метод</span><span class="sxs-lookup"><span data-stu-id="26f37-128">Method</span></span> |
| [<span data-ttu-id="26f37-129">convertToUtcClientTime</span><span class="sxs-lookup"><span data-stu-id="26f37-129">convertToUtcClientTime</span></span>](#converttoutcclienttimeinput--date) | <span data-ttu-id="26f37-130">Метод</span><span class="sxs-lookup"><span data-stu-id="26f37-130">Method</span></span> |
| [<span data-ttu-id="26f37-131">displayAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="26f37-131">displayAppointmentForm</span></span>](#displayappointmentformitemid) | <span data-ttu-id="26f37-132">Метод</span><span class="sxs-lookup"><span data-stu-id="26f37-132">Method</span></span> |
| [<span data-ttu-id="26f37-133">displayMessageForm</span><span class="sxs-lookup"><span data-stu-id="26f37-133">displayMessageForm</span></span>](#displaymessageformitemid) | <span data-ttu-id="26f37-134">Метод</span><span class="sxs-lookup"><span data-stu-id="26f37-134">Method</span></span> |
| [<span data-ttu-id="26f37-135">displayNewAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="26f37-135">displayNewAppointmentForm</span></span>](#displaynewappointmentformparameters) | <span data-ttu-id="26f37-136">Метод</span><span class="sxs-lookup"><span data-stu-id="26f37-136">Method</span></span> |
| [<span data-ttu-id="26f37-137">дисплайневмессажеформ</span><span class="sxs-lookup"><span data-stu-id="26f37-137">displayNewMessageForm</span></span>](#displaynewmessageformparameters) | <span data-ttu-id="26f37-138">Метод</span><span class="sxs-lookup"><span data-stu-id="26f37-138">Method</span></span> |
| [<span data-ttu-id="26f37-139">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="26f37-139">getCallbackTokenAsync</span></span>](#getcallbacktokenasyncoptions-callback) | <span data-ttu-id="26f37-140">Метод</span><span class="sxs-lookup"><span data-stu-id="26f37-140">Method</span></span> |
| [<span data-ttu-id="26f37-141">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="26f37-141">getCallbackTokenAsync</span></span>](#getcallbacktokenasynccallback-usercontext) | <span data-ttu-id="26f37-142">Метод</span><span class="sxs-lookup"><span data-stu-id="26f37-142">Method</span></span> |
| [<span data-ttu-id="26f37-143">getUserIdentityTokenAsync</span><span class="sxs-lookup"><span data-stu-id="26f37-143">getUserIdentityTokenAsync</span></span>](#getuseridentitytokenasynccallback-usercontext) | <span data-ttu-id="26f37-144">Метод</span><span class="sxs-lookup"><span data-stu-id="26f37-144">Method</span></span> |
| [<span data-ttu-id="26f37-145">makeEwsRequestAsync</span><span class="sxs-lookup"><span data-stu-id="26f37-145">makeEwsRequestAsync</span></span>](#makeewsrequestasyncdata-callback-usercontext) | <span data-ttu-id="26f37-146">Метод</span><span class="sxs-lookup"><span data-stu-id="26f37-146">Method</span></span> |
| [<span data-ttu-id="26f37-147">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="26f37-147">removeHandlerAsync</span></span>](#removehandlerasynceventtype-options-callback) | <span data-ttu-id="26f37-148">Метод</span><span class="sxs-lookup"><span data-stu-id="26f37-148">Method</span></span> |

### <a name="namespaces"></a><span data-ttu-id="26f37-149">Пространства имен</span><span class="sxs-lookup"><span data-stu-id="26f37-149">Namespaces</span></span>

<span data-ttu-id="26f37-150">[diagnostics](Office.context.mailbox.diagnostics.md). Предоставляет надстройке Outlook диагностические сведения.</span><span class="sxs-lookup"><span data-stu-id="26f37-150">[diagnostics](Office.context.mailbox.diagnostics.md): Provides diagnostic information to an Outlook add-in.</span></span>

<span data-ttu-id="26f37-151">[item](Office.context.mailbox.item.md). Предоставляет методы и свойства для доступа к сообщению или встрече в надстройке Outlook.</span><span class="sxs-lookup"><span data-stu-id="26f37-151">[item](Office.context.mailbox.item.md): Provides methods and properties for accessing a message or appointment in an Outlook add-in.</span></span>

<span data-ttu-id="26f37-152">[userProfile](Office.context.mailbox.userProfile.md). Предоставляет сведения о пользователе в надстройке Outlook.</span><span class="sxs-lookup"><span data-stu-id="26f37-152">[userProfile](Office.context.mailbox.userProfile.md): Provides information about the user in an Outlook add-in.</span></span>

### <a name="members"></a><span data-ttu-id="26f37-153">Members</span><span class="sxs-lookup"><span data-stu-id="26f37-153">Members</span></span>

#### <a name="ewsurl-string"></a><span data-ttu-id="26f37-154">ewsUrl: String</span><span class="sxs-lookup"><span data-stu-id="26f37-154">ewsUrl: String</span></span>

<span data-ttu-id="26f37-p101">Получает URL-адрес конечной точки веб-служб Exchange (EWS) для этой учетной записи электронной почты. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="26f37-p101">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="26f37-157">Этот элемент не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="26f37-157">This member is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="26f37-p102">Удаленная служба может использовать значение `ewsUrl`, чтобы выполнять вызовы EWS для почтового ящика пользователя. Например, вы можете создать удаленную службу, чтобы [получить вложения из выбранного элемента](/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="26f37-p102">The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="26f37-160">Чтобы вызвать элемент `ewsUrl` в режиме чтения, в манифесте приложения должно быть указано разрешение **ReadItem**.</span><span class="sxs-lookup"><span data-stu-id="26f37-160">Your app must have the **ReadItem** permission specified in its manifest to call the `ewsUrl` member in read mode.</span></span>

<span data-ttu-id="26f37-p103">Перед использованием элемента `ewsUrl` в режиме создания необходимо вызвать метод [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback). Для вызова метода `saveAsync` приложение должно иметь разрешения **ReadWriteItem**.</span><span class="sxs-lookup"><span data-stu-id="26f37-p103">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `ewsUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="26f37-163">Тип</span><span class="sxs-lookup"><span data-stu-id="26f37-163">Type</span></span>

*   <span data-ttu-id="26f37-164">String</span><span class="sxs-lookup"><span data-stu-id="26f37-164">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="26f37-165">Требования</span><span class="sxs-lookup"><span data-stu-id="26f37-165">Requirements</span></span>

|<span data-ttu-id="26f37-166">Требование</span><span class="sxs-lookup"><span data-stu-id="26f37-166">Requirement</span></span>| <span data-ttu-id="26f37-167">Значение</span><span class="sxs-lookup"><span data-stu-id="26f37-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="26f37-168">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="26f37-168">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="26f37-169">1.0</span><span class="sxs-lookup"><span data-stu-id="26f37-169">1.0</span></span>|
|[<span data-ttu-id="26f37-170">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="26f37-170">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="26f37-171">ReadItem</span><span class="sxs-lookup"><span data-stu-id="26f37-171">ReadItem</span></span>|
|[<span data-ttu-id="26f37-172">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="26f37-172">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="26f37-173">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="26f37-173">Compose or Read</span></span>|

<br>

---
---

#### <a name="resturl-string"></a><span data-ttu-id="26f37-174">restUrl: String</span><span class="sxs-lookup"><span data-stu-id="26f37-174">restUrl: String</span></span>

<span data-ttu-id="26f37-175">Возвращает URL-адрес конечной точки REST для этой учетной записи электронной почты.</span><span class="sxs-lookup"><span data-stu-id="26f37-175">Gets the URL of the REST endpoint for this email account.</span></span>

<span data-ttu-id="26f37-176">С помощью значения `restUrl` можно выполнять вызовы [REST API](/outlook/rest/) для почтового ящика пользователя.</span><span class="sxs-lookup"><span data-stu-id="26f37-176">The `restUrl` value can be used to make [REST API](/outlook/rest/) calls to the user's mailbox.</span></span>

<span data-ttu-id="26f37-177">Чтобы вызвать элемент `restUrl` в режиме чтения, в манифесте приложения необходимо указать разрешение **ReadItem**.</span><span class="sxs-lookup"><span data-stu-id="26f37-177">Your app must have the **ReadItem** permission specified in its manifest to call the `restUrl` member in read mode.</span></span>

<span data-ttu-id="26f37-p104">Перед использованием элемента `restUrl` в режиме создания необходимо вызвать метод [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback). Для вызова метода `saveAsync` приложение должно иметь разрешения **ReadWriteItem**.</span><span class="sxs-lookup"><span data-stu-id="26f37-p104">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `restUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="26f37-180">Тип</span><span class="sxs-lookup"><span data-stu-id="26f37-180">Type</span></span>

*   <span data-ttu-id="26f37-181">String</span><span class="sxs-lookup"><span data-stu-id="26f37-181">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="26f37-182">Требования</span><span class="sxs-lookup"><span data-stu-id="26f37-182">Requirements</span></span>

|<span data-ttu-id="26f37-183">Требование</span><span class="sxs-lookup"><span data-stu-id="26f37-183">Requirement</span></span>| <span data-ttu-id="26f37-184">Значение</span><span class="sxs-lookup"><span data-stu-id="26f37-184">Value</span></span>|
|---|---|
|[<span data-ttu-id="26f37-185">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="26f37-185">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="26f37-186">1.5</span><span class="sxs-lookup"><span data-stu-id="26f37-186">1.5</span></span> |
|[<span data-ttu-id="26f37-187">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="26f37-187">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="26f37-188">ReadItem</span><span class="sxs-lookup"><span data-stu-id="26f37-188">ReadItem</span></span>|
|[<span data-ttu-id="26f37-189">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="26f37-189">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="26f37-190">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="26f37-190">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="26f37-191">Методы</span><span class="sxs-lookup"><span data-stu-id="26f37-191">Methods</span></span>

#### <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="26f37-192">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="26f37-192">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="26f37-193">Добавляет обработчик для поддерживаемого события.</span><span class="sxs-lookup"><span data-stu-id="26f37-193">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="26f37-194">В настоящее время поддерживаются только события типа `Office.EventType.ItemChanged`, которые вызываются, когда пользователь выбирает новый элемент.</span><span class="sxs-lookup"><span data-stu-id="26f37-194">Currently the only supported event type is `Office.EventType.ItemChanged`, which is invoked when the user selects a new item.</span></span> <span data-ttu-id="26f37-195">Это событие используется надстройками, реализующими закрепляемую область задач, и позволяет надстройке обновлять пользовательский интерфейс области задач в соответствии с выбранным в данный момент элементом.</span><span class="sxs-lookup"><span data-stu-id="26f37-195">This event is used by add-ins that implement a pinnable task pane, and allows the add-in to refresh the task pane UI based on the currently selected item.</span></span>

##### <a name="parameters"></a><span data-ttu-id="26f37-196">Параметры</span><span class="sxs-lookup"><span data-stu-id="26f37-196">Parameters</span></span>

| <span data-ttu-id="26f37-197">Имя</span><span class="sxs-lookup"><span data-stu-id="26f37-197">Name</span></span> | <span data-ttu-id="26f37-198">Тип</span><span class="sxs-lookup"><span data-stu-id="26f37-198">Type</span></span> | <span data-ttu-id="26f37-199">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="26f37-199">Attributes</span></span> | <span data-ttu-id="26f37-200">Описание</span><span class="sxs-lookup"><span data-stu-id="26f37-200">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="26f37-201">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="26f37-201">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="26f37-202">Событие, которое должно вызвать обработчик.</span><span class="sxs-lookup"><span data-stu-id="26f37-202">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="26f37-203">Function</span><span class="sxs-lookup"><span data-stu-id="26f37-203">Function</span></span> || <span data-ttu-id="26f37-p106">Функция для обработки события. Функция должна принимать один параметр, представляющий собой объектный литерал. Значение свойства `type` параметра совпадет со значением параметра `eventType`, переданного методу `addHandlerAsync`.</span><span class="sxs-lookup"><span data-stu-id="26f37-p106">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="26f37-207">Объект</span><span class="sxs-lookup"><span data-stu-id="26f37-207">Object</span></span> | <span data-ttu-id="26f37-208">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="26f37-208">&lt;optional&gt;</span></span> | <span data-ttu-id="26f37-209">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="26f37-209">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="26f37-210">Object</span><span class="sxs-lookup"><span data-stu-id="26f37-210">Object</span></span> | <span data-ttu-id="26f37-211">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="26f37-211">&lt;optional&gt;</span></span> | <span data-ttu-id="26f37-212">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="26f37-212">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="26f37-213">функция</span><span class="sxs-lookup"><span data-stu-id="26f37-213">function</span></span>| <span data-ttu-id="26f37-214">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="26f37-214">&lt;optional&gt;</span></span>|<span data-ttu-id="26f37-215">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="26f37-215">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="26f37-216">Требования</span><span class="sxs-lookup"><span data-stu-id="26f37-216">Requirements</span></span>

|<span data-ttu-id="26f37-217">Требование</span><span class="sxs-lookup"><span data-stu-id="26f37-217">Requirement</span></span>| <span data-ttu-id="26f37-218">Значение</span><span class="sxs-lookup"><span data-stu-id="26f37-218">Value</span></span>|
|---|---|
|[<span data-ttu-id="26f37-219">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="26f37-219">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="26f37-220">1.5</span><span class="sxs-lookup"><span data-stu-id="26f37-220">1.5</span></span> |
|[<span data-ttu-id="26f37-221">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="26f37-221">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="26f37-222">ReadItem</span><span class="sxs-lookup"><span data-stu-id="26f37-222">ReadItem</span></span> |
|[<span data-ttu-id="26f37-223">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="26f37-223">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="26f37-224">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="26f37-224">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="26f37-225">Пример</span><span class="sxs-lookup"><span data-stu-id="26f37-225">Example</span></span>

```js
Office.initialize = function (reason) {
  $(document).ready(function () {
    Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, loadNewItem, function (result) {
      if (result.status === Office.AsyncResultStatus.Failed) {
        // Handle error.
      }
    });
  });
};

function loadNewItem(eventArgs) {
  // Load the properties of the newly selected item.
  loadProps(Office.context.mailbox.item);
};
```

<br>

---
---

#### <a name="converttoewsiditemid-restversion--string"></a><span data-ttu-id="26f37-226">convertToEwsId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="26f37-226">convertToEwsId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="26f37-227">Преобразовывает идентификатор элемента из формата REST в формат EWS.</span><span class="sxs-lookup"><span data-stu-id="26f37-227">Converts an item ID formatted for REST into EWS format.</span></span>

> [!NOTE]
> <span data-ttu-id="26f37-228">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="26f37-228">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="26f37-p107">Формат идентификаторов, извлекаемых через API REST (например, [API Почты Outlook](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) или [Microsoft Graph](https://graph.microsoft.io/)), отличается от формата веб-служб Exchange (EWS). Метод `convertToEwsId` преобразовывает идентификатор в формате REST в формат EWS.</span><span class="sxs-lookup"><span data-stu-id="26f37-p107">Item IDs retrieved via a REST API (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)) use a different format than the format used by Exchange Web Services (EWS). The `convertToEwsId` method converts a REST-formatted ID into the proper format for EWS.</span></span>

##### <a name="parameters"></a><span data-ttu-id="26f37-231">Параметры</span><span class="sxs-lookup"><span data-stu-id="26f37-231">Parameters</span></span>

|<span data-ttu-id="26f37-232">Имя</span><span class="sxs-lookup"><span data-stu-id="26f37-232">Name</span></span>| <span data-ttu-id="26f37-233">Тип</span><span class="sxs-lookup"><span data-stu-id="26f37-233">Type</span></span>| <span data-ttu-id="26f37-234">Описание</span><span class="sxs-lookup"><span data-stu-id="26f37-234">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="26f37-235">String</span><span class="sxs-lookup"><span data-stu-id="26f37-235">String</span></span>|<span data-ttu-id="26f37-236">Идентификатор элемента в формате REST API для Outlook</span><span class="sxs-lookup"><span data-stu-id="26f37-236">An item ID formatted for the Outlook REST APIs</span></span>|
|`restVersion`| [<span data-ttu-id="26f37-237">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="26f37-237">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion?view=outlook-js-1.6)|<span data-ttu-id="26f37-238">Значение, определяющее версию REST API для Outlook, которая используется для извлечения идентификатора элемента.</span><span class="sxs-lookup"><span data-stu-id="26f37-238">A value indicating the version of the Outlook REST API used to retrieve the item ID.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="26f37-239">Требования</span><span class="sxs-lookup"><span data-stu-id="26f37-239">Requirements</span></span>

|<span data-ttu-id="26f37-240">Требование</span><span class="sxs-lookup"><span data-stu-id="26f37-240">Requirement</span></span>| <span data-ttu-id="26f37-241">Значение</span><span class="sxs-lookup"><span data-stu-id="26f37-241">Value</span></span>|
|---|---|
|[<span data-ttu-id="26f37-242">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="26f37-242">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="26f37-243">1.3</span><span class="sxs-lookup"><span data-stu-id="26f37-243">1.3</span></span>|
|[<span data-ttu-id="26f37-244">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="26f37-244">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="26f37-245">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="26f37-245">Restricted</span></span>|
|[<span data-ttu-id="26f37-246">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="26f37-246">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="26f37-247">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="26f37-247">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="26f37-248">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="26f37-248">Returns:</span></span>

<span data-ttu-id="26f37-249">Тип: String</span><span class="sxs-lookup"><span data-stu-id="26f37-249">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="26f37-250">Пример</span><span class="sxs-lookup"><span data-stu-id="26f37-250">Example</span></span>

```js
// Get an item's ID from a REST API.
var restId = 'AAMkAGVlOTZjNTM3LW...';

// Treat restId as coming from the v2.0 version of the Outlook Mail API.
var ewsId = Office.context.mailbox.convertToEwsId(restId, Office.MailboxEnums.RestVersion.v2_0);
```

<br>

---
---

#### <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttimeviewoutlook-js-16"></a><span data-ttu-id="26f37-251">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.6)}</span><span class="sxs-lookup"><span data-stu-id="26f37-251">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.6)}</span></span>

<span data-ttu-id="26f37-252">Получает словарь, содержащий сведения о локальном времени клиента.</span><span class="sxs-lookup"><span data-stu-id="26f37-252">Gets a dictionary containing time information in local client time.</span></span>

<span data-ttu-id="26f37-p108">Почтовое приложение для классической версии Outlook или версии в Интернете может использовать разные часовые пояса для дат и времени. Классическое приложение Outlook использует часовой пояс клиентского компьютера. Outlook в Интернете использует часовой пояс, заданный в Центре администрирования Exchange (EAC). Значения даты и времени должны обрабатываться так, чтобы значения в пользовательском интерфейсе всегда согласовывались с часовым поясом, ожидаемым пользователем.</span><span class="sxs-lookup"><span data-stu-id="26f37-p108">A mail app for Outlook on a desktop or on the web can use different time zones for the dates and times. Outlook on a desktop uses the client computer time zone; Outlook on the web uses the time zone set on the Exchange Admin Center (EAC). You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that the user expects.</span></span>

<span data-ttu-id="26f37-p109">Если почтовое приложение работает в классическом клиенте Outlook, метод `convertToLocalClientTime` вернет объект словаря со значениями часового пояса клиентского компьютера. Если почтовое приложение работает в Outlook в Интернете, метод `convertToLocalClientTime` вернет объект словаря со значениями часового пояса, заданного в Центре администрирования Exchange.</span><span class="sxs-lookup"><span data-stu-id="26f37-p109">If the mail app is running in Outlook on a desktop client, the `convertToLocalClientTime` method will return a dictionary object with the values set to the client computer time zone. If the mail app is running in Outlook on the web, the `convertToLocalClientTime` method will return a dictionary object with the values set to the time zone specified in the EAC.</span></span>

##### <a name="parameters"></a><span data-ttu-id="26f37-258">Параметры</span><span class="sxs-lookup"><span data-stu-id="26f37-258">Parameters</span></span>

|<span data-ttu-id="26f37-259">Имя</span><span class="sxs-lookup"><span data-stu-id="26f37-259">Name</span></span>| <span data-ttu-id="26f37-260">Тип</span><span class="sxs-lookup"><span data-stu-id="26f37-260">Type</span></span>| <span data-ttu-id="26f37-261">Описание</span><span class="sxs-lookup"><span data-stu-id="26f37-261">Description</span></span>|
|---|---|---|
|`timeValue`| <span data-ttu-id="26f37-262">Date</span><span class="sxs-lookup"><span data-stu-id="26f37-262">Date</span></span>|<span data-ttu-id="26f37-263">Объект Date</span><span class="sxs-lookup"><span data-stu-id="26f37-263">A Date object</span></span>|

##### <a name="requirements"></a><span data-ttu-id="26f37-264">Требования</span><span class="sxs-lookup"><span data-stu-id="26f37-264">Requirements</span></span>

|<span data-ttu-id="26f37-265">Требование</span><span class="sxs-lookup"><span data-stu-id="26f37-265">Requirement</span></span>| <span data-ttu-id="26f37-266">Значение</span><span class="sxs-lookup"><span data-stu-id="26f37-266">Value</span></span>|
|---|---|
|[<span data-ttu-id="26f37-267">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="26f37-267">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="26f37-268">1.0</span><span class="sxs-lookup"><span data-stu-id="26f37-268">1.0</span></span>|
|[<span data-ttu-id="26f37-269">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="26f37-269">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="26f37-270">ReadItem</span><span class="sxs-lookup"><span data-stu-id="26f37-270">ReadItem</span></span>|
|[<span data-ttu-id="26f37-271">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="26f37-271">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="26f37-272">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="26f37-272">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="26f37-273">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="26f37-273">Returns:</span></span>

<span data-ttu-id="26f37-274">Тип: [LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="26f37-274">Type: [LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.6)</span></span>

<br>

---
---

#### <a name="converttorestiditemid-restversion--string"></a><span data-ttu-id="26f37-275">convertToRestId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="26f37-275">convertToRestId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="26f37-276">Преобразовывает идентификатор элемента в формате EWS в формат REST.</span><span class="sxs-lookup"><span data-stu-id="26f37-276">Converts an item ID formatted for EWS into REST format.</span></span>

> [!NOTE]
> <span data-ttu-id="26f37-277">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="26f37-277">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="26f37-p110">Формат идентификаторов, извлекаемых через EWS или свойство `itemId`, отличается от формата API REST (таких как [API Почты Outlook](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) или [Microsoft Graph](https://graph.microsoft.io/)). Метод `convertToRestId` преобразовывает идентификатор в формате EWS в формат REST.</span><span class="sxs-lookup"><span data-stu-id="26f37-p110">Item IDs retrieved via EWS or via the `itemId` property use a different format than the format used by REST APIs (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)). The `convertToRestId` method converts an EWS-formatted ID into the proper format for REST.</span></span>

##### <a name="parameters"></a><span data-ttu-id="26f37-280">Параметры</span><span class="sxs-lookup"><span data-stu-id="26f37-280">Parameters</span></span>

|<span data-ttu-id="26f37-281">Имя</span><span class="sxs-lookup"><span data-stu-id="26f37-281">Name</span></span>| <span data-ttu-id="26f37-282">Тип</span><span class="sxs-lookup"><span data-stu-id="26f37-282">Type</span></span>| <span data-ttu-id="26f37-283">Описание</span><span class="sxs-lookup"><span data-stu-id="26f37-283">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="26f37-284">String</span><span class="sxs-lookup"><span data-stu-id="26f37-284">String</span></span>|<span data-ttu-id="26f37-285">Идентификатор элемента в формате EWS</span><span class="sxs-lookup"><span data-stu-id="26f37-285">An item ID formatted for Exchange Web Services (EWS)</span></span>|
|`restVersion`| [<span data-ttu-id="26f37-286">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="26f37-286">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion?view=outlook-js-1.6)|<span data-ttu-id="26f37-287">Значение, определяющее версию REST API для Outlook, с которой будет использоваться преобразованный идентификатор.</span><span class="sxs-lookup"><span data-stu-id="26f37-287">A value indicating the version of the Outlook REST API that the converted ID will be used with.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="26f37-288">Требования</span><span class="sxs-lookup"><span data-stu-id="26f37-288">Requirements</span></span>

|<span data-ttu-id="26f37-289">Требование</span><span class="sxs-lookup"><span data-stu-id="26f37-289">Requirement</span></span>| <span data-ttu-id="26f37-290">Значение</span><span class="sxs-lookup"><span data-stu-id="26f37-290">Value</span></span>|
|---|---|
|[<span data-ttu-id="26f37-291">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="26f37-291">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="26f37-292">1.3</span><span class="sxs-lookup"><span data-stu-id="26f37-292">1.3</span></span>|
|[<span data-ttu-id="26f37-293">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="26f37-293">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="26f37-294">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="26f37-294">Restricted</span></span>|
|[<span data-ttu-id="26f37-295">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="26f37-295">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="26f37-296">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="26f37-296">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="26f37-297">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="26f37-297">Returns:</span></span>

<span data-ttu-id="26f37-298">Тип: String</span><span class="sxs-lookup"><span data-stu-id="26f37-298">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="26f37-299">Пример</span><span class="sxs-lookup"><span data-stu-id="26f37-299">Example</span></span>

```js
// Get the currently selected item's ID.
var ewsId = Office.context.mailbox.item.itemId;

// Convert to a REST ID for the v2.0 version of the Outlook Mail API.
var restId = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
```

<br>

---
---

#### <a name="converttoutcclienttimeinput--date"></a><span data-ttu-id="26f37-300">convertToUtcClientTime(input) → {Date}</span><span class="sxs-lookup"><span data-stu-id="26f37-300">convertToUtcClientTime(input) → {Date}</span></span>

<span data-ttu-id="26f37-301">Получает объект Date из словаря, содержащего сведения о времени.</span><span class="sxs-lookup"><span data-stu-id="26f37-301">Gets a Date object from a dictionary containing time information.</span></span>

<span data-ttu-id="26f37-302">Метод `convertToUtcClientTime` преобразует словарь, содержащий локальную дату и время, в объект Date с правильными значениями локальной даты и времени.</span><span class="sxs-lookup"><span data-stu-id="26f37-302">The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a Date object with the correct values for the local date and time.</span></span>

##### <a name="parameters"></a><span data-ttu-id="26f37-303">Параметры</span><span class="sxs-lookup"><span data-stu-id="26f37-303">Parameters</span></span>

|<span data-ttu-id="26f37-304">Имя</span><span class="sxs-lookup"><span data-stu-id="26f37-304">Name</span></span>| <span data-ttu-id="26f37-305">Тип</span><span class="sxs-lookup"><span data-stu-id="26f37-305">Type</span></span>| <span data-ttu-id="26f37-306">Описание</span><span class="sxs-lookup"><span data-stu-id="26f37-306">Description</span></span>|
|---|---|---|
|`input`| [<span data-ttu-id="26f37-307">LocalClientTime</span><span class="sxs-lookup"><span data-stu-id="26f37-307">LocalClientTime</span></span>](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.6)|<span data-ttu-id="26f37-308">Значение локального времени для преобразования.</span><span class="sxs-lookup"><span data-stu-id="26f37-308">The local time value to convert.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="26f37-309">Требования</span><span class="sxs-lookup"><span data-stu-id="26f37-309">Requirements</span></span>

|<span data-ttu-id="26f37-310">Требование</span><span class="sxs-lookup"><span data-stu-id="26f37-310">Requirement</span></span>| <span data-ttu-id="26f37-311">Значение</span><span class="sxs-lookup"><span data-stu-id="26f37-311">Value</span></span>|
|---|---|
|[<span data-ttu-id="26f37-312">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="26f37-312">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="26f37-313">1.0</span><span class="sxs-lookup"><span data-stu-id="26f37-313">1.0</span></span>|
|[<span data-ttu-id="26f37-314">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="26f37-314">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="26f37-315">ReadItem</span><span class="sxs-lookup"><span data-stu-id="26f37-315">ReadItem</span></span>|
|[<span data-ttu-id="26f37-316">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="26f37-316">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="26f37-317">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="26f37-317">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="26f37-318">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="26f37-318">Returns:</span></span>

<span data-ttu-id="26f37-319">Объект Date со временем в формате UTC.</span><span class="sxs-lookup"><span data-stu-id="26f37-319">A Date object with the time expressed in UTC.</span></span>

<span data-ttu-id="26f37-320">Тип: Date</span><span class="sxs-lookup"><span data-stu-id="26f37-320">Type: Date</span></span>

##### <a name="example"></a><span data-ttu-id="26f37-321">Пример</span><span class="sxs-lookup"><span data-stu-id="26f37-321">Example</span></span>

```js
// Represents 3:37 PM PDT on Monday, August 26, 2019.
var input = {
  date: 26,
  hours: 15,
  milliseconds: 2,
  minutes: 37,
  month: 7,
  seconds: 2,
  timezoneOffset: -420,
  year: 2019
};

// result should be a Date object.
var result = Office.context.mailbox.convertToUtcClientTime(input);

// Output should be "2019-08-26T22:37:02.002Z".
console.log(result.toISOString());
```

<br>

---
---

#### <a name="displayappointmentformitemid"></a><span data-ttu-id="26f37-322">displayAppointmentForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="26f37-322">displayAppointmentForm(itemId)</span></span>

<span data-ttu-id="26f37-323">Отображает имеющуюся встречу из календаря.</span><span class="sxs-lookup"><span data-stu-id="26f37-323">Displays an existing calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="26f37-324">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="26f37-324">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="26f37-325">Метод `displayAppointmentForm` открывает новое окно на компьютере или диалоговое окно на мобильном устройстве, содержащее сведения календаря о существующей встрече.</span><span class="sxs-lookup"><span data-stu-id="26f37-325">The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="26f37-p111">В Outlook для Mac с помощью этого метода можно отобразить одну встречу, которая не является частью повторяющегося ряда, или основную встречу такого ряда, но не экземпляр из него, так как в Outlook для Mac невозможно получить доступ к свойствам экземпляра повторяющегося ряда (в том числе к идентификатору элемента).</span><span class="sxs-lookup"><span data-stu-id="26f37-p111">In Outlook on Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment of a recurring series, but you cannot display an instance of the series. This is because in Outlook on Mac, you cannot access the properties (including the item ID) of instances of a recurring series.</span></span>

<span data-ttu-id="26f37-328">В Outlook в Интернете этот метод открывает указанную форму, только если текст формы содержит символы размером не более 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="26f37-328">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32KB number of characters.</span></span>

<span data-ttu-id="26f37-329">Если указанный идентификатор элемента не определяет существующую встречу, на клиентском компьютере или устройстве открывается пустая страница, и сообщение об ошибке не возвращается.</span><span class="sxs-lookup"><span data-stu-id="26f37-329">If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and no error message will be returned.</span></span>

##### <a name="parameters"></a><span data-ttu-id="26f37-330">Параметры</span><span class="sxs-lookup"><span data-stu-id="26f37-330">Parameters</span></span>

|<span data-ttu-id="26f37-331">Имя</span><span class="sxs-lookup"><span data-stu-id="26f37-331">Name</span></span>| <span data-ttu-id="26f37-332">Тип</span><span class="sxs-lookup"><span data-stu-id="26f37-332">Type</span></span>| <span data-ttu-id="26f37-333">Описание</span><span class="sxs-lookup"><span data-stu-id="26f37-333">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="26f37-334">String</span><span class="sxs-lookup"><span data-stu-id="26f37-334">String</span></span>|<span data-ttu-id="26f37-335">Идентификатор веб-служб Exchange для существующей встречи в календаре.</span><span class="sxs-lookup"><span data-stu-id="26f37-335">The Exchange Web Services (EWS) identifier for an existing calendar appointment.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="26f37-336">Требования</span><span class="sxs-lookup"><span data-stu-id="26f37-336">Requirements</span></span>

|<span data-ttu-id="26f37-337">Требование</span><span class="sxs-lookup"><span data-stu-id="26f37-337">Requirement</span></span>| <span data-ttu-id="26f37-338">Значение</span><span class="sxs-lookup"><span data-stu-id="26f37-338">Value</span></span>|
|---|---|
|[<span data-ttu-id="26f37-339">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="26f37-339">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="26f37-340">1.0</span><span class="sxs-lookup"><span data-stu-id="26f37-340">1.0</span></span>|
|[<span data-ttu-id="26f37-341">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="26f37-341">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="26f37-342">ReadItem</span><span class="sxs-lookup"><span data-stu-id="26f37-342">ReadItem</span></span>|
|[<span data-ttu-id="26f37-343">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="26f37-343">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="26f37-344">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="26f37-344">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="26f37-345">Пример</span><span class="sxs-lookup"><span data-stu-id="26f37-345">Example</span></span>

```js
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

<br>

---
---

#### <a name="displaymessageformitemid"></a><span data-ttu-id="26f37-346">displayMessageForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="26f37-346">displayMessageForm(itemId)</span></span>

<span data-ttu-id="26f37-347">Отображает имеющееся сообщение.</span><span class="sxs-lookup"><span data-stu-id="26f37-347">Displays an existing message.</span></span>

> [!NOTE]
> <span data-ttu-id="26f37-348">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="26f37-348">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="26f37-349">Метод `displayMessageForm` открывает новое окно на компьютере или диалоговое окно на мобильном устройстве, содержащее существующее сообщение.</span><span class="sxs-lookup"><span data-stu-id="26f37-349">The `displayMessageForm` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="26f37-350">В Outlook в Интернете этот метод открывает указанную форму, только если текст формы содержит символы размером не более 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="26f37-350">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32 KB number of characters.</span></span>

<span data-ttu-id="26f37-351">Если указанный идентификатор элемента не определяет существующее сообщение, окно на клиентском компьютере не открывается и сообщение об ошибке не возвращается.</span><span class="sxs-lookup"><span data-stu-id="26f37-351">If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and no error message will be returned.</span></span>

<span data-ttu-id="26f37-p112">Не используйте `displayMessageForm` с параметром `itemId`, который представляет собой встречу. Используйте метод `displayAppointmentForm`, чтобы отобразить сведения о существующей встрече, а метод `displayNewAppointmentForm` — для отображения формы создания встречи.</span><span class="sxs-lookup"><span data-stu-id="26f37-p112">Do not use the `displayMessageForm` with an `itemId` that represents an appointment. Use the `displayAppointmentForm` method to display an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.</span></span>

##### <a name="parameters"></a><span data-ttu-id="26f37-354">Параметры</span><span class="sxs-lookup"><span data-stu-id="26f37-354">Parameters</span></span>

|<span data-ttu-id="26f37-355">Имя</span><span class="sxs-lookup"><span data-stu-id="26f37-355">Name</span></span>| <span data-ttu-id="26f37-356">Тип</span><span class="sxs-lookup"><span data-stu-id="26f37-356">Type</span></span>| <span data-ttu-id="26f37-357">Описание</span><span class="sxs-lookup"><span data-stu-id="26f37-357">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="26f37-358">String</span><span class="sxs-lookup"><span data-stu-id="26f37-358">String</span></span>|<span data-ttu-id="26f37-359">Идентификатор веб-служб Exchange для существующего сообщения.</span><span class="sxs-lookup"><span data-stu-id="26f37-359">The Exchange Web Services (EWS) identifier for an existing message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="26f37-360">Требования</span><span class="sxs-lookup"><span data-stu-id="26f37-360">Requirements</span></span>

|<span data-ttu-id="26f37-361">Требование</span><span class="sxs-lookup"><span data-stu-id="26f37-361">Requirement</span></span>| <span data-ttu-id="26f37-362">Значение</span><span class="sxs-lookup"><span data-stu-id="26f37-362">Value</span></span>|
|---|---|
|[<span data-ttu-id="26f37-363">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="26f37-363">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="26f37-364">1.0</span><span class="sxs-lookup"><span data-stu-id="26f37-364">1.0</span></span>|
|[<span data-ttu-id="26f37-365">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="26f37-365">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="26f37-366">ReadItem</span><span class="sxs-lookup"><span data-stu-id="26f37-366">ReadItem</span></span>|
|[<span data-ttu-id="26f37-367">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="26f37-367">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="26f37-368">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="26f37-368">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="26f37-369">Пример</span><span class="sxs-lookup"><span data-stu-id="26f37-369">Example</span></span>

```js
Office.context.mailbox.displayMessageForm(messageId);
```

<br>

---
---

#### <a name="displaynewappointmentformparameters"></a><span data-ttu-id="26f37-370">displayNewAppointmentForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="26f37-370">displayNewAppointmentForm(parameters)</span></span>

<span data-ttu-id="26f37-371">Отображает форму для создания новой встречи в календаре.</span><span class="sxs-lookup"><span data-stu-id="26f37-371">Displays a form for creating a new calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="26f37-372">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="26f37-372">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="26f37-p113">Метод `displayNewAppointmentForm` открывает форму, в которой пользователь может создать встречу или собрание. Если параметры заданы, поля формы встречи автоматически заполняются их содержимым.</span><span class="sxs-lookup"><span data-stu-id="26f37-p113">The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting. If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="26f37-p114">В Outlook в Интернете и на мобильных устройствах этот метод всегда отображает форму с полем участников. Если вы не укажете участников в качестве входных аргументов, метод отображает форму с кнопкой **Сохранить**. Если вы укажете участников, форма будет включать участников и кнопку **Отправить**.</span><span class="sxs-lookup"><span data-stu-id="26f37-p114">In Outlook on the web and mobile devices, this method always displays a form with an attendees field. If you do not specify any attendees as input arguments, the method displays a form with a **Save** button. If you have specified attendees, the form would include the attendees and a **Send** button.</span></span>

<span data-ttu-id="26f37-p115">Если вы укажете участников или ресурсы с помощью параметра `requiredAttendees`, `optionalAttendees` или `resources` в клиенте Outlook с расширенными возможностями и Outlook RT, этот метод отобразит форму собрания с кнопкой **Отправить**. Если не указать получателей, этот метод отобразит форму встречи с кнопкой **Сохранить и закрыть**.</span><span class="sxs-lookup"><span data-stu-id="26f37-p115">In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or `resources` parameter, this method displays a meeting form with a **Send** button. If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.</span></span>

<span data-ttu-id="26f37-380">Если параметры превышают указанные ограничения размера или если указано неизвестное имя параметра, вызывается исключение.</span><span class="sxs-lookup"><span data-stu-id="26f37-380">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="26f37-381">Параметры</span><span class="sxs-lookup"><span data-stu-id="26f37-381">Parameters</span></span>

> [!NOTE]
> <span data-ttu-id="26f37-382">Все параметры являются необязательными.</span><span class="sxs-lookup"><span data-stu-id="26f37-382">All parameters are optional.</span></span>

|<span data-ttu-id="26f37-383">Имя</span><span class="sxs-lookup"><span data-stu-id="26f37-383">Name</span></span>| <span data-ttu-id="26f37-384">Тип</span><span class="sxs-lookup"><span data-stu-id="26f37-384">Type</span></span>| <span data-ttu-id="26f37-385">Описание</span><span class="sxs-lookup"><span data-stu-id="26f37-385">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="26f37-386">Object</span><span class="sxs-lookup"><span data-stu-id="26f37-386">Object</span></span> | <span data-ttu-id="26f37-387">Словарь параметров, описывающий новую встречу.</span><span class="sxs-lookup"><span data-stu-id="26f37-387">A dictionary of parameters describing the new appointment.</span></span> |
| `parameters.requiredAttendees` | <span data-ttu-id="26f37-388">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)&gt;</span><span class="sxs-lookup"><span data-stu-id="26f37-388">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)&gt;</span></span> | <span data-ttu-id="26f37-p116">Массив строк, содержащий электронные адреса, или массив, содержащий объекты `EmailAddressDetails` для каждого из обязательных участников встречи. Массив может включать не более 100 записей.</span><span class="sxs-lookup"><span data-stu-id="26f37-p116">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the required attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.optionalAttendees` | <span data-ttu-id="26f37-391">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)&gt;</span><span class="sxs-lookup"><span data-stu-id="26f37-391">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)&gt;</span></span> | <span data-ttu-id="26f37-p117">Массив строк, содержащий электронные адреса, или массив, содержащий объекты `EmailAddressDetails` для каждого из необязательных участников встречи. Массив может включать не более 100 записей.</span><span class="sxs-lookup"><span data-stu-id="26f37-p117">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the optional attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.start` | <span data-ttu-id="26f37-394">Date</span><span class="sxs-lookup"><span data-stu-id="26f37-394">Date</span></span> | <span data-ttu-id="26f37-395">Объект `Date`, указывающий дату и время начала встречи.</span><span class="sxs-lookup"><span data-stu-id="26f37-395">A `Date` object specifying the start date and time of the appointment.</span></span> |
| `parameters.end` | <span data-ttu-id="26f37-396">Date</span><span class="sxs-lookup"><span data-stu-id="26f37-396">Date</span></span> | <span data-ttu-id="26f37-397">Объект `Date`, указывающий дату и время окончания встречи.</span><span class="sxs-lookup"><span data-stu-id="26f37-397">A `Date` object specifying the end date and time of the appointment.</span></span> |
| `parameters.location` | <span data-ttu-id="26f37-398">String</span><span class="sxs-lookup"><span data-stu-id="26f37-398">String</span></span> | <span data-ttu-id="26f37-p118">Строка со сведениями о месте встречи. Максимальное количество символов в строке — 255.</span><span class="sxs-lookup"><span data-stu-id="26f37-p118">A string containing the location of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.resources` | <span data-ttu-id="26f37-401">Array.&lt;String&gt;</span><span class="sxs-lookup"><span data-stu-id="26f37-401">Array.&lt;String&gt;</span></span> | <span data-ttu-id="26f37-p119">Массив строк, содержащий необходимые для встречи ресурсы. Массив может включать не более 100 записей.</span><span class="sxs-lookup"><span data-stu-id="26f37-p119">An array of strings containing the resources required for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="26f37-404">String</span><span class="sxs-lookup"><span data-stu-id="26f37-404">String</span></span> | <span data-ttu-id="26f37-p120">Строка с темой встречи. Максимальное количество символов в строке — 255.</span><span class="sxs-lookup"><span data-stu-id="26f37-p120">A string containing the subject of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.body` | <span data-ttu-id="26f37-407">String</span><span class="sxs-lookup"><span data-stu-id="26f37-407">String</span></span> | <span data-ttu-id="26f37-p121">Текст сообщения о встрече. Максимальный размер содержимого сообщения — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="26f37-p121">The body of the appointment. The body content is limited to a maximum size of 32 KB.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="26f37-410">Требования</span><span class="sxs-lookup"><span data-stu-id="26f37-410">Requirements</span></span>

|<span data-ttu-id="26f37-411">Требование</span><span class="sxs-lookup"><span data-stu-id="26f37-411">Requirement</span></span>| <span data-ttu-id="26f37-412">Значение</span><span class="sxs-lookup"><span data-stu-id="26f37-412">Value</span></span>|
|---|---|
|[<span data-ttu-id="26f37-413">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="26f37-413">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="26f37-414">1.0</span><span class="sxs-lookup"><span data-stu-id="26f37-414">1.0</span></span>|
|[<span data-ttu-id="26f37-415">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="26f37-415">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="26f37-416">ReadItem</span><span class="sxs-lookup"><span data-stu-id="26f37-416">ReadItem</span></span>|
|[<span data-ttu-id="26f37-417">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="26f37-417">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="26f37-418">Чтение</span><span class="sxs-lookup"><span data-stu-id="26f37-418">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="26f37-419">Пример</span><span class="sxs-lookup"><span data-stu-id="26f37-419">Example</span></span>

```js
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

<br>

---
---

#### <a name="displaynewmessageformparameters"></a><span data-ttu-id="26f37-420">Дисплайневмессажеформ (Parameters)</span><span class="sxs-lookup"><span data-stu-id="26f37-420">displayNewMessageForm(parameters)</span></span>

<span data-ttu-id="26f37-421">Отображает форму для создания нового сообщения.</span><span class="sxs-lookup"><span data-stu-id="26f37-421">Displays a form for creating a new message.</span></span>

<span data-ttu-id="26f37-422">`displayNewMessageForm` Метод открывает форму, которая позволяет пользователю создать новое сообщение.</span><span class="sxs-lookup"><span data-stu-id="26f37-422">The `displayNewMessageForm` method opens a form that enables the user to create a new message.</span></span> <span data-ttu-id="26f37-423">Если указаны параметры, поля формы сообщения автоматически заполняются содержимым параметров.</span><span class="sxs-lookup"><span data-stu-id="26f37-423">If parameters are specified, the message form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="26f37-424">Если параметры превышают указанные ограничения размера или если указано неизвестное имя параметра, вызывается исключение.</span><span class="sxs-lookup"><span data-stu-id="26f37-424">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="26f37-425">Параметры</span><span class="sxs-lookup"><span data-stu-id="26f37-425">Parameters</span></span>

> [!NOTE]
> <span data-ttu-id="26f37-426">Все параметры являются необязательными.</span><span class="sxs-lookup"><span data-stu-id="26f37-426">All parameters are optional.</span></span>

|<span data-ttu-id="26f37-427">Имя</span><span class="sxs-lookup"><span data-stu-id="26f37-427">Name</span></span>| <span data-ttu-id="26f37-428">Тип</span><span class="sxs-lookup"><span data-stu-id="26f37-428">Type</span></span>| <span data-ttu-id="26f37-429">Описание</span><span class="sxs-lookup"><span data-stu-id="26f37-429">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="26f37-430">Object</span><span class="sxs-lookup"><span data-stu-id="26f37-430">Object</span></span> | <span data-ttu-id="26f37-431">Словарь параметров, описывающих новое сообщение.</span><span class="sxs-lookup"><span data-stu-id="26f37-431">A dictionary of parameters describing the new message.</span></span> |
| `parameters.toRecipients` | <span data-ttu-id="26f37-432">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)&gt;</span><span class="sxs-lookup"><span data-stu-id="26f37-432">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)&gt;</span></span> | <span data-ttu-id="26f37-433">Массив строк, содержащий адреса электронной почты или массив, содержащий `EmailAddressDetails` объект для каждого из получателей в строке "Кому".</span><span class="sxs-lookup"><span data-stu-id="26f37-433">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the To line.</span></span> <span data-ttu-id="26f37-434">Массив может включать не более 100 записей.</span><span class="sxs-lookup"><span data-stu-id="26f37-434">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.ccRecipients` | <span data-ttu-id="26f37-435">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)&gt;</span><span class="sxs-lookup"><span data-stu-id="26f37-435">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)&gt;</span></span> | <span data-ttu-id="26f37-436">Массив строк, содержащий адреса электронной почты или массив, содержащий `EmailAddressDetails` объект для каждого получателя в строке "копия".</span><span class="sxs-lookup"><span data-stu-id="26f37-436">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Cc line.</span></span> <span data-ttu-id="26f37-437">Массив может включать не более 100 записей.</span><span class="sxs-lookup"><span data-stu-id="26f37-437">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.bccRecipients` | <span data-ttu-id="26f37-438">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)&gt;</span><span class="sxs-lookup"><span data-stu-id="26f37-438">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)&gt;</span></span> | <span data-ttu-id="26f37-439">Массив строк, содержащий адреса электронной почты или массив, содержащий `EmailAddressDetails` объект для каждого из получателей, указанных в строке "СК".</span><span class="sxs-lookup"><span data-stu-id="26f37-439">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Bcc line.</span></span> <span data-ttu-id="26f37-440">Массив может включать не более 100 записей.</span><span class="sxs-lookup"><span data-stu-id="26f37-440">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="26f37-441">String</span><span class="sxs-lookup"><span data-stu-id="26f37-441">String</span></span> | <span data-ttu-id="26f37-442">Строка, содержащая тему сообщения.</span><span class="sxs-lookup"><span data-stu-id="26f37-442">A string containing the subject of the message.</span></span> <span data-ttu-id="26f37-443">Максимальное количество символов в строке — 255.</span><span class="sxs-lookup"><span data-stu-id="26f37-443">The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.htmlBody` | <span data-ttu-id="26f37-444">String</span><span class="sxs-lookup"><span data-stu-id="26f37-444">String</span></span> | <span data-ttu-id="26f37-445">Текст сообщения в формате HTML.</span><span class="sxs-lookup"><span data-stu-id="26f37-445">The HTML body of the message.</span></span> <span data-ttu-id="26f37-446">Максимальный размер содержимого сообщения — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="26f37-446">The body content is limited to a maximum size of 32 KB.</span></span> |
| `parameters.attachments` | <span data-ttu-id="26f37-447">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="26f37-447">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="26f37-448">Массив объектов JSON, представляющих собой вложенные файлы или элементы.</span><span class="sxs-lookup"><span data-stu-id="26f37-448">An array of JSON objects that are either file or item attachments.</span></span> |
| `parameters.attachments.type` | <span data-ttu-id="26f37-449">String</span><span class="sxs-lookup"><span data-stu-id="26f37-449">String</span></span> | <span data-ttu-id="26f37-p128">Указывает тип вложения. Допустимые значения: `file` для вложенного файла и `item` для вложенного элемента.</span><span class="sxs-lookup"><span data-stu-id="26f37-p128">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `parameters.attachments.name` | <span data-ttu-id="26f37-452">Строка</span><span class="sxs-lookup"><span data-stu-id="26f37-452">String</span></span> | <span data-ttu-id="26f37-453">Строка, содержащая имя вложения, длиной до 255 символов.</span><span class="sxs-lookup"><span data-stu-id="26f37-453">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `parameters.attachments.url` | <span data-ttu-id="26f37-454">String</span><span class="sxs-lookup"><span data-stu-id="26f37-454">String</span></span> | <span data-ttu-id="26f37-p129">Используется, только если свойству `type` задано значение `file`. URI расположения файла.</span><span class="sxs-lookup"><span data-stu-id="26f37-p129">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `parameters.attachments.isInline` | <span data-ttu-id="26f37-457">Логический</span><span class="sxs-lookup"><span data-stu-id="26f37-457">Boolean</span></span> | <span data-ttu-id="26f37-p130">Используется, только если свойству `type` задано значение `file`. Значение `true` указывает на то, что вложение будет встроено в текст сообщения и не должно отображаться в списке вложений.</span><span class="sxs-lookup"><span data-stu-id="26f37-p130">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `parameters.attachments.itemId` | <span data-ttu-id="26f37-460">Строка</span><span class="sxs-lookup"><span data-stu-id="26f37-460">String</span></span> | <span data-ttu-id="26f37-461">Используется, только если свойству `type` присвоено значение `item`.</span><span class="sxs-lookup"><span data-stu-id="26f37-461">Only used if `type` is set to `item`.</span></span> <span data-ttu-id="26f37-462">Идентификатор элемента EWS существующего сообщения электронной почты, которое необходимо присоединить к новому сообщению.</span><span class="sxs-lookup"><span data-stu-id="26f37-462">The EWS item id of the existing e-mail you want to attach to the new message.</span></span> <span data-ttu-id="26f37-463">Это строка длиной до 100 символов.</span><span class="sxs-lookup"><span data-stu-id="26f37-463">This is a string up to 100 characters.</span></span> |


##### <a name="requirements"></a><span data-ttu-id="26f37-464">Требования</span><span class="sxs-lookup"><span data-stu-id="26f37-464">Requirements</span></span>

|<span data-ttu-id="26f37-465">Требование</span><span class="sxs-lookup"><span data-stu-id="26f37-465">Requirement</span></span>| <span data-ttu-id="26f37-466">Значение</span><span class="sxs-lookup"><span data-stu-id="26f37-466">Value</span></span>|
|---|---|
|[<span data-ttu-id="26f37-467">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="26f37-467">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="26f37-468">1.6</span><span class="sxs-lookup"><span data-stu-id="26f37-468">1.6</span></span> |
|[<span data-ttu-id="26f37-469">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="26f37-469">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="26f37-470">ReadItem</span><span class="sxs-lookup"><span data-stu-id="26f37-470">ReadItem</span></span>|
|[<span data-ttu-id="26f37-471">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="26f37-471">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="26f37-472">Чтение</span><span class="sxs-lookup"><span data-stu-id="26f37-472">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="26f37-473">Пример</span><span class="sxs-lookup"><span data-stu-id="26f37-473">Example</span></span>

```js
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

<br>

---
---

#### <a name="getcallbacktokenasyncoptions-callback"></a><span data-ttu-id="26f37-474">getCallbackTokenAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="26f37-474">getCallbackTokenAsync([options], callback)</span></span>

<span data-ttu-id="26f37-475">Возвращает строку, содержащую маркер, который используется для вызова интерфейсов REST API или веб-служб Exchange.</span><span class="sxs-lookup"><span data-stu-id="26f37-475">Gets a string that contains a token used to call REST APIs or Exchange Web Services.</span></span>

<span data-ttu-id="26f37-p132">Метод `getCallbackTokenAsync` совершает асинхронный вызов, чтобы получить непрозрачный маркер с сервера Exchange Server, на котором размещен почтовый ящик пользователя. Время существования маркера обратного вызова составляет 5 минут.</span><span class="sxs-lookup"><span data-stu-id="26f37-p132">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

> [!NOTE]
> <span data-ttu-id="26f37-478">Рекомендуем сделать так, чтобы по мере возможности надстройки использовали интерфейсы REST API, а не веб-службы Exchange.</span><span class="sxs-lookup"><span data-stu-id="26f37-478">It is recommended that add-ins use the REST APIs instead of Exchange Web Services whenever possible.</span></span>

<span data-ttu-id="26f37-479">Для вызова `getCallbackTokenAsync` метода в режиме чтения требуется минимальный уровень разрешений **ReadItem**.</span><span class="sxs-lookup"><span data-stu-id="26f37-479">Calling the `getCallbackTokenAsync` method in read mode requires a minimum permission level of **ReadItem**.</span></span>

<span data-ttu-id="26f37-480">Для `getCallbackTokenAsync` вызова в режиме создания необходимо сохранить элемент.</span><span class="sxs-lookup"><span data-stu-id="26f37-480">Calling `getCallbackTokenAsync` in compose mode requires you to have saved the item.</span></span> <span data-ttu-id="26f37-481">[`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) Метод требует наличия минимального уровня разрешений **ReadWriteItem**.</span><span class="sxs-lookup"><span data-stu-id="26f37-481">The [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method requires a minimum permission level of **ReadWriteItem**.</span></span>

<span data-ttu-id="26f37-482">**Маркеры REST**</span><span class="sxs-lookup"><span data-stu-id="26f37-482">**REST Tokens**</span></span>

<span data-ttu-id="26f37-p134">Если запрашивается маркер REST (`options.isRest = true`), полученный маркер не подойдет для проверки подлинности при вызовах веб-служб Exchange. Область действия маркера будет ограничена доступом только для чтения к текущему элементу и его вложениям, если в манифесте надстройки не указано разрешение [`ReadWriteMailbox`](/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission). Если указано разрешение `ReadWriteMailbox`, полученный маркер предоставит доступ на чтение и запись к почте, календарю и контактам, включая возможность отправки почты.</span><span class="sxs-lookup"><span data-stu-id="26f37-p134">When a REST token is requested (`options.isRest = true`), the resulting token will not work to authenticate Exchange Web Services calls. The token will be limited in scope to read-only access to the current item and its attachments, unless the add-in has specified the [`ReadWriteMailbox`](/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) permission in its manifest. If the `ReadWriteMailbox` permission is specified, the resulting token will grant read/write access to mail, calendar, and contacts, including the ability to send mail.</span></span>

<span data-ttu-id="26f37-486">С помощью свойства `restUrl` надстройка должна определить правильный URL-адрес для вызовов REST API.</span><span class="sxs-lookup"><span data-stu-id="26f37-486">The add-in should use the `restUrl` property to determine the correct URL to use when making REST API calls.</span></span>

<span data-ttu-id="26f37-487">**Маркеры EWS**</span><span class="sxs-lookup"><span data-stu-id="26f37-487">**EWS Tokens**</span></span>

<span data-ttu-id="26f37-p135">Если запрашивается маркер EWS (`options.isRest = false`), полученный маркер не подойдет для проверки подлинности при вызовах REST API. Область действия маркера будет ограничена доступом к текущему элементу.</span><span class="sxs-lookup"><span data-stu-id="26f37-p135">When an EWS token is requested (`options.isRest = false`), the resulting token will not work to authenticate REST API calls. The token will be limited in scope to accessing the current item.</span></span>

<span data-ttu-id="26f37-490">С помощью свойства `ewsUrl` надстройка должна определить правильный URL-адрес для вызовов EWS.</span><span class="sxs-lookup"><span data-stu-id="26f37-490">The add-in should use the `ewsUrl` property to determine the correct URL to use when making EWS calls.</span></span>

<span data-ttu-id="26f37-491">Можно передать как маркер, так и идентификатор вложения или идентификатор элемента в систему стороннего производителя.</span><span class="sxs-lookup"><span data-stu-id="26f37-491">You can pass both the token and either an attachment identifier or item identifier to a third-party system.</span></span> <span data-ttu-id="26f37-492">Третья система использует маркер в качестве маркера авторизации носителя, чтобы вызвать операцию [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) [или GetItem](/exchange/client-developer/web-service-reference/getitem-operation) веб-служб Exchange (EWS) для возврата вложения или элемента.</span><span class="sxs-lookup"><span data-stu-id="26f37-492">The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) operation or [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item.</span></span> <span data-ttu-id="26f37-493">Например, вы можете создать удаленную службу, чтобы [получить вложения из выбранного элемента](/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="26f37-493">For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

##### <a name="parameters"></a><span data-ttu-id="26f37-494">Параметры</span><span class="sxs-lookup"><span data-stu-id="26f37-494">Parameters</span></span>

|<span data-ttu-id="26f37-495">Имя</span><span class="sxs-lookup"><span data-stu-id="26f37-495">Name</span></span>| <span data-ttu-id="26f37-496">Тип</span><span class="sxs-lookup"><span data-stu-id="26f37-496">Type</span></span>| <span data-ttu-id="26f37-497">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="26f37-497">Attributes</span></span>| <span data-ttu-id="26f37-498">Описание</span><span class="sxs-lookup"><span data-stu-id="26f37-498">Description</span></span>|
|---|---|---|---|
| `options` | <span data-ttu-id="26f37-499">Object</span><span class="sxs-lookup"><span data-stu-id="26f37-499">Object</span></span> | <span data-ttu-id="26f37-500">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="26f37-500">&lt;optional&gt;</span></span> | <span data-ttu-id="26f37-501">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="26f37-501">An object literal that contains one or more of the following properties.</span></span> |
| `options.isRest` | <span data-ttu-id="26f37-502">Boolean</span><span class="sxs-lookup"><span data-stu-id="26f37-502">Boolean</span></span> |  <span data-ttu-id="26f37-503">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="26f37-503">&lt;optional&gt;</span></span> | <span data-ttu-id="26f37-p137">Определяет, будет ли предоставленный маркер использоваться для интерфейсов REST API Outlook или веб-служб Exchange. Значение по умолчанию: `false`.</span><span class="sxs-lookup"><span data-stu-id="26f37-p137">Determines if the token provided will be used for the Outlook REST APIs or Exchange Web Services. Default value is `false`.</span></span> |
| `options.asyncContext` | <span data-ttu-id="26f37-506">Object</span><span class="sxs-lookup"><span data-stu-id="26f37-506">Object</span></span> |  <span data-ttu-id="26f37-507">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="26f37-507">&lt;optional&gt;</span></span> | <span data-ttu-id="26f37-508">Данные о состоянии, передаваемые в асинхронный метод.</span><span class="sxs-lookup"><span data-stu-id="26f37-508">Any state data that is passed to the asynchronous method.</span></span> |
|`callback`| <span data-ttu-id="26f37-509">функция</span><span class="sxs-lookup"><span data-stu-id="26f37-509">function</span></span>||<span data-ttu-id="26f37-510">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="26f37-510">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="26f37-511">Маркер указывается в виде строки в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="26f37-511">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="26f37-512">При наличии ошибки свойства `asyncResult.error` и `asyncResult.diagnostics` могут предоставлять дополнительные сведения.</span><span class="sxs-lookup"><span data-stu-id="26f37-512">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|

##### <a name="errors"></a><span data-ttu-id="26f37-513">Ошибки</span><span class="sxs-lookup"><span data-stu-id="26f37-513">Errors</span></span>

|<span data-ttu-id="26f37-514">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="26f37-514">Error code</span></span>|<span data-ttu-id="26f37-515">Описание</span><span class="sxs-lookup"><span data-stu-id="26f37-515">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="26f37-516">Не удалось выполнить запрос.</span><span class="sxs-lookup"><span data-stu-id="26f37-516">The request has failed.</span></span> <span data-ttu-id="26f37-517">Просмотрите объект диагностики для кода ошибки HTTP.</span><span class="sxs-lookup"><span data-stu-id="26f37-517">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="26f37-518">Сервер Exchange Server вернул ошибку.</span><span class="sxs-lookup"><span data-stu-id="26f37-518">The Exchange server returned an error.</span></span> <span data-ttu-id="26f37-519">Для получения дополнительных сведений просмотрите объект диагностики.</span><span class="sxs-lookup"><span data-stu-id="26f37-519">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="26f37-520">Пользователь отключен от сети.</span><span class="sxs-lookup"><span data-stu-id="26f37-520">The user is no longer connected to the network.</span></span> <span data-ttu-id="26f37-521">Проверьте сетевое подключение и повторите попытку.</span><span class="sxs-lookup"><span data-stu-id="26f37-521">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="26f37-522">Требования</span><span class="sxs-lookup"><span data-stu-id="26f37-522">Requirements</span></span>

|<span data-ttu-id="26f37-523">Требование</span><span class="sxs-lookup"><span data-stu-id="26f37-523">Requirement</span></span>| <span data-ttu-id="26f37-524">Значение</span><span class="sxs-lookup"><span data-stu-id="26f37-524">Value</span></span>|
|---|---|
|[<span data-ttu-id="26f37-525">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="26f37-525">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="26f37-526">1.5</span><span class="sxs-lookup"><span data-stu-id="26f37-526">1.5</span></span> |
|[<span data-ttu-id="26f37-527">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="26f37-527">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="26f37-528">ReadItem</span><span class="sxs-lookup"><span data-stu-id="26f37-528">ReadItem</span></span>|
|[<span data-ttu-id="26f37-529">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="26f37-529">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="26f37-530">Создание и чтение</span><span class="sxs-lookup"><span data-stu-id="26f37-530">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="26f37-531">Пример</span><span class="sxs-lookup"><span data-stu-id="26f37-531">Example</span></span>

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

<br>

---
---

#### <a name="getcallbacktokenasynccallback-usercontext"></a><span data-ttu-id="26f37-532">getCallbackTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="26f37-532">getCallbackTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="26f37-533">Получает строку, содержащую маркер, используемый для получения вложения или элемента с Exchange Server.</span><span class="sxs-lookup"><span data-stu-id="26f37-533">Gets a string that contains a token used to get an attachment or item from an Exchange Server.</span></span>

<span data-ttu-id="26f37-p141">Метод `getCallbackTokenAsync` совершает асинхронный вызов, чтобы получить непрозрачный маркер с сервера Exchange Server, на котором размещен почтовый ящик пользователя. Время существования маркера обратного вызова составляет 5 минут.</span><span class="sxs-lookup"><span data-stu-id="26f37-p141">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

<span data-ttu-id="26f37-536">Можно передать как маркер, так и идентификатор вложения или идентификатор элемента в систему стороннего производителя.</span><span class="sxs-lookup"><span data-stu-id="26f37-536">You can pass both the token and either an attachment identifier or item identifier to a third-party system.</span></span> <span data-ttu-id="26f37-537">Третья система использует маркер в качестве маркера авторизации носителя, чтобы вызвать операцию [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) [или GetItem](/exchange/client-developer/web-service-reference/getitem-operation) веб-служб Exchange (EWS) для возврата вложения или элемента.</span><span class="sxs-lookup"><span data-stu-id="26f37-537">The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) operation or [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item.</span></span> <span data-ttu-id="26f37-538">Например, вы можете создать удаленную службу, чтобы [получить вложения из выбранного элемента](/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="26f37-538">For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="26f37-539">Для вызова `getCallbackTokenAsync` метода в режиме чтения требуется минимальный уровень разрешений **ReadItem**.</span><span class="sxs-lookup"><span data-stu-id="26f37-539">Calling the `getCallbackTokenAsync` method in read mode requires a minimum permission level of **ReadItem**.</span></span>

<span data-ttu-id="26f37-540">Для `getCallbackTokenAsync` вызова в режиме создания необходимо сохранить элемент.</span><span class="sxs-lookup"><span data-stu-id="26f37-540">Calling `getCallbackTokenAsync` in compose mode requires you to have saved the item.</span></span> <span data-ttu-id="26f37-541">[`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) Метод требует наличия минимального уровня разрешений **ReadWriteItem**.</span><span class="sxs-lookup"><span data-stu-id="26f37-541">The [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method requires a minimum permission level of **ReadWriteItem**.</span></span>

##### <a name="parameters"></a><span data-ttu-id="26f37-542">Параметры</span><span class="sxs-lookup"><span data-stu-id="26f37-542">Parameters</span></span>

|<span data-ttu-id="26f37-543">Имя</span><span class="sxs-lookup"><span data-stu-id="26f37-543">Name</span></span>| <span data-ttu-id="26f37-544">Тип</span><span class="sxs-lookup"><span data-stu-id="26f37-544">Type</span></span>| <span data-ttu-id="26f37-545">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="26f37-545">Attributes</span></span>| <span data-ttu-id="26f37-546">Описание</span><span class="sxs-lookup"><span data-stu-id="26f37-546">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="26f37-547">функция</span><span class="sxs-lookup"><span data-stu-id="26f37-547">function</span></span>||<span data-ttu-id="26f37-548">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="26f37-548">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="26f37-549">Маркер указывается в виде строки в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="26f37-549">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="26f37-550">При наличии ошибки свойства `asyncResult.error` и `asyncResult.diagnostics` могут предоставлять дополнительные сведения.</span><span class="sxs-lookup"><span data-stu-id="26f37-550">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|
|`userContext`| <span data-ttu-id="26f37-551">Объект</span><span class="sxs-lookup"><span data-stu-id="26f37-551">Object</span></span>| <span data-ttu-id="26f37-552">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="26f37-552">&lt;optional&gt;</span></span>|<span data-ttu-id="26f37-553">Данные о состоянии, передаваемые в асинхронный метод.</span><span class="sxs-lookup"><span data-stu-id="26f37-553">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="errors"></a><span data-ttu-id="26f37-554">Ошибки</span><span class="sxs-lookup"><span data-stu-id="26f37-554">Errors</span></span>

|<span data-ttu-id="26f37-555">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="26f37-555">Error code</span></span>|<span data-ttu-id="26f37-556">Описание</span><span class="sxs-lookup"><span data-stu-id="26f37-556">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="26f37-557">Не удалось выполнить запрос.</span><span class="sxs-lookup"><span data-stu-id="26f37-557">The request has failed.</span></span> <span data-ttu-id="26f37-558">Просмотрите объект диагностики для кода ошибки HTTP.</span><span class="sxs-lookup"><span data-stu-id="26f37-558">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="26f37-559">Сервер Exchange Server вернул ошибку.</span><span class="sxs-lookup"><span data-stu-id="26f37-559">The Exchange server returned an error.</span></span> <span data-ttu-id="26f37-560">Для получения дополнительных сведений просмотрите объект диагностики.</span><span class="sxs-lookup"><span data-stu-id="26f37-560">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="26f37-561">Пользователь отключен от сети.</span><span class="sxs-lookup"><span data-stu-id="26f37-561">The user is no longer connected to the network.</span></span> <span data-ttu-id="26f37-562">Проверьте сетевое подключение и повторите попытку.</span><span class="sxs-lookup"><span data-stu-id="26f37-562">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="26f37-563">Требования</span><span class="sxs-lookup"><span data-stu-id="26f37-563">Requirements</span></span>

|<span data-ttu-id="26f37-564">Требование</span><span class="sxs-lookup"><span data-stu-id="26f37-564">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="26f37-565">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="26f37-565">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="26f37-566">1.0</span><span class="sxs-lookup"><span data-stu-id="26f37-566">1.0</span></span> | <span data-ttu-id="26f37-567">1.3</span><span class="sxs-lookup"><span data-stu-id="26f37-567">1.3</span></span> |
|[<span data-ttu-id="26f37-568">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="26f37-568">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="26f37-569">ReadItem</span><span class="sxs-lookup"><span data-stu-id="26f37-569">ReadItem</span></span> | <span data-ttu-id="26f37-570">ReadItem</span><span class="sxs-lookup"><span data-stu-id="26f37-570">ReadItem</span></span> |
|[<span data-ttu-id="26f37-571">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="26f37-571">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="26f37-572">Чтение</span><span class="sxs-lookup"><span data-stu-id="26f37-572">Read</span></span> | <span data-ttu-id="26f37-573">Создание</span><span class="sxs-lookup"><span data-stu-id="26f37-573">Compose</span></span> |

##### <a name="example"></a><span data-ttu-id="26f37-574">Пример</span><span class="sxs-lookup"><span data-stu-id="26f37-574">Example</span></span>

```js
function getCallbackToken() {
  Office.context.mailbox.getCallbackTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

<br>

---
---

#### <a name="getuseridentitytokenasynccallback-usercontext"></a><span data-ttu-id="26f37-575">getUserIdentityTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="26f37-575">getUserIdentityTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="26f37-576">Получает маркер, идентифицирующий пользователя и надстройку Office.</span><span class="sxs-lookup"><span data-stu-id="26f37-576">Gets a token identifying the user and the Office Add-in.</span></span>

<span data-ttu-id="26f37-577">Метод `getUserIdentityTokenAsync` возвращает маркер, который можно использовать для идентификации, а также [проверки подлинности надстройки и пользователя в сторонней системе](/outlook/add-ins/authentication).</span><span class="sxs-lookup"><span data-stu-id="26f37-577">The `getUserIdentityTokenAsync` method returns a token that you can use to identify and [authenticate the add-in and user with a third-party system](/outlook/add-ins/authentication).</span></span>

##### <a name="parameters"></a><span data-ttu-id="26f37-578">Параметры</span><span class="sxs-lookup"><span data-stu-id="26f37-578">Parameters</span></span>

|<span data-ttu-id="26f37-579">Имя</span><span class="sxs-lookup"><span data-stu-id="26f37-579">Name</span></span>| <span data-ttu-id="26f37-580">Тип</span><span class="sxs-lookup"><span data-stu-id="26f37-580">Type</span></span>| <span data-ttu-id="26f37-581">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="26f37-581">Attributes</span></span>| <span data-ttu-id="26f37-582">Описание</span><span class="sxs-lookup"><span data-stu-id="26f37-582">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="26f37-583">функция</span><span class="sxs-lookup"><span data-stu-id="26f37-583">function</span></span>||<span data-ttu-id="26f37-584">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="26f37-584">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="26f37-585">Маркер указывается в виде строки в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="26f37-585">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="26f37-586">При наличии ошибки свойства `asyncResult.error` и `asyncResult.diagnostics` могут предоставлять дополнительные сведения.</span><span class="sxs-lookup"><span data-stu-id="26f37-586">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|
|`userContext`| <span data-ttu-id="26f37-587">Объект</span><span class="sxs-lookup"><span data-stu-id="26f37-587">Object</span></span>| <span data-ttu-id="26f37-588">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="26f37-588">&lt;optional&gt;</span></span>|<span data-ttu-id="26f37-589">Данные о состоянии, передаваемые в асинхронный метод.</span><span class="sxs-lookup"><span data-stu-id="26f37-589">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="errors"></a><span data-ttu-id="26f37-590">Ошибки</span><span class="sxs-lookup"><span data-stu-id="26f37-590">Errors</span></span>

|<span data-ttu-id="26f37-591">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="26f37-591">Error code</span></span>|<span data-ttu-id="26f37-592">Описание</span><span class="sxs-lookup"><span data-stu-id="26f37-592">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="26f37-593">Не удалось выполнить запрос.</span><span class="sxs-lookup"><span data-stu-id="26f37-593">The request has failed.</span></span> <span data-ttu-id="26f37-594">Просмотрите объект диагностики для кода ошибки HTTP.</span><span class="sxs-lookup"><span data-stu-id="26f37-594">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="26f37-595">Сервер Exchange Server вернул ошибку.</span><span class="sxs-lookup"><span data-stu-id="26f37-595">The Exchange server returned an error.</span></span> <span data-ttu-id="26f37-596">Для получения дополнительных сведений просмотрите объект диагностики.</span><span class="sxs-lookup"><span data-stu-id="26f37-596">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="26f37-597">Пользователь отключен от сети.</span><span class="sxs-lookup"><span data-stu-id="26f37-597">The user is no longer connected to the network.</span></span> <span data-ttu-id="26f37-598">Проверьте сетевое подключение и повторите попытку.</span><span class="sxs-lookup"><span data-stu-id="26f37-598">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="26f37-599">Требования</span><span class="sxs-lookup"><span data-stu-id="26f37-599">Requirements</span></span>

|<span data-ttu-id="26f37-600">Требование</span><span class="sxs-lookup"><span data-stu-id="26f37-600">Requirement</span></span>| <span data-ttu-id="26f37-601">Значение</span><span class="sxs-lookup"><span data-stu-id="26f37-601">Value</span></span>|
|---|---|
|[<span data-ttu-id="26f37-602">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="26f37-602">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="26f37-603">1.0</span><span class="sxs-lookup"><span data-stu-id="26f37-603">1.0</span></span>|
|[<span data-ttu-id="26f37-604">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="26f37-604">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="26f37-605">ReadItem</span><span class="sxs-lookup"><span data-stu-id="26f37-605">ReadItem</span></span>|
|[<span data-ttu-id="26f37-606">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="26f37-606">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="26f37-607">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="26f37-607">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="26f37-608">Пример</span><span class="sxs-lookup"><span data-stu-id="26f37-608">Example</span></span>

```js
function getIdentityToken() {
  Office.context.mailbox.getUserIdentityTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

<br>

---
---

#### <a name="makeewsrequestasyncdata-callback-usercontext"></a><span data-ttu-id="26f37-609">makeEwsRequestAsync(data, callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="26f37-609">makeEwsRequestAsync(data, callback, [userContext])</span></span>

<span data-ttu-id="26f37-610">Выполняет асинхронный запрос для веб-служб Exchange (EWS) на сервере Exchange Server, на котором размещен почтовый ящик пользователя.</span><span class="sxs-lookup"><span data-stu-id="26f37-610">Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="26f37-611">Этот метод не поддерживается в следующих сценариях:</span><span class="sxs-lookup"><span data-stu-id="26f37-611">This method is not supported in the following scenarios.</span></span>
> - <span data-ttu-id="26f37-612">В Outlook для iOS и Android</span><span class="sxs-lookup"><span data-stu-id="26f37-612">In Outlook on iOS or Android</span></span>
> - <span data-ttu-id="26f37-613">Если надстройка загружается в почтовый ящик Gmail.</span><span class="sxs-lookup"><span data-stu-id="26f37-613">When the add-in is loaded in a Gmail mailbox</span></span>
> 
> <span data-ttu-id="26f37-614">В таких случаях надстройка должна [использовать REST API](/outlook/add-ins/use-rest-api) для доступа к почтовому ящику пользователя.</span><span class="sxs-lookup"><span data-stu-id="26f37-614">In these cases, add-ins should [use REST APIs](/outlook/add-ins/use-rest-api) to access the user's mailbox instead.</span></span>

<span data-ttu-id="26f37-615">Метод `makeEwsRequestAsync` отправляет запрос EWS от имени надстройки в Exchange.</span><span class="sxs-lookup"><span data-stu-id="26f37-615">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange.</span></span> <span data-ttu-id="26f37-616">Список поддерживаемых операций EWS см. в статье [Вызов веб-служб из надстройки Outlook](/outlook/add-ins/web-services#ews-operations-that-add-ins-support).</span><span class="sxs-lookup"><span data-stu-id="26f37-616">See [Call web services from an Outlook add-in](/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span></span>

<span data-ttu-id="26f37-617">С помощью метода `makeEwsRequestAsync` невозможно запрашивать элементы, связанные с папкой.</span><span class="sxs-lookup"><span data-stu-id="26f37-617">You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.</span></span>

<span data-ttu-id="26f37-618">В запросе XML должна быть указана кодировка UTF-8.</span><span class="sxs-lookup"><span data-stu-id="26f37-618">The XML request must specify UTF-8 encoding.</span></span>

```xml
<?xml version="1.0" encoding="utf-8"?>
```

<span data-ttu-id="26f37-p151">У вашей надстройки должно быть разрешение **ReadWriteMailbox** для использования метода `makeEwsRequestAsync`. Сведения об использовании разрешения **ReadWriteMailbox** и операций EWS, которые можно вызывать с помощью метода `makeEwsRequestAsync`, см. в статье [Указание разрешений для доступа почтовой надстройки к почтовому ящику пользователя](/outlook/add-ins/understanding-outlook-add-in-permissions).</span><span class="sxs-lookup"><span data-stu-id="26f37-p151">Your add-in must have the **ReadWriteMailbox** permission to use the `makeEwsRequestAsync` method. For information about using the **ReadWriteMailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method, see [Specify permissions for mail add-in access to the user's mailbox](/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>

> [!NOTE]
> <span data-ttu-id="26f37-621">Администратор сервера должен установить значение true для параметра `OAuthAuthentication` в каталоге сервера клиентского доступа EWS, чтобы метод `makeEwsRequestAsync` мог выполнять запросы EWS.</span><span class="sxs-lookup"><span data-stu-id="26f37-621">The server administrator must set `OAuthAuthentication` to true on the Client Access Server EWS directory to enable the `makeEwsRequestAsync` method to make EWS requests.</span></span>

##### <a name="version-differences"></a><span data-ttu-id="26f37-622">Различия версий</span><span class="sxs-lookup"><span data-stu-id="26f37-622">Version differences</span></span>

<span data-ttu-id="26f37-623">Если вы используете метод `makeEwsRequestAsync` в почтовых приложениях, которые выполняются в Outlook версии более ранней, чем 15.0.4535.1004, указывайте кодировку `ISO-8859-1`.</span><span class="sxs-lookup"><span data-stu-id="26f37-623">When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set the encoding value to `ISO-8859-1`.</span></span>

```xml
<?xml version="1.0" encoding="iso-8859-1"?>
```

<span data-ttu-id="26f37-p152">Значение кодировки не нужно указывать, если почтовое приложение выполняется в Outlook в Интернете. Чтобы определить, выполняется ли приложение в Outlook или Outlook в Интернете, используйте свойство mailbox.diagnostics.hostName. Используемую версию Outlook можно определить с помощью свойства mailbox.diagnostics.hostVersion.</span><span class="sxs-lookup"><span data-stu-id="26f37-p152">You do not need to set the encoding value when your mail app is running in Outlook on the web. You can determine whether your mail app is running in Outlook or Outlook on the web by using the mailbox.diagnostics.hostName property. You can determine what version of Outlook is running by using the mailbox.diagnostics.hostVersion property.</span></span>

##### <a name="parameters"></a><span data-ttu-id="26f37-627">Параметры</span><span class="sxs-lookup"><span data-stu-id="26f37-627">Parameters</span></span>

|<span data-ttu-id="26f37-628">Имя</span><span class="sxs-lookup"><span data-stu-id="26f37-628">Name</span></span>| <span data-ttu-id="26f37-629">Тип</span><span class="sxs-lookup"><span data-stu-id="26f37-629">Type</span></span>| <span data-ttu-id="26f37-630">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="26f37-630">Attributes</span></span>| <span data-ttu-id="26f37-631">Описание</span><span class="sxs-lookup"><span data-stu-id="26f37-631">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="26f37-632">String</span><span class="sxs-lookup"><span data-stu-id="26f37-632">String</span></span>||<span data-ttu-id="26f37-633">Запрос EWS.</span><span class="sxs-lookup"><span data-stu-id="26f37-633">The EWS request.</span></span>|
|`callback`| <span data-ttu-id="26f37-634">function</span><span class="sxs-lookup"><span data-stu-id="26f37-634">function</span></span>||<span data-ttu-id="26f37-635">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="26f37-635">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="26f37-636">Результат XML вызова EWS указывается в виде строки в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="26f37-636">The XML result of the EWS call is provided as a string in the `asyncResult.value` property.</span></span> <span data-ttu-id="26f37-637">Если размер результата превышает 1 МБ, возвращается сообщение об ошибке.</span><span class="sxs-lookup"><span data-stu-id="26f37-637">If the result exceeds 1 MB in size, an error message is returned instead.</span></span>|
|`userContext`| <span data-ttu-id="26f37-638">Object</span><span class="sxs-lookup"><span data-stu-id="26f37-638">Object</span></span>| <span data-ttu-id="26f37-639">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="26f37-639">&lt;optional&gt;</span></span>|<span data-ttu-id="26f37-640">Данные о состоянии, передаваемые в асинхронный метод.</span><span class="sxs-lookup"><span data-stu-id="26f37-640">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="26f37-641">Требования</span><span class="sxs-lookup"><span data-stu-id="26f37-641">Requirements</span></span>

|<span data-ttu-id="26f37-642">Требование</span><span class="sxs-lookup"><span data-stu-id="26f37-642">Requirement</span></span>| <span data-ttu-id="26f37-643">Значение</span><span class="sxs-lookup"><span data-stu-id="26f37-643">Value</span></span>|
|---|---|
|[<span data-ttu-id="26f37-644">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="26f37-644">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="26f37-645">1.0</span><span class="sxs-lookup"><span data-stu-id="26f37-645">1.0</span></span>|
|[<span data-ttu-id="26f37-646">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="26f37-646">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="26f37-647">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="26f37-647">ReadWriteMailbox</span></span>|
|[<span data-ttu-id="26f37-648">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="26f37-648">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="26f37-649">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="26f37-649">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="26f37-650">Пример</span><span class="sxs-lookup"><span data-stu-id="26f37-650">Example</span></span>

<span data-ttu-id="26f37-651">В приведенном ниже примере вызывается `makeEwsRequestAsync` для получения темы элемента с помощью операции `GetItem`.</span><span class="sxs-lookup"><span data-stu-id="26f37-651">The following example calls `makeEwsRequestAsync` to use the `GetItem` operation to get the subject of an item.</span></span>

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

<br>

---
---

#### <a name="removehandlerasynceventtype-options-callback"></a><span data-ttu-id="26f37-652">removeHandlerAsync(eventType, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="26f37-652">removeHandlerAsync(eventType, [options], [callback])</span></span>

<span data-ttu-id="26f37-653">Удаляет обработчиков для поддерживаемого типа события.</span><span class="sxs-lookup"><span data-stu-id="26f37-653">Removes the event handlers for a supported event type.</span></span>

<span data-ttu-id="26f37-654">В настоящее время единственный поддерживаемый тип события — `Office.EventType.ItemChanged`.</span><span class="sxs-lookup"><span data-stu-id="26f37-654">Currently, the only supported event type is `Office.EventType.ItemChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="26f37-655">Параметры</span><span class="sxs-lookup"><span data-stu-id="26f37-655">Parameters</span></span>

| <span data-ttu-id="26f37-656">Имя</span><span class="sxs-lookup"><span data-stu-id="26f37-656">Name</span></span> | <span data-ttu-id="26f37-657">Тип</span><span class="sxs-lookup"><span data-stu-id="26f37-657">Type</span></span> | <span data-ttu-id="26f37-658">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="26f37-658">Attributes</span></span> | <span data-ttu-id="26f37-659">Описание</span><span class="sxs-lookup"><span data-stu-id="26f37-659">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="26f37-660">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="26f37-660">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="26f37-661">Событие, которое должно отменить обработчик.</span><span class="sxs-lookup"><span data-stu-id="26f37-661">The event that should revoke the handler.</span></span> |
| `options` | <span data-ttu-id="26f37-662">Объект</span><span class="sxs-lookup"><span data-stu-id="26f37-662">Object</span></span> | <span data-ttu-id="26f37-663">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="26f37-663">&lt;optional&gt;</span></span> | <span data-ttu-id="26f37-664">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="26f37-664">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="26f37-665">Object</span><span class="sxs-lookup"><span data-stu-id="26f37-665">Object</span></span> | <span data-ttu-id="26f37-666">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="26f37-666">&lt;optional&gt;</span></span> | <span data-ttu-id="26f37-667">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="26f37-667">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="26f37-668">функция</span><span class="sxs-lookup"><span data-stu-id="26f37-668">function</span></span>| <span data-ttu-id="26f37-669">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="26f37-669">&lt;optional&gt;</span></span>|<span data-ttu-id="26f37-670">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="26f37-670">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="26f37-671">Требования</span><span class="sxs-lookup"><span data-stu-id="26f37-671">Requirements</span></span>

|<span data-ttu-id="26f37-672">Требование</span><span class="sxs-lookup"><span data-stu-id="26f37-672">Requirement</span></span>| <span data-ttu-id="26f37-673">Значение</span><span class="sxs-lookup"><span data-stu-id="26f37-673">Value</span></span>|
|---|---|
|[<span data-ttu-id="26f37-674">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="26f37-674">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="26f37-675">1.5</span><span class="sxs-lookup"><span data-stu-id="26f37-675">1.5</span></span> |
|[<span data-ttu-id="26f37-676">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="26f37-676">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="26f37-677">ReadItem</span><span class="sxs-lookup"><span data-stu-id="26f37-677">ReadItem</span></span> |
|[<span data-ttu-id="26f37-678">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="26f37-678">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="26f37-679">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="26f37-679">Compose or Read</span></span>|
