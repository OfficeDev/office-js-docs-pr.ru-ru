---
title: Office. Context. Mailbox — набор обязательных элементов 1,7
description: ''
ms.date: 10/21/2019
localization_priority: Normal
ms.openlocfilehash: 87e5334879bb4b5fa84700a03f6da86d4c72e7d2
ms.sourcegitcommit: 499bf49b41205f8034c501d4db5fe4b02dab205e
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 10/22/2019
ms.locfileid: "37627078"
---
# <a name="mailbox"></a><span data-ttu-id="dde3f-102">mailbox</span><span class="sxs-lookup"><span data-stu-id="dde3f-102">mailbox</span></span>

### <a name="officeofficemdcontextofficecontextmdmailbox"></a><span data-ttu-id="dde3f-103">[Office](Office.md)[.context](Office.context.md).mailbox</span><span class="sxs-lookup"><span data-stu-id="dde3f-103">[Office](Office.md)[.context](Office.context.md).mailbox</span></span>

<span data-ttu-id="dde3f-104">Предоставляет для Microsoft Outlook доступ к объектной модели надстройки Outlook.</span><span class="sxs-lookup"><span data-stu-id="dde3f-104">Provides access to the Outlook add-in object model for Microsoft Outlook.</span></span>

##### <a name="requirements"></a><span data-ttu-id="dde3f-105">Требования</span><span class="sxs-lookup"><span data-stu-id="dde3f-105">Requirements</span></span>

|<span data-ttu-id="dde3f-106">Требование</span><span class="sxs-lookup"><span data-stu-id="dde3f-106">Requirement</span></span>| <span data-ttu-id="dde3f-107">Значение</span><span class="sxs-lookup"><span data-stu-id="dde3f-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="dde3f-108">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="dde3f-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dde3f-109">1.0</span><span class="sxs-lookup"><span data-stu-id="dde3f-109">1.0</span></span>|
|[<span data-ttu-id="dde3f-110">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="dde3f-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dde3f-111">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="dde3f-111">Restricted</span></span>|
|[<span data-ttu-id="dde3f-112">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="dde3f-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="dde3f-113">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="dde3f-113">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="dde3f-114">Элементы и методы</span><span class="sxs-lookup"><span data-stu-id="dde3f-114">Members and methods</span></span>

| <span data-ttu-id="dde3f-115">Элемент</span><span class="sxs-lookup"><span data-stu-id="dde3f-115">Member</span></span> | <span data-ttu-id="dde3f-116">Тип</span><span class="sxs-lookup"><span data-stu-id="dde3f-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="dde3f-117">ewsUrl</span><span class="sxs-lookup"><span data-stu-id="dde3f-117">ewsUrl</span></span>](#ewsurl-string) | <span data-ttu-id="dde3f-118">Элемент</span><span class="sxs-lookup"><span data-stu-id="dde3f-118">Member</span></span> |
| [<span data-ttu-id="dde3f-119">restUrl</span><span class="sxs-lookup"><span data-stu-id="dde3f-119">restUrl</span></span>](#resturl-string) | <span data-ttu-id="dde3f-120">Элемент</span><span class="sxs-lookup"><span data-stu-id="dde3f-120">Member</span></span> |
| [<span data-ttu-id="dde3f-121">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="dde3f-121">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="dde3f-122">Метод</span><span class="sxs-lookup"><span data-stu-id="dde3f-122">Method</span></span> |
| [<span data-ttu-id="dde3f-123">convertToEwsId</span><span class="sxs-lookup"><span data-stu-id="dde3f-123">convertToEwsId</span></span>](#converttoewsiditemid-restversion--string) | <span data-ttu-id="dde3f-124">Метод</span><span class="sxs-lookup"><span data-stu-id="dde3f-124">Method</span></span> |
| [<span data-ttu-id="dde3f-125">convertToLocalClientTime</span><span class="sxs-lookup"><span data-stu-id="dde3f-125">convertToLocalClientTime</span></span>](#converttolocalclienttimetimevalue--localclienttime) | <span data-ttu-id="dde3f-126">Метод</span><span class="sxs-lookup"><span data-stu-id="dde3f-126">Method</span></span> |
| [<span data-ttu-id="dde3f-127">convertToRestId</span><span class="sxs-lookup"><span data-stu-id="dde3f-127">convertToRestId</span></span>](#converttorestiditemid-restversion--string) | <span data-ttu-id="dde3f-128">Метод</span><span class="sxs-lookup"><span data-stu-id="dde3f-128">Method</span></span> |
| [<span data-ttu-id="dde3f-129">convertToUtcClientTime</span><span class="sxs-lookup"><span data-stu-id="dde3f-129">convertToUtcClientTime</span></span>](#converttoutcclienttimeinput--date) | <span data-ttu-id="dde3f-130">Метод</span><span class="sxs-lookup"><span data-stu-id="dde3f-130">Method</span></span> |
| [<span data-ttu-id="dde3f-131">displayAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="dde3f-131">displayAppointmentForm</span></span>](#displayappointmentformitemid) | <span data-ttu-id="dde3f-132">Метод</span><span class="sxs-lookup"><span data-stu-id="dde3f-132">Method</span></span> |
| [<span data-ttu-id="dde3f-133">displayMessageForm</span><span class="sxs-lookup"><span data-stu-id="dde3f-133">displayMessageForm</span></span>](#displaymessageformitemid) | <span data-ttu-id="dde3f-134">Метод</span><span class="sxs-lookup"><span data-stu-id="dde3f-134">Method</span></span> |
| [<span data-ttu-id="dde3f-135">displayNewAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="dde3f-135">displayNewAppointmentForm</span></span>](#displaynewappointmentformparameters) | <span data-ttu-id="dde3f-136">Метод</span><span class="sxs-lookup"><span data-stu-id="dde3f-136">Method</span></span> |
| [<span data-ttu-id="dde3f-137">дисплайневмессажеформ</span><span class="sxs-lookup"><span data-stu-id="dde3f-137">displayNewMessageForm</span></span>](#displaynewmessageformparameters) | <span data-ttu-id="dde3f-138">Метод</span><span class="sxs-lookup"><span data-stu-id="dde3f-138">Method</span></span> |
| [<span data-ttu-id="dde3f-139">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="dde3f-139">getCallbackTokenAsync</span></span>](#getcallbacktokenasyncoptions-callback) | <span data-ttu-id="dde3f-140">Метод</span><span class="sxs-lookup"><span data-stu-id="dde3f-140">Method</span></span> |
| [<span data-ttu-id="dde3f-141">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="dde3f-141">getCallbackTokenAsync</span></span>](#getcallbacktokenasynccallback-usercontext) | <span data-ttu-id="dde3f-142">Метод</span><span class="sxs-lookup"><span data-stu-id="dde3f-142">Method</span></span> |
| [<span data-ttu-id="dde3f-143">getUserIdentityTokenAsync</span><span class="sxs-lookup"><span data-stu-id="dde3f-143">getUserIdentityTokenAsync</span></span>](#getuseridentitytokenasynccallback-usercontext) | <span data-ttu-id="dde3f-144">Метод</span><span class="sxs-lookup"><span data-stu-id="dde3f-144">Method</span></span> |
| [<span data-ttu-id="dde3f-145">makeEwsRequestAsync</span><span class="sxs-lookup"><span data-stu-id="dde3f-145">makeEwsRequestAsync</span></span>](#makeewsrequestasyncdata-callback-usercontext) | <span data-ttu-id="dde3f-146">Метод</span><span class="sxs-lookup"><span data-stu-id="dde3f-146">Method</span></span> |
| [<span data-ttu-id="dde3f-147">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="dde3f-147">removeHandlerAsync</span></span>](#removehandlerasynceventtype-options-callback) | <span data-ttu-id="dde3f-148">Метод</span><span class="sxs-lookup"><span data-stu-id="dde3f-148">Method</span></span> |

### <a name="namespaces"></a><span data-ttu-id="dde3f-149">Пространства имен</span><span class="sxs-lookup"><span data-stu-id="dde3f-149">Namespaces</span></span>

<span data-ttu-id="dde3f-150">[diagnostics](Office.context.mailbox.diagnostics.md). Предоставляет надстройке Outlook диагностические сведения.</span><span class="sxs-lookup"><span data-stu-id="dde3f-150">[diagnostics](Office.context.mailbox.diagnostics.md): Provides diagnostic information to an Outlook add-in.</span></span>

<span data-ttu-id="dde3f-151">[item](Office.context.mailbox.item.md). Предоставляет методы и свойства для доступа к сообщению или встрече в надстройке Outlook.</span><span class="sxs-lookup"><span data-stu-id="dde3f-151">[item](Office.context.mailbox.item.md): Provides methods and properties for accessing a message or appointment in an Outlook add-in.</span></span>

<span data-ttu-id="dde3f-152">[userProfile](Office.context.mailbox.userProfile.md). Предоставляет сведения о пользователе в надстройке Outlook.</span><span class="sxs-lookup"><span data-stu-id="dde3f-152">[userProfile](Office.context.mailbox.userProfile.md): Provides information about the user in an Outlook add-in.</span></span>

### <a name="members"></a><span data-ttu-id="dde3f-153">Members</span><span class="sxs-lookup"><span data-stu-id="dde3f-153">Members</span></span>

#### <a name="ewsurl-string"></a><span data-ttu-id="dde3f-154">ewsUrl: String</span><span class="sxs-lookup"><span data-stu-id="dde3f-154">ewsUrl: String</span></span>

<span data-ttu-id="dde3f-p101">Получает URL-адрес конечной точки веб-служб Exchange (EWS) для этой учетной записи электронной почты. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="dde3f-p101">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="dde3f-157">Этот элемент не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="dde3f-157">This member is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="dde3f-p102">Удаленная служба может использовать значение `ewsUrl`, чтобы выполнять вызовы EWS для почтового ящика пользователя. Например, вы можете создать удаленную службу, чтобы [получить вложения из выбранного элемента](/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="dde3f-p102">The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="dde3f-160">Чтобы вызвать элемент `ewsUrl` в режиме чтения, в манифесте приложения должно быть указано разрешение **ReadItem**.</span><span class="sxs-lookup"><span data-stu-id="dde3f-160">Your app must have the **ReadItem** permission specified in its manifest to call the `ewsUrl` member in read mode.</span></span>

<span data-ttu-id="dde3f-p103">Перед использованием элемента `ewsUrl` в режиме создания необходимо вызвать метод [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback). Для вызова метода `saveAsync` приложение должно иметь разрешения **ReadWriteItem**.</span><span class="sxs-lookup"><span data-stu-id="dde3f-p103">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `ewsUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="dde3f-163">Тип</span><span class="sxs-lookup"><span data-stu-id="dde3f-163">Type</span></span>

*   <span data-ttu-id="dde3f-164">String</span><span class="sxs-lookup"><span data-stu-id="dde3f-164">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="dde3f-165">Требования</span><span class="sxs-lookup"><span data-stu-id="dde3f-165">Requirements</span></span>

|<span data-ttu-id="dde3f-166">Требование</span><span class="sxs-lookup"><span data-stu-id="dde3f-166">Requirement</span></span>| <span data-ttu-id="dde3f-167">Значение</span><span class="sxs-lookup"><span data-stu-id="dde3f-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="dde3f-168">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="dde3f-168">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dde3f-169">1.0</span><span class="sxs-lookup"><span data-stu-id="dde3f-169">1.0</span></span>|
|[<span data-ttu-id="dde3f-170">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="dde3f-170">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dde3f-171">ReadItem</span><span class="sxs-lookup"><span data-stu-id="dde3f-171">ReadItem</span></span>|
|[<span data-ttu-id="dde3f-172">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="dde3f-172">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="dde3f-173">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="dde3f-173">Compose or Read</span></span>|

<br>

---
---

#### <a name="resturl-string"></a><span data-ttu-id="dde3f-174">restUrl: String</span><span class="sxs-lookup"><span data-stu-id="dde3f-174">restUrl: String</span></span>

<span data-ttu-id="dde3f-175">Возвращает URL-адрес конечной точки REST для этой учетной записи электронной почты.</span><span class="sxs-lookup"><span data-stu-id="dde3f-175">Gets the URL of the REST endpoint for this email account.</span></span>

<span data-ttu-id="dde3f-176">С помощью значения `restUrl` можно выполнять вызовы [REST API](/outlook/rest/) для почтового ящика пользователя.</span><span class="sxs-lookup"><span data-stu-id="dde3f-176">The `restUrl` value can be used to make [REST API](/outlook/rest/) calls to the user's mailbox.</span></span>

<span data-ttu-id="dde3f-177">Чтобы вызвать элемент `restUrl` в режиме чтения, в манифесте приложения необходимо указать разрешение **ReadItem**.</span><span class="sxs-lookup"><span data-stu-id="dde3f-177">Your app must have the **ReadItem** permission specified in its manifest to call the `restUrl` member in read mode.</span></span>

<span data-ttu-id="dde3f-p104">Перед использованием элемента `restUrl` в режиме создания необходимо вызвать метод [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback). Для вызова метода `saveAsync` приложение должно иметь разрешения **ReadWriteItem**.</span><span class="sxs-lookup"><span data-stu-id="dde3f-p104">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `restUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="dde3f-180">Тип</span><span class="sxs-lookup"><span data-stu-id="dde3f-180">Type</span></span>

*   <span data-ttu-id="dde3f-181">String</span><span class="sxs-lookup"><span data-stu-id="dde3f-181">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="dde3f-182">Требования</span><span class="sxs-lookup"><span data-stu-id="dde3f-182">Requirements</span></span>

|<span data-ttu-id="dde3f-183">Требование</span><span class="sxs-lookup"><span data-stu-id="dde3f-183">Requirement</span></span>| <span data-ttu-id="dde3f-184">Значение</span><span class="sxs-lookup"><span data-stu-id="dde3f-184">Value</span></span>|
|---|---|
|[<span data-ttu-id="dde3f-185">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="dde3f-185">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dde3f-186">1.5</span><span class="sxs-lookup"><span data-stu-id="dde3f-186">1.5</span></span> |
|[<span data-ttu-id="dde3f-187">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="dde3f-187">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dde3f-188">ReadItem</span><span class="sxs-lookup"><span data-stu-id="dde3f-188">ReadItem</span></span>|
|[<span data-ttu-id="dde3f-189">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="dde3f-189">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="dde3f-190">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="dde3f-190">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="dde3f-191">Методы</span><span class="sxs-lookup"><span data-stu-id="dde3f-191">Methods</span></span>

#### <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="dde3f-192">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="dde3f-192">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="dde3f-193">Добавляет обработчик для поддерживаемого события.</span><span class="sxs-lookup"><span data-stu-id="dde3f-193">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="dde3f-194">В настоящее время единственный поддерживаемый тип события — `Office.EventType.ItemChanged`.</span><span class="sxs-lookup"><span data-stu-id="dde3f-194">Currently, the only supported event type is `Office.EventType.ItemChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="dde3f-195">Параметры</span><span class="sxs-lookup"><span data-stu-id="dde3f-195">Parameters</span></span>

| <span data-ttu-id="dde3f-196">Имя</span><span class="sxs-lookup"><span data-stu-id="dde3f-196">Name</span></span> | <span data-ttu-id="dde3f-197">Тип</span><span class="sxs-lookup"><span data-stu-id="dde3f-197">Type</span></span> | <span data-ttu-id="dde3f-198">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="dde3f-198">Attributes</span></span> | <span data-ttu-id="dde3f-199">Описание</span><span class="sxs-lookup"><span data-stu-id="dde3f-199">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="dde3f-200">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="dde3f-200">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="dde3f-201">Событие, которое должно вызвать обработчик.</span><span class="sxs-lookup"><span data-stu-id="dde3f-201">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="dde3f-202">Function</span><span class="sxs-lookup"><span data-stu-id="dde3f-202">Function</span></span> || <span data-ttu-id="dde3f-p105">Функция для обработки события. Функция должна принимать один параметр, представляющий собой объектный литерал. Значение свойства `type` параметра совпадет со значением параметра `eventType`, переданного методу `addHandlerAsync`.</span><span class="sxs-lookup"><span data-stu-id="dde3f-p105">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="dde3f-206">Объект</span><span class="sxs-lookup"><span data-stu-id="dde3f-206">Object</span></span> | <span data-ttu-id="dde3f-207">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="dde3f-207">&lt;optional&gt;</span></span> | <span data-ttu-id="dde3f-208">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="dde3f-208">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="dde3f-209">Object</span><span class="sxs-lookup"><span data-stu-id="dde3f-209">Object</span></span> | <span data-ttu-id="dde3f-210">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="dde3f-210">&lt;optional&gt;</span></span> | <span data-ttu-id="dde3f-211">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="dde3f-211">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="dde3f-212">функция</span><span class="sxs-lookup"><span data-stu-id="dde3f-212">function</span></span>| <span data-ttu-id="dde3f-213">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="dde3f-213">&lt;optional&gt;</span></span>|<span data-ttu-id="dde3f-214">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="dde3f-214">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="dde3f-215">Требования</span><span class="sxs-lookup"><span data-stu-id="dde3f-215">Requirements</span></span>

|<span data-ttu-id="dde3f-216">Требование</span><span class="sxs-lookup"><span data-stu-id="dde3f-216">Requirement</span></span>| <span data-ttu-id="dde3f-217">Значение</span><span class="sxs-lookup"><span data-stu-id="dde3f-217">Value</span></span>|
|---|---|
|[<span data-ttu-id="dde3f-218">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="dde3f-218">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dde3f-219">1.5</span><span class="sxs-lookup"><span data-stu-id="dde3f-219">1.5</span></span> |
|[<span data-ttu-id="dde3f-220">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="dde3f-220">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dde3f-221">ReadItem</span><span class="sxs-lookup"><span data-stu-id="dde3f-221">ReadItem</span></span> |
|[<span data-ttu-id="dde3f-222">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="dde3f-222">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="dde3f-223">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="dde3f-223">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="dde3f-224">Пример</span><span class="sxs-lookup"><span data-stu-id="dde3f-224">Example</span></span>

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

#### <a name="converttoewsiditemid-restversion--string"></a><span data-ttu-id="dde3f-225">convertToEwsId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="dde3f-225">convertToEwsId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="dde3f-226">Преобразовывает идентификатор элемента из формата REST в формат EWS.</span><span class="sxs-lookup"><span data-stu-id="dde3f-226">Converts an item ID formatted for REST into EWS format.</span></span>

> [!NOTE]
> <span data-ttu-id="dde3f-227">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="dde3f-227">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="dde3f-p106">Формат идентификаторов, извлекаемых через API REST (например, [API Почты Outlook](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) или [Microsoft Graph](https://graph.microsoft.io/)), отличается от формата веб-служб Exchange (EWS). Метод `convertToEwsId` преобразовывает идентификатор в формате REST в формат EWS.</span><span class="sxs-lookup"><span data-stu-id="dde3f-p106">Item IDs retrieved via a REST API (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)) use a different format than the format used by Exchange Web Services (EWS). The `convertToEwsId` method converts a REST-formatted ID into the proper format for EWS.</span></span>

##### <a name="parameters"></a><span data-ttu-id="dde3f-230">Параметры</span><span class="sxs-lookup"><span data-stu-id="dde3f-230">Parameters</span></span>

|<span data-ttu-id="dde3f-231">Имя</span><span class="sxs-lookup"><span data-stu-id="dde3f-231">Name</span></span>| <span data-ttu-id="dde3f-232">Тип</span><span class="sxs-lookup"><span data-stu-id="dde3f-232">Type</span></span>| <span data-ttu-id="dde3f-233">Описание</span><span class="sxs-lookup"><span data-stu-id="dde3f-233">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="dde3f-234">String</span><span class="sxs-lookup"><span data-stu-id="dde3f-234">String</span></span>|<span data-ttu-id="dde3f-235">Идентификатор элемента в формате REST API для Outlook</span><span class="sxs-lookup"><span data-stu-id="dde3f-235">An item ID formatted for the Outlook REST APIs</span></span>|
|`restVersion`| [<span data-ttu-id="dde3f-236">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="dde3f-236">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion?view=outlook-js-1.7)|<span data-ttu-id="dde3f-237">Значение, определяющее версию REST API для Outlook, которая используется для извлечения идентификатора элемента.</span><span class="sxs-lookup"><span data-stu-id="dde3f-237">A value indicating the version of the Outlook REST API used to retrieve the item ID.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="dde3f-238">Требования</span><span class="sxs-lookup"><span data-stu-id="dde3f-238">Requirements</span></span>

|<span data-ttu-id="dde3f-239">Требование</span><span class="sxs-lookup"><span data-stu-id="dde3f-239">Requirement</span></span>| <span data-ttu-id="dde3f-240">Значение</span><span class="sxs-lookup"><span data-stu-id="dde3f-240">Value</span></span>|
|---|---|
|[<span data-ttu-id="dde3f-241">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="dde3f-241">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dde3f-242">1.3</span><span class="sxs-lookup"><span data-stu-id="dde3f-242">1.3</span></span>|
|[<span data-ttu-id="dde3f-243">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="dde3f-243">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dde3f-244">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="dde3f-244">Restricted</span></span>|
|[<span data-ttu-id="dde3f-245">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="dde3f-245">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="dde3f-246">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="dde3f-246">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="dde3f-247">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="dde3f-247">Returns:</span></span>

<span data-ttu-id="dde3f-248">Тип: String</span><span class="sxs-lookup"><span data-stu-id="dde3f-248">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="dde3f-249">Пример</span><span class="sxs-lookup"><span data-stu-id="dde3f-249">Example</span></span>

```js
// Get an item's ID from a REST API.
var restId = 'AAMkAGVlOTZjNTM3LW...';

// Treat restId as coming from the v2.0 version of the Outlook Mail API.
var ewsId = Office.context.mailbox.convertToEwsId(restId, Office.MailboxEnums.RestVersion.v2_0);
```

<br>

---
---

#### <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttimeviewoutlook-js-17"></a><span data-ttu-id="dde3f-250">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.7)}</span><span class="sxs-lookup"><span data-stu-id="dde3f-250">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.7)}</span></span>

<span data-ttu-id="dde3f-251">Получает словарь, содержащий сведения о локальном времени клиента.</span><span class="sxs-lookup"><span data-stu-id="dde3f-251">Gets a dictionary containing time information in local client time.</span></span>

<span data-ttu-id="dde3f-p107">Почтовое приложение для классической версии Outlook или версии в Интернете может использовать разные часовые пояса для дат и времени. Классическое приложение Outlook использует часовой пояс клиентского компьютера. Outlook в Интернете использует часовой пояс, заданный в Центре администрирования Exchange (EAC). Значения даты и времени должны обрабатываться так, чтобы значения в пользовательском интерфейсе всегда согласовывались с часовым поясом, ожидаемым пользователем.</span><span class="sxs-lookup"><span data-stu-id="dde3f-p107">A mail app for Outlook on a desktop or on the web can use different time zones for the dates and times. Outlook on a desktop uses the client computer time zone; Outlook on the web uses the time zone set on the Exchange Admin Center (EAC). You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that the user expects.</span></span>

<span data-ttu-id="dde3f-p108">Если почтовое приложение работает в классическом клиенте Outlook, метод `convertToLocalClientTime` вернет объект словаря со значениями часового пояса клиентского компьютера. Если почтовое приложение работает в Outlook в Интернете, метод `convertToLocalClientTime` вернет объект словаря со значениями часового пояса, заданного в Центре администрирования Exchange.</span><span class="sxs-lookup"><span data-stu-id="dde3f-p108">If the mail app is running in Outlook on a desktop client, the `convertToLocalClientTime` method will return a dictionary object with the values set to the client computer time zone. If the mail app is running in Outlook on the web, the `convertToLocalClientTime` method will return a dictionary object with the values set to the time zone specified in the EAC.</span></span>

##### <a name="parameters"></a><span data-ttu-id="dde3f-257">Параметры</span><span class="sxs-lookup"><span data-stu-id="dde3f-257">Parameters</span></span>

|<span data-ttu-id="dde3f-258">Имя</span><span class="sxs-lookup"><span data-stu-id="dde3f-258">Name</span></span>| <span data-ttu-id="dde3f-259">Тип</span><span class="sxs-lookup"><span data-stu-id="dde3f-259">Type</span></span>| <span data-ttu-id="dde3f-260">Описание</span><span class="sxs-lookup"><span data-stu-id="dde3f-260">Description</span></span>|
|---|---|---|
|`timeValue`| <span data-ttu-id="dde3f-261">Date</span><span class="sxs-lookup"><span data-stu-id="dde3f-261">Date</span></span>|<span data-ttu-id="dde3f-262">Объект Date</span><span class="sxs-lookup"><span data-stu-id="dde3f-262">A Date object</span></span>|

##### <a name="requirements"></a><span data-ttu-id="dde3f-263">Требования</span><span class="sxs-lookup"><span data-stu-id="dde3f-263">Requirements</span></span>

|<span data-ttu-id="dde3f-264">Требование</span><span class="sxs-lookup"><span data-stu-id="dde3f-264">Requirement</span></span>| <span data-ttu-id="dde3f-265">Значение</span><span class="sxs-lookup"><span data-stu-id="dde3f-265">Value</span></span>|
|---|---|
|[<span data-ttu-id="dde3f-266">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="dde3f-266">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dde3f-267">1.0</span><span class="sxs-lookup"><span data-stu-id="dde3f-267">1.0</span></span>|
|[<span data-ttu-id="dde3f-268">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="dde3f-268">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dde3f-269">ReadItem</span><span class="sxs-lookup"><span data-stu-id="dde3f-269">ReadItem</span></span>|
|[<span data-ttu-id="dde3f-270">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="dde3f-270">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="dde3f-271">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="dde3f-271">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="dde3f-272">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="dde3f-272">Returns:</span></span>

<span data-ttu-id="dde3f-273">Тип: [LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="dde3f-273">Type: [LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.7)</span></span>

<br>

---
---

#### <a name="converttorestiditemid-restversion--string"></a><span data-ttu-id="dde3f-274">convertToRestId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="dde3f-274">convertToRestId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="dde3f-275">Преобразовывает идентификатор элемента в формате EWS в формат REST.</span><span class="sxs-lookup"><span data-stu-id="dde3f-275">Converts an item ID formatted for EWS into REST format.</span></span>

> [!NOTE]
> <span data-ttu-id="dde3f-276">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="dde3f-276">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="dde3f-p109">Формат идентификаторов, извлекаемых через EWS или свойство `itemId`, отличается от формата API REST (таких как [API Почты Outlook](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) или [Microsoft Graph](https://graph.microsoft.io/)). Метод `convertToRestId` преобразовывает идентификатор в формате EWS в формат REST.</span><span class="sxs-lookup"><span data-stu-id="dde3f-p109">Item IDs retrieved via EWS or via the `itemId` property use a different format than the format used by REST APIs (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)). The `convertToRestId` method converts an EWS-formatted ID into the proper format for REST.</span></span>

##### <a name="parameters"></a><span data-ttu-id="dde3f-279">Параметры</span><span class="sxs-lookup"><span data-stu-id="dde3f-279">Parameters</span></span>

|<span data-ttu-id="dde3f-280">Имя</span><span class="sxs-lookup"><span data-stu-id="dde3f-280">Name</span></span>| <span data-ttu-id="dde3f-281">Тип</span><span class="sxs-lookup"><span data-stu-id="dde3f-281">Type</span></span>| <span data-ttu-id="dde3f-282">Описание</span><span class="sxs-lookup"><span data-stu-id="dde3f-282">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="dde3f-283">String</span><span class="sxs-lookup"><span data-stu-id="dde3f-283">String</span></span>|<span data-ttu-id="dde3f-284">Идентификатор элемента в формате EWS</span><span class="sxs-lookup"><span data-stu-id="dde3f-284">An item ID formatted for Exchange Web Services (EWS)</span></span>|
|`restVersion`| [<span data-ttu-id="dde3f-285">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="dde3f-285">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion?view=outlook-js-1.7)|<span data-ttu-id="dde3f-286">Значение, определяющее версию REST API для Outlook, с которой будет использоваться преобразованный идентификатор.</span><span class="sxs-lookup"><span data-stu-id="dde3f-286">A value indicating the version of the Outlook REST API that the converted ID will be used with.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="dde3f-287">Требования</span><span class="sxs-lookup"><span data-stu-id="dde3f-287">Requirements</span></span>

|<span data-ttu-id="dde3f-288">Требование</span><span class="sxs-lookup"><span data-stu-id="dde3f-288">Requirement</span></span>| <span data-ttu-id="dde3f-289">Значение</span><span class="sxs-lookup"><span data-stu-id="dde3f-289">Value</span></span>|
|---|---|
|[<span data-ttu-id="dde3f-290">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="dde3f-290">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dde3f-291">1.3</span><span class="sxs-lookup"><span data-stu-id="dde3f-291">1.3</span></span>|
|[<span data-ttu-id="dde3f-292">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="dde3f-292">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dde3f-293">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="dde3f-293">Restricted</span></span>|
|[<span data-ttu-id="dde3f-294">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="dde3f-294">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="dde3f-295">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="dde3f-295">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="dde3f-296">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="dde3f-296">Returns:</span></span>

<span data-ttu-id="dde3f-297">Тип: String</span><span class="sxs-lookup"><span data-stu-id="dde3f-297">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="dde3f-298">Пример</span><span class="sxs-lookup"><span data-stu-id="dde3f-298">Example</span></span>

```js
// Get the currently selected item's ID.
var ewsId = Office.context.mailbox.item.itemId;

// Convert to a REST ID for the v2.0 version of the Outlook Mail API.
var restId = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
```

<br>

---
---

#### <a name="converttoutcclienttimeinput--date"></a><span data-ttu-id="dde3f-299">convertToUtcClientTime(input) → {Date}</span><span class="sxs-lookup"><span data-stu-id="dde3f-299">convertToUtcClientTime(input) → {Date}</span></span>

<span data-ttu-id="dde3f-300">Получает объект Date из словаря, содержащего сведения о времени.</span><span class="sxs-lookup"><span data-stu-id="dde3f-300">Gets a Date object from a dictionary containing time information.</span></span>

<span data-ttu-id="dde3f-301">Метод `convertToUtcClientTime` преобразует словарь, содержащий локальную дату и время, в объект Date с правильными значениями локальной даты и времени.</span><span class="sxs-lookup"><span data-stu-id="dde3f-301">The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a Date object with the correct values for the local date and time.</span></span>

##### <a name="parameters"></a><span data-ttu-id="dde3f-302">Параметры</span><span class="sxs-lookup"><span data-stu-id="dde3f-302">Parameters</span></span>

|<span data-ttu-id="dde3f-303">Имя</span><span class="sxs-lookup"><span data-stu-id="dde3f-303">Name</span></span>| <span data-ttu-id="dde3f-304">Тип</span><span class="sxs-lookup"><span data-stu-id="dde3f-304">Type</span></span>| <span data-ttu-id="dde3f-305">Описание</span><span class="sxs-lookup"><span data-stu-id="dde3f-305">Description</span></span>|
|---|---|---|
|`input`| [<span data-ttu-id="dde3f-306">LocalClientTime</span><span class="sxs-lookup"><span data-stu-id="dde3f-306">LocalClientTime</span></span>](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.7)|<span data-ttu-id="dde3f-307">Значение локального времени для преобразования.</span><span class="sxs-lookup"><span data-stu-id="dde3f-307">The local time value to convert.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="dde3f-308">Требования</span><span class="sxs-lookup"><span data-stu-id="dde3f-308">Requirements</span></span>

|<span data-ttu-id="dde3f-309">Требование</span><span class="sxs-lookup"><span data-stu-id="dde3f-309">Requirement</span></span>| <span data-ttu-id="dde3f-310">Значение</span><span class="sxs-lookup"><span data-stu-id="dde3f-310">Value</span></span>|
|---|---|
|[<span data-ttu-id="dde3f-311">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="dde3f-311">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dde3f-312">1.0</span><span class="sxs-lookup"><span data-stu-id="dde3f-312">1.0</span></span>|
|[<span data-ttu-id="dde3f-313">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="dde3f-313">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dde3f-314">ReadItem</span><span class="sxs-lookup"><span data-stu-id="dde3f-314">ReadItem</span></span>|
|[<span data-ttu-id="dde3f-315">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="dde3f-315">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="dde3f-316">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="dde3f-316">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="dde3f-317">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="dde3f-317">Returns:</span></span>

<span data-ttu-id="dde3f-318">Объект Date со временем в формате UTC.</span><span class="sxs-lookup"><span data-stu-id="dde3f-318">A Date object with the time expressed in UTC.</span></span>

<span data-ttu-id="dde3f-319">Тип: Date</span><span class="sxs-lookup"><span data-stu-id="dde3f-319">Type: Date</span></span>

##### <a name="example"></a><span data-ttu-id="dde3f-320">Пример</span><span class="sxs-lookup"><span data-stu-id="dde3f-320">Example</span></span>

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

#### <a name="displayappointmentformitemid"></a><span data-ttu-id="dde3f-321">displayAppointmentForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="dde3f-321">displayAppointmentForm(itemId)</span></span>

<span data-ttu-id="dde3f-322">Отображает имеющуюся встречу из календаря.</span><span class="sxs-lookup"><span data-stu-id="dde3f-322">Displays an existing calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="dde3f-323">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="dde3f-323">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="dde3f-324">Метод `displayAppointmentForm` открывает новое окно на компьютере или диалоговое окно на мобильном устройстве, содержащее сведения календаря о существующей встрече.</span><span class="sxs-lookup"><span data-stu-id="dde3f-324">The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="dde3f-p110">В Outlook для Mac с помощью этого метода можно отобразить одну встречу, которая не является частью повторяющегося ряда, или основную встречу такого ряда, но не экземпляр из него, так как в Outlook для Mac невозможно получить доступ к свойствам экземпляра повторяющегося ряда (в том числе к идентификатору элемента).</span><span class="sxs-lookup"><span data-stu-id="dde3f-p110">In Outlook on Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment of a recurring series, but you cannot display an instance of the series. This is because in Outlook on Mac, you cannot access the properties (including the item ID) of instances of a recurring series.</span></span>

<span data-ttu-id="dde3f-327">В Outlook в Интернете этот метод открывает указанную форму, только если текст формы содержит символы размером не более 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="dde3f-327">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32KB number of characters.</span></span>

<span data-ttu-id="dde3f-328">Если указанный идентификатор элемента не определяет существующую встречу, на клиентском компьютере или устройстве открывается пустая страница, и сообщение об ошибке не возвращается.</span><span class="sxs-lookup"><span data-stu-id="dde3f-328">If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and no error message will be returned.</span></span>

##### <a name="parameters"></a><span data-ttu-id="dde3f-329">Параметры</span><span class="sxs-lookup"><span data-stu-id="dde3f-329">Parameters</span></span>

|<span data-ttu-id="dde3f-330">Имя</span><span class="sxs-lookup"><span data-stu-id="dde3f-330">Name</span></span>| <span data-ttu-id="dde3f-331">Тип</span><span class="sxs-lookup"><span data-stu-id="dde3f-331">Type</span></span>| <span data-ttu-id="dde3f-332">Описание</span><span class="sxs-lookup"><span data-stu-id="dde3f-332">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="dde3f-333">String</span><span class="sxs-lookup"><span data-stu-id="dde3f-333">String</span></span>|<span data-ttu-id="dde3f-334">Идентификатор веб-служб Exchange для существующей встречи в календаре.</span><span class="sxs-lookup"><span data-stu-id="dde3f-334">The Exchange Web Services (EWS) identifier for an existing calendar appointment.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="dde3f-335">Требования</span><span class="sxs-lookup"><span data-stu-id="dde3f-335">Requirements</span></span>

|<span data-ttu-id="dde3f-336">Требование</span><span class="sxs-lookup"><span data-stu-id="dde3f-336">Requirement</span></span>| <span data-ttu-id="dde3f-337">Значение</span><span class="sxs-lookup"><span data-stu-id="dde3f-337">Value</span></span>|
|---|---|
|[<span data-ttu-id="dde3f-338">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="dde3f-338">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dde3f-339">1.0</span><span class="sxs-lookup"><span data-stu-id="dde3f-339">1.0</span></span>|
|[<span data-ttu-id="dde3f-340">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="dde3f-340">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dde3f-341">ReadItem</span><span class="sxs-lookup"><span data-stu-id="dde3f-341">ReadItem</span></span>|
|[<span data-ttu-id="dde3f-342">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="dde3f-342">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="dde3f-343">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="dde3f-343">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="dde3f-344">Пример</span><span class="sxs-lookup"><span data-stu-id="dde3f-344">Example</span></span>

```js
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

<br>

---
---

#### <a name="displaymessageformitemid"></a><span data-ttu-id="dde3f-345">displayMessageForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="dde3f-345">displayMessageForm(itemId)</span></span>

<span data-ttu-id="dde3f-346">Отображает имеющееся сообщение.</span><span class="sxs-lookup"><span data-stu-id="dde3f-346">Displays an existing message.</span></span>

> [!NOTE]
> <span data-ttu-id="dde3f-347">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="dde3f-347">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="dde3f-348">Метод `displayMessageForm` открывает новое окно на компьютере или диалоговое окно на мобильном устройстве, содержащее существующее сообщение.</span><span class="sxs-lookup"><span data-stu-id="dde3f-348">The `displayMessageForm` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="dde3f-349">В Outlook в Интернете этот метод открывает указанную форму, только если текст формы содержит символы размером не более 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="dde3f-349">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32 KB number of characters.</span></span>

<span data-ttu-id="dde3f-350">Если указанный идентификатор элемента не определяет существующее сообщение, окно на клиентском компьютере не открывается и сообщение об ошибке не возвращается.</span><span class="sxs-lookup"><span data-stu-id="dde3f-350">If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and no error message will be returned.</span></span>

<span data-ttu-id="dde3f-p111">Не используйте `displayMessageForm` с параметром `itemId`, который представляет собой встречу. Используйте метод `displayAppointmentForm`, чтобы отобразить сведения о существующей встрече, а метод `displayNewAppointmentForm` — для отображения формы создания встречи.</span><span class="sxs-lookup"><span data-stu-id="dde3f-p111">Do not use the `displayMessageForm` with an `itemId` that represents an appointment. Use the `displayAppointmentForm` method to display an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.</span></span>

##### <a name="parameters"></a><span data-ttu-id="dde3f-353">Параметры</span><span class="sxs-lookup"><span data-stu-id="dde3f-353">Parameters</span></span>

|<span data-ttu-id="dde3f-354">Имя</span><span class="sxs-lookup"><span data-stu-id="dde3f-354">Name</span></span>| <span data-ttu-id="dde3f-355">Тип</span><span class="sxs-lookup"><span data-stu-id="dde3f-355">Type</span></span>| <span data-ttu-id="dde3f-356">Описание</span><span class="sxs-lookup"><span data-stu-id="dde3f-356">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="dde3f-357">String</span><span class="sxs-lookup"><span data-stu-id="dde3f-357">String</span></span>|<span data-ttu-id="dde3f-358">Идентификатор веб-служб Exchange для существующего сообщения.</span><span class="sxs-lookup"><span data-stu-id="dde3f-358">The Exchange Web Services (EWS) identifier for an existing message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="dde3f-359">Требования</span><span class="sxs-lookup"><span data-stu-id="dde3f-359">Requirements</span></span>

|<span data-ttu-id="dde3f-360">Требование</span><span class="sxs-lookup"><span data-stu-id="dde3f-360">Requirement</span></span>| <span data-ttu-id="dde3f-361">Значение</span><span class="sxs-lookup"><span data-stu-id="dde3f-361">Value</span></span>|
|---|---|
|[<span data-ttu-id="dde3f-362">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="dde3f-362">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dde3f-363">1.0</span><span class="sxs-lookup"><span data-stu-id="dde3f-363">1.0</span></span>|
|[<span data-ttu-id="dde3f-364">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="dde3f-364">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dde3f-365">ReadItem</span><span class="sxs-lookup"><span data-stu-id="dde3f-365">ReadItem</span></span>|
|[<span data-ttu-id="dde3f-366">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="dde3f-366">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="dde3f-367">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="dde3f-367">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="dde3f-368">Пример</span><span class="sxs-lookup"><span data-stu-id="dde3f-368">Example</span></span>

```js
Office.context.mailbox.displayMessageForm(messageId);
```

<br>

---
---

#### <a name="displaynewappointmentformparameters"></a><span data-ttu-id="dde3f-369">displayNewAppointmentForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="dde3f-369">displayNewAppointmentForm(parameters)</span></span>

<span data-ttu-id="dde3f-370">Отображает форму для создания новой встречи в календаре.</span><span class="sxs-lookup"><span data-stu-id="dde3f-370">Displays a form for creating a new calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="dde3f-371">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="dde3f-371">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="dde3f-p112">Метод `displayNewAppointmentForm` открывает форму, в которой пользователь может создать встречу или собрание. Если параметры заданы, поля формы встречи автоматически заполняются их содержимым.</span><span class="sxs-lookup"><span data-stu-id="dde3f-p112">The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting. If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="dde3f-p113">В Outlook в Интернете и на мобильных устройствах этот метод всегда отображает форму с полем участников. Если вы не укажете участников в качестве входных аргументов, метод отображает форму с кнопкой **Сохранить**. Если вы укажете участников, форма будет включать участников и кнопку **Отправить**.</span><span class="sxs-lookup"><span data-stu-id="dde3f-p113">In Outlook on the web and mobile devices, this method always displays a form with an attendees field. If you do not specify any attendees as input arguments, the method displays a form with a **Save** button. If you have specified attendees, the form would include the attendees and a **Send** button.</span></span>

<span data-ttu-id="dde3f-p114">Если вы укажете участников или ресурсы с помощью параметра `requiredAttendees`, `optionalAttendees` или `resources` в клиенте Outlook с расширенными возможностями и Outlook RT, этот метод отобразит форму собрания с кнопкой **Отправить**. Если не указать получателей, этот метод отобразит форму встречи с кнопкой **Сохранить и закрыть**.</span><span class="sxs-lookup"><span data-stu-id="dde3f-p114">In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or `resources` parameter, this method displays a meeting form with a **Send** button. If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.</span></span>

<span data-ttu-id="dde3f-379">Если параметры превышают указанные ограничения размера или если указано неизвестное имя параметра, вызывается исключение.</span><span class="sxs-lookup"><span data-stu-id="dde3f-379">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="dde3f-380">Параметры</span><span class="sxs-lookup"><span data-stu-id="dde3f-380">Parameters</span></span>

> [!NOTE]
> <span data-ttu-id="dde3f-381">Все параметры являются необязательными.</span><span class="sxs-lookup"><span data-stu-id="dde3f-381">All parameters are optional.</span></span>

|<span data-ttu-id="dde3f-382">Имя</span><span class="sxs-lookup"><span data-stu-id="dde3f-382">Name</span></span>| <span data-ttu-id="dde3f-383">Тип</span><span class="sxs-lookup"><span data-stu-id="dde3f-383">Type</span></span>| <span data-ttu-id="dde3f-384">Описание</span><span class="sxs-lookup"><span data-stu-id="dde3f-384">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="dde3f-385">Object</span><span class="sxs-lookup"><span data-stu-id="dde3f-385">Object</span></span> | <span data-ttu-id="dde3f-386">Словарь параметров, описывающий новую встречу.</span><span class="sxs-lookup"><span data-stu-id="dde3f-386">A dictionary of parameters describing the new appointment.</span></span> |
| `parameters.requiredAttendees` | <span data-ttu-id="dde3f-387">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)&gt;</span><span class="sxs-lookup"><span data-stu-id="dde3f-387">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)&gt;</span></span> | <span data-ttu-id="dde3f-p115">Массив строк, содержащий электронные адреса, или массив, содержащий объекты `EmailAddressDetails` для каждого из обязательных участников встречи. Массив может включать не более 100 записей.</span><span class="sxs-lookup"><span data-stu-id="dde3f-p115">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the required attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.optionalAttendees` | <span data-ttu-id="dde3f-390">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)&gt;</span><span class="sxs-lookup"><span data-stu-id="dde3f-390">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)&gt;</span></span> | <span data-ttu-id="dde3f-p116">Массив строк, содержащий электронные адреса, или массив, содержащий объекты `EmailAddressDetails` для каждого из необязательных участников встречи. Массив может включать не более 100 записей.</span><span class="sxs-lookup"><span data-stu-id="dde3f-p116">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the optional attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.start` | <span data-ttu-id="dde3f-393">Date</span><span class="sxs-lookup"><span data-stu-id="dde3f-393">Date</span></span> | <span data-ttu-id="dde3f-394">Объект `Date`, указывающий дату и время начала встречи.</span><span class="sxs-lookup"><span data-stu-id="dde3f-394">A `Date` object specifying the start date and time of the appointment.</span></span> |
| `parameters.end` | <span data-ttu-id="dde3f-395">Date</span><span class="sxs-lookup"><span data-stu-id="dde3f-395">Date</span></span> | <span data-ttu-id="dde3f-396">Объект `Date`, указывающий дату и время окончания встречи.</span><span class="sxs-lookup"><span data-stu-id="dde3f-396">A `Date` object specifying the end date and time of the appointment.</span></span> |
| `parameters.location` | <span data-ttu-id="dde3f-397">String</span><span class="sxs-lookup"><span data-stu-id="dde3f-397">String</span></span> | <span data-ttu-id="dde3f-p117">Строка со сведениями о месте встречи. Максимальное количество символов в строке — 255.</span><span class="sxs-lookup"><span data-stu-id="dde3f-p117">A string containing the location of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.resources` | <span data-ttu-id="dde3f-400">Array.&lt;String&gt;</span><span class="sxs-lookup"><span data-stu-id="dde3f-400">Array.&lt;String&gt;</span></span> | <span data-ttu-id="dde3f-p118">Массив строк, содержащий необходимые для встречи ресурсы. Массив может включать не более 100 записей.</span><span class="sxs-lookup"><span data-stu-id="dde3f-p118">An array of strings containing the resources required for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="dde3f-403">String</span><span class="sxs-lookup"><span data-stu-id="dde3f-403">String</span></span> | <span data-ttu-id="dde3f-p119">Строка с темой встречи. Максимальное количество символов в строке — 255.</span><span class="sxs-lookup"><span data-stu-id="dde3f-p119">A string containing the subject of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.body` | <span data-ttu-id="dde3f-406">String</span><span class="sxs-lookup"><span data-stu-id="dde3f-406">String</span></span> | <span data-ttu-id="dde3f-p120">Текст сообщения о встрече. Максимальный размер содержимого сообщения — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="dde3f-p120">The body of the appointment. The body content is limited to a maximum size of 32 KB.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="dde3f-409">Требования</span><span class="sxs-lookup"><span data-stu-id="dde3f-409">Requirements</span></span>

|<span data-ttu-id="dde3f-410">Требование</span><span class="sxs-lookup"><span data-stu-id="dde3f-410">Requirement</span></span>| <span data-ttu-id="dde3f-411">Значение</span><span class="sxs-lookup"><span data-stu-id="dde3f-411">Value</span></span>|
|---|---|
|[<span data-ttu-id="dde3f-412">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="dde3f-412">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dde3f-413">1.0</span><span class="sxs-lookup"><span data-stu-id="dde3f-413">1.0</span></span>|
|[<span data-ttu-id="dde3f-414">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="dde3f-414">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dde3f-415">ReadItem</span><span class="sxs-lookup"><span data-stu-id="dde3f-415">ReadItem</span></span>|
|[<span data-ttu-id="dde3f-416">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="dde3f-416">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="dde3f-417">Чтение</span><span class="sxs-lookup"><span data-stu-id="dde3f-417">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="dde3f-418">Пример</span><span class="sxs-lookup"><span data-stu-id="dde3f-418">Example</span></span>

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

#### <a name="displaynewmessageformparameters"></a><span data-ttu-id="dde3f-419">Дисплайневмессажеформ (Parameters)</span><span class="sxs-lookup"><span data-stu-id="dde3f-419">displayNewMessageForm(parameters)</span></span>

<span data-ttu-id="dde3f-420">Отображает форму для создания нового сообщения.</span><span class="sxs-lookup"><span data-stu-id="dde3f-420">Displays a form for creating a new message.</span></span>

<span data-ttu-id="dde3f-421">`displayNewMessageForm` Метод открывает форму, которая позволяет пользователю создать новое сообщение.</span><span class="sxs-lookup"><span data-stu-id="dde3f-421">The `displayNewMessageForm` method opens a form that enables the user to create a new message.</span></span> <span data-ttu-id="dde3f-422">Если указаны параметры, поля формы сообщения автоматически заполняются содержимым параметров.</span><span class="sxs-lookup"><span data-stu-id="dde3f-422">If parameters are specified, the message form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="dde3f-423">Если параметры превышают указанные ограничения размера или если указано неизвестное имя параметра, вызывается исключение.</span><span class="sxs-lookup"><span data-stu-id="dde3f-423">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="dde3f-424">Параметры</span><span class="sxs-lookup"><span data-stu-id="dde3f-424">Parameters</span></span>

> [!NOTE]
> <span data-ttu-id="dde3f-425">Все параметры являются необязательными.</span><span class="sxs-lookup"><span data-stu-id="dde3f-425">All parameters are optional.</span></span>

|<span data-ttu-id="dde3f-426">Имя</span><span class="sxs-lookup"><span data-stu-id="dde3f-426">Name</span></span>| <span data-ttu-id="dde3f-427">Тип</span><span class="sxs-lookup"><span data-stu-id="dde3f-427">Type</span></span>| <span data-ttu-id="dde3f-428">Описание</span><span class="sxs-lookup"><span data-stu-id="dde3f-428">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="dde3f-429">Object</span><span class="sxs-lookup"><span data-stu-id="dde3f-429">Object</span></span> | <span data-ttu-id="dde3f-430">Словарь параметров, описывающих новое сообщение.</span><span class="sxs-lookup"><span data-stu-id="dde3f-430">A dictionary of parameters describing the new message.</span></span> |
| `parameters.toRecipients` | <span data-ttu-id="dde3f-431">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)&gt;</span><span class="sxs-lookup"><span data-stu-id="dde3f-431">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)&gt;</span></span> | <span data-ttu-id="dde3f-432">Массив строк, содержащий адреса электронной почты или массив, содержащий `EmailAddressDetails` объект для каждого из получателей в строке "Кому".</span><span class="sxs-lookup"><span data-stu-id="dde3f-432">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the To line.</span></span> <span data-ttu-id="dde3f-433">Массив может включать не более 100 записей.</span><span class="sxs-lookup"><span data-stu-id="dde3f-433">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.ccRecipients` | <span data-ttu-id="dde3f-434">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)&gt;</span><span class="sxs-lookup"><span data-stu-id="dde3f-434">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)&gt;</span></span> | <span data-ttu-id="dde3f-435">Массив строк, содержащий адреса электронной почты или массив, содержащий `EmailAddressDetails` объект для каждого получателя в строке "копия".</span><span class="sxs-lookup"><span data-stu-id="dde3f-435">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Cc line.</span></span> <span data-ttu-id="dde3f-436">Массив может включать не более 100 записей.</span><span class="sxs-lookup"><span data-stu-id="dde3f-436">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.bccRecipients` | <span data-ttu-id="dde3f-437">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)&gt;</span><span class="sxs-lookup"><span data-stu-id="dde3f-437">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)&gt;</span></span> | <span data-ttu-id="dde3f-438">Массив строк, содержащий адреса электронной почты или массив, содержащий `EmailAddressDetails` объект для каждого из получателей, указанных в строке "СК".</span><span class="sxs-lookup"><span data-stu-id="dde3f-438">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Bcc line.</span></span> <span data-ttu-id="dde3f-439">Массив может включать не более 100 записей.</span><span class="sxs-lookup"><span data-stu-id="dde3f-439">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="dde3f-440">String</span><span class="sxs-lookup"><span data-stu-id="dde3f-440">String</span></span> | <span data-ttu-id="dde3f-441">Строка, содержащая тему сообщения.</span><span class="sxs-lookup"><span data-stu-id="dde3f-441">A string containing the subject of the message.</span></span> <span data-ttu-id="dde3f-442">Максимальное количество символов в строке — 255.</span><span class="sxs-lookup"><span data-stu-id="dde3f-442">The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.htmlBody` | <span data-ttu-id="dde3f-443">String</span><span class="sxs-lookup"><span data-stu-id="dde3f-443">String</span></span> | <span data-ttu-id="dde3f-444">Текст сообщения в формате HTML.</span><span class="sxs-lookup"><span data-stu-id="dde3f-444">The HTML body of the message.</span></span> <span data-ttu-id="dde3f-445">Максимальный размер содержимого сообщения — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="dde3f-445">The body content is limited to a maximum size of 32 KB.</span></span> |
| `parameters.attachments` | <span data-ttu-id="dde3f-446">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="dde3f-446">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="dde3f-447">Массив объектов JSON, представляющих собой вложенные файлы или элементы.</span><span class="sxs-lookup"><span data-stu-id="dde3f-447">An array of JSON objects that are either file or item attachments.</span></span> |
| `parameters.attachments.type` | <span data-ttu-id="dde3f-448">String</span><span class="sxs-lookup"><span data-stu-id="dde3f-448">String</span></span> | <span data-ttu-id="dde3f-p127">Указывает тип вложения. Допустимые значения: `file` для вложенного файла и `item` для вложенного элемента.</span><span class="sxs-lookup"><span data-stu-id="dde3f-p127">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `parameters.attachments.name` | <span data-ttu-id="dde3f-451">Строка</span><span class="sxs-lookup"><span data-stu-id="dde3f-451">String</span></span> | <span data-ttu-id="dde3f-452">Строка, содержащая имя вложения, длиной до 255 символов.</span><span class="sxs-lookup"><span data-stu-id="dde3f-452">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `parameters.attachments.url` | <span data-ttu-id="dde3f-453">String</span><span class="sxs-lookup"><span data-stu-id="dde3f-453">String</span></span> | <span data-ttu-id="dde3f-p128">Используется, только если свойству `type` задано значение `file`. URI расположения файла.</span><span class="sxs-lookup"><span data-stu-id="dde3f-p128">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `parameters.attachments.isInline` | <span data-ttu-id="dde3f-456">Логический</span><span class="sxs-lookup"><span data-stu-id="dde3f-456">Boolean</span></span> | <span data-ttu-id="dde3f-p129">Используется, только если свойству `type` задано значение `file`. Значение `true` указывает на то, что вложение будет встроено в текст сообщения и не должно отображаться в списке вложений.</span><span class="sxs-lookup"><span data-stu-id="dde3f-p129">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `parameters.attachments.itemId` | <span data-ttu-id="dde3f-459">Строка</span><span class="sxs-lookup"><span data-stu-id="dde3f-459">String</span></span> | <span data-ttu-id="dde3f-460">Используется, только если свойству `type` присвоено значение `item`.</span><span class="sxs-lookup"><span data-stu-id="dde3f-460">Only used if `type` is set to `item`.</span></span> <span data-ttu-id="dde3f-461">Идентификатор элемента EWS существующего сообщения электронной почты, которое необходимо присоединить к новому сообщению.</span><span class="sxs-lookup"><span data-stu-id="dde3f-461">The EWS item id of the existing e-mail you want to attach to the new message.</span></span> <span data-ttu-id="dde3f-462">Это строка длиной до 100 символов.</span><span class="sxs-lookup"><span data-stu-id="dde3f-462">This is a string up to 100 characters.</span></span> |


##### <a name="requirements"></a><span data-ttu-id="dde3f-463">Требования</span><span class="sxs-lookup"><span data-stu-id="dde3f-463">Requirements</span></span>

|<span data-ttu-id="dde3f-464">Требование</span><span class="sxs-lookup"><span data-stu-id="dde3f-464">Requirement</span></span>| <span data-ttu-id="dde3f-465">Значение</span><span class="sxs-lookup"><span data-stu-id="dde3f-465">Value</span></span>|
|---|---|
|[<span data-ttu-id="dde3f-466">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="dde3f-466">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dde3f-467">1.6</span><span class="sxs-lookup"><span data-stu-id="dde3f-467">1.6</span></span> |
|[<span data-ttu-id="dde3f-468">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="dde3f-468">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dde3f-469">ReadItem</span><span class="sxs-lookup"><span data-stu-id="dde3f-469">ReadItem</span></span>|
|[<span data-ttu-id="dde3f-470">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="dde3f-470">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="dde3f-471">Чтение</span><span class="sxs-lookup"><span data-stu-id="dde3f-471">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="dde3f-472">Пример</span><span class="sxs-lookup"><span data-stu-id="dde3f-472">Example</span></span>

```js
Office.context.mailbox.displayNewMessageForm(
  {
    // Copy the To line from current item.
    toRecipients: Office.context.mailbox.item.to
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

#### <a name="getcallbacktokenasyncoptions-callback"></a><span data-ttu-id="dde3f-473">getCallbackTokenAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="dde3f-473">getCallbackTokenAsync([options], callback)</span></span>

<span data-ttu-id="dde3f-474">Возвращает строку, содержащую маркер, который используется для вызова интерфейсов REST API или веб-служб Exchange.</span><span class="sxs-lookup"><span data-stu-id="dde3f-474">Gets a string that contains a token used to call REST APIs or Exchange Web Services.</span></span>

<span data-ttu-id="dde3f-p131">Метод `getCallbackTokenAsync` совершает асинхронный вызов, чтобы получить непрозрачный маркер с сервера Exchange Server, на котором размещен почтовый ящик пользователя. Время существования маркера обратного вызова составляет 5 минут.</span><span class="sxs-lookup"><span data-stu-id="dde3f-p131">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

> [!NOTE]
> <span data-ttu-id="dde3f-477">Рекомендуем сделать так, чтобы по мере возможности надстройки использовали интерфейсы REST API, а не веб-службы Exchange.</span><span class="sxs-lookup"><span data-stu-id="dde3f-477">It is recommended that add-ins use the REST APIs instead of Exchange Web Services whenever possible.</span></span>

<span data-ttu-id="dde3f-478">Для вызова `getCallbackTokenAsync` метода в режиме чтения требуется минимальный уровень разрешений **ReadItem**.</span><span class="sxs-lookup"><span data-stu-id="dde3f-478">Calling the `getCallbackTokenAsync` method in read mode requires a minimum permission level of **ReadItem**.</span></span>

<span data-ttu-id="dde3f-479">Для `getCallbackTokenAsync` вызова в режиме создания необходимо сохранить элемент.</span><span class="sxs-lookup"><span data-stu-id="dde3f-479">Calling `getCallbackTokenAsync` in compose mode requires you to have saved the item.</span></span> <span data-ttu-id="dde3f-480">[`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) Метод требует наличия минимального уровня разрешений **ReadWriteItem**.</span><span class="sxs-lookup"><span data-stu-id="dde3f-480">The [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method requires a minimum permission level of **ReadWriteItem**.</span></span>

<span data-ttu-id="dde3f-481">**Маркеры REST**</span><span class="sxs-lookup"><span data-stu-id="dde3f-481">**REST Tokens**</span></span>

<span data-ttu-id="dde3f-p133">Если запрашивается маркер REST (`options.isRest = true`), полученный маркер не подойдет для проверки подлинности при вызовах веб-служб Exchange. Область действия маркера будет ограничена доступом только для чтения к текущему элементу и его вложениям, если в манифесте надстройки не указано разрешение [`ReadWriteMailbox`](/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission). Если указано разрешение `ReadWriteMailbox`, полученный маркер предоставит доступ на чтение и запись к почте, календарю и контактам, включая возможность отправки почты.</span><span class="sxs-lookup"><span data-stu-id="dde3f-p133">When a REST token is requested (`options.isRest = true`), the resulting token will not work to authenticate Exchange Web Services calls. The token will be limited in scope to read-only access to the current item and its attachments, unless the add-in has specified the [`ReadWriteMailbox`](/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) permission in its manifest. If the `ReadWriteMailbox` permission is specified, the resulting token will grant read/write access to mail, calendar, and contacts, including the ability to send mail.</span></span>

<span data-ttu-id="dde3f-485">С помощью свойства `restUrl` надстройка должна определить правильный URL-адрес для вызовов REST API.</span><span class="sxs-lookup"><span data-stu-id="dde3f-485">The add-in should use the `restUrl` property to determine the correct URL to use when making REST API calls.</span></span>

<span data-ttu-id="dde3f-486">**Маркеры EWS**</span><span class="sxs-lookup"><span data-stu-id="dde3f-486">**EWS Tokens**</span></span>

<span data-ttu-id="dde3f-p134">Если запрашивается маркер EWS (`options.isRest = false`), полученный маркер не подойдет для проверки подлинности при вызовах REST API. Область действия маркера будет ограничена доступом к текущему элементу.</span><span class="sxs-lookup"><span data-stu-id="dde3f-p134">When an EWS token is requested (`options.isRest = false`), the resulting token will not work to authenticate REST API calls. The token will be limited in scope to accessing the current item.</span></span>

<span data-ttu-id="dde3f-489">С помощью свойства `ewsUrl` надстройка должна определить правильный URL-адрес для вызовов EWS.</span><span class="sxs-lookup"><span data-stu-id="dde3f-489">The add-in should use the `ewsUrl` property to determine the correct URL to use when making EWS calls.</span></span>

<span data-ttu-id="dde3f-490">Можно передать как маркер, так и идентификатор вложения или идентификатор элемента в систему стороннего производителя.</span><span class="sxs-lookup"><span data-stu-id="dde3f-490">You can pass both the token and either an attachment identifier or item identifier to a third-party system.</span></span> <span data-ttu-id="dde3f-491">Третья система использует маркер в качестве маркера авторизации носителя, чтобы вызвать операцию [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) [или GetItem](/exchange/client-developer/web-service-reference/getitem-operation) веб-служб Exchange (EWS) для возврата вложения или элемента.</span><span class="sxs-lookup"><span data-stu-id="dde3f-491">The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) operation or [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item.</span></span> <span data-ttu-id="dde3f-492">Например, вы можете создать удаленную службу, чтобы [получить вложения из выбранного элемента](/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="dde3f-492">For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

##### <a name="parameters"></a><span data-ttu-id="dde3f-493">Параметры</span><span class="sxs-lookup"><span data-stu-id="dde3f-493">Parameters</span></span>

|<span data-ttu-id="dde3f-494">Имя</span><span class="sxs-lookup"><span data-stu-id="dde3f-494">Name</span></span>| <span data-ttu-id="dde3f-495">Тип</span><span class="sxs-lookup"><span data-stu-id="dde3f-495">Type</span></span>| <span data-ttu-id="dde3f-496">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="dde3f-496">Attributes</span></span>| <span data-ttu-id="dde3f-497">Описание</span><span class="sxs-lookup"><span data-stu-id="dde3f-497">Description</span></span>|
|---|---|---|---|
| `options` | <span data-ttu-id="dde3f-498">Object</span><span class="sxs-lookup"><span data-stu-id="dde3f-498">Object</span></span> | <span data-ttu-id="dde3f-499">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="dde3f-499">&lt;optional&gt;</span></span> | <span data-ttu-id="dde3f-500">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="dde3f-500">An object literal that contains one or more of the following properties.</span></span> |
| `options.isRest` | <span data-ttu-id="dde3f-501">Boolean</span><span class="sxs-lookup"><span data-stu-id="dde3f-501">Boolean</span></span> |  <span data-ttu-id="dde3f-502">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="dde3f-502">&lt;optional&gt;</span></span> | <span data-ttu-id="dde3f-p136">Определяет, будет ли предоставленный маркер использоваться для интерфейсов REST API Outlook или веб-служб Exchange. Значение по умолчанию: `false`.</span><span class="sxs-lookup"><span data-stu-id="dde3f-p136">Determines if the token provided will be used for the Outlook REST APIs or Exchange Web Services. Default value is `false`.</span></span> |
| `options.asyncContext` | <span data-ttu-id="dde3f-505">Object</span><span class="sxs-lookup"><span data-stu-id="dde3f-505">Object</span></span> |  <span data-ttu-id="dde3f-506">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="dde3f-506">&lt;optional&gt;</span></span> | <span data-ttu-id="dde3f-507">Данные о состоянии, передаваемые в асинхронный метод.</span><span class="sxs-lookup"><span data-stu-id="dde3f-507">Any state data that is passed to the asynchronous method.</span></span> |
|`callback`| <span data-ttu-id="dde3f-508">функция</span><span class="sxs-lookup"><span data-stu-id="dde3f-508">function</span></span>||<span data-ttu-id="dde3f-509">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="dde3f-509">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="dde3f-510">Маркер указывается в виде строки в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="dde3f-510">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="dde3f-511">При наличии ошибки свойства `asyncResult.error` и `asyncResult.diagnostics` могут предоставлять дополнительные сведения.</span><span class="sxs-lookup"><span data-stu-id="dde3f-511">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|

##### <a name="errors"></a><span data-ttu-id="dde3f-512">Ошибки</span><span class="sxs-lookup"><span data-stu-id="dde3f-512">Errors</span></span>

|<span data-ttu-id="dde3f-513">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="dde3f-513">Error code</span></span>|<span data-ttu-id="dde3f-514">Описание</span><span class="sxs-lookup"><span data-stu-id="dde3f-514">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="dde3f-515">Не удалось выполнить запрос.</span><span class="sxs-lookup"><span data-stu-id="dde3f-515">The request has failed.</span></span> <span data-ttu-id="dde3f-516">Просмотрите объект диагностики для кода ошибки HTTP.</span><span class="sxs-lookup"><span data-stu-id="dde3f-516">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="dde3f-517">Сервер Exchange Server вернул ошибку.</span><span class="sxs-lookup"><span data-stu-id="dde3f-517">The Exchange server returned an error.</span></span> <span data-ttu-id="dde3f-518">Для получения дополнительных сведений просмотрите объект диагностики.</span><span class="sxs-lookup"><span data-stu-id="dde3f-518">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="dde3f-519">Пользователь отключен от сети.</span><span class="sxs-lookup"><span data-stu-id="dde3f-519">The user is no longer connected to the network.</span></span> <span data-ttu-id="dde3f-520">Проверьте сетевое подключение и повторите попытку.</span><span class="sxs-lookup"><span data-stu-id="dde3f-520">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="dde3f-521">Требования</span><span class="sxs-lookup"><span data-stu-id="dde3f-521">Requirements</span></span>

|<span data-ttu-id="dde3f-522">Требование</span><span class="sxs-lookup"><span data-stu-id="dde3f-522">Requirement</span></span>| <span data-ttu-id="dde3f-523">Значение</span><span class="sxs-lookup"><span data-stu-id="dde3f-523">Value</span></span>|
|---|---|
|[<span data-ttu-id="dde3f-524">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="dde3f-524">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dde3f-525">1.5</span><span class="sxs-lookup"><span data-stu-id="dde3f-525">1.5</span></span> |
|[<span data-ttu-id="dde3f-526">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="dde3f-526">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dde3f-527">ReadItem</span><span class="sxs-lookup"><span data-stu-id="dde3f-527">ReadItem</span></span>|
|[<span data-ttu-id="dde3f-528">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="dde3f-528">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="dde3f-529">Создание и чтение</span><span class="sxs-lookup"><span data-stu-id="dde3f-529">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="dde3f-530">Пример</span><span class="sxs-lookup"><span data-stu-id="dde3f-530">Example</span></span>

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

#### <a name="getcallbacktokenasynccallback-usercontext"></a><span data-ttu-id="dde3f-531">getCallbackTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="dde3f-531">getCallbackTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="dde3f-532">Получает строку, содержащую маркер, используемый для получения вложения или элемента с Exchange Server.</span><span class="sxs-lookup"><span data-stu-id="dde3f-532">Gets a string that contains a token used to get an attachment or item from an Exchange Server.</span></span>

<span data-ttu-id="dde3f-p140">Метод `getCallbackTokenAsync` совершает асинхронный вызов, чтобы получить непрозрачный маркер с сервера Exchange Server, на котором размещен почтовый ящик пользователя. Время существования маркера обратного вызова составляет 5 минут.</span><span class="sxs-lookup"><span data-stu-id="dde3f-p140">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

<span data-ttu-id="dde3f-535">Можно передать как маркер, так и идентификатор вложения или идентификатор элемента в систему стороннего производителя.</span><span class="sxs-lookup"><span data-stu-id="dde3f-535">You can pass both the token and either an attachment identifier or item identifier to a third-party system.</span></span> <span data-ttu-id="dde3f-536">Третья система использует маркер в качестве маркера авторизации носителя, чтобы вызвать операцию [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) [или GetItem](/exchange/client-developer/web-service-reference/getitem-operation) веб-служб Exchange (EWS) для возврата вложения или элемента.</span><span class="sxs-lookup"><span data-stu-id="dde3f-536">The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) operation or [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item.</span></span> <span data-ttu-id="dde3f-537">Например, вы можете создать удаленную службу, чтобы [получить вложения из выбранного элемента](/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="dde3f-537">For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="dde3f-538">Для вызова `getCallbackTokenAsync` метода в режиме чтения требуется минимальный уровень разрешений **ReadItem**.</span><span class="sxs-lookup"><span data-stu-id="dde3f-538">Calling the `getCallbackTokenAsync` method in read mode requires a minimum permission level of **ReadItem**.</span></span>

<span data-ttu-id="dde3f-539">Для `getCallbackTokenAsync` вызова в режиме создания необходимо сохранить элемент.</span><span class="sxs-lookup"><span data-stu-id="dde3f-539">Calling `getCallbackTokenAsync` in compose mode requires you to have saved the item.</span></span> <span data-ttu-id="dde3f-540">[`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) Метод требует наличия минимального уровня разрешений **ReadWriteItem**.</span><span class="sxs-lookup"><span data-stu-id="dde3f-540">The [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method requires a minimum permission level of **ReadWriteItem**.</span></span>

##### <a name="parameters"></a><span data-ttu-id="dde3f-541">Параметры</span><span class="sxs-lookup"><span data-stu-id="dde3f-541">Parameters</span></span>

|<span data-ttu-id="dde3f-542">Имя</span><span class="sxs-lookup"><span data-stu-id="dde3f-542">Name</span></span>| <span data-ttu-id="dde3f-543">Тип</span><span class="sxs-lookup"><span data-stu-id="dde3f-543">Type</span></span>| <span data-ttu-id="dde3f-544">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="dde3f-544">Attributes</span></span>| <span data-ttu-id="dde3f-545">Описание</span><span class="sxs-lookup"><span data-stu-id="dde3f-545">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="dde3f-546">функция</span><span class="sxs-lookup"><span data-stu-id="dde3f-546">function</span></span>||<span data-ttu-id="dde3f-547">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="dde3f-547">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="dde3f-548">Маркер указывается в виде строки в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="dde3f-548">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="dde3f-549">При наличии ошибки свойства `asyncResult.error` и `asyncResult.diagnostics` могут предоставлять дополнительные сведения.</span><span class="sxs-lookup"><span data-stu-id="dde3f-549">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|
|`userContext`| <span data-ttu-id="dde3f-550">Объект</span><span class="sxs-lookup"><span data-stu-id="dde3f-550">Object</span></span>| <span data-ttu-id="dde3f-551">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="dde3f-551">&lt;optional&gt;</span></span>|<span data-ttu-id="dde3f-552">Данные о состоянии, передаваемые в асинхронный метод.</span><span class="sxs-lookup"><span data-stu-id="dde3f-552">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="errors"></a><span data-ttu-id="dde3f-553">Ошибки</span><span class="sxs-lookup"><span data-stu-id="dde3f-553">Errors</span></span>

|<span data-ttu-id="dde3f-554">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="dde3f-554">Error code</span></span>|<span data-ttu-id="dde3f-555">Описание</span><span class="sxs-lookup"><span data-stu-id="dde3f-555">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="dde3f-556">Не удалось выполнить запрос.</span><span class="sxs-lookup"><span data-stu-id="dde3f-556">The request has failed.</span></span> <span data-ttu-id="dde3f-557">Просмотрите объект диагностики для кода ошибки HTTP.</span><span class="sxs-lookup"><span data-stu-id="dde3f-557">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="dde3f-558">Сервер Exchange Server вернул ошибку.</span><span class="sxs-lookup"><span data-stu-id="dde3f-558">The Exchange server returned an error.</span></span> <span data-ttu-id="dde3f-559">Для получения дополнительных сведений просмотрите объект диагностики.</span><span class="sxs-lookup"><span data-stu-id="dde3f-559">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="dde3f-560">Пользователь отключен от сети.</span><span class="sxs-lookup"><span data-stu-id="dde3f-560">The user is no longer connected to the network.</span></span> <span data-ttu-id="dde3f-561">Проверьте сетевое подключение и повторите попытку.</span><span class="sxs-lookup"><span data-stu-id="dde3f-561">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="dde3f-562">Требования</span><span class="sxs-lookup"><span data-stu-id="dde3f-562">Requirements</span></span>

|<span data-ttu-id="dde3f-563">Требование</span><span class="sxs-lookup"><span data-stu-id="dde3f-563">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="dde3f-564">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="dde3f-564">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dde3f-565">1.0</span><span class="sxs-lookup"><span data-stu-id="dde3f-565">1.0</span></span> | <span data-ttu-id="dde3f-566">1.3</span><span class="sxs-lookup"><span data-stu-id="dde3f-566">1.3</span></span> |
|[<span data-ttu-id="dde3f-567">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="dde3f-567">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dde3f-568">ReadItem</span><span class="sxs-lookup"><span data-stu-id="dde3f-568">ReadItem</span></span> | <span data-ttu-id="dde3f-569">ReadItem</span><span class="sxs-lookup"><span data-stu-id="dde3f-569">ReadItem</span></span> |
|[<span data-ttu-id="dde3f-570">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="dde3f-570">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="dde3f-571">Чтение</span><span class="sxs-lookup"><span data-stu-id="dde3f-571">Read</span></span> | <span data-ttu-id="dde3f-572">Создание</span><span class="sxs-lookup"><span data-stu-id="dde3f-572">Compose</span></span> |

##### <a name="example"></a><span data-ttu-id="dde3f-573">Пример</span><span class="sxs-lookup"><span data-stu-id="dde3f-573">Example</span></span>

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

#### <a name="getuseridentitytokenasynccallback-usercontext"></a><span data-ttu-id="dde3f-574">getUserIdentityTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="dde3f-574">getUserIdentityTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="dde3f-575">Получает маркер, идентифицирующий пользователя и надстройку Office.</span><span class="sxs-lookup"><span data-stu-id="dde3f-575">Gets a token identifying the user and the Office Add-in.</span></span>

<span data-ttu-id="dde3f-576">Метод `getUserIdentityTokenAsync` возвращает маркер, который можно использовать для идентификации, а также [проверки подлинности надстройки и пользователя в сторонней системе](/outlook/add-ins/authentication).</span><span class="sxs-lookup"><span data-stu-id="dde3f-576">The `getUserIdentityTokenAsync` method returns a token that you can use to identify and [authenticate the add-in and user with a third-party system](/outlook/add-ins/authentication).</span></span>

##### <a name="parameters"></a><span data-ttu-id="dde3f-577">Параметры</span><span class="sxs-lookup"><span data-stu-id="dde3f-577">Parameters</span></span>

|<span data-ttu-id="dde3f-578">Имя</span><span class="sxs-lookup"><span data-stu-id="dde3f-578">Name</span></span>| <span data-ttu-id="dde3f-579">Тип</span><span class="sxs-lookup"><span data-stu-id="dde3f-579">Type</span></span>| <span data-ttu-id="dde3f-580">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="dde3f-580">Attributes</span></span>| <span data-ttu-id="dde3f-581">Описание</span><span class="sxs-lookup"><span data-stu-id="dde3f-581">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="dde3f-582">функция</span><span class="sxs-lookup"><span data-stu-id="dde3f-582">function</span></span>||<span data-ttu-id="dde3f-583">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="dde3f-583">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="dde3f-584">Маркер указывается в виде строки в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="dde3f-584">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="dde3f-585">При наличии ошибки свойства `asyncResult.error` и `asyncResult.diagnostics` могут предоставлять дополнительные сведения.</span><span class="sxs-lookup"><span data-stu-id="dde3f-585">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|
|`userContext`| <span data-ttu-id="dde3f-586">Объект</span><span class="sxs-lookup"><span data-stu-id="dde3f-586">Object</span></span>| <span data-ttu-id="dde3f-587">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="dde3f-587">&lt;optional&gt;</span></span>|<span data-ttu-id="dde3f-588">Данные о состоянии, передаваемые в асинхронный метод.</span><span class="sxs-lookup"><span data-stu-id="dde3f-588">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="errors"></a><span data-ttu-id="dde3f-589">Ошибки</span><span class="sxs-lookup"><span data-stu-id="dde3f-589">Errors</span></span>

|<span data-ttu-id="dde3f-590">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="dde3f-590">Error code</span></span>|<span data-ttu-id="dde3f-591">Описание</span><span class="sxs-lookup"><span data-stu-id="dde3f-591">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="dde3f-592">Не удалось выполнить запрос.</span><span class="sxs-lookup"><span data-stu-id="dde3f-592">The request has failed.</span></span> <span data-ttu-id="dde3f-593">Просмотрите объект диагностики для кода ошибки HTTP.</span><span class="sxs-lookup"><span data-stu-id="dde3f-593">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="dde3f-594">Сервер Exchange Server вернул ошибку.</span><span class="sxs-lookup"><span data-stu-id="dde3f-594">The Exchange server returned an error.</span></span> <span data-ttu-id="dde3f-595">Для получения дополнительных сведений просмотрите объект диагностики.</span><span class="sxs-lookup"><span data-stu-id="dde3f-595">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="dde3f-596">Пользователь отключен от сети.</span><span class="sxs-lookup"><span data-stu-id="dde3f-596">The user is no longer connected to the network.</span></span> <span data-ttu-id="dde3f-597">Проверьте сетевое подключение и повторите попытку.</span><span class="sxs-lookup"><span data-stu-id="dde3f-597">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="dde3f-598">Требования</span><span class="sxs-lookup"><span data-stu-id="dde3f-598">Requirements</span></span>

|<span data-ttu-id="dde3f-599">Требование</span><span class="sxs-lookup"><span data-stu-id="dde3f-599">Requirement</span></span>| <span data-ttu-id="dde3f-600">Значение</span><span class="sxs-lookup"><span data-stu-id="dde3f-600">Value</span></span>|
|---|---|
|[<span data-ttu-id="dde3f-601">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="dde3f-601">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dde3f-602">1.0</span><span class="sxs-lookup"><span data-stu-id="dde3f-602">1.0</span></span>|
|[<span data-ttu-id="dde3f-603">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="dde3f-603">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dde3f-604">ReadItem</span><span class="sxs-lookup"><span data-stu-id="dde3f-604">ReadItem</span></span>|
|[<span data-ttu-id="dde3f-605">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="dde3f-605">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="dde3f-606">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="dde3f-606">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="dde3f-607">Пример</span><span class="sxs-lookup"><span data-stu-id="dde3f-607">Example</span></span>

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

#### <a name="makeewsrequestasyncdata-callback-usercontext"></a><span data-ttu-id="dde3f-608">makeEwsRequestAsync(data, callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="dde3f-608">makeEwsRequestAsync(data, callback, [userContext])</span></span>

<span data-ttu-id="dde3f-609">Выполняет асинхронный запрос для веб-служб Exchange (EWS) на сервере Exchange Server, на котором размещен почтовый ящик пользователя.</span><span class="sxs-lookup"><span data-stu-id="dde3f-609">Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="dde3f-610">Этот метод не поддерживается в следующих сценариях:</span><span class="sxs-lookup"><span data-stu-id="dde3f-610">This method is not supported in the following scenarios.</span></span>
> - <span data-ttu-id="dde3f-611">В Outlook для iOS и Android</span><span class="sxs-lookup"><span data-stu-id="dde3f-611">In Outlook on iOS or Android</span></span>
> - <span data-ttu-id="dde3f-612">Если надстройка загружается в почтовый ящик Gmail.</span><span class="sxs-lookup"><span data-stu-id="dde3f-612">When the add-in is loaded in a Gmail mailbox</span></span>
> 
> <span data-ttu-id="dde3f-613">В таких случаях надстройка должна [использовать REST API](/outlook/add-ins/use-rest-api) для доступа к почтовому ящику пользователя.</span><span class="sxs-lookup"><span data-stu-id="dde3f-613">In these cases, add-ins should [use REST APIs](/outlook/add-ins/use-rest-api) to access the user's mailbox instead.</span></span>

<span data-ttu-id="dde3f-614">Метод `makeEwsRequestAsync` отправляет запрос EWS от имени надстройки в Exchange.</span><span class="sxs-lookup"><span data-stu-id="dde3f-614">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange.</span></span> <span data-ttu-id="dde3f-615">Список поддерживаемых операций EWS см. в статье [Вызов веб-служб из надстройки Outlook](/outlook/add-ins/web-services#ews-operations-that-add-ins-support).</span><span class="sxs-lookup"><span data-stu-id="dde3f-615">See [Call web services from an Outlook add-in](/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span></span>

<span data-ttu-id="dde3f-616">С помощью метода `makeEwsRequestAsync` невозможно запрашивать элементы, связанные с папкой.</span><span class="sxs-lookup"><span data-stu-id="dde3f-616">You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.</span></span>

<span data-ttu-id="dde3f-617">В запросе XML должна быть указана кодировка UTF-8.</span><span class="sxs-lookup"><span data-stu-id="dde3f-617">The XML request must specify UTF-8 encoding.</span></span>

```xml
<?xml version="1.0" encoding="utf-8"?>
```

<span data-ttu-id="dde3f-p150">У вашей надстройки должно быть разрешение **ReadWriteMailbox** для использования метода `makeEwsRequestAsync`. Сведения об использовании разрешения **ReadWriteMailbox** и операций EWS, которые можно вызывать с помощью метода `makeEwsRequestAsync`, см. в статье [Указание разрешений для доступа почтовой надстройки к почтовому ящику пользователя](/outlook/add-ins/understanding-outlook-add-in-permissions).</span><span class="sxs-lookup"><span data-stu-id="dde3f-p150">Your add-in must have the **ReadWriteMailbox** permission to use the `makeEwsRequestAsync` method. For information about using the **ReadWriteMailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method, see [Specify permissions for mail add-in access to the user's mailbox](/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>

> [!NOTE]
> <span data-ttu-id="dde3f-620">Администратор сервера должен установить значение true для параметра `OAuthAuthentication` в каталоге сервера клиентского доступа EWS, чтобы метод `makeEwsRequestAsync` мог выполнять запросы EWS.</span><span class="sxs-lookup"><span data-stu-id="dde3f-620">The server administrator must set `OAuthAuthentication` to true on the Client Access Server EWS directory to enable the `makeEwsRequestAsync` method to make EWS requests.</span></span>

##### <a name="version-differences"></a><span data-ttu-id="dde3f-621">Различия версий</span><span class="sxs-lookup"><span data-stu-id="dde3f-621">Version differences</span></span>

<span data-ttu-id="dde3f-622">Если вы используете метод `makeEwsRequestAsync` в почтовых приложениях, которые выполняются в Outlook версии более ранней, чем 15.0.4535.1004, указывайте кодировку `ISO-8859-1`.</span><span class="sxs-lookup"><span data-stu-id="dde3f-622">When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set the encoding value to `ISO-8859-1`.</span></span>

```xml
<?xml version="1.0" encoding="iso-8859-1"?>
```

<span data-ttu-id="dde3f-p151">Значение кодировки не нужно указывать, если почтовое приложение выполняется в Outlook в Интернете. Чтобы определить, выполняется ли приложение в Outlook или Outlook в Интернете, используйте свойство mailbox.diagnostics.hostName. Используемую версию Outlook можно определить с помощью свойства mailbox.diagnostics.hostVersion.</span><span class="sxs-lookup"><span data-stu-id="dde3f-p151">You do not need to set the encoding value when your mail app is running in Outlook on the web. You can determine whether your mail app is running in Outlook or Outlook on the web by using the mailbox.diagnostics.hostName property. You can determine what version of Outlook is running by using the mailbox.diagnostics.hostVersion property.</span></span>

##### <a name="parameters"></a><span data-ttu-id="dde3f-626">Параметры</span><span class="sxs-lookup"><span data-stu-id="dde3f-626">Parameters</span></span>

|<span data-ttu-id="dde3f-627">Имя</span><span class="sxs-lookup"><span data-stu-id="dde3f-627">Name</span></span>| <span data-ttu-id="dde3f-628">Тип</span><span class="sxs-lookup"><span data-stu-id="dde3f-628">Type</span></span>| <span data-ttu-id="dde3f-629">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="dde3f-629">Attributes</span></span>| <span data-ttu-id="dde3f-630">Описание</span><span class="sxs-lookup"><span data-stu-id="dde3f-630">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="dde3f-631">String</span><span class="sxs-lookup"><span data-stu-id="dde3f-631">String</span></span>||<span data-ttu-id="dde3f-632">Запрос EWS.</span><span class="sxs-lookup"><span data-stu-id="dde3f-632">The EWS request.</span></span>|
|`callback`| <span data-ttu-id="dde3f-633">function</span><span class="sxs-lookup"><span data-stu-id="dde3f-633">function</span></span>||<span data-ttu-id="dde3f-634">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="dde3f-634">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="dde3f-635">Результат XML вызова EWS указывается в виде строки в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="dde3f-635">The XML result of the EWS call is provided as a string in the `asyncResult.value` property.</span></span> <span data-ttu-id="dde3f-636">Если размер результата превышает 1 МБ, возвращается сообщение об ошибке.</span><span class="sxs-lookup"><span data-stu-id="dde3f-636">If the result exceeds 1 MB in size, an error message is returned instead.</span></span>|
|`userContext`| <span data-ttu-id="dde3f-637">Object</span><span class="sxs-lookup"><span data-stu-id="dde3f-637">Object</span></span>| <span data-ttu-id="dde3f-638">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="dde3f-638">&lt;optional&gt;</span></span>|<span data-ttu-id="dde3f-639">Данные о состоянии, передаваемые в асинхронный метод.</span><span class="sxs-lookup"><span data-stu-id="dde3f-639">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="dde3f-640">Требования</span><span class="sxs-lookup"><span data-stu-id="dde3f-640">Requirements</span></span>

|<span data-ttu-id="dde3f-641">Требование</span><span class="sxs-lookup"><span data-stu-id="dde3f-641">Requirement</span></span>| <span data-ttu-id="dde3f-642">Значение</span><span class="sxs-lookup"><span data-stu-id="dde3f-642">Value</span></span>|
|---|---|
|[<span data-ttu-id="dde3f-643">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="dde3f-643">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dde3f-644">1.0</span><span class="sxs-lookup"><span data-stu-id="dde3f-644">1.0</span></span>|
|[<span data-ttu-id="dde3f-645">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="dde3f-645">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dde3f-646">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="dde3f-646">ReadWriteMailbox</span></span>|
|[<span data-ttu-id="dde3f-647">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="dde3f-647">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="dde3f-648">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="dde3f-648">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="dde3f-649">Пример</span><span class="sxs-lookup"><span data-stu-id="dde3f-649">Example</span></span>

<span data-ttu-id="dde3f-650">В приведенном ниже примере вызывается `makeEwsRequestAsync` для получения темы элемента с помощью операции `GetItem`.</span><span class="sxs-lookup"><span data-stu-id="dde3f-650">The following example calls `makeEwsRequestAsync` to use the `GetItem` operation to get the subject of an item.</span></span>

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

#### <a name="removehandlerasynceventtype-options-callback"></a><span data-ttu-id="dde3f-651">removeHandlerAsync(eventType, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="dde3f-651">removeHandlerAsync(eventType, [options], [callback])</span></span>

<span data-ttu-id="dde3f-652">Удаляет обработчиков для поддерживаемого типа события.</span><span class="sxs-lookup"><span data-stu-id="dde3f-652">Removes the event handlers for a supported event type.</span></span>

<span data-ttu-id="dde3f-653">В настоящее время единственный поддерживаемый тип события — `Office.EventType.ItemChanged`.</span><span class="sxs-lookup"><span data-stu-id="dde3f-653">Currently, the only supported event type is `Office.EventType.ItemChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="dde3f-654">Параметры</span><span class="sxs-lookup"><span data-stu-id="dde3f-654">Parameters</span></span>

| <span data-ttu-id="dde3f-655">Имя</span><span class="sxs-lookup"><span data-stu-id="dde3f-655">Name</span></span> | <span data-ttu-id="dde3f-656">Тип</span><span class="sxs-lookup"><span data-stu-id="dde3f-656">Type</span></span> | <span data-ttu-id="dde3f-657">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="dde3f-657">Attributes</span></span> | <span data-ttu-id="dde3f-658">Описание</span><span class="sxs-lookup"><span data-stu-id="dde3f-658">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="dde3f-659">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="dde3f-659">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="dde3f-660">Событие, которое должно отменить обработчик.</span><span class="sxs-lookup"><span data-stu-id="dde3f-660">The event that should revoke the handler.</span></span> |
| `options` | <span data-ttu-id="dde3f-661">Объект</span><span class="sxs-lookup"><span data-stu-id="dde3f-661">Object</span></span> | <span data-ttu-id="dde3f-662">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="dde3f-662">&lt;optional&gt;</span></span> | <span data-ttu-id="dde3f-663">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="dde3f-663">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="dde3f-664">Object</span><span class="sxs-lookup"><span data-stu-id="dde3f-664">Object</span></span> | <span data-ttu-id="dde3f-665">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="dde3f-665">&lt;optional&gt;</span></span> | <span data-ttu-id="dde3f-666">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="dde3f-666">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="dde3f-667">функция</span><span class="sxs-lookup"><span data-stu-id="dde3f-667">function</span></span>| <span data-ttu-id="dde3f-668">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="dde3f-668">&lt;optional&gt;</span></span>|<span data-ttu-id="dde3f-669">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="dde3f-669">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="dde3f-670">Требования</span><span class="sxs-lookup"><span data-stu-id="dde3f-670">Requirements</span></span>

|<span data-ttu-id="dde3f-671">Требование</span><span class="sxs-lookup"><span data-stu-id="dde3f-671">Requirement</span></span>| <span data-ttu-id="dde3f-672">Значение</span><span class="sxs-lookup"><span data-stu-id="dde3f-672">Value</span></span>|
|---|---|
|[<span data-ttu-id="dde3f-673">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="dde3f-673">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dde3f-674">1.5</span><span class="sxs-lookup"><span data-stu-id="dde3f-674">1.5</span></span> |
|[<span data-ttu-id="dde3f-675">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="dde3f-675">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dde3f-676">ReadItem</span><span class="sxs-lookup"><span data-stu-id="dde3f-676">ReadItem</span></span> |
|[<span data-ttu-id="dde3f-677">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="dde3f-677">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="dde3f-678">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="dde3f-678">Compose or Read</span></span>|
