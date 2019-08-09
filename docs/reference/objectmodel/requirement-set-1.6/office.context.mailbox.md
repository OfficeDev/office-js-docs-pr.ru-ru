---
title: Office. Context. Mailbox — набор обязательных элементов 1,6
description: ''
ms.date: 08/06/2019
localization_priority: Normal
ms.openlocfilehash: f394c23cf9e35e3798de1fe7559bc8083478cc6b
ms.sourcegitcommit: 654ac1a0c477413662b48cffc0faee5cb65fc25f
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/09/2019
ms.locfileid: "36268364"
---
# <a name="mailbox"></a><span data-ttu-id="3f914-102">mailbox</span><span class="sxs-lookup"><span data-stu-id="3f914-102">mailbox</span></span>

### <a name="officeofficemdcontextofficecontextmdmailbox"></a><span data-ttu-id="3f914-103">[Office](Office.md)[.context](Office.context.md).mailbox</span><span class="sxs-lookup"><span data-stu-id="3f914-103">[Office](Office.md)[.context](Office.context.md).mailbox</span></span>

<span data-ttu-id="3f914-104">Предоставляет доступ к объектной модели надстройки Outlook для Microsoft Outlook.</span><span class="sxs-lookup"><span data-stu-id="3f914-104">Provides access to the Outlook add-in object model for Microsoft Outlook.</span></span>

##### <a name="requirements"></a><span data-ttu-id="3f914-105">Требования</span><span class="sxs-lookup"><span data-stu-id="3f914-105">Requirements</span></span>

|<span data-ttu-id="3f914-106">Требование</span><span class="sxs-lookup"><span data-stu-id="3f914-106">Requirement</span></span>| <span data-ttu-id="3f914-107">Значение</span><span class="sxs-lookup"><span data-stu-id="3f914-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="3f914-108">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="3f914-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3f914-109">1.0</span><span class="sxs-lookup"><span data-stu-id="3f914-109">1.0</span></span>|
|[<span data-ttu-id="3f914-110">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="3f914-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3f914-111">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="3f914-111">Restricted</span></span>|
|[<span data-ttu-id="3f914-112">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="3f914-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3f914-113">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="3f914-113">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="3f914-114">Элементы и методы</span><span class="sxs-lookup"><span data-stu-id="3f914-114">Members and methods</span></span>

| <span data-ttu-id="3f914-115">Элемент</span><span class="sxs-lookup"><span data-stu-id="3f914-115">Member</span></span> | <span data-ttu-id="3f914-116">Тип</span><span class="sxs-lookup"><span data-stu-id="3f914-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="3f914-117">ewsUrl</span><span class="sxs-lookup"><span data-stu-id="3f914-117">ewsUrl</span></span>](#ewsurl-string) | <span data-ttu-id="3f914-118">Элемент</span><span class="sxs-lookup"><span data-stu-id="3f914-118">Member</span></span> |
| [<span data-ttu-id="3f914-119">restUrl</span><span class="sxs-lookup"><span data-stu-id="3f914-119">restUrl</span></span>](#resturl-string) | <span data-ttu-id="3f914-120">Элемент</span><span class="sxs-lookup"><span data-stu-id="3f914-120">Member</span></span> |
| [<span data-ttu-id="3f914-121">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="3f914-121">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="3f914-122">Метод</span><span class="sxs-lookup"><span data-stu-id="3f914-122">Method</span></span> |
| [<span data-ttu-id="3f914-123">convertToEwsId</span><span class="sxs-lookup"><span data-stu-id="3f914-123">convertToEwsId</span></span>](#converttoewsiditemid-restversion--string) | <span data-ttu-id="3f914-124">Метод</span><span class="sxs-lookup"><span data-stu-id="3f914-124">Method</span></span> |
| [<span data-ttu-id="3f914-125">convertToLocalClientTime</span><span class="sxs-lookup"><span data-stu-id="3f914-125">convertToLocalClientTime</span></span>](#converttolocalclienttimetimevalue--localclienttime) | <span data-ttu-id="3f914-126">Метод</span><span class="sxs-lookup"><span data-stu-id="3f914-126">Method</span></span> |
| [<span data-ttu-id="3f914-127">convertToRestId</span><span class="sxs-lookup"><span data-stu-id="3f914-127">convertToRestId</span></span>](#converttorestiditemid-restversion--string) | <span data-ttu-id="3f914-128">Метод</span><span class="sxs-lookup"><span data-stu-id="3f914-128">Method</span></span> |
| [<span data-ttu-id="3f914-129">convertToUtcClientTime</span><span class="sxs-lookup"><span data-stu-id="3f914-129">convertToUtcClientTime</span></span>](#converttoutcclienttimeinput--date) | <span data-ttu-id="3f914-130">Метод</span><span class="sxs-lookup"><span data-stu-id="3f914-130">Method</span></span> |
| [<span data-ttu-id="3f914-131">displayAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="3f914-131">displayAppointmentForm</span></span>](#displayappointmentformitemid) | <span data-ttu-id="3f914-132">Метод</span><span class="sxs-lookup"><span data-stu-id="3f914-132">Method</span></span> |
| [<span data-ttu-id="3f914-133">displayMessageForm</span><span class="sxs-lookup"><span data-stu-id="3f914-133">displayMessageForm</span></span>](#displaymessageformitemid) | <span data-ttu-id="3f914-134">Метод</span><span class="sxs-lookup"><span data-stu-id="3f914-134">Method</span></span> |
| [<span data-ttu-id="3f914-135">displayNewAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="3f914-135">displayNewAppointmentForm</span></span>](#displaynewappointmentformparameters) | <span data-ttu-id="3f914-136">Метод</span><span class="sxs-lookup"><span data-stu-id="3f914-136">Method</span></span> |
| [<span data-ttu-id="3f914-137">дисплайневмессажеформ</span><span class="sxs-lookup"><span data-stu-id="3f914-137">displayNewMessageForm</span></span>](#displaynewmessageformparameters) | <span data-ttu-id="3f914-138">Метод</span><span class="sxs-lookup"><span data-stu-id="3f914-138">Method</span></span> |
| [<span data-ttu-id="3f914-139">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="3f914-139">getCallbackTokenAsync</span></span>](#getcallbacktokenasyncoptions-callback) | <span data-ttu-id="3f914-140">Метод</span><span class="sxs-lookup"><span data-stu-id="3f914-140">Method</span></span> |
| [<span data-ttu-id="3f914-141">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="3f914-141">getCallbackTokenAsync</span></span>](#getcallbacktokenasynccallback-usercontext) | <span data-ttu-id="3f914-142">Метод</span><span class="sxs-lookup"><span data-stu-id="3f914-142">Method</span></span> |
| [<span data-ttu-id="3f914-143">getUserIdentityTokenAsync</span><span class="sxs-lookup"><span data-stu-id="3f914-143">getUserIdentityTokenAsync</span></span>](#getuseridentitytokenasynccallback-usercontext) | <span data-ttu-id="3f914-144">Метод</span><span class="sxs-lookup"><span data-stu-id="3f914-144">Method</span></span> |
| [<span data-ttu-id="3f914-145">makeEwsRequestAsync</span><span class="sxs-lookup"><span data-stu-id="3f914-145">makeEwsRequestAsync</span></span>](#makeewsrequestasyncdata-callback-usercontext) | <span data-ttu-id="3f914-146">Метод</span><span class="sxs-lookup"><span data-stu-id="3f914-146">Method</span></span> |
| [<span data-ttu-id="3f914-147">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="3f914-147">removeHandlerAsync</span></span>](#removehandlerasynceventtype-options-callback) | <span data-ttu-id="3f914-148">Метод</span><span class="sxs-lookup"><span data-stu-id="3f914-148">Method</span></span> |

### <a name="namespaces"></a><span data-ttu-id="3f914-149">Пространства имен</span><span class="sxs-lookup"><span data-stu-id="3f914-149">Namespaces</span></span>

<span data-ttu-id="3f914-150">[diagnostics](Office.context.mailbox.diagnostics.md). Предоставляет надстройке Outlook диагностические сведения.</span><span class="sxs-lookup"><span data-stu-id="3f914-150">[diagnostics](Office.context.mailbox.diagnostics.md): Provides diagnostic information to an Outlook add-in.</span></span>

<span data-ttu-id="3f914-151">[item](Office.context.mailbox.item.md). Предоставляет методы и свойства для доступа к сообщению или встрече в надстройке Outlook.</span><span class="sxs-lookup"><span data-stu-id="3f914-151">[item](Office.context.mailbox.item.md): Provides methods and properties for accessing a message or appointment in an Outlook add-in.</span></span>

<span data-ttu-id="3f914-152">[userProfile](Office.context.mailbox.userProfile.md). Предоставляет сведения о пользователе в надстройке Outlook.</span><span class="sxs-lookup"><span data-stu-id="3f914-152">[userProfile](Office.context.mailbox.userProfile.md): Provides information about the user in an Outlook add-in.</span></span>

### <a name="members"></a><span data-ttu-id="3f914-153">Элементы</span><span class="sxs-lookup"><span data-stu-id="3f914-153">Members</span></span>

#### <a name="ewsurl-string"></a><span data-ttu-id="3f914-154">ewsUrl: строка</span><span class="sxs-lookup"><span data-stu-id="3f914-154">ewsUrl: String</span></span>

<span data-ttu-id="3f914-155">Получает URL-адрес конечной точки веб-служб Exchange (EWS) для этой учетной записи электронной почты.</span><span class="sxs-lookup"><span data-stu-id="3f914-155">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account.</span></span> <span data-ttu-id="3f914-156">Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="3f914-156">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="3f914-157">Этот элемент не поддерживается в Outlook на iOS или Android.</span><span class="sxs-lookup"><span data-stu-id="3f914-157">This member is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="3f914-p102">Удаленная служба может использовать значение `ewsUrl`, чтобы выполнять вызовы EWS для почтового ящика пользователя. Например, вы можете создать удаленную службу, чтобы [получить вложения из выбранного элемента](/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="3f914-p102">The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="3f914-160">Чтобы вызвать элемент `ewsUrl` в режиме чтения, в манифесте приложения должно быть указано разрешение **ReadItem**.</span><span class="sxs-lookup"><span data-stu-id="3f914-160">Your app must have the **ReadItem** permission specified in its manifest to call the `ewsUrl` member in read mode.</span></span>

<span data-ttu-id="3f914-p103">Перед использованием элемента `ewsUrl` в режиме создания необходимо вызвать метод [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback). Для вызова метода `saveAsync` приложение должно иметь разрешения **ReadWriteItem**.</span><span class="sxs-lookup"><span data-stu-id="3f914-p103">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `ewsUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="3f914-163">Тип</span><span class="sxs-lookup"><span data-stu-id="3f914-163">Type</span></span>

*   <span data-ttu-id="3f914-164">String</span><span class="sxs-lookup"><span data-stu-id="3f914-164">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="3f914-165">Требования</span><span class="sxs-lookup"><span data-stu-id="3f914-165">Requirements</span></span>

|<span data-ttu-id="3f914-166">Требование</span><span class="sxs-lookup"><span data-stu-id="3f914-166">Requirement</span></span>| <span data-ttu-id="3f914-167">Значение</span><span class="sxs-lookup"><span data-stu-id="3f914-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="3f914-168">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="3f914-168">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3f914-169">1.0</span><span class="sxs-lookup"><span data-stu-id="3f914-169">1.0</span></span>|
|[<span data-ttu-id="3f914-170">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="3f914-170">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3f914-171">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3f914-171">ReadItem</span></span>|
|[<span data-ttu-id="3f914-172">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="3f914-172">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3f914-173">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="3f914-173">Compose or Read</span></span>|

#### <a name="resturl-string"></a><span data-ttu-id="3f914-174">Рестурл: строка</span><span class="sxs-lookup"><span data-stu-id="3f914-174">restUrl: String</span></span>

<span data-ttu-id="3f914-175">Возвращает URL-адрес конечной точки REST для этой учетной записи электронной почты.</span><span class="sxs-lookup"><span data-stu-id="3f914-175">Gets the URL of the REST endpoint for this email account.</span></span>

<span data-ttu-id="3f914-176">С помощью значения `restUrl` можно выполнять вызовы [REST API](/outlook/rest/) для почтового ящика пользователя.</span><span class="sxs-lookup"><span data-stu-id="3f914-176">The `restUrl` value can be used to make [REST API](/outlook/rest/) calls to the user's mailbox.</span></span>

<span data-ttu-id="3f914-177">Чтобы вызвать элемент `restUrl` в режиме чтения, в манифесте приложения необходимо указать разрешение **ReadItem**.</span><span class="sxs-lookup"><span data-stu-id="3f914-177">Your app must have the **ReadItem** permission specified in its manifest to call the `restUrl` member in read mode.</span></span>

<span data-ttu-id="3f914-p104">Перед использованием элемента `restUrl` в режиме создания необходимо вызвать метод [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback). Для вызова метода `saveAsync` приложение должно иметь разрешения **ReadWriteItem**.</span><span class="sxs-lookup"><span data-stu-id="3f914-p104">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `restUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="3f914-180">Тип</span><span class="sxs-lookup"><span data-stu-id="3f914-180">Type</span></span>

*   <span data-ttu-id="3f914-181">String</span><span class="sxs-lookup"><span data-stu-id="3f914-181">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="3f914-182">Требования</span><span class="sxs-lookup"><span data-stu-id="3f914-182">Requirements</span></span>

|<span data-ttu-id="3f914-183">Требование</span><span class="sxs-lookup"><span data-stu-id="3f914-183">Requirement</span></span>| <span data-ttu-id="3f914-184">Значение</span><span class="sxs-lookup"><span data-stu-id="3f914-184">Value</span></span>|
|---|---|
|[<span data-ttu-id="3f914-185">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="3f914-185">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3f914-186">1.5</span><span class="sxs-lookup"><span data-stu-id="3f914-186">1.5</span></span> |
|[<span data-ttu-id="3f914-187">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="3f914-187">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3f914-188">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3f914-188">ReadItem</span></span>|
|[<span data-ttu-id="3f914-189">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="3f914-189">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3f914-190">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="3f914-190">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="3f914-191">Методы</span><span class="sxs-lookup"><span data-stu-id="3f914-191">Methods</span></span>

#### <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="3f914-192">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="3f914-192">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="3f914-193">Добавляет обработчик для поддерживаемого события.</span><span class="sxs-lookup"><span data-stu-id="3f914-193">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="3f914-194">В настоящее время поддерживаются только события типа `Office.EventType.ItemChanged`, которые вызываются, когда пользователь выбирает новый элемент.</span><span class="sxs-lookup"><span data-stu-id="3f914-194">Currently the only supported event type is `Office.EventType.ItemChanged`, which is invoked when the user selects a new item.</span></span> <span data-ttu-id="3f914-195">Это событие используется надстройками, реализующими закрепляемую область задач, и позволяет надстройке обновлять пользовательский интерфейс области задач в соответствии с выбранным в данный момент элементом.</span><span class="sxs-lookup"><span data-stu-id="3f914-195">This event is used by add-ins that implement a pinnable task pane, and allows the add-in to refresh the task pane UI based on the currently selected item.</span></span>

##### <a name="parameters"></a><span data-ttu-id="3f914-196">Параметры</span><span class="sxs-lookup"><span data-stu-id="3f914-196">Parameters</span></span>

| <span data-ttu-id="3f914-197">Имя</span><span class="sxs-lookup"><span data-stu-id="3f914-197">Name</span></span> | <span data-ttu-id="3f914-198">Тип</span><span class="sxs-lookup"><span data-stu-id="3f914-198">Type</span></span> | <span data-ttu-id="3f914-199">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="3f914-199">Attributes</span></span> | <span data-ttu-id="3f914-200">Описание</span><span class="sxs-lookup"><span data-stu-id="3f914-200">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="3f914-201">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="3f914-201">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="3f914-202">Событие, которое должно вызвать обработчик.</span><span class="sxs-lookup"><span data-stu-id="3f914-202">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="3f914-203">Function</span><span class="sxs-lookup"><span data-stu-id="3f914-203">Function</span></span> || <span data-ttu-id="3f914-p106">Функция для обработки события. Функция должна принимать один параметр, представляющий собой объектный литерал. Значение свойства `type` параметра совпадет со значением параметра `eventType`, переданного методу `addHandlerAsync`.</span><span class="sxs-lookup"><span data-stu-id="3f914-p106">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="3f914-207">Объект</span><span class="sxs-lookup"><span data-stu-id="3f914-207">Object</span></span> | <span data-ttu-id="3f914-208">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="3f914-208">&lt;optional&gt;</span></span> | <span data-ttu-id="3f914-209">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="3f914-209">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="3f914-210">Object</span><span class="sxs-lookup"><span data-stu-id="3f914-210">Object</span></span> | <span data-ttu-id="3f914-211">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="3f914-211">&lt;optional&gt;</span></span> | <span data-ttu-id="3f914-212">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="3f914-212">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="3f914-213">функция</span><span class="sxs-lookup"><span data-stu-id="3f914-213">function</span></span>| <span data-ttu-id="3f914-214">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="3f914-214">&lt;optional&gt;</span></span>|<span data-ttu-id="3f914-215">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="3f914-215">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="3f914-216">Требования</span><span class="sxs-lookup"><span data-stu-id="3f914-216">Requirements</span></span>

|<span data-ttu-id="3f914-217">Требование</span><span class="sxs-lookup"><span data-stu-id="3f914-217">Requirement</span></span>| <span data-ttu-id="3f914-218">Значение</span><span class="sxs-lookup"><span data-stu-id="3f914-218">Value</span></span>|
|---|---|
|[<span data-ttu-id="3f914-219">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="3f914-219">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3f914-220">1.5</span><span class="sxs-lookup"><span data-stu-id="3f914-220">1.5</span></span> |
|[<span data-ttu-id="3f914-221">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="3f914-221">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3f914-222">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3f914-222">ReadItem</span></span> |
|[<span data-ttu-id="3f914-223">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="3f914-223">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3f914-224">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="3f914-224">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="3f914-225">Пример</span><span class="sxs-lookup"><span data-stu-id="3f914-225">Example</span></span>

```javascript
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

#### <a name="converttoewsiditemid-restversion--string"></a><span data-ttu-id="3f914-226">convertToEwsId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="3f914-226">convertToEwsId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="3f914-227">Преобразовывает идентификатор элемента из формата REST в формат EWS.</span><span class="sxs-lookup"><span data-stu-id="3f914-227">Converts an item ID formatted for REST into EWS format.</span></span>

> [!NOTE]
> <span data-ttu-id="3f914-228">Этот метод не поддерживается в Outlook на iOS или Android.</span><span class="sxs-lookup"><span data-stu-id="3f914-228">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="3f914-p107">Формат идентификаторов, извлекаемых через API REST (например, [API Почты Outlook](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) или [Microsoft Graph](https://graph.microsoft.io/)), отличается от формата веб-служб Exchange (EWS). Метод `convertToEwsId` преобразовывает идентификатор в формате REST в формат EWS.</span><span class="sxs-lookup"><span data-stu-id="3f914-p107">Item IDs retrieved via a REST API (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)) use a different format than the format used by Exchange Web Services (EWS). The `convertToEwsId` method converts a REST-formatted ID into the proper format for EWS.</span></span>

##### <a name="parameters"></a><span data-ttu-id="3f914-231">Параметры</span><span class="sxs-lookup"><span data-stu-id="3f914-231">Parameters</span></span>

|<span data-ttu-id="3f914-232">Имя</span><span class="sxs-lookup"><span data-stu-id="3f914-232">Name</span></span>| <span data-ttu-id="3f914-233">Тип</span><span class="sxs-lookup"><span data-stu-id="3f914-233">Type</span></span>| <span data-ttu-id="3f914-234">Описание</span><span class="sxs-lookup"><span data-stu-id="3f914-234">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="3f914-235">String</span><span class="sxs-lookup"><span data-stu-id="3f914-235">String</span></span>|<span data-ttu-id="3f914-236">Идентификатор элемента в формате REST API для Outlook</span><span class="sxs-lookup"><span data-stu-id="3f914-236">An item ID formatted for the Outlook REST APIs</span></span>|
|`restVersion`| [<span data-ttu-id="3f914-237">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="3f914-237">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion?view=outlook-js-1.6)|<span data-ttu-id="3f914-238">Значение, определяющее версию REST API для Outlook, которая используется для извлечения идентификатора элемента.</span><span class="sxs-lookup"><span data-stu-id="3f914-238">A value indicating the version of the Outlook REST API used to retrieve the item ID.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="3f914-239">Требования</span><span class="sxs-lookup"><span data-stu-id="3f914-239">Requirements</span></span>

|<span data-ttu-id="3f914-240">Требование</span><span class="sxs-lookup"><span data-stu-id="3f914-240">Requirement</span></span>| <span data-ttu-id="3f914-241">Значение</span><span class="sxs-lookup"><span data-stu-id="3f914-241">Value</span></span>|
|---|---|
|[<span data-ttu-id="3f914-242">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="3f914-242">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3f914-243">1.3</span><span class="sxs-lookup"><span data-stu-id="3f914-243">1.3</span></span>|
|[<span data-ttu-id="3f914-244">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="3f914-244">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3f914-245">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="3f914-245">Restricted</span></span>|
|[<span data-ttu-id="3f914-246">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="3f914-246">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3f914-247">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="3f914-247">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="3f914-248">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="3f914-248">Returns:</span></span>

<span data-ttu-id="3f914-249">Тип: String</span><span class="sxs-lookup"><span data-stu-id="3f914-249">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="3f914-250">Пример</span><span class="sxs-lookup"><span data-stu-id="3f914-250">Example</span></span>

```javascript
// Get an item's ID from a REST API.
var restId = 'AAMkAGVlOTZjNTM3LW...';

// Treat restId as coming from the v2.0 version of the Outlook Mail API.
var ewsId = Office.context.mailbox.convertToEwsId(restId, Office.MailboxEnums.RestVersion.v2_0);
```

#### <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttimeviewoutlook-js-16"></a><span data-ttu-id="3f914-251">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.6)}</span><span class="sxs-lookup"><span data-stu-id="3f914-251">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.6)}</span></span>

<span data-ttu-id="3f914-252">Получает словарь, содержащий сведения о локальном времени клиента.</span><span class="sxs-lookup"><span data-stu-id="3f914-252">Gets a dictionary containing time information in local client time.</span></span>

<span data-ttu-id="3f914-253">Почтовое приложение для Outlook на настольном компьютере или в Интернете может использовать разные часовые пояса для дат и времени.</span><span class="sxs-lookup"><span data-stu-id="3f914-253">A mail app for Outlook on a desktop or on the web can use different time zones for the dates and times.</span></span> <span data-ttu-id="3f914-254">Outlook на рабочем столе использует часовой пояс клиентского компьютера; В Outlook в Интернете используется часовой пояс, установленный в центре администрирования Exchange.</span><span class="sxs-lookup"><span data-stu-id="3f914-254">Outlook on a desktop uses the client computer time zone; Outlook on the web uses the time zone set on the Exchange Admin Center (EAC).</span></span> <span data-ttu-id="3f914-255">Значения даты и времени должны обрабатываться таким образом, чтобы значения, отображаемые в интерфейсе пользователя, всегда согласовывались с часовым поясом, ожидаемым пользователем.</span><span class="sxs-lookup"><span data-stu-id="3f914-255">You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that the user expects.</span></span>

<span data-ttu-id="3f914-256">Если почтовое приложение запущено в Outlook на настольном клиенте `convertToLocalClientTime` , метод возвратит объект Dictionary со значениями, заданными для часового пояса клиентского компьютера.</span><span class="sxs-lookup"><span data-stu-id="3f914-256">If the mail app is running in Outlook on a desktop client, the `convertToLocalClientTime` method will return a dictionary object with the values set to the client computer time zone.</span></span> <span data-ttu-id="3f914-257">Если почтовое приложение запущено в Outlook в Интернете, `convertToLocalClientTime` метод возвратит объект Dictionary со значениями, заданными в часовом поясе, заданном в центре администрирования Exchange.</span><span class="sxs-lookup"><span data-stu-id="3f914-257">If the mail app is running in Outlook on the web, the `convertToLocalClientTime` method will return a dictionary object with the values set to the time zone specified in the EAC.</span></span>

##### <a name="parameters"></a><span data-ttu-id="3f914-258">Параметры</span><span class="sxs-lookup"><span data-stu-id="3f914-258">Parameters</span></span>

|<span data-ttu-id="3f914-259">Имя</span><span class="sxs-lookup"><span data-stu-id="3f914-259">Name</span></span>| <span data-ttu-id="3f914-260">Тип</span><span class="sxs-lookup"><span data-stu-id="3f914-260">Type</span></span>| <span data-ttu-id="3f914-261">Описание</span><span class="sxs-lookup"><span data-stu-id="3f914-261">Description</span></span>|
|---|---|---|
|`timeValue`| <span data-ttu-id="3f914-262">Date</span><span class="sxs-lookup"><span data-stu-id="3f914-262">Date</span></span>|<span data-ttu-id="3f914-263">Объект Date</span><span class="sxs-lookup"><span data-stu-id="3f914-263">A Date object</span></span>|

##### <a name="requirements"></a><span data-ttu-id="3f914-264">Требования</span><span class="sxs-lookup"><span data-stu-id="3f914-264">Requirements</span></span>

|<span data-ttu-id="3f914-265">Требование</span><span class="sxs-lookup"><span data-stu-id="3f914-265">Requirement</span></span>| <span data-ttu-id="3f914-266">Значение</span><span class="sxs-lookup"><span data-stu-id="3f914-266">Value</span></span>|
|---|---|
|[<span data-ttu-id="3f914-267">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="3f914-267">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3f914-268">1.0</span><span class="sxs-lookup"><span data-stu-id="3f914-268">1.0</span></span>|
|[<span data-ttu-id="3f914-269">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="3f914-269">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3f914-270">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3f914-270">ReadItem</span></span>|
|[<span data-ttu-id="3f914-271">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="3f914-271">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3f914-272">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="3f914-272">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="3f914-273">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="3f914-273">Returns:</span></span>

<span data-ttu-id="3f914-274">Тип: [LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="3f914-274">Type: [LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.6)</span></span>

#### <a name="converttorestiditemid-restversion--string"></a><span data-ttu-id="3f914-275">convertToRestId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="3f914-275">convertToRestId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="3f914-276">Преобразовывает идентификатор элемента в формате EWS в формат REST.</span><span class="sxs-lookup"><span data-stu-id="3f914-276">Converts an item ID formatted for EWS into REST format.</span></span>

> [!NOTE]
> <span data-ttu-id="3f914-277">Этот метод не поддерживается в Outlook на iOS или Android.</span><span class="sxs-lookup"><span data-stu-id="3f914-277">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="3f914-p110">Формат идентификаторов, извлекаемых через EWS или свойство `itemId`, отличается от формата API REST (таких как [API Почты Outlook](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) или [Microsoft Graph](https://graph.microsoft.io/)). Метод `convertToRestId` преобразовывает идентификатор в формате EWS в формат REST.</span><span class="sxs-lookup"><span data-stu-id="3f914-p110">Item IDs retrieved via EWS or via the `itemId` property use a different format than the format used by REST APIs (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)). The `convertToRestId` method converts an EWS-formatted ID into the proper format for REST.</span></span>

##### <a name="parameters"></a><span data-ttu-id="3f914-280">Параметры</span><span class="sxs-lookup"><span data-stu-id="3f914-280">Parameters</span></span>

|<span data-ttu-id="3f914-281">Имя</span><span class="sxs-lookup"><span data-stu-id="3f914-281">Name</span></span>| <span data-ttu-id="3f914-282">Тип</span><span class="sxs-lookup"><span data-stu-id="3f914-282">Type</span></span>| <span data-ttu-id="3f914-283">Описание</span><span class="sxs-lookup"><span data-stu-id="3f914-283">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="3f914-284">String</span><span class="sxs-lookup"><span data-stu-id="3f914-284">String</span></span>|<span data-ttu-id="3f914-285">Идентификатор элемента в формате EWS</span><span class="sxs-lookup"><span data-stu-id="3f914-285">An item ID formatted for Exchange Web Services (EWS)</span></span>|
|`restVersion`| [<span data-ttu-id="3f914-286">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="3f914-286">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion?view=outlook-js-1.6)|<span data-ttu-id="3f914-287">Значение, определяющее версию REST API для Outlook, с которой будет использоваться преобразованный идентификатор.</span><span class="sxs-lookup"><span data-stu-id="3f914-287">A value indicating the version of the Outlook REST API that the converted ID will be used with.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="3f914-288">Требования</span><span class="sxs-lookup"><span data-stu-id="3f914-288">Requirements</span></span>

|<span data-ttu-id="3f914-289">Требование</span><span class="sxs-lookup"><span data-stu-id="3f914-289">Requirement</span></span>| <span data-ttu-id="3f914-290">Значение</span><span class="sxs-lookup"><span data-stu-id="3f914-290">Value</span></span>|
|---|---|
|[<span data-ttu-id="3f914-291">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="3f914-291">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3f914-292">1.3</span><span class="sxs-lookup"><span data-stu-id="3f914-292">1.3</span></span>|
|[<span data-ttu-id="3f914-293">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="3f914-293">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3f914-294">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="3f914-294">Restricted</span></span>|
|[<span data-ttu-id="3f914-295">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="3f914-295">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3f914-296">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="3f914-296">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="3f914-297">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="3f914-297">Returns:</span></span>

<span data-ttu-id="3f914-298">Тип: String</span><span class="sxs-lookup"><span data-stu-id="3f914-298">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="3f914-299">Пример</span><span class="sxs-lookup"><span data-stu-id="3f914-299">Example</span></span>

```javascript
// Get the currently selected item's ID.
var ewsId = Office.context.mailbox.item.itemId;

// Convert to a REST ID for the v2.0 version of the Outlook Mail API.
var restId = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
```

#### <a name="converttoutcclienttimeinput--date"></a><span data-ttu-id="3f914-300">convertToUtcClientTime(input) → {Date}</span><span class="sxs-lookup"><span data-stu-id="3f914-300">convertToUtcClientTime(input) → {Date}</span></span>

<span data-ttu-id="3f914-301">Получает объект Date из словаря, содержащего сведения о времени.</span><span class="sxs-lookup"><span data-stu-id="3f914-301">Gets a Date object from a dictionary containing time information.</span></span>

<span data-ttu-id="3f914-302">Метод `convertToUtcClientTime` преобразует словарь, содержащий локальную дату и время, в объект Date с правильными значениями локальной даты и времени.</span><span class="sxs-lookup"><span data-stu-id="3f914-302">The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a Date object with the correct values for the local date and time.</span></span>

##### <a name="parameters"></a><span data-ttu-id="3f914-303">Параметры</span><span class="sxs-lookup"><span data-stu-id="3f914-303">Parameters</span></span>

|<span data-ttu-id="3f914-304">Имя</span><span class="sxs-lookup"><span data-stu-id="3f914-304">Name</span></span>| <span data-ttu-id="3f914-305">Тип</span><span class="sxs-lookup"><span data-stu-id="3f914-305">Type</span></span>| <span data-ttu-id="3f914-306">Описание</span><span class="sxs-lookup"><span data-stu-id="3f914-306">Description</span></span>|
|---|---|---|
|`input`| [<span data-ttu-id="3f914-307">LocalClientTime</span><span class="sxs-lookup"><span data-stu-id="3f914-307">LocalClientTime</span></span>](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.6)|<span data-ttu-id="3f914-308">Значение локального времени для преобразования.</span><span class="sxs-lookup"><span data-stu-id="3f914-308">The local time value to convert.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="3f914-309">Требования</span><span class="sxs-lookup"><span data-stu-id="3f914-309">Requirements</span></span>

|<span data-ttu-id="3f914-310">Требование</span><span class="sxs-lookup"><span data-stu-id="3f914-310">Requirement</span></span>| <span data-ttu-id="3f914-311">Значение</span><span class="sxs-lookup"><span data-stu-id="3f914-311">Value</span></span>|
|---|---|
|[<span data-ttu-id="3f914-312">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="3f914-312">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3f914-313">1.0</span><span class="sxs-lookup"><span data-stu-id="3f914-313">1.0</span></span>|
|[<span data-ttu-id="3f914-314">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="3f914-314">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3f914-315">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3f914-315">ReadItem</span></span>|
|[<span data-ttu-id="3f914-316">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="3f914-316">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3f914-317">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="3f914-317">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="3f914-318">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="3f914-318">Returns:</span></span>

<span data-ttu-id="3f914-319">Объект Date со временем в формате UTC.</span><span class="sxs-lookup"><span data-stu-id="3f914-319">A Date object with the time expressed in UTC.</span></span>

<dl class="param-type"><span data-ttu-id="3f914-320">

<dt>Тип</dt>

</span><span class="sxs-lookup"><span data-stu-id="3f914-320">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="3f914-321">Date</span><span class="sxs-lookup"><span data-stu-id="3f914-321">Date</span></span></dd>

</dl>

#### <a name="displayappointmentformitemid"></a><span data-ttu-id="3f914-322">displayAppointmentForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="3f914-322">displayAppointmentForm(itemId)</span></span>

<span data-ttu-id="3f914-323">Отображает имеющуюся встречу из календаря.</span><span class="sxs-lookup"><span data-stu-id="3f914-323">Displays an existing calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="3f914-324">Этот метод не поддерживается в Outlook на iOS или Android.</span><span class="sxs-lookup"><span data-stu-id="3f914-324">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="3f914-325">Метод `displayAppointmentForm` открывает новое окно на компьютере или диалоговое окно на мобильном устройстве, содержащее сведения календаря о существующей встрече.</span><span class="sxs-lookup"><span data-stu-id="3f914-325">The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="3f914-326">В Outlook на Mac Этот метод можно использовать для отображения одной встречи, которая не является частью повторяющегося ряда, или главной встречи повторяющейся серии, но невозможно отобразить экземпляр ряда.</span><span class="sxs-lookup"><span data-stu-id="3f914-326">In Outlook on Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment of a recurring series, but you cannot display an instance of the series.</span></span> <span data-ttu-id="3f914-327">Это связано с тем, что в Outlook на Mac-адресе невозможно получить доступ к свойствам (включая идентификатор элемента) повторяющихся рядов.</span><span class="sxs-lookup"><span data-stu-id="3f914-327">This is because in Outlook on Mac, you cannot access the properties (including the item ID) of instances of a recurring series.</span></span>

<span data-ttu-id="3f914-328">В Outlook в Интернете этот метод открывает указанную форму, только если текст формы меньше или равен 32 КБ числу символов.</span><span class="sxs-lookup"><span data-stu-id="3f914-328">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32KB number of characters.</span></span>

<span data-ttu-id="3f914-329">Если указанный идентификатор элемента не определяет существующую встречу, на клиентском компьютере или устройстве открывается пустая страница, и сообщение об ошибке не возвращается.</span><span class="sxs-lookup"><span data-stu-id="3f914-329">If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and no error message will be returned.</span></span>

##### <a name="parameters"></a><span data-ttu-id="3f914-330">Параметры</span><span class="sxs-lookup"><span data-stu-id="3f914-330">Parameters</span></span>

|<span data-ttu-id="3f914-331">Имя</span><span class="sxs-lookup"><span data-stu-id="3f914-331">Name</span></span>| <span data-ttu-id="3f914-332">Тип</span><span class="sxs-lookup"><span data-stu-id="3f914-332">Type</span></span>| <span data-ttu-id="3f914-333">Описание</span><span class="sxs-lookup"><span data-stu-id="3f914-333">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="3f914-334">String</span><span class="sxs-lookup"><span data-stu-id="3f914-334">String</span></span>|<span data-ttu-id="3f914-335">Идентификатор веб-служб Exchange для существующей встречи в календаре.</span><span class="sxs-lookup"><span data-stu-id="3f914-335">The Exchange Web Services (EWS) identifier for an existing calendar appointment.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="3f914-336">Требования</span><span class="sxs-lookup"><span data-stu-id="3f914-336">Requirements</span></span>

|<span data-ttu-id="3f914-337">Требование</span><span class="sxs-lookup"><span data-stu-id="3f914-337">Requirement</span></span>| <span data-ttu-id="3f914-338">Значение</span><span class="sxs-lookup"><span data-stu-id="3f914-338">Value</span></span>|
|---|---|
|[<span data-ttu-id="3f914-339">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="3f914-339">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3f914-340">1.0</span><span class="sxs-lookup"><span data-stu-id="3f914-340">1.0</span></span>|
|[<span data-ttu-id="3f914-341">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="3f914-341">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3f914-342">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3f914-342">ReadItem</span></span>|
|[<span data-ttu-id="3f914-343">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="3f914-343">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3f914-344">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="3f914-344">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="3f914-345">Пример</span><span class="sxs-lookup"><span data-stu-id="3f914-345">Example</span></span>

```javascript
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

#### <a name="displaymessageformitemid"></a><span data-ttu-id="3f914-346">displayMessageForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="3f914-346">displayMessageForm(itemId)</span></span>

<span data-ttu-id="3f914-347">Отображает имеющееся сообщение.</span><span class="sxs-lookup"><span data-stu-id="3f914-347">Displays an existing message.</span></span>

> [!NOTE]
> <span data-ttu-id="3f914-348">Этот метод не поддерживается в Outlook на iOS или Android.</span><span class="sxs-lookup"><span data-stu-id="3f914-348">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="3f914-349">Метод `displayMessageForm` открывает новое окно на компьютере или диалоговое окно на мобильном устройстве, содержащее существующее сообщение.</span><span class="sxs-lookup"><span data-stu-id="3f914-349">The `displayMessageForm` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="3f914-350">В Outlook в Интернете этот метод открывает указанную форму только в том случае, если размер текста формы меньше или равен 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="3f914-350">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32 KB number of characters.</span></span>

<span data-ttu-id="3f914-351">Если указанный идентификатор элемента не определяет существующее сообщение, окно на клиентском компьютере не открывается и сообщение об ошибке не возвращается.</span><span class="sxs-lookup"><span data-stu-id="3f914-351">If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and no error message will be returned.</span></span>

<span data-ttu-id="3f914-p112">Не используйте `displayMessageForm` с параметром `itemId`, который представляет собой встречу. Используйте метод `displayAppointmentForm`, чтобы отобразить сведения о существующей встрече, а метод `displayNewAppointmentForm` — для отображения формы создания встречи.</span><span class="sxs-lookup"><span data-stu-id="3f914-p112">Do not use the `displayMessageForm` with an `itemId` that represents an appointment. Use the `displayAppointmentForm` method to display an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.</span></span>

##### <a name="parameters"></a><span data-ttu-id="3f914-354">Параметры</span><span class="sxs-lookup"><span data-stu-id="3f914-354">Parameters</span></span>

|<span data-ttu-id="3f914-355">Имя</span><span class="sxs-lookup"><span data-stu-id="3f914-355">Name</span></span>| <span data-ttu-id="3f914-356">Тип</span><span class="sxs-lookup"><span data-stu-id="3f914-356">Type</span></span>| <span data-ttu-id="3f914-357">Описание</span><span class="sxs-lookup"><span data-stu-id="3f914-357">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="3f914-358">String</span><span class="sxs-lookup"><span data-stu-id="3f914-358">String</span></span>|<span data-ttu-id="3f914-359">Идентификатор веб-служб Exchange для существующего сообщения.</span><span class="sxs-lookup"><span data-stu-id="3f914-359">The Exchange Web Services (EWS) identifier for an existing message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="3f914-360">Требования</span><span class="sxs-lookup"><span data-stu-id="3f914-360">Requirements</span></span>

|<span data-ttu-id="3f914-361">Требование</span><span class="sxs-lookup"><span data-stu-id="3f914-361">Requirement</span></span>| <span data-ttu-id="3f914-362">Значение</span><span class="sxs-lookup"><span data-stu-id="3f914-362">Value</span></span>|
|---|---|
|[<span data-ttu-id="3f914-363">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="3f914-363">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3f914-364">1.0</span><span class="sxs-lookup"><span data-stu-id="3f914-364">1.0</span></span>|
|[<span data-ttu-id="3f914-365">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="3f914-365">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3f914-366">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3f914-366">ReadItem</span></span>|
|[<span data-ttu-id="3f914-367">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="3f914-367">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3f914-368">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="3f914-368">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="3f914-369">Пример</span><span class="sxs-lookup"><span data-stu-id="3f914-369">Example</span></span>

```javascript
Office.context.mailbox.displayMessageForm(messageId);
```

#### <a name="displaynewappointmentformparameters"></a><span data-ttu-id="3f914-370">displayNewAppointmentForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="3f914-370">displayNewAppointmentForm(parameters)</span></span>

<span data-ttu-id="3f914-371">Отображает форму для создания новой встречи в календаре.</span><span class="sxs-lookup"><span data-stu-id="3f914-371">Displays a form for creating a new calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="3f914-372">Этот метод не поддерживается в Outlook на iOS или Android.</span><span class="sxs-lookup"><span data-stu-id="3f914-372">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="3f914-p113">Метод `displayNewAppointmentForm` открывает форму, в которой пользователь может создать встречу или собрание. Если параметры заданы, поля формы встречи автоматически заполняются их содержимым.</span><span class="sxs-lookup"><span data-stu-id="3f914-p113">The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting. If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="3f914-375">В Outlook в Интернете и на мобильных устройствах этот метод всегда отображает форму с полем участников.</span><span class="sxs-lookup"><span data-stu-id="3f914-375">In Outlook on the web and mobile devices, this method always displays a form with an attendees field.</span></span> <span data-ttu-id="3f914-376">Если вы не укажете участников в качестве входных аргументов, метод отображает форму с кнопкой **Сохранить**.</span><span class="sxs-lookup"><span data-stu-id="3f914-376">If you do not specify any attendees as input arguments, the method displays a form with a **Save** button.</span></span> <span data-ttu-id="3f914-377">Если вы укажете участников, форма будет включать участников и кнопку **Отправить**.</span><span class="sxs-lookup"><span data-stu-id="3f914-377">If you have specified attendees, the form would include the attendees and a **Send** button.</span></span>

<span data-ttu-id="3f914-p115">Если вы укажете участников или ресурсы с помощью параметра `requiredAttendees`, `optionalAttendees` или `resources` в клиенте Outlook с расширенными возможностями и Outlook RT, этот метод отобразит форму собрания с кнопкой **Отправить**. Если не указать получателей, этот метод отобразит форму встречи с кнопкой **Сохранить и закрыть**.</span><span class="sxs-lookup"><span data-stu-id="3f914-p115">In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or `resources` parameter, this method displays a meeting form with a **Send** button. If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.</span></span>

<span data-ttu-id="3f914-380">Если параметры превышают указанные ограничения размера или если указано неизвестное имя параметра, вызывается исключение.</span><span class="sxs-lookup"><span data-stu-id="3f914-380">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="3f914-381">Параметры</span><span class="sxs-lookup"><span data-stu-id="3f914-381">Parameters</span></span>

> [!NOTE]
> <span data-ttu-id="3f914-382">Все параметры являются необязательными.</span><span class="sxs-lookup"><span data-stu-id="3f914-382">All parameters are optional.</span></span>

|<span data-ttu-id="3f914-383">Имя</span><span class="sxs-lookup"><span data-stu-id="3f914-383">Name</span></span>| <span data-ttu-id="3f914-384">Тип</span><span class="sxs-lookup"><span data-stu-id="3f914-384">Type</span></span>| <span data-ttu-id="3f914-385">Описание</span><span class="sxs-lookup"><span data-stu-id="3f914-385">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="3f914-386">Object</span><span class="sxs-lookup"><span data-stu-id="3f914-386">Object</span></span> | <span data-ttu-id="3f914-387">Словарь параметров, описывающий новую встречу.</span><span class="sxs-lookup"><span data-stu-id="3f914-387">A dictionary of parameters describing the new appointment.</span></span> |
| `parameters.requiredAttendees` | <span data-ttu-id="3f914-388">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)&gt;</span><span class="sxs-lookup"><span data-stu-id="3f914-388">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)&gt;</span></span> | <span data-ttu-id="3f914-p116">Массив строк, содержащий электронные адреса, или массив, содержащий объекты `EmailAddressDetails` для каждого из обязательных участников встречи. Массив может включать не более 100 записей.</span><span class="sxs-lookup"><span data-stu-id="3f914-p116">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the required attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.optionalAttendees` | <span data-ttu-id="3f914-391">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)&gt;</span><span class="sxs-lookup"><span data-stu-id="3f914-391">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)&gt;</span></span> | <span data-ttu-id="3f914-p117">Массив строк, содержащий электронные адреса, или массив, содержащий объекты `EmailAddressDetails` для каждого из необязательных участников встречи. Массив может включать не более 100 записей.</span><span class="sxs-lookup"><span data-stu-id="3f914-p117">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the optional attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.start` | <span data-ttu-id="3f914-394">Date</span><span class="sxs-lookup"><span data-stu-id="3f914-394">Date</span></span> | <span data-ttu-id="3f914-395">Объект `Date`, указывающий дату и время начала встречи.</span><span class="sxs-lookup"><span data-stu-id="3f914-395">A `Date` object specifying the start date and time of the appointment.</span></span> |
| `parameters.end` | <span data-ttu-id="3f914-396">Date</span><span class="sxs-lookup"><span data-stu-id="3f914-396">Date</span></span> | <span data-ttu-id="3f914-397">Объект `Date`, указывающий дату и время окончания встречи.</span><span class="sxs-lookup"><span data-stu-id="3f914-397">A `Date` object specifying the end date and time of the appointment.</span></span> |
| `parameters.location` | <span data-ttu-id="3f914-398">String</span><span class="sxs-lookup"><span data-stu-id="3f914-398">String</span></span> | <span data-ttu-id="3f914-p118">Строка со сведениями о месте встречи. Максимальное количество символов в строке — 255.</span><span class="sxs-lookup"><span data-stu-id="3f914-p118">A string containing the location of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.resources` | <span data-ttu-id="3f914-401">Array.&lt;String&gt;</span><span class="sxs-lookup"><span data-stu-id="3f914-401">Array.&lt;String&gt;</span></span> | <span data-ttu-id="3f914-p119">Массив строк, содержащий необходимые для встречи ресурсы. Массив может включать не более 100 записей.</span><span class="sxs-lookup"><span data-stu-id="3f914-p119">An array of strings containing the resources required for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="3f914-404">String</span><span class="sxs-lookup"><span data-stu-id="3f914-404">String</span></span> | <span data-ttu-id="3f914-p120">Строка с темой встречи. Максимальное количество символов в строке — 255.</span><span class="sxs-lookup"><span data-stu-id="3f914-p120">A string containing the subject of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.body` | <span data-ttu-id="3f914-407">String</span><span class="sxs-lookup"><span data-stu-id="3f914-407">String</span></span> | <span data-ttu-id="3f914-p121">Текст сообщения о встрече. Максимальный размер содержимого сообщения — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="3f914-p121">The body of the appointment. The body content is limited to a maximum size of 32 KB.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="3f914-410">Требования</span><span class="sxs-lookup"><span data-stu-id="3f914-410">Requirements</span></span>

|<span data-ttu-id="3f914-411">Требование</span><span class="sxs-lookup"><span data-stu-id="3f914-411">Requirement</span></span>| <span data-ttu-id="3f914-412">Значение</span><span class="sxs-lookup"><span data-stu-id="3f914-412">Value</span></span>|
|---|---|
|[<span data-ttu-id="3f914-413">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="3f914-413">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3f914-414">1.0</span><span class="sxs-lookup"><span data-stu-id="3f914-414">1.0</span></span>|
|[<span data-ttu-id="3f914-415">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="3f914-415">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3f914-416">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3f914-416">ReadItem</span></span>|
|[<span data-ttu-id="3f914-417">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="3f914-417">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3f914-418">Чтение</span><span class="sxs-lookup"><span data-stu-id="3f914-418">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="3f914-419">Пример</span><span class="sxs-lookup"><span data-stu-id="3f914-419">Example</span></span>

```javascript
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

#### <a name="displaynewmessageformparameters"></a><span data-ttu-id="3f914-420">Дисплайневмессажеформ (Parameters)</span><span class="sxs-lookup"><span data-stu-id="3f914-420">displayNewMessageForm(parameters)</span></span>

<span data-ttu-id="3f914-421">Отображает форму для создания нового сообщения.</span><span class="sxs-lookup"><span data-stu-id="3f914-421">Displays a form for creating a new message.</span></span>

<span data-ttu-id="3f914-422">`displayNewMessageForm` Метод открывает форму, которая позволяет пользователю создать новое сообщение.</span><span class="sxs-lookup"><span data-stu-id="3f914-422">The `displayNewMessageForm` method opens a form that enables the user to create a new message.</span></span> <span data-ttu-id="3f914-423">Если указаны параметры, поля формы сообщения автоматически заполняются содержимым параметров.</span><span class="sxs-lookup"><span data-stu-id="3f914-423">If parameters are specified, the message form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="3f914-424">Если параметры превышают указанные ограничения размера или если указано неизвестное имя параметра, вызывается исключение.</span><span class="sxs-lookup"><span data-stu-id="3f914-424">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="3f914-425">Параметры</span><span class="sxs-lookup"><span data-stu-id="3f914-425">Parameters</span></span>

> [!NOTE]
> <span data-ttu-id="3f914-426">Все параметры являются необязательными.</span><span class="sxs-lookup"><span data-stu-id="3f914-426">All parameters are optional.</span></span>

|<span data-ttu-id="3f914-427">Имя</span><span class="sxs-lookup"><span data-stu-id="3f914-427">Name</span></span>| <span data-ttu-id="3f914-428">Тип</span><span class="sxs-lookup"><span data-stu-id="3f914-428">Type</span></span>| <span data-ttu-id="3f914-429">Описание</span><span class="sxs-lookup"><span data-stu-id="3f914-429">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="3f914-430">Object</span><span class="sxs-lookup"><span data-stu-id="3f914-430">Object</span></span> | <span data-ttu-id="3f914-431">Словарь параметров, описывающих новое сообщение.</span><span class="sxs-lookup"><span data-stu-id="3f914-431">A dictionary of parameters describing the new message.</span></span> |
| `parameters.toRecipients` | <span data-ttu-id="3f914-432">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)&gt;</span><span class="sxs-lookup"><span data-stu-id="3f914-432">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)&gt;</span></span> | <span data-ttu-id="3f914-433">Массив строк, содержащий адреса электронной почты или массив, содержащий `EmailAddressDetails` объект для каждого из получателей в строке "Кому".</span><span class="sxs-lookup"><span data-stu-id="3f914-433">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the To line.</span></span> <span data-ttu-id="3f914-434">Массив может включать не более 100 записей.</span><span class="sxs-lookup"><span data-stu-id="3f914-434">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.ccRecipients` | <span data-ttu-id="3f914-435">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)&gt;</span><span class="sxs-lookup"><span data-stu-id="3f914-435">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)&gt;</span></span> | <span data-ttu-id="3f914-436">Массив строк, содержащий адреса электронной почты или массив, содержащий `EmailAddressDetails` объект для каждого получателя в строке "копия".</span><span class="sxs-lookup"><span data-stu-id="3f914-436">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Cc line.</span></span> <span data-ttu-id="3f914-437">Массив может включать не более 100 записей.</span><span class="sxs-lookup"><span data-stu-id="3f914-437">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.bccRecipients` | <span data-ttu-id="3f914-438">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)&gt;</span><span class="sxs-lookup"><span data-stu-id="3f914-438">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)&gt;</span></span> | <span data-ttu-id="3f914-439">Массив строк, содержащий адреса электронной почты или массив, содержащий `EmailAddressDetails` объект для каждого из получателей, указанных в строке "СК".</span><span class="sxs-lookup"><span data-stu-id="3f914-439">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Bcc line.</span></span> <span data-ttu-id="3f914-440">Массив может включать не более 100 записей.</span><span class="sxs-lookup"><span data-stu-id="3f914-440">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="3f914-441">String</span><span class="sxs-lookup"><span data-stu-id="3f914-441">String</span></span> | <span data-ttu-id="3f914-442">Строка, содержащая тему сообщения.</span><span class="sxs-lookup"><span data-stu-id="3f914-442">A string containing the subject of the message.</span></span> <span data-ttu-id="3f914-443">Максимальное количество символов в строке — 255.</span><span class="sxs-lookup"><span data-stu-id="3f914-443">The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.htmlBody` | <span data-ttu-id="3f914-444">String</span><span class="sxs-lookup"><span data-stu-id="3f914-444">String</span></span> | <span data-ttu-id="3f914-445">Текст сообщения в формате HTML.</span><span class="sxs-lookup"><span data-stu-id="3f914-445">The HTML body of the message.</span></span> <span data-ttu-id="3f914-446">Максимальный размер содержимого сообщения — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="3f914-446">The body content is limited to a maximum size of 32 KB.</span></span> |
| `parameters.attachments` | <span data-ttu-id="3f914-447">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="3f914-447">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="3f914-448">Массив объектов JSON, представляющих собой вложенные файлы или элементы.</span><span class="sxs-lookup"><span data-stu-id="3f914-448">An array of JSON objects that are either file or item attachments.</span></span> |
| `parameters.attachments.type` | <span data-ttu-id="3f914-449">String</span><span class="sxs-lookup"><span data-stu-id="3f914-449">String</span></span> | <span data-ttu-id="3f914-p128">Указывает тип вложения. Допустимые значения: `file` для вложенного файла и `item` для вложенного элемента.</span><span class="sxs-lookup"><span data-stu-id="3f914-p128">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `parameters.attachments.name` | <span data-ttu-id="3f914-452">Строка</span><span class="sxs-lookup"><span data-stu-id="3f914-452">String</span></span> | <span data-ttu-id="3f914-453">Строка, содержащая имя вложения, длиной до 255 символов.</span><span class="sxs-lookup"><span data-stu-id="3f914-453">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `parameters.attachments.url` | <span data-ttu-id="3f914-454">String</span><span class="sxs-lookup"><span data-stu-id="3f914-454">String</span></span> | <span data-ttu-id="3f914-p129">Используется, только если свойству `type` задано значение `file`. URI расположения файла.</span><span class="sxs-lookup"><span data-stu-id="3f914-p129">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `parameters.attachments.isInline` | <span data-ttu-id="3f914-457">Логический</span><span class="sxs-lookup"><span data-stu-id="3f914-457">Boolean</span></span> | <span data-ttu-id="3f914-p130">Используется, только если свойству `type` задано значение `file`. Значение `true` указывает на то, что вложение будет встроено в текст сообщения и не должно отображаться в списке вложений.</span><span class="sxs-lookup"><span data-stu-id="3f914-p130">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `parameters.attachments.itemId` | <span data-ttu-id="3f914-460">Строка</span><span class="sxs-lookup"><span data-stu-id="3f914-460">String</span></span> | <span data-ttu-id="3f914-461">Используется, только если свойству `type` присвоено значение `item`.</span><span class="sxs-lookup"><span data-stu-id="3f914-461">Only used if `type` is set to `item`.</span></span> <span data-ttu-id="3f914-462">Идентификатор элемента EWS существующего сообщения электронной почты, которое необходимо присоединить к новому сообщению.</span><span class="sxs-lookup"><span data-stu-id="3f914-462">The EWS item id of the existing e-mail you want to attach to the new message.</span></span> <span data-ttu-id="3f914-463">Это строка длиной до 100 символов.</span><span class="sxs-lookup"><span data-stu-id="3f914-463">This is a string up to 100 characters.</span></span> |


##### <a name="requirements"></a><span data-ttu-id="3f914-464">Требования</span><span class="sxs-lookup"><span data-stu-id="3f914-464">Requirements</span></span>

|<span data-ttu-id="3f914-465">Требование</span><span class="sxs-lookup"><span data-stu-id="3f914-465">Requirement</span></span>| <span data-ttu-id="3f914-466">Значение</span><span class="sxs-lookup"><span data-stu-id="3f914-466">Value</span></span>|
|---|---|
|[<span data-ttu-id="3f914-467">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="3f914-467">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3f914-468">1.6</span><span class="sxs-lookup"><span data-stu-id="3f914-468">1.6</span></span> |
|[<span data-ttu-id="3f914-469">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="3f914-469">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3f914-470">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3f914-470">ReadItem</span></span>|
|[<span data-ttu-id="3f914-471">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="3f914-471">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3f914-472">Чтение</span><span class="sxs-lookup"><span data-stu-id="3f914-472">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="3f914-473">Пример</span><span class="sxs-lookup"><span data-stu-id="3f914-473">Example</span></span>

```javascript
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

#### <a name="getcallbacktokenasyncoptions-callback"></a><span data-ttu-id="3f914-474">getCallbackTokenAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="3f914-474">getCallbackTokenAsync([options], callback)</span></span>

<span data-ttu-id="3f914-475">Возвращает строку, содержащую маркер, который используется для вызова интерфейсов REST API или веб-служб Exchange.</span><span class="sxs-lookup"><span data-stu-id="3f914-475">Gets a string that contains a token used to call REST APIs or Exchange Web Services.</span></span>

<span data-ttu-id="3f914-p132">Метод `getCallbackTokenAsync` совершает асинхронный вызов, чтобы получить непрозрачный маркер с сервера Exchange Server, на котором размещен почтовый ящик пользователя. Время существования маркера обратного вызова составляет 5 минут.</span><span class="sxs-lookup"><span data-stu-id="3f914-p132">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

> [!NOTE]
> <span data-ttu-id="3f914-478">Рекомендуем сделать так, чтобы по мере возможности надстройки использовали интерфейсы REST API, а не веб-службы Exchange.</span><span class="sxs-lookup"><span data-stu-id="3f914-478">It is recommended that add-ins use the REST APIs instead of Exchange Web Services whenever possible.</span></span> 

<span data-ttu-id="3f914-479">**Маркеры REST**</span><span class="sxs-lookup"><span data-stu-id="3f914-479">**REST Tokens**</span></span>

<span data-ttu-id="3f914-p133">Если запрашивается маркер REST (`options.isRest = true`), полученный маркер не подойдет для проверки подлинности при вызовах веб-служб Exchange. Область действия маркера будет ограничена доступом только для чтения к текущему элементу и его вложениям, если в манифесте надстройки не указано разрешение [`ReadWriteMailbox`](/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission). Если указано разрешение `ReadWriteMailbox`, полученный маркер предоставит доступ на чтение и запись к почте, календарю и контактам, включая возможность отправки почты.</span><span class="sxs-lookup"><span data-stu-id="3f914-p133">When a REST token is requested (`options.isRest = true`), the resulting token will not work to authenticate Exchange Web Services calls. The token will be limited in scope to read-only access to the current item and its attachments, unless the add-in has specified the [`ReadWriteMailbox`](/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) permission in its manifest. If the `ReadWriteMailbox` permission is specified, the resulting token will grant read/write access to mail, calendar, and contacts, including the ability to send mail.</span></span>

<span data-ttu-id="3f914-483">С помощью свойства `restUrl` надстройка должна определить правильный URL-адрес для вызовов REST API.</span><span class="sxs-lookup"><span data-stu-id="3f914-483">The add-in should use the `restUrl` property to determine the correct URL to use when making REST API calls.</span></span>

<span data-ttu-id="3f914-484">**Маркеры EWS**</span><span class="sxs-lookup"><span data-stu-id="3f914-484">**EWS Tokens**</span></span>

<span data-ttu-id="3f914-p134">Если запрашивается маркер EWS (`options.isRest = false`), полученный маркер не подойдет для проверки подлинности при вызовах REST API. Область действия маркера будет ограничена доступом к текущему элементу.</span><span class="sxs-lookup"><span data-stu-id="3f914-p134">When an EWS token is requested (`options.isRest = false`), the resulting token will not work to authenticate REST API calls. The token will be limited in scope to accessing the current item.</span></span>

<span data-ttu-id="3f914-487">С помощью свойства `ewsUrl` надстройка должна определить правильный URL-адрес для вызовов EWS.</span><span class="sxs-lookup"><span data-stu-id="3f914-487">The add-in should use the `ewsUrl` property to determine the correct URL to use when making EWS calls.</span></span>

##### <a name="parameters"></a><span data-ttu-id="3f914-488">Параметры</span><span class="sxs-lookup"><span data-stu-id="3f914-488">Parameters</span></span>

|<span data-ttu-id="3f914-489">Имя</span><span class="sxs-lookup"><span data-stu-id="3f914-489">Name</span></span>| <span data-ttu-id="3f914-490">Тип</span><span class="sxs-lookup"><span data-stu-id="3f914-490">Type</span></span>| <span data-ttu-id="3f914-491">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="3f914-491">Attributes</span></span>| <span data-ttu-id="3f914-492">Описание</span><span class="sxs-lookup"><span data-stu-id="3f914-492">Description</span></span>|
|---|---|---|---|
| `options` | <span data-ttu-id="3f914-493">Object</span><span class="sxs-lookup"><span data-stu-id="3f914-493">Object</span></span> | <span data-ttu-id="3f914-494">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="3f914-494">&lt;optional&gt;</span></span> | <span data-ttu-id="3f914-495">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="3f914-495">An object literal that contains one or more of the following properties.</span></span> |
| `options.isRest` | <span data-ttu-id="3f914-496">Boolean</span><span class="sxs-lookup"><span data-stu-id="3f914-496">Boolean</span></span> |  <span data-ttu-id="3f914-497">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="3f914-497">&lt;optional&gt;</span></span> | <span data-ttu-id="3f914-p135">Определяет, будет ли предоставленный маркер использоваться для интерфейсов REST API Outlook или веб-служб Exchange. Значение по умолчанию: `false`.</span><span class="sxs-lookup"><span data-stu-id="3f914-p135">Determines if the token provided will be used for the Outlook REST APIs or Exchange Web Services. Default value is `false`.</span></span> |
| `options.asyncContext` | <span data-ttu-id="3f914-500">Объект</span><span class="sxs-lookup"><span data-stu-id="3f914-500">Object</span></span> |  <span data-ttu-id="3f914-501">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="3f914-501">&lt;optional&gt;</span></span> | <span data-ttu-id="3f914-502">Данные о состоянии, передаваемые в асинхронный метод.</span><span class="sxs-lookup"><span data-stu-id="3f914-502">Any state data that is passed to the asynchronous method.</span></span> |
|`callback`| <span data-ttu-id="3f914-503">function</span><span class="sxs-lookup"><span data-stu-id="3f914-503">function</span></span>||<span data-ttu-id="3f914-504">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="3f914-504">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="3f914-505">Маркер указывается в виде строки в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="3f914-505">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="3f914-506">При возникновении ошибки свойства `asyncResult.error` и `asyncResult.diagnostics` могут содержать дополнительные сведения.</span><span class="sxs-lookup"><span data-stu-id="3f914-506">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|

##### <a name="errors"></a><span data-ttu-id="3f914-507">Ошибки</span><span class="sxs-lookup"><span data-stu-id="3f914-507">Errors</span></span>

|<span data-ttu-id="3f914-508">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="3f914-508">Error code</span></span>|<span data-ttu-id="3f914-509">Описание</span><span class="sxs-lookup"><span data-stu-id="3f914-509">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="3f914-510">Запрос не выполнен.</span><span class="sxs-lookup"><span data-stu-id="3f914-510">The request has failed.</span></span> <span data-ttu-id="3f914-511">Просмотрите объект Diagnostics для кода ошибки HTTP.</span><span class="sxs-lookup"><span data-stu-id="3f914-511">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="3f914-512">Сервер Exchange возвратил ошибку.</span><span class="sxs-lookup"><span data-stu-id="3f914-512">The Exchange server returned an error.</span></span> <span data-ttu-id="3f914-513">Дополнительные сведения можно найти в объекте диагностики.</span><span class="sxs-lookup"><span data-stu-id="3f914-513">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="3f914-514">Пользователь больше не подключен к сети.</span><span class="sxs-lookup"><span data-stu-id="3f914-514">The user is no longer connected to the network.</span></span> <span data-ttu-id="3f914-515">Проверьте сетевое подключение и повторите попытку.</span><span class="sxs-lookup"><span data-stu-id="3f914-515">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="3f914-516">Требования</span><span class="sxs-lookup"><span data-stu-id="3f914-516">Requirements</span></span>

|<span data-ttu-id="3f914-517">Требование</span><span class="sxs-lookup"><span data-stu-id="3f914-517">Requirement</span></span>| <span data-ttu-id="3f914-518">Значение</span><span class="sxs-lookup"><span data-stu-id="3f914-518">Value</span></span>|
|---|---|
|[<span data-ttu-id="3f914-519">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="3f914-519">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3f914-520">1.5</span><span class="sxs-lookup"><span data-stu-id="3f914-520">1.5</span></span> |
|[<span data-ttu-id="3f914-521">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="3f914-521">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3f914-522">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3f914-522">ReadItem</span></span>|
|[<span data-ttu-id="3f914-523">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="3f914-523">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3f914-524">Создание и чтение</span><span class="sxs-lookup"><span data-stu-id="3f914-524">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="3f914-525">Пример</span><span class="sxs-lookup"><span data-stu-id="3f914-525">Example</span></span>

```javascript
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

#### <a name="getcallbacktokenasynccallback-usercontext"></a><span data-ttu-id="3f914-526">getCallbackTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="3f914-526">getCallbackTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="3f914-527">Получает строку, содержащую маркер, используемый для получения вложения или элемента с Exchange Server.</span><span class="sxs-lookup"><span data-stu-id="3f914-527">Gets a string that contains a token used to get an attachment or item from an Exchange Server.</span></span>

<span data-ttu-id="3f914-p139">Метод `getCallbackTokenAsync` совершает асинхронный вызов, чтобы получить непрозрачный маркер с сервера Exchange Server, на котором размещен почтовый ящик пользователя. Время существования маркера обратного вызова составляет 5 минут.</span><span class="sxs-lookup"><span data-stu-id="3f914-p139">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

<span data-ttu-id="3f914-p140">Вы можете передать сторонней системе маркер и идентификатор вложения или элемента. Сторонняя система использует этот маркер как маркер авторизации, чтобы вызвать операцию [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) или [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) веб-служб Exchange для возврата вложения или элемента. Например, вы можете создать удаленную службу, чтобы [получить вложения из выбранного элемента](/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="3f914-p140">You can pass the token and an attachment identifier or item identifier to a third-party system. The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) or [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="3f914-533">Для вызова метода `getCallbackTokenAsync` в режиме чтения манифесте приложения должно быть указано разрешение **ReadItem**.</span><span class="sxs-lookup"><span data-stu-id="3f914-533">Your app must have the **ReadItem** permission specified in its manifest to call the `getCallbackTokenAsync` method in read mode.</span></span>

<span data-ttu-id="3f914-p141">Чтобы получить идентификатор элемента для передачи в метод `getCallbackTokenAsync`, в режиме создания необходимо вызвать метод [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback). Для вызова метода `saveAsync` приложение должно иметь разрешения **ReadWriteItem**.</span><span class="sxs-lookup"><span data-stu-id="3f914-p141">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method to get an item identifier to pass to the `getCallbackTokenAsync` method. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="parameters"></a><span data-ttu-id="3f914-536">Параметры</span><span class="sxs-lookup"><span data-stu-id="3f914-536">Parameters</span></span>

|<span data-ttu-id="3f914-537">Имя</span><span class="sxs-lookup"><span data-stu-id="3f914-537">Name</span></span>| <span data-ttu-id="3f914-538">Тип</span><span class="sxs-lookup"><span data-stu-id="3f914-538">Type</span></span>| <span data-ttu-id="3f914-539">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="3f914-539">Attributes</span></span>| <span data-ttu-id="3f914-540">Описание</span><span class="sxs-lookup"><span data-stu-id="3f914-540">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="3f914-541">функция</span><span class="sxs-lookup"><span data-stu-id="3f914-541">function</span></span>||<span data-ttu-id="3f914-542">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="3f914-542">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="3f914-543">Маркер указывается в виде строки в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="3f914-543">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="3f914-544">При возникновении ошибки свойства `asyncResult.error` и `asyncResult.diagnostics` могут содержать дополнительные сведения.</span><span class="sxs-lookup"><span data-stu-id="3f914-544">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|
|`userContext`| <span data-ttu-id="3f914-545">Объект</span><span class="sxs-lookup"><span data-stu-id="3f914-545">Object</span></span>| <span data-ttu-id="3f914-546">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="3f914-546">&lt;optional&gt;</span></span>|<span data-ttu-id="3f914-547">Данные о состоянии, передаваемые в асинхронный метод.</span><span class="sxs-lookup"><span data-stu-id="3f914-547">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="errors"></a><span data-ttu-id="3f914-548">Ошибки</span><span class="sxs-lookup"><span data-stu-id="3f914-548">Errors</span></span>

|<span data-ttu-id="3f914-549">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="3f914-549">Error code</span></span>|<span data-ttu-id="3f914-550">Описание</span><span class="sxs-lookup"><span data-stu-id="3f914-550">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="3f914-551">Запрос не выполнен.</span><span class="sxs-lookup"><span data-stu-id="3f914-551">The request has failed.</span></span> <span data-ttu-id="3f914-552">Просмотрите объект Diagnostics для кода ошибки HTTP.</span><span class="sxs-lookup"><span data-stu-id="3f914-552">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="3f914-553">Сервер Exchange возвратил ошибку.</span><span class="sxs-lookup"><span data-stu-id="3f914-553">The Exchange server returned an error.</span></span> <span data-ttu-id="3f914-554">Дополнительные сведения можно найти в объекте диагностики.</span><span class="sxs-lookup"><span data-stu-id="3f914-554">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="3f914-555">Пользователь больше не подключен к сети.</span><span class="sxs-lookup"><span data-stu-id="3f914-555">The user is no longer connected to the network.</span></span> <span data-ttu-id="3f914-556">Проверьте сетевое подключение и повторите попытку.</span><span class="sxs-lookup"><span data-stu-id="3f914-556">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="3f914-557">Требования</span><span class="sxs-lookup"><span data-stu-id="3f914-557">Requirements</span></span>

|<span data-ttu-id="3f914-558">Требование</span><span class="sxs-lookup"><span data-stu-id="3f914-558">Requirement</span></span>| <span data-ttu-id="3f914-559">Значение</span><span class="sxs-lookup"><span data-stu-id="3f914-559">Value</span></span>|
|---|---|
|[<span data-ttu-id="3f914-560">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="3f914-560">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3f914-561">1.0</span><span class="sxs-lookup"><span data-stu-id="3f914-561">1.0</span></span>|
|[<span data-ttu-id="3f914-562">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="3f914-562">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3f914-563">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3f914-563">ReadItem</span></span>|
|[<span data-ttu-id="3f914-564">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="3f914-564">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3f914-565">Создание и чтение</span><span class="sxs-lookup"><span data-stu-id="3f914-565">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="3f914-566">Пример</span><span class="sxs-lookup"><span data-stu-id="3f914-566">Example</span></span>

```javascript
function getCallbackToken() {
  Office.context.mailbox.getCallbackTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

#### <a name="getuseridentitytokenasynccallback-usercontext"></a><span data-ttu-id="3f914-567">getUserIdentityTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="3f914-567">getUserIdentityTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="3f914-568">Получает маркер, идентифицирующий пользователя и надстройку Office.</span><span class="sxs-lookup"><span data-stu-id="3f914-568">Gets a token identifying the user and the Office Add-in.</span></span>

<span data-ttu-id="3f914-569">Метод `getUserIdentityTokenAsync` возвращает маркер, который можно использовать для идентификации, а также [проверки подлинности надстройки и пользователя в сторонней системе](/outlook/add-ins/authentication).</span><span class="sxs-lookup"><span data-stu-id="3f914-569">The `getUserIdentityTokenAsync` method returns a token that you can use to identify and [authenticate the add-in and user with a third-party system](/outlook/add-ins/authentication).</span></span>

##### <a name="parameters"></a><span data-ttu-id="3f914-570">Параметры</span><span class="sxs-lookup"><span data-stu-id="3f914-570">Parameters</span></span>

|<span data-ttu-id="3f914-571">Имя</span><span class="sxs-lookup"><span data-stu-id="3f914-571">Name</span></span>| <span data-ttu-id="3f914-572">Тип</span><span class="sxs-lookup"><span data-stu-id="3f914-572">Type</span></span>| <span data-ttu-id="3f914-573">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="3f914-573">Attributes</span></span>| <span data-ttu-id="3f914-574">Описание</span><span class="sxs-lookup"><span data-stu-id="3f914-574">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="3f914-575">функция</span><span class="sxs-lookup"><span data-stu-id="3f914-575">function</span></span>||<span data-ttu-id="3f914-576">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="3f914-576">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="3f914-577">Маркер указывается в виде строки в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="3f914-577">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="3f914-578">При возникновении ошибки свойства `asyncResult.error` и `asyncResult.diagnostics` могут содержать дополнительные сведения.</span><span class="sxs-lookup"><span data-stu-id="3f914-578">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|
|`userContext`| <span data-ttu-id="3f914-579">Объект</span><span class="sxs-lookup"><span data-stu-id="3f914-579">Object</span></span>| <span data-ttu-id="3f914-580">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="3f914-580">&lt;optional&gt;</span></span>|<span data-ttu-id="3f914-581">Данные о состоянии, передаваемые в асинхронный метод.</span><span class="sxs-lookup"><span data-stu-id="3f914-581">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="errors"></a><span data-ttu-id="3f914-582">Ошибки</span><span class="sxs-lookup"><span data-stu-id="3f914-582">Errors</span></span>

|<span data-ttu-id="3f914-583">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="3f914-583">Error code</span></span>|<span data-ttu-id="3f914-584">Описание</span><span class="sxs-lookup"><span data-stu-id="3f914-584">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="3f914-585">Запрос не выполнен.</span><span class="sxs-lookup"><span data-stu-id="3f914-585">The request has failed.</span></span> <span data-ttu-id="3f914-586">Просмотрите объект Diagnostics для кода ошибки HTTP.</span><span class="sxs-lookup"><span data-stu-id="3f914-586">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="3f914-587">Сервер Exchange возвратил ошибку.</span><span class="sxs-lookup"><span data-stu-id="3f914-587">The Exchange server returned an error.</span></span> <span data-ttu-id="3f914-588">Дополнительные сведения можно найти в объекте диагностики.</span><span class="sxs-lookup"><span data-stu-id="3f914-588">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="3f914-589">Пользователь больше не подключен к сети.</span><span class="sxs-lookup"><span data-stu-id="3f914-589">The user is no longer connected to the network.</span></span> <span data-ttu-id="3f914-590">Проверьте сетевое подключение и повторите попытку.</span><span class="sxs-lookup"><span data-stu-id="3f914-590">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="3f914-591">Требования</span><span class="sxs-lookup"><span data-stu-id="3f914-591">Requirements</span></span>

|<span data-ttu-id="3f914-592">Требование</span><span class="sxs-lookup"><span data-stu-id="3f914-592">Requirement</span></span>| <span data-ttu-id="3f914-593">Значение</span><span class="sxs-lookup"><span data-stu-id="3f914-593">Value</span></span>|
|---|---|
|[<span data-ttu-id="3f914-594">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="3f914-594">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3f914-595">1.0</span><span class="sxs-lookup"><span data-stu-id="3f914-595">1.0</span></span>|
|[<span data-ttu-id="3f914-596">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="3f914-596">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3f914-597">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3f914-597">ReadItem</span></span>|
|[<span data-ttu-id="3f914-598">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="3f914-598">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3f914-599">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="3f914-599">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="3f914-600">Пример</span><span class="sxs-lookup"><span data-stu-id="3f914-600">Example</span></span>

```javascript
function getIdentityToken() {
  Office.context.mailbox.getUserIdentityTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

#### <a name="makeewsrequestasyncdata-callback-usercontext"></a><span data-ttu-id="3f914-601">makeEwsRequestAsync(data, callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="3f914-601">makeEwsRequestAsync(data, callback, [userContext])</span></span>

<span data-ttu-id="3f914-602">Выполняет асинхронный запрос для веб-служб Exchange (EWS) на сервере Exchange Server, на котором размещен почтовый ящик пользователя.</span><span class="sxs-lookup"><span data-stu-id="3f914-602">Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="3f914-603">Этот метод не поддерживается в следующих сценариях:</span><span class="sxs-lookup"><span data-stu-id="3f914-603">This method is not supported in the following scenarios.</span></span>
> - <span data-ttu-id="3f914-604">В Outlook на iOS или Android</span><span class="sxs-lookup"><span data-stu-id="3f914-604">In Outlook on iOS or Android</span></span>
> - <span data-ttu-id="3f914-605">Если надстройка загружается в почтовый ящик Gmail.</span><span class="sxs-lookup"><span data-stu-id="3f914-605">When the add-in is loaded in a Gmail mailbox</span></span>
> 
> <span data-ttu-id="3f914-606">В таких случаях надстройка должна [использовать REST API](/outlook/add-ins/use-rest-api) для доступа к почтовому ящику пользователя.</span><span class="sxs-lookup"><span data-stu-id="3f914-606">In these cases, add-ins should [use REST APIs](/outlook/add-ins/use-rest-api) to access the user's mailbox instead.</span></span>

<span data-ttu-id="3f914-607">Метод `makeEwsRequestAsync` отправляет запрос EWS от имени надстройки в Exchange.</span><span class="sxs-lookup"><span data-stu-id="3f914-607">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange.</span></span> <span data-ttu-id="3f914-608">Список поддерживаемых операций EWS см. в статье [Вызов веб-служб из надстройки Outlook](/outlook/add-ins/web-services#ews-operations-that-add-ins-support).</span><span class="sxs-lookup"><span data-stu-id="3f914-608">See [Call web services from an Outlook add-in](/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span></span>

<span data-ttu-id="3f914-609">С помощью метода `makeEwsRequestAsync` невозможно запрашивать элементы, связанные с папкой.</span><span class="sxs-lookup"><span data-stu-id="3f914-609">You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.</span></span>

<span data-ttu-id="3f914-610">В запросе XML должна быть указана кодировка UTF-8.</span><span class="sxs-lookup"><span data-stu-id="3f914-610">The XML request must specify UTF-8 encoding.</span></span>

```xml
<?xml version="1.0" encoding="utf-8"?>
```

<span data-ttu-id="3f914-p149">У вашей надстройки должно быть разрешение **ReadWriteMailbox** для использования метода `makeEwsRequestAsync`. Сведения об использовании разрешения **ReadWriteMailbox** и операций EWS, которые можно вызывать с помощью метода `makeEwsRequestAsync`, см. в статье [Указание разрешений для доступа почтовой надстройки к почтовому ящику пользователя](/outlook/add-ins/understanding-outlook-add-in-permissions).</span><span class="sxs-lookup"><span data-stu-id="3f914-p149">Your add-in must have the **ReadWriteMailbox** permission to use the `makeEwsRequestAsync` method. For information about using the **ReadWriteMailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method, see [Specify permissions for mail add-in access to the user's mailbox](/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>

> [!NOTE]
> <span data-ttu-id="3f914-613">Администратор сервера должен установить значение true для параметра `OAuthAuthentication` в каталоге сервера клиентского доступа EWS, чтобы метод `makeEwsRequestAsync` мог выполнять запросы EWS.</span><span class="sxs-lookup"><span data-stu-id="3f914-613">The server administrator must set `OAuthAuthentication` to true on the Client Access Server EWS directory to enable the `makeEwsRequestAsync` method to make EWS requests.</span></span>

##### <a name="version-differences"></a><span data-ttu-id="3f914-614">Различия версий</span><span class="sxs-lookup"><span data-stu-id="3f914-614">Version differences</span></span>

<span data-ttu-id="3f914-615">Если вы используете метод `makeEwsRequestAsync` в почтовых приложениях, которые выполняются в Outlook версии более ранней, чем 15.0.4535.1004, указывайте кодировку `ISO-8859-1`.</span><span class="sxs-lookup"><span data-stu-id="3f914-615">When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set the encoding value to `ISO-8859-1`.</span></span>

```xml
<?xml version="1.0" encoding="iso-8859-1"?>
```

<span data-ttu-id="3f914-p150">Значение кодировки не нужно указывать, если почтовое приложение выполняется в Outlook в Интернете. Чтобы определить, выполняется ли приложение в Outlook или Outlook в Интернете, используйте свойство mailbox.diagnostics.hostName. Используемую версию Outlook можно определить с помощью свойства mailbox.diagnostics.hostVersion.</span><span class="sxs-lookup"><span data-stu-id="3f914-p150">You do not need to set the encoding value when your mail app is running in Outlook on the web. You can determine whether your mail app is running in Outlook or Outlook on the web by using the mailbox.diagnostics.hostName property. You can determine what version of Outlook is running by using the mailbox.diagnostics.hostVersion property.</span></span>

##### <a name="parameters"></a><span data-ttu-id="3f914-619">Параметры</span><span class="sxs-lookup"><span data-stu-id="3f914-619">Parameters</span></span>

|<span data-ttu-id="3f914-620">Имя</span><span class="sxs-lookup"><span data-stu-id="3f914-620">Name</span></span>| <span data-ttu-id="3f914-621">Тип</span><span class="sxs-lookup"><span data-stu-id="3f914-621">Type</span></span>| <span data-ttu-id="3f914-622">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="3f914-622">Attributes</span></span>| <span data-ttu-id="3f914-623">Описание</span><span class="sxs-lookup"><span data-stu-id="3f914-623">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="3f914-624">String</span><span class="sxs-lookup"><span data-stu-id="3f914-624">String</span></span>||<span data-ttu-id="3f914-625">Запрос EWS.</span><span class="sxs-lookup"><span data-stu-id="3f914-625">The EWS request.</span></span>|
|`callback`| <span data-ttu-id="3f914-626">function</span><span class="sxs-lookup"><span data-stu-id="3f914-626">function</span></span>||<span data-ttu-id="3f914-627">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="3f914-627">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="3f914-628">Результат XML вызова EWS указывается в виде строки в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="3f914-628">The XML result of the EWS call is provided as a string in the `asyncResult.value` property.</span></span> <span data-ttu-id="3f914-629">Если размер результата превышает 1 МБ, возвращается сообщение об ошибке.</span><span class="sxs-lookup"><span data-stu-id="3f914-629">If the result exceeds 1 MB in size, an error message is returned instead.</span></span>|
|`userContext`| <span data-ttu-id="3f914-630">Объект</span><span class="sxs-lookup"><span data-stu-id="3f914-630">Object</span></span>| <span data-ttu-id="3f914-631">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="3f914-631">&lt;optional&gt;</span></span>|<span data-ttu-id="3f914-632">Данные о состоянии, передаваемые в асинхронный метод.</span><span class="sxs-lookup"><span data-stu-id="3f914-632">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="3f914-633">Требования</span><span class="sxs-lookup"><span data-stu-id="3f914-633">Requirements</span></span>

|<span data-ttu-id="3f914-634">Требование</span><span class="sxs-lookup"><span data-stu-id="3f914-634">Requirement</span></span>| <span data-ttu-id="3f914-635">Значение</span><span class="sxs-lookup"><span data-stu-id="3f914-635">Value</span></span>|
|---|---|
|[<span data-ttu-id="3f914-636">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="3f914-636">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3f914-637">1.0</span><span class="sxs-lookup"><span data-stu-id="3f914-637">1.0</span></span>|
|[<span data-ttu-id="3f914-638">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="3f914-638">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3f914-639">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="3f914-639">ReadWriteMailbox</span></span>|
|[<span data-ttu-id="3f914-640">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="3f914-640">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3f914-641">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="3f914-641">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="3f914-642">Пример</span><span class="sxs-lookup"><span data-stu-id="3f914-642">Example</span></span>

<span data-ttu-id="3f914-643">В приведенном ниже примере вызывается `makeEwsRequestAsync` для получения темы элемента с помощью операции `GetItem`.</span><span class="sxs-lookup"><span data-stu-id="3f914-643">The following example calls `makeEwsRequestAsync` to use the `GetItem` operation to get the subject of an item.</span></span>

```javascript
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

#### <a name="removehandlerasynceventtype-options-callback"></a><span data-ttu-id="3f914-644">removeHandlerAsync(eventType, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="3f914-644">removeHandlerAsync(eventType, [options], [callback])</span></span>

<span data-ttu-id="3f914-645">Удаляет обработчиков для поддерживаемого типа события.</span><span class="sxs-lookup"><span data-stu-id="3f914-645">Removes the event handlers for a supported event type.</span></span>

<span data-ttu-id="3f914-646">В настоящее время единственный поддерживаемый тип события — `Office.EventType.ItemChanged`.</span><span class="sxs-lookup"><span data-stu-id="3f914-646">Currently, the only supported event type is `Office.EventType.ItemChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="3f914-647">Параметры</span><span class="sxs-lookup"><span data-stu-id="3f914-647">Parameters</span></span>

| <span data-ttu-id="3f914-648">Имя</span><span class="sxs-lookup"><span data-stu-id="3f914-648">Name</span></span> | <span data-ttu-id="3f914-649">Тип</span><span class="sxs-lookup"><span data-stu-id="3f914-649">Type</span></span> | <span data-ttu-id="3f914-650">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="3f914-650">Attributes</span></span> | <span data-ttu-id="3f914-651">Описание</span><span class="sxs-lookup"><span data-stu-id="3f914-651">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="3f914-652">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="3f914-652">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="3f914-653">Событие, которое должно отменить обработчик.</span><span class="sxs-lookup"><span data-stu-id="3f914-653">The event that should revoke the handler.</span></span> |
| `options` | <span data-ttu-id="3f914-654">Объект</span><span class="sxs-lookup"><span data-stu-id="3f914-654">Object</span></span> | <span data-ttu-id="3f914-655">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="3f914-655">&lt;optional&gt;</span></span> | <span data-ttu-id="3f914-656">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="3f914-656">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="3f914-657">Object</span><span class="sxs-lookup"><span data-stu-id="3f914-657">Object</span></span> | <span data-ttu-id="3f914-658">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="3f914-658">&lt;optional&gt;</span></span> | <span data-ttu-id="3f914-659">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="3f914-659">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="3f914-660">функция</span><span class="sxs-lookup"><span data-stu-id="3f914-660">function</span></span>| <span data-ttu-id="3f914-661">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="3f914-661">&lt;optional&gt;</span></span>|<span data-ttu-id="3f914-662">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="3f914-662">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="3f914-663">Требования</span><span class="sxs-lookup"><span data-stu-id="3f914-663">Requirements</span></span>

|<span data-ttu-id="3f914-664">Требование</span><span class="sxs-lookup"><span data-stu-id="3f914-664">Requirement</span></span>| <span data-ttu-id="3f914-665">Значение</span><span class="sxs-lookup"><span data-stu-id="3f914-665">Value</span></span>|
|---|---|
|[<span data-ttu-id="3f914-666">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="3f914-666">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3f914-667">1.5</span><span class="sxs-lookup"><span data-stu-id="3f914-667">1.5</span></span> |
|[<span data-ttu-id="3f914-668">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="3f914-668">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3f914-669">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3f914-669">ReadItem</span></span> |
|[<span data-ttu-id="3f914-670">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="3f914-670">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3f914-671">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="3f914-671">Compose or Read</span></span>|
